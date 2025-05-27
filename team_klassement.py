import os
import pandas as pd
from utils import load_deelnemers, load_result, get_current_week, MAX_POINTS

TEAM_KLASSEMENT_FILE = "team_klassement_2025.xlsx"
IS_SECOND_PERIOD_STARTED = False  # Toggle this to True when 2nd period starts

def calculate_team_klassement():
    deelnemers = load_deelnemers()
    uitslag = load_result()
    current_week = get_current_week(TEAM_KLASSEMENT_FILE, sheet_name="TEAMS STA")
    week_col = str(current_week)

    if 'team' not in deelnemers.columns:
        raise ValueError("Deelnemers data must have a 'team' column")

    # Filter out team '0' from deelnemers immediately
    deelnemers = deelnemers[deelnemers['team'] != '0']

    # Calculate points per rider for current week
    punten = {}
    for klasse in deelnemers['klasse'].unique():
        klasse_deelnemers = deelnemers[deelnemers['klasse'] == klasse]
        klasse_bibs = set(klasse_deelnemers['bib'])
        klasse_result = uitslag[uitslag['bib'].isin(klasse_bibs)].copy()
        klasse_result['rank_in_class'] = klasse_result['plaats'].rank(method='first').astype(int)

        for row in klasse_result.itertuples():
            punten[row.bib] = row.rank_in_class if row.rank_in_class < 60 else 60

    punten_per_rijder = []
    for _, rijder in deelnemers.iterrows():
        punten_per_rijder.append({
            'bib': rijder['bib'],
            'team': rijder['team'],
            week_col: punten.get(rijder['bib'], MAX_POINTS)
        })

    punten_df = pd.DataFrame(punten_per_rijder)

    # Load or create team klassement df
    if os.path.isfile(TEAM_KLASSEMENT_FILE):
        team_klassement_df = pd.read_excel(TEAM_KLASSEMENT_FILE, sheet_name="TEAMS STA")
        # Filter out team '0' from existing klassement data
        team_klassement_df = team_klassement_df[team_klassement_df['team'].astype(str) != '0']
    else:
        teams = deelnemers['team'].unique()
        team_klassement_df = pd.DataFrame({'team': teams})

    # Find existing week columns named like '1T', '2T', ...
    week_cols = [col for col in team_klassement_df.columns if col.endswith('T') and col[:-1].isdigit()]
    week_nums = sorted([int(col[:-1]) for col in week_cols])

    # New week column
    new_week_col = f"{current_week}T"

    # Sum points per team for this week
    team_points_this_week = punten_df.groupby('team')[week_col].sum().reset_index()

    # Exclude team '0' BEFORE ranking
    team_points_this_week = team_points_this_week[team_points_this_week['team'] != '0']

    # Rename the sum column to the new week column name
    team_points_this_week.rename(columns={week_col: new_week_col}, inplace=True)

    # Calculate ranking: highest rider points get 1 team point (rank 1)
    team_points_this_week[new_week_col] = team_points_this_week[new_week_col].rank(method='min', ascending=False).astype(int)

    # Merge with existing data
    team_klassement_df = team_klassement_df.merge(team_points_this_week, on='team', how='outer')

    # Filter out any '0' teams after merge, just to be sure
    team_klassement_df = team_klassement_df[team_klassement_df['team'].astype(str) != '0']

    # Drop rows where 'team' is empty or null to remove empty rows
    team_klassement_df = team_klassement_df[team_klassement_df['team'].notna() & (team_klassement_df['team'].astype(str) != '')]

    # Fill missing week points with 0
    week_cols = [col for col in team_klassement_df.columns if col.endswith('T') and col[:-1].isdigit()]
    team_klassement_df[week_cols] = team_klassement_df[week_cols].fillna(0)

    # Calculate total points (sum of team points across all weeks)
    team_klassement_df['Totaal'] = team_klassement_df[week_cols].sum(axis=1)

    # Calculate 1e Periode and 2e Periode sums
    if IS_SECOND_PERIOD_STARTED and week_cols:
        second_period_start = max([int(col[:-1]) for col in week_cols])
        first_period_weeks = [col for col in week_cols if int(col[:-1]) < second_period_start]
        second_period_weeks = [col for col in week_cols if int(col[:-1]) >= second_period_start]
    else:
        first_period_weeks = week_cols
        second_period_weeks = []

    team_klassement_df['1e Periode'] = team_klassement_df[first_period_weeks].sum(axis=1) if first_period_weeks else 0
    team_klassement_df['2e Periode'] = team_klassement_df[second_period_weeks].sum(axis=1) if second_period_weeks else 0

    # Rank teams by total team points ascending (lowest rank = best, so smallest team points number)
    team_klassement_df['Plaats'] = team_klassement_df['Totaal'].rank(method='min', ascending=True).astype(int)

    # Sort by rank (Plaats)
    team_klassement_df = team_klassement_df.sort_values('Plaats')

    # Rename 'team' to 'Team' for output
    team_klassement_df.rename(columns={'team': 'Team'}, inplace=True)

    # Arrange columns order
    cols_order = ['Plaats', 'Team', '1e Periode', '2e Periode', 'Totaal'] + sorted(week_cols, key=lambda c: int(c[:-1]))
    team_klassement_df = team_klassement_df[cols_order]

    # Save to Excel
    with pd.ExcelWriter(TEAM_KLASSEMENT_FILE, engine='openpyxl', mode='w') as writer:
        team_klassement_df.to_excel(writer, sheet_name="TEAMS STA", index=False)

    print(f"✅ Team klassement updated with week {current_week} in {TEAM_KLASSEMENT_FILE}")

if __name__ == '__main__':
    calculate_team_klassement()
