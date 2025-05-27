import os
import pandas as pd
from utils import load_deelnemers, load_result, get_current_week, MAX_POINTS

TEAM_KLASSEMENT_FILE = "team_klassement_2025_DAM_only.xlsx"
IS_SECOND_PERIOD_STARTED = False  # Toggle this to True when 2nd period starts

def calculate_team_klassement():
    deelnemers = load_deelnemers()
    uitslag = load_result()
    current_week = get_current_week(TEAM_KLASSEMENT_FILE, sheet_name="TeamKlassement")
    week_col = str(current_week)

    if 'team' not in deelnemers.columns or 'categorie' not in deelnemers.columns:
        raise ValueError("Deelnemers data must have 'team' and 'cat' columns")

    # Filter out team '0'
    deelnemers = deelnemers[deelnemers['team'] != '0']

    # Filter teams with at least one DAM rider (CAT == 'DAM')
    dam_teams = deelnemers[deelnemers['categorie'] == 'DAM']['team'].unique()
    deelnemers = deelnemers[deelnemers['team'].isin(dam_teams)]

    # Calculate points per rider for current week
    punten = {}
    for cat in deelnemers['categorie'].unique():
        cat_deelnemers = deelnemers[deelnemers['categorie'] == cat]
        cat_bibs = set(cat_deelnemers['bib'])
        cat_result = uitslag[uitslag['bib'].isin(cat_bibs)].copy()
        cat_result['rank_in_class'] = cat_result['plaats'].rank(method='first').astype(int)

        for row in cat_result.itertuples():
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
        team_klassement_df = pd.read_excel(TEAM_KLASSEMENT_FILE, sheet_name="TeamKlassement")
        team_klassement_df = team_klassement_df[team_klassement_df['team'].astype(str) != '0']
    else:
        teams = deelnemers['team'].unique()
        team_klassement_df = pd.DataFrame({'team': teams})

    # Identify week columns
    week_cols = [col for col in team_klassement_df.columns if col.endswith('T') and col[:-1].isdigit()]
    new_week_col = f"{current_week}T"

    # Sum points per team for this week
    team_points_this_week = punten_df.groupby('team')[week_col].sum().reset_index()
    team_points_this_week = team_points_this_week[team_points_this_week['team'] != '0']
    team_points_this_week.rename(columns={week_col: new_week_col}, inplace=True)

    # Ranking: more rider points = better = lower team points (1 is best)
    team_points_this_week[new_week_col] = (
        team_points_this_week[new_week_col]
        .rank(method='min', ascending=False)
        .fillna(len(team_points_this_week))
        .astype(int)
    )

    # Merge with klassement
    team_klassement_df = team_klassement_df.merge(team_points_this_week, on='team', how='outer')
    team_klassement_df = team_klassement_df[team_klassement_df['team'].astype(str) != '0']
    team_klassement_df = team_klassement_df[team_klassement_df['team'].notna() & (team_klassement_df['team'].astype(str) != '')]

    # Fill NA
    week_cols = [col for col in team_klassement_df.columns if col.endswith('T') and col[:-1].isdigit()]
    team_klassement_df[week_cols] = team_klassement_df[week_cols].fillna(0)

    # Totals
    team_klassement_df['Totaal'] = team_klassement_df[week_cols].sum(axis=1)

    # Periods
    if IS_SECOND_PERIOD_STARTED and week_cols:
        second_period_start = max([int(col[:-1]) for col in week_cols])
        first_period_weeks = [col for col in week_cols if int(col[:-1]) < second_period_start]
        second_period_weeks = [col for col in week_cols if int(col[:-1]) >= second_period_start]
    else:
        first_period_weeks = week_cols
        second_period_weeks = []

    team_klassement_df['1e Periode'] = team_klassement_df[first_period_weeks].sum(axis=1) if first_period_weeks else 0
    team_klassement_df['2e Periode'] = team_klassement_df[second_period_weeks].sum(axis=1) if second_period_weeks else 0

    # Final ranking: lowest total team points is best
    team_klassement_df['Plaats'] = team_klassement_df['Totaal'].rank(method='min', ascending=True).astype(int)
    team_klassement_df = team_klassement_df.sort_values('Plaats')

    # Output
    team_klassement_df.rename(columns={'team': 'Team'}, inplace=True)
    cols_order = ['Plaats', 'Team', '1e Periode', '2e Periode', 'Totaal'] + sorted(week_cols, key=lambda c: int(c[:-1]))
    team_klassement_df = team_klassement_df[cols_order]

    # Save
    with pd.ExcelWriter(TEAM_KLASSEMENT_FILE, engine='openpyxl', mode='w') as writer:
        team_klassement_df.to_excel(writer, sheet_name="TeamKlassement", index=False)

    print(f"✅ DAM-only team klassement updated with week {current_week} in {TEAM_KLASSEMENT_FILE}")

if __name__ == '__main__':
    calculate_team_klassement()
