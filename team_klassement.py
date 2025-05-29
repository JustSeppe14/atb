import os
import pandas as pd
from utils import load_deelnemers, load_result, MAX_POINTS

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

TEAM_KLASSEMENT_FILE = os.path.join(OUTPUT_DIR, "team_klassement_2025.xlsx")
IS_SECOND_PERIOD_STARTED = False  # Toggle this to True when 2nd period starts

def calculate_team_klassement():
    deelnemers = load_deelnemers()
    uitslag = load_result()

    # Load existing klassement or start fresh
    if os.path.isfile(TEAM_KLASSEMENT_FILE):
        team_klassement_df = pd.read_excel(TEAM_KLASSEMENT_FILE, sheet_name="TEAMS STA")
        # Filter out team '0' just in case
        team_klassement_df = team_klassement_df[
            team_klassement_df['team'].notna() &
            (team_klassement_df['team'].astype(str).str.strip() != '') &
            (team_klassement_df['team'].astype(str).str.strip() != '0')
        ]
        # Identify existing week columns like '1T', '2T', etc.
        week_cols = [col for col in team_klassement_df.columns if col.endswith('T') and col[:-1].isdigit()]
        if week_cols:
            max_week_num = max(int(col[:-1]) for col in week_cols)
            current_week = max_week_num + 1  # Next week to add
        else:
            current_week = 1  # Start from week 1 if none exist
    else:
        current_week = 1
        teams = deelnemers['team'].unique()
        team_klassement_df = pd.DataFrame({'team': teams})
        week_cols = []

    new_week_col = f"{current_week}T"

    if 'team' not in deelnemers.columns:
        raise ValueError("Deelnemers data must have a 'team' column")

    # Filter out team '0'
    deelnemers = deelnemers[
        deelnemers['team'].notna() &
        (deelnemers['team'].astype(str).str.strip() != '') &
        (deelnemers['team'].astype(str).str.strip() != '0')
    ]

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
            current_week: punten.get(rijder['bib'], MAX_POINTS)
        })

    punten_df = pd.DataFrame(punten_per_rijder)

    # Sum points per team for this week (use numeric week_col for grouping)
    team_points_this_week = punten_df.groupby('team')[current_week].sum().reset_index()

    # Exclude team '0' before ranking
    team_points_this_week = team_points_this_week[
        (team_points_this_week['team'].astype(str).str.strip() != '') &
        (team_points_this_week['team'].astype(str).str.strip() != '0')
    ]

    # Rename sum column to 'nT' format
    team_points_this_week.rename(columns={current_week: new_week_col}, inplace=True)

    # Rank teams by points: highest points => rank 1 (lowest rank number)
    team_points_this_week[new_week_col] = team_points_this_week[new_week_col].rank(method='min', ascending=False).astype(int)

    # Merge with existing klassement dataframe
    team_klassement_df = team_klassement_df.merge(team_points_this_week[['team', new_week_col]], on='team', how='outer')

    # Clean teams again just in case
    team_klassement_df = team_klassement_df[
        team_klassement_df['team'].notna() &
        (team_klassement_df['team'].astype(str).str.strip() != '') &
        (team_klassement_df['team'].astype(str).str.strip() != '0')
    ]

    # Update week_cols (now including the newly added week)
    week_cols = [col for col in team_klassement_df.columns if col.endswith('T') and col[:-1].isdigit()]
    team_klassement_df[week_cols] = team_klassement_df[week_cols].fillna(0)

    # Calculate totals
    team_klassement_df['Totaal'] = team_klassement_df[week_cols].sum(axis=1)

    # Calculate periods sums
    if IS_SECOND_PERIOD_STARTED and week_cols:
        second_period_start = max([int(col[:-1]) for col in week_cols])
        first_period_weeks = [col for col in week_cols if int(col[:-1]) < second_period_start]
        second_period_weeks = [col for col in week_cols if int(col[:-1]) >= second_period_start]
    else:
        first_period_weeks = week_cols
        second_period_weeks = []

    team_klassement_df['1e Periode'] = team_klassement_df[first_period_weeks].sum(axis=1) if first_period_weeks else 0
    team_klassement_df['2e Periode'] = team_klassement_df[second_period_weeks].sum(axis=1) if second_period_weeks else 0

    # Final ranking: sort by total ascending and assign unique places 1,2,3...
    team_klassement_df = team_klassement_df.sort_values('Totaal')
    team_klassement_df['Plaats'] = range(1, len(team_klassement_df) + 1)

    # Column order for output
    cols_order = ['Plaats', 'team', '1e Periode', '2e Periode', 'Totaal'] + sorted(week_cols, key=lambda c: int(c[:-1]))
    team_klassement_df = team_klassement_df[cols_order]

    # Save to Excel
    with pd.ExcelWriter(TEAM_KLASSEMENT_FILE, engine='openpyxl', mode='w') as writer:
        team_klassement_df.to_excel(writer, sheet_name="TEAMS STA", index=False)

    print(f"✅ Team klassement updated with week {current_week} (column {new_week_col}) in {TEAM_KLASSEMENT_FILE}")

if __name__ == '__main__':
    calculate_team_klassement()
