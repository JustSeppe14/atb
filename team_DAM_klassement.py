import os
import pandas as pd
from utils import load_deelnemers, load_result, MAX_POINTS

import logging
# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

TEAM_KLASSEMENT_FILE = os.path.join(OUTPUT_DIR, "team_klassement_2025_DAM_only.xlsx")
IS_SECOND_PERIOD_STARTED = False  # Toggle this to True when 2nd period starts

def calculate_team_klassement():
    try:
        deelnemers = load_deelnemers()
        uitslag = load_result()
        logger.info("Calculating DAM-only team klassement")


        # Filter out team '0' and empty
        deelnemers = deelnemers[
            deelnemers['team'].notna() &
            (deelnemers['team'].astype(str).str.strip() != '') &
            (deelnemers['team'].astype(str).str.strip() != '0')
        ]

        # Only include teams with at least one DAM rider
        dam_teams = deelnemers[deelnemers['categorie'] == 'DAM']['team'].unique()
        deelnemers = deelnemers[deelnemers['team'].isin(dam_teams)]

        # Load existing klassement or start fresh
        if os.path.isfile(TEAM_KLASSEMENT_FILE):
            team_klassement_df = pd.read_excel(TEAM_KLASSEMENT_FILE, sheet_name="TEAMS MIXED")
            team_klassement_df = team_klassement_df[
                team_klassement_df['team'].notna() &
                (team_klassement_df['team'].astype(str).str.strip() != '') &
                (team_klassement_df['team'].astype(str).str.strip() != '0')
            ]
            week_cols = [col for col in team_klassement_df.columns if col.endswith('T') and col[:-1].isdigit()]
            if week_cols:
                max_week_num = max(int(col[:-1]) for col in week_cols)
                current_week = max_week_num + 1
            else:
                current_week = 1
        else:
            current_week = 1
            teams = deelnemers['team'].unique()
            team_klassement_df = pd.DataFrame({'team': teams})
            week_cols = []

        new_week_col = f"{current_week}T"

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
                current_week: punten.get(rijder['bib'], MAX_POINTS)
            })

        punten_df = pd.DataFrame(punten_per_rijder)

        # Sum points per team
        team_points_this_week = punten_df.groupby('team')[current_week].sum().reset_index()
        team_points_this_week = team_points_this_week[
            (team_points_this_week['team'].astype(str).str.strip() != '') &
            (team_points_this_week['team'].astype(str).str.strip() != '0')
        ]
        team_points_this_week.rename(columns={current_week: new_week_col}, inplace=True)

        # Rank teams: highest points = best = lower rank number
        team_points_this_week[new_week_col] = team_points_this_week[new_week_col].rank(method='min', ascending=False).astype(int)

        # Merge week results into klassement
        team_klassement_df = team_klassement_df.merge(team_points_this_week[['team', new_week_col]], on='team', how='outer')

        # Clean up again
        team_klassement_df = team_klassement_df[
            team_klassement_df['team'].notna() &
            (team_klassement_df['team'].astype(str).str.strip() != '') &
            (team_klassement_df['team'].astype(str).str.strip() != '0')
        ]

        # Update week_cols
        week_cols = [col for col in team_klassement_df.columns if col.endswith('T') and col[:-1].isdigit()]
        team_klassement_df[week_cols] = team_klassement_df[week_cols].fillna(0)

        # Totals and periods
        team_klassement_df['Totaal'] = team_klassement_df[week_cols].sum(axis=1)

        if IS_SECOND_PERIOD_STARTED and week_cols:
            second_period_start = max([int(col[:-1]) for col in week_cols])
            first_period_weeks = [col for col in week_cols if int(col[:-1]) < second_period_start]
            second_period_weeks = [col for col in week_cols if int(col[:-1]) >= second_period_start]
        else:
            first_period_weeks = week_cols
            second_period_weeks = []

        team_klassement_df['1e Periode'] = team_klassement_df[first_period_weeks].sum(axis=1) if first_period_weeks else 0
        team_klassement_df['2e Periode'] = team_klassement_df[second_period_weeks].sum(axis=1) if second_period_weeks else 0

        # Final ranking
        team_klassement_df = team_klassement_df.sort_values('Totaal')
        team_klassement_df['Plaats'] = range(1, len(team_klassement_df) + 1)

        # Reorder columns
        cols_order = ['Plaats', 'team', '1e Periode', '2e Periode', 'Totaal'] + sorted(week_cols, key=lambda c: int(c[:-1]))
        team_klassement_df = team_klassement_df[cols_order]

        # Save
        with pd.ExcelWriter(TEAM_KLASSEMENT_FILE, engine='openpyxl', mode='w') as writer:
            team_klassement_df.to_excel(writer, sheet_name="TEAMS MIXED", index=False)

        logger.info(f"✅ DAM-only team klassement updated with week {current_week} (column {new_week_col}) in {TEAM_KLASSEMENT_FILE}")
    except Exception as e:
        logger.error(f"❌ Error in calculate_team_klassement: {e}")
        raise
if __name__ == '__main__':
    calculate_team_klassement()
