import os
import pandas as pd
from utils import load_deelnemers, load_result, MAX_POINTS, backup_file

import logging
import shutil
from datetime import datetime

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

TEAM_KLASSEMENT_FILE = os.path.join(OUTPUT_DIR, "team_klassement_2025.xlsx")
IS_SECOND_PERIOD_STARTED = os.environ.get('IS_SECOND_PERIOD_STARTED', 'False').lower() == 'true'

def calculate_team_klassement():
    try:
        deelnemers = load_deelnemers()
        uitslag = load_result()
        logger.info("Calculating team klassement")

        if os.path.isfile(TEAM_KLASSEMENT_FILE):
            team_klassement_df = pd.read_excel(TEAM_KLASSEMENT_FILE, sheet_name="TEAMS STA")
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

        if 'team' not in deelnemers.columns:
            raise ValueError("Deelnemers data must have a 'team' column")

        deelnemers = deelnemers[
            deelnemers['team'].notna() &
            (deelnemers['team'].astype(str).str.strip() != '') &
            (deelnemers['team'].astype(str).str.strip() != '0')
        ]

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

        # Only count top 4 best riders per team
        top_riders_per_team = (
            punten_df
            .sort_values(by=current_week)  # Lower rank = better position
            .groupby('team')
            .head(4)  # Select top 4 riders per team
        )

        team_points_this_week = (
            top_riders_per_team
            .groupby('team')[current_week]
            .sum()
            .reset_index()
        )

        team_points_this_week = team_points_this_week[
            (team_points_this_week['team'].astype(str).str.strip() != '') &
            (team_points_this_week['team'].astype(str).str.strip() != '0')
        ]

        team_points_this_week.rename(columns={current_week: new_week_col}, inplace=True)

        # Rank teams (lower total = better, hence rank ascending)
        team_points_this_week[new_week_col] = team_points_this_week[new_week_col].rank(method='min', ascending=False).astype(int)

        team_klassement_df = team_klassement_df.merge(team_points_this_week[['team', new_week_col]], on='team', how='outer')

        team_klassement_df = team_klassement_df[
            team_klassement_df['team'].notna() &
            (team_klassement_df['team'].astype(str).str.strip() != '') &
            (team_klassement_df['team'].astype(str).str.strip() != '0')
        ]

        week_cols = [col for col in team_klassement_df.columns if col.endswith('T') and col[:-1].isdigit()]
        team_klassement_df[week_cols] = team_klassement_df[week_cols].fillna(0)

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

        team_klassement_df = team_klassement_df.sort_values('Totaal')
        team_klassement_df['Plaats'] = range(1, len(team_klassement_df) + 1)

        cols_order = ['Plaats', 'team', '1e Periode', '2e Periode', 'Totaal'] + sorted(week_cols, key=lambda c: int(c[:-1]))
        team_klassement_df = team_klassement_df[cols_order]

        with pd.ExcelWriter(TEAM_KLASSEMENT_FILE, engine='openpyxl', mode='w') as writer:
            team_klassement_df.to_excel(writer, sheet_name="TEAMS STA", index=False)

         # --- Save backup using shared backup system ---
        backup_path = backup_file(TEAM_KLASSEMENT_FILE, f"team_klassement_2025_week_{current_week}.xlsx")
        logger.info(f"📁 Backup saved to {backup_path}")

        logger.info(f"✅ Team klassement updated with week {current_week} (column {new_week_col}) in {TEAM_KLASSEMENT_FILE}")
    except Exception as e:
        logger.error(f"❌ Error in calculate_team_klassement: {e}")
        raise

if __name__ == '__main__':
    calculate_team_klassement()