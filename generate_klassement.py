import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from utils import (
    load_deelnemers,
    load_result,
    get_current_week,
    load_template_column_order,
    MAX_POINTS,
    DEELNEMERS_FILE,
    RESULT_FILE,
    TEMPLATE_FILE
)

KLASSEMENT_FILE = "klassement_2025.xlsx"
IS_SECOND_PERIOD_STARTED = False

def generate_klassement():
    deelnemers = load_deelnemers()
    uitslag = load_result()
    current_week = get_current_week(KLASSEMENT_FILE, sheet_name="Klassement")
    week_col = str(current_week)

    # Load existing klassement data if available, else create from deelnemers
    if os.path.isfile(KLASSEMENT_FILE):
        klassement_df = pd.read_excel(KLASSEMENT_FILE, sheet_name="Klassement")
        klassement_df = klassement_df.rename(columns={
            'Nr.': 'bib',
            'Naam': 'naam',
            'Klasse': 'klasse',
            'Cat.': 'categorie'
        })
    else:
        klassement_df = deelnemers[['naam', 'bib', 'klasse', 'categorie']].copy()

    # Ensure all participants from deelnemers are included
    klassement_df = deelnemers[['naam', 'bib', 'klasse', 'categorie']].merge(
        klassement_df, on=['bib', 'naam', 'klasse', 'categorie'], how='left'
    ).fillna(MAX_POINTS)

    # Calculate points for the current week only
    punten_per_rijder = []
    for klasse in deelnemers['klasse'].unique():
        klasse_deelnemers = deelnemers[deelnemers['klasse'] == klasse]
        klasse_bibs = set(klasse_deelnemers['bib'])
        klasse_result = uitslag[uitslag['bib'].isin(klasse_bibs)].copy()
        klasse_result['rank_in_class'] = klasse_result['plaats'].rank(method='first').astype(int)

        punten = {}
        for row in klasse_result.itertuples():
            punten[row.bib] = row.rank_in_class if row.rank_in_class < 60 else 60

        for _, rijder in klasse_deelnemers.iterrows():
            punten_per_rijder.append({
                'bib': rijder['bib'],
                week_col: punten.get(rijder['bib'], MAX_POINTS)
            })

    current_week_points_df = pd.DataFrame(punten_per_rijder)

    # Update or add the current week column in klassement_df
    klassement_df = klassement_df.merge(current_week_points_df, on='bib', how='left')
    klassement_df[week_col] = klassement_df[week_col].fillna(klassement_df.get(week_col, MAX_POINTS))

    # Get all week columns as digits, excluding non-numeric columns
    week_cols = [col for col in klassement_df.columns if str(col).isdigit()]
    week_cols = sorted(week_cols, key=int)

    # Fill missing week columns with MAX_POINTS to avoid NaNs
    klassement_df[week_cols] = klassement_df[week_cols].fillna(MAX_POINTS)

    # Calculate totals and period sums
    klassement_df['Totaal'] = klassement_df[week_cols].sum(axis=1)

    if IS_SECOND_PERIOD_STARTED:
        # Assuming second period starts at a certain week number
        # Here you can define the exact week number where period 2 starts
        second_period_start = week_cols[-1]  # Or any custom logic
        first_period_weeks = [col for col in week_cols if int(col) < int(second_period_start)]
        second_period_weeks = [col for col in week_cols if int(col) >= int(second_period_start)]
    else:
        first_period_weeks = week_cols
        second_period_weeks = []

    klassement_df['1e Periode'] = klassement_df[first_period_weeks].sum(axis=1)
    klassement_df['2e Periode'] = klassement_df[second_period_weeks].sum(axis=1)

    # Compute rank per klasse
    klassement_df['Plaats Klasse'] = (
        klassement_df.sort_values(by=['klasse', 'Totaal'])
        .groupby('klasse')
        .cumcount() + 1
    )

    # Sort by klasse and total points
    klassement_df = klassement_df.sort_values(by=['klasse', 'Totaal'])

    # Rename columns back for Excel export
    klassement_df = klassement_df.rename(columns={
        'bib': 'Nr.',
        'naam': 'Naam',
        'klasse': 'Klasse',
        'categorie': 'Cat.'
    })

    # Final column order
    final_column_order = load_template_column_order()
    if 'Plaats Klasse' not in final_column_order:
        try:
            klasse_idx = final_column_order.index('Klasse')
            final_column_order.insert(klasse_idx + 1, 'Plaats Klasse')
        except ValueError:
            final_column_order.append('Plaats Klasse')

    # Add any week columns not in template order at the end
    final_cols = final_column_order + [col for col in week_cols if col not in final_column_order]
    klassement_df = klassement_df[[col for col in final_cols if col in klassement_df.columns]]

    # Write updated sheet, replacing only the 'Klassement' sheet
    with pd.ExcelWriter(KLASSEMENT_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        klassement_df.to_excel(writer, sheet_name="Klassement", index=False)

    # Formatting
    wb = load_workbook(KLASSEMENT_FILE)
    sheet = wb['Klassement']

    pink_fill = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        if row[3].value == 'DAM':
            for cell in row:
                cell.fill = pink_fill

    header = [cell.value for cell in sheet[1]]
    for idx, col_name in enumerate(header, start=1):
        if str(col_name).isdigit():
            fill = green_fill if int(col_name) <= 4 else blue_fill
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                row[idx - 1].fill = fill

    wb.save(KLASSEMENT_FILE)
    print(f"✅ Klassement updated with week {current_week} in {KLASSEMENT_FILE}")

if __name__ == '__main__':
    generate_klassement()
