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

    punten_per_rijder = []

    for klasse in deelnemers['klasse'].unique():
        klasse_deelnemers = deelnemers[deelnemers['klasse'] == klasse]
        klasse_bibs = set(klasse_deelnemers['bib'])
        klasse_result = uitslag[uitslag['bib'].isin(klasse_bibs)].sort_values(by='plaats')

        punten = {}
        for i, row in enumerate(klasse_result.itertuples(), start=1):
            punten[row.bib] = i if i < MAX_POINTS else MAX_POINTS

        for _, rijder in klasse_deelnemers.iterrows():
            punten_per_rijder.append({
                'bib': rijder['bib'],
                'klasse': rijder['klasse'],
                'naam': rijder['naam'],
                'categorie': rijder['categorie'],
                week_col: punten.get(rijder['bib'], MAX_POINTS)
            })

    punten_df = pd.DataFrame(punten_per_rijder)

    # Accumulate all weeks
    all_weeks = list(range(1, current_week + 1))
    for week in all_weeks[1:]:  # From week 2 onwards
        result_path = f"Results/finish.xlsx"
        if not os.path.exists(result_path):
            continue
        global RESULT_FILE
        RESULT_FILE = result_path
        extra_result = load_result()

        week_col = str(week)
        for klasse in deelnemers['klasse'].unique():
            klasse_deelnemers = deelnemers[deelnemers['klasse'] == klasse]
            klasse_bibs = set(klasse_deelnemers['bib'])
            klasse_result = extra_result[extra_result['bib'].isin(klasse_bibs)].sort_values(by='plaats')

            punten = {}
            for i, row in enumerate(klasse_result.itertuples(), start=1):
                punten[row.bib] = i if i < MAX_POINTS else MAX_POINTS

            for _, rijder in klasse_deelnemers.iterrows():
                mask = (punten_df['bib'] == rijder['bib']) & (punten_df['klasse'] == rijder['klasse'])
                punten_df.loc[mask, week_col] = punten.get(rijder['bib'], MAX_POINTS)

    week_cols = [col for col in punten_df.columns if str(col).isdigit()]
    punten_df[week_cols] = punten_df[week_cols].fillna(MAX_POINTS)

    # Total and period sums
    punten_df['Totaal'] = punten_df[week_cols].sum(axis=1)

    if IS_SECOND_PERIOD_STARTED:
        first_period_weeks = [col for col in week_cols if int(col) < week_cols[-1]]
        second_period_weeks = [col for col in week_cols if int(col) >= week_cols[-1]]
    else:
        first_period_weeks = week_cols
        second_period_weeks = []

    punten_df['1e Periode'] = punten_df[first_period_weeks].sum(axis=1)
    punten_df['2e Periode'] = punten_df[second_period_weeks].sum(axis=1)

    # Split full vs incomplete
    full_mask = punten_df[week_cols].apply(lambda row: all(v < MAX_POINTS for v in row), axis=1)
    full_df = punten_df[full_mask].copy()
    incomplete_df = punten_df[~full_mask].copy()

    # Rank full participants
    full_df['Plaats Klasse'] = (
        full_df.sort_values(by=['klasse', 'Totaal'])
        .groupby('klasse')
        .cumcount() + 1
    )

    # Reassign incomplete participants to pseudo-classes
    unique_klasses = sorted(punten_df['klasse'].unique())
    klasse_to_letter = {k: f"X{chr(65+i)}" for i, k in enumerate(unique_klasses)}
    incomplete_df['klasse'] = incomplete_df['klasse'].map(klasse_to_letter)

    # Rank incomplete participants
    incomplete_df['Plaats Klasse'] = (
        incomplete_df.sort_values(by=['klasse', 'Totaal'])
        .groupby('klasse')
        .cumcount() + 1
    )

    # Combine and sort
    combined_df = pd.concat([full_df, incomplete_df], ignore_index=True)
    combined_df = combined_df.sort_values(by=['klasse', 'Totaal'])

    # Rename for Excel
    combined_df = combined_df.rename(columns={
        'bib': 'Nr.',
        'naam': 'Naam',
        'klasse': 'Klasse',
        'categorie': 'Cat.'
    })

    # Column order
    final_column_order = load_template_column_order()
    if 'Plaats Klasse' not in final_column_order:
        try:
            klasse_idx = final_column_order.index('Klasse')
            final_column_order.insert(klasse_idx + 1, 'Plaats Klasse')
        except ValueError:
            final_column_order.append('Plaats Klasse')

    final_cols = final_column_order + [col for col in week_cols if col not in final_column_order]
    combined_df = combined_df[[col for col in final_cols if col in combined_df.columns]]

    # Write to Excel
    with pd.ExcelWriter(KLASSEMENT_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        combined_df.to_excel(writer, sheet_name="Klassement", index=False)

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

    for idx, col_name in enumerate(sheet[1], start=1):
        if str(col_name.value).isdigit():
            fill = green_fill if int(col_name.value) <= 4 else blue_fill
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                row[idx - 1].fill = fill

    wb.save(KLASSEMENT_FILE)
    print(f"✅ Klassement-tab met volledige én onvolledige deelnemers toegevoegd aan {KLASSEMENT_FILE}")

if __name__ == '__main__':
    generate_klassement()
