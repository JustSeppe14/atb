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

def generate_regelmatigheidscriterium():
    deelnemers = load_deelnemers()
    uitslag = load_result()
    week_num = get_current_week(KLASSEMENT_FILE, sheet_name="Regelmatigheidscriterium")
    week_col = str(week_num)

    punten_per_rijder = []

    for klasse in deelnemers['klasse'].unique():
        klasse_deelnemers = deelnemers[deelnemers['klasse'] == klasse]
        klasse_bibs = set(klasse_deelnemers['bib'])

        klasse_result = uitslag[uitslag['bib'].isin(klasse_bibs)].sort_values(by='plaats')

        punten = {}
        for i, row in enumerate(klasse_result.itertuples(), start=1):
            punten[row.bib] = i if i < 60 else 60

        for _, rijder in klasse_deelnemers.iterrows():
            punten_per_rijder.append({
                'bib': rijder['bib'],
                week_col: punten.get(rijder['bib'], MAX_POINTS)
            })

    punten_df = pd.DataFrame(punten_per_rijder)

    if os.path.isfile(KLASSEMENT_FILE):
        klassement_df = pd.read_excel(KLASSEMENT_FILE, sheet_name="Regelmatigheidscriterium")
        klassement_df = klassement_df.rename(columns={
            'Nr.': 'bib',
            'Naam': 'naam',
            'Klasse': 'klasse',
            'Cat.': 'categorie',
            'Totaal': 'total',
            '1e Periode': 'eerst_heft',
            '2e Periode': 'tweede_heft'
        })
    else:
        klassement_df = deelnemers[['naam', 'bib', 'klasse', 'categorie']].copy()

    klassement_df = klassement_df.merge(punten_df, on='bib', how='left')
    klassement_df[week_col] = klassement_df[week_col].fillna(MAX_POINTS)

    week_cols = [col for col in klassement_df.columns if str(col).isdigit()]
    week_cols = sorted(week_cols, key=int)

    klassement_df['total'] = klassement_df[week_cols].sum(axis=1)

    if week_cols:
        if IS_SECOND_PERIOD_STARTED:
            # All weeks after the second period starts go into '2e Periode'
            first_period_weeks = [col for col in week_cols if int(col) < week_cols[-1]]  # All weeks before the last week
            second_period_weeks = [col for col in week_cols if int(col) >= week_cols[-1]]  # All weeks after the second period start
        else:
            # If the second period has not started, all weeks go into '1e Periode'
            first_period_weeks = week_cols
            second_period_weeks = []

        # Assign sums to the respective periods
        klassement_df['eerst_heft'] = klassement_df[first_period_weeks].sum(axis=1)
        klassement_df['tweede_heft'] = klassement_df[second_period_weeks].sum(axis=1)

    # Compute rank per klasse and insert as 'Plaats Klasse'
    klassement_df['Plaats Klasse'] = (klassement_df.sort_values(by=['klasse', 'total']).groupby('klasse').cumcount() +1)

    # # Compute per-category ranks (STA, SEN, DAM)
    # for cat in ['STA', 'SEN', 'DAM']:
    #     col_name = f'Plaats {cat}'
    #     mask = klassement_df['categorie'] == cat
    #     df_cat = klassement_df[mask].copy()
    #     df_cat = df_cat.sort_values(by=['total']).reset_index()
    #     df_cat[col_name] = range(1, len(df_cat) + 1)
    #     klassement_df.loc[df_cat['index'], col_name] = df_cat[col_name]





    klassement_df = klassement_df.sort_values(by=['klasse', 'total'])

    # Rename to final Excel names
    klassement_df = klassement_df.rename(columns={
        'bib': 'Nr.',
        'naam': 'Naam',
        'klasse': 'Klasse',
        'categorie': 'Cat.',
        'total': 'Totaal',
        'eerst_heft': '1e Periode',
        'tweede_heft': '2e Periode'
    })

    # Determine column order
    final_column_order = load_template_column_order()
    if 'Plaats Klasse' not in final_column_order:
        try:
            klasse_idx = final_column_order.index('Klasse')
            final_column_order.insert(klasse_idx + 1, 'Plaats Klasse')
        except ValueError:
            final_column_order.append('Plaats Klasse')

    week_cols_in_output = [col for col in klassement_df.columns if str(col).isdigit()]
    final_cols = final_column_order + [col for col in week_cols_in_output if col not in final_column_order]
    klassement_df = klassement_df[[col for col in final_cols if col in klassement_df.columns]]

    # Write to Excel
    with pd.ExcelWriter(KLASSEMENT_FILE, engine='openpyxl', mode='w') as writer:
        klassement_df.to_excel(writer, sheet_name="Regelmatigheidscriterium", index=False)

    # Formatting
    wb = load_workbook(KLASSEMENT_FILE)
    sheet = wb['Regelmatigheidscriterium']

    pink_fill = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")

    # Highlight DAM rows
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        if row[3].value == 'DAM':
            for cell in row:
                cell.fill = pink_fill

    # Highlight week columns
    header = [cell.value for cell in sheet[1]]
    for idx, col_name in enumerate(header, start=1):
        if str(col_name).isdigit():
            fill = green_fill if int(col_name) <= 4 else blue_fill
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                row[idx - 1].fill = fill

    wb.save(KLASSEMENT_FILE)
    print(f"✅ Week {week_num} toegevoegd aan {KLASSEMENT_FILE}")

if __name__ == '__main__':
    generate_regelmatigheidscriterium()
