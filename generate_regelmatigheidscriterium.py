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
    klasse_per_rijder = []
    categorie_per_rijder = []

    for klasse in deelnemers['klasse'].unique():
        klasse_deelnemers = deelnemers[deelnemers['klasse'] == klasse]
        klasse_bibs = set(klasse_deelnemers['bib'])

        klasse_result = uitslag[uitslag['bib'].isin(klasse_bibs)].copy()
        klasse_result['rank_in_class'] = klasse_result['plaats'].rank(method='first').astype(int)

        punten = {}
        for i, row in enumerate(klasse_result.itertuples(), start=1):
            punten[row.bib] = i if i < 60 else 60

        for _, rijder in klasse_deelnemers.iterrows():
            bib = rijder['bib']
            punten_per_rijder.append({'bib': bib, week_col: punten.get(bib, MAX_POINTS)})
            klasse_per_rijder.append({'bib': bib, f'Klasse_{week_col}': rijder['klasse']})
            categorie_per_rijder.append({'bib': bib, f'Cat_{week_col}': rijder['categorie']})

    punten_df = pd.DataFrame(punten_per_rijder)
    klasse_df = pd.DataFrame(klasse_per_rijder)
    categorie_df = pd.DataFrame(categorie_per_rijder)

    if os.path.isfile(KLASSEMENT_FILE):
        klassement_df = pd.read_excel(KLASSEMENT_FILE, sheet_name="Regelmatigheidscriterium")

        # Rename columns for internal consistency if needed
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

    # Merge new week points and klasse/categorie separately without overwriting old weeks
    klassement_df = klassement_df.merge(punten_df, on='bib', how='left')
    klassement_df = klassement_df.merge(klasse_df, on='bib', how='left')
    klassement_df = klassement_df.merge(categorie_df, on='bib', how='left')

    # Fill missing points for new week with MAX_POINTS
    klassement_df[week_col] = klassement_df[week_col].fillna(MAX_POINTS)

    # Make sure all previous klasse/categorie columns are preserved and new ones added
    # Fill missing new week klasse/categorie from previous values or leave as is
    klassement_df[f'Klasse_{week_col}'] = klassement_df[f'Klasse_{week_col}'].fillna('Unknown')
    klassement_df[f'Cat_{week_col}'] = klassement_df[f'Cat_{week_col}'].fillna('Unknown')

    # Get all week number columns (only point columns)
    week_cols = [col for col in klassement_df.columns if col.isdigit()]
    week_cols = sorted(week_cols, key=int)

    # Total points per rider
    klassement_df['total'] = klassement_df[week_cols].sum(axis=1)

    if week_cols:
        if IS_SECOND_PERIOD_STARTED:
            first_period_weeks = [col for col in week_cols if int(col) < int(week_cols[-1])]
            second_period_weeks = [col for col in week_cols if int(col) >= int(week_cols[-1])]
        else:
            first_period_weeks = week_cols
            second_period_weeks = []

        klassement_df['eerst_heft'] = klassement_df[first_period_weeks].sum(axis=1)
        klassement_df['tweede_heft'] = klassement_df[second_period_weeks].sum(axis=1)

    # To rank riders by klasse per their current klasse (for the latest week)
    # You can choose how to rank: by last known klasse or by baseline klasse column

    # Here we rank by their klasse in the latest week:
    # Extract latest klasse for each rider
    klassement_df['current_klasse'] = klassement_df[f'Klasse_{week_cols[-1]}'] if week_cols else klassement_df['klasse']

    klassement_df['Plaats Klasse'] = (
        klassement_df.sort_values(by=['current_klasse', 'total'])
        .groupby('current_klasse')
        .cumcount() + 1
    )

    # Sort by klasse and total points
    klassement_df = klassement_df.sort_values(by=['current_klasse', 'total'])

    # Rename columns back to final Excel output names
    klassement_df = klassement_df.rename(columns={
        'bib': 'Nr.',
        'naam': 'Naam',
        'klasse': 'Klasse',
        'categorie': 'Cat.',
        'total': 'Totaal',
        'eerst_heft': '1e Periode',
        'tweede_heft': '2e Periode'
    })

    # Add the current_klasse column for debugging or remove it before saving
    klassement_df.drop(columns=['current_klasse'], inplace=True)

    # Determine final column order
    final_column_order = load_template_column_order()
    if 'Plaats Klasse' not in final_column_order:
        try:
            klasse_idx = final_column_order.index('Klasse')
            final_column_order.insert(klasse_idx + 1, 'Plaats Klasse')
        except ValueError:
            final_column_order.append('Plaats Klasse')

    week_cols_in_output = [col for col in klassement_df.columns if col.isdigit()]
    final_cols = final_column_order + [col for col in week_cols_in_output if col not in final_column_order]
    klassement_df = klassement_df[[col for col in final_cols if col in klassement_df.columns]]

    # Write to Excel (overwrite with updated sheet)
    with pd.ExcelWriter(KLASSEMENT_FILE, engine='openpyxl', mode='w') as writer:
        klassement_df.to_excel(writer, sheet_name="Regelmatigheidscriterium", index=False)

    # Formatting
    wb = load_workbook(KLASSEMENT_FILE)
    sheet = wb['Regelmatigheidscriterium']

    pink_fill = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")

    # Highlight DAM rows pink
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        if row[3].value == 'DAM':
            for cell in row:
                cell.fill = pink_fill

    # Highlight week columns green for <=4 and blue otherwise
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
