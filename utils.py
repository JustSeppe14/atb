import pandas as pd
import os
from openpyxl import load_workbook


DEELNEMERS_FILE = "Deelnemers/deelnemerslijst 2025.xlsx"
RESULT_FILE = "Result/finish.xlsx"
TEMPLATE_FILE = "Template/klassement.xlsx"
MAX_POINTS = 80

def load_deelnemers():
    df = pd.read_excel(DEELNEMERS_FILE, header=4)
    df.columns = df.columns.str.strip().str.lower()
    df = df.rename(columns={
        'number': 'bib',
        'name': 'naam',
        'klasse': 'klasse',
        'cat': 'categorie',
        'team': 'team'
    })
    df = df.dropna(subset=['bib', 'naam', 'klasse'])
    df['naam'] = df['naam'].str.strip().str.lower()
    df['bib'] = df['bib'].astype(int)
    return df

def load_result():
    df = pd.read_excel(RESULT_FILE)
    df.columns = df.columns.str.strip().str.lower()
    df = df.rename(columns={'pl': 'plaats', 'bib': 'bib', 'naam': 'naam'})
    df['bib'] = pd.to_numeric(df['bib'], errors='coerce').astype('Int64')
    df = df.dropna(subset=['bib', 'plaats'])
    return df

def get_current_week(overall_path, sheet_name):
    # If file doesn't exist, don't create anything — just start at week 1
    if not os.path.isfile(overall_path):
        return 1

    wb = load_workbook(overall_path)

    # Create the sheet if it's missing
    if sheet_name not in wb.sheetnames:
        wb.create_sheet(title=sheet_name)
        wb.save(overall_path)
        return 1

    # Sheet exists — check existing week columns
    df = pd.read_excel(overall_path, sheet_name=sheet_name)
    week_cols = [col for col in df.columns if str(col).isdigit()]
    return len(week_cols) + 1

def load_template_column_order():
    template_df = pd.read_excel(TEMPLATE_FILE, sheet_name=0, nrows=0)
    return [col for col in template_df.columns if not str(col).startswith("Unnamed")]
