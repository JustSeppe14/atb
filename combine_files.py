import pandas as pd
import os

os.makedirs("output", exist_ok=True)
# Bestanden en bijhorende bladen
KLASSEMENT_BESTAND = "output/klassement_totaal_2025.xlsx"
REGELMATIGHEID_BESTAND = "output/klassement_2025.xlsx"
TEAMSTTA_BESTAND = "output/team_klassement_2025.xlsx"
TEAMSDAM_BESTAND = "output/team_klassement_2025_DAM_only.xlsx"
UITVOER_BESTAND = "wedstrijd_data_2025.xlsx"

# Laad de gewenste sheets
regelmatigheid_df = pd.read_excel(REGELMATIGHEID_BESTAND, sheet_name="REGELMATIGHEIDSCRITERIUM")
klassement_df = pd.read_excel(KLASSEMENT_BESTAND, sheet_name="KLASSEMENT")
teamsta_df = pd.read_excel(TEAMSTTA_BESTAND, sheet_name="TEAMS STA")
teamdam_df = pd.read_excel(TEAMSDAM_BESTAND, sheet_name="TEAMS MIXED")


# Combineer in één bestand met twee bladen
with pd.ExcelWriter(UITVOER_BESTAND, engine='openpyxl') as writer:
    regelmatigheid_df.to_excel(writer, sheet_name="REGELMATIGHEIDSCRITERIUM", index=False)
    klassement_df.to_excel(writer, sheet_name="KLASSEMENT", index=False)
    teamsta_df.to_excel(writer, sheet_name="TEAMS STA", index=False)
    teamdam_df.to_excel(writer, sheet_name="TEAMS MIXED", index=False)
   
print(f"✅ Alle bestanden zijn succesvol samengevoegd in '{UITVOER_BESTAND}' met 4 bladen.")
