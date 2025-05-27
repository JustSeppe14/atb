import pandas as pd

# Bestanden en bijhorende bladen
REGELMATIGHEID_BESTAND = "klassement_2025.xlsx"
KLASSEMENT_BESTAND = "klassement_totaal_2025.xlsx"
TEAMSTTA_BESTAND = "team_klassement_2025.xlsx"
UITVOER_BESTAND = "wedstrijd_data_2025.xlsx"

# Laad de gewenste sheets
klassement_df = pd.read_excel(KLASSEMENT_BESTAND, sheet_name="KLASSEMENT")
regelmatigheid_df = pd.read_excel(REGELMATIGHEID_BESTAND, sheet_name="REGELMATIGHEIDSCRITERIUM")
teamsta_df = pd.read_excel(TEAMSTTA_BESTAND, sheet_name="TEAMS STA")


# Combineer in één bestand met twee bladen
with pd.ExcelWriter(UITVOER_BESTAND, engine='openpyxl') as writer:
    klassement_df.to_excel(writer, sheet_name="KLASSEMENT", index=False)
    regelmatigheid_df.to_excel(writer, sheet_name="REGELMATIGHEIDSCRITERIUM", index=False)
    teamsta_df.to_excel(writer, sheet_name="TEAMS STA", index=False)
   
print(f"✅ Beide bestanden zijn succesvol samengevoegd in '{UITVOER_BESTAND}' met twee bladen.")
