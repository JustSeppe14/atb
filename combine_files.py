import pandas as pd

# Bestanden en bijhorende bladen
REGELMATIGHEID_BESTAND = "klassement_2025.xlsx"
KLASSEMENT_BESTAND = "klassement_totaal_2025.xlsx"
UITVOER_BESTAND = "wedstrijd_data_2025.xlsx"

# Laad de gewenste sheets
regelmatigheid_df = pd.read_excel(REGELMATIGHEID_BESTAND, sheet_name="Regelmatigheidscriterium")
klassement_df = pd.read_excel(KLASSEMENT_BESTAND, sheet_name="Klassement")

# Combineer in één bestand met twee bladen
with pd.ExcelWriter(UITVOER_BESTAND, engine='openpyxl') as writer:
    regelmatigheid_df.to_excel(writer, sheet_name="Regelmatigheidscriterium", index=False)
    klassement_df.to_excel(writer, sheet_name="Klassement", index=False)

print(f"✅ Beide bestanden zijn succesvol samengevoegd in '{UITVOER_BESTAND}' met twee bladen.")
