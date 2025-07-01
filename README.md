# ATB Klassement Automatisering

Dit project automatiseert het samenvoegen van verschillende Excel-bestanden tot één overzichtelijk bestand en verstuurt het resultaat automatisch per e-mail naar de gewenste ontvangers.

## Functionaliteit

- Combineert verschillende klassement- en team-bestanden tot één Excel-bestand (`wedstrijd_data_2025.xlsx`) met meerdere bladen.
- Verstuurt het samengestelde bestand automatisch als bijlage via e-mail naar één of meerdere ontvangers.

## Vereisten

- Node.js & npm
- Een e-mailaccount (bijvoorbeeld Gmail) met SMTP-toegang

## Installatie

1. **Clone deze repository**  
   Download of clone dit project naar je lokale machine.
   ```
   git clone https://github.com/JustSeppe14/atb.git
   cd atb
   ```

2. **Installeer de benodigde packages**  
   Voer in de terminal uit:
   ```
   npm install
   ```

3. **.env bestand instellen**
   - Er is een voorbeeldbestand `.env copy` aanwezig.
   - Hernoem dit bestand naar `.env` en vul de nodige informatie in (zie onder).
   - ```
     EMAIL_ACCOUNT=je.email@provider.com
     EMAIL_PASSWORD=je_app_wachtwoord_of_email_wachtwoord
     EMAIL_RECIPIENTS=ontvanger1@provider.com,ontvanger2@provider.com
     GOOGLE_SHEETS_ID=
     GOOGLE_SHEETS_GID=
     ```
     > **LET OP!** Gebruik een App Password als je 2FA hebt ingeschakeld.

4. **Voeg je Excel-bestanden toe**  
   Zorg dat je je deelnemersbestand in de juiste map hebt staan.

## Gebruik

Voer het volgende script uit om alles automatisch te laten verlopen:
```
python generate_all.py
```

## Opmerkingen

- Controleer altijd of je `.env` bestand niet wordt meegestuurd in versiebeheer (staat in `.gitignore`).
- Je kunt meerdere ontvangers opgeven door e-mailadressen te scheiden met een komma in de `EMAIL_RECIPIENTS` variabele.
- Logging van de scripts is zichtbaar in de terminal voor eenvoudige foutopsporing.

---

Voor vragen of problemen, neem contact op met de beheerder van dit project.