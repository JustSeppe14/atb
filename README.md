# ATB Klassement Automatisering

Dit project automatiseert het samenvoegen van verschillende Excel-bestanden tot één overzichtelijk bestand en verstuurt het resultaat automatisch per e-mail naar de gewenste ontvangers.

## Functionaliteit

- Combineert verschillende klassement- en team-bestanden tot één Excel-bestand (`wedstrijd_data_2025.xlsx`) met meerdere bladen.
- Verstuurt het samengestelde bestand automatisch als bijlage via e-mail naar één of meerdere ontvangers.

## Vereisten

- Python 3.8+
- [pip](https://pip.pypa.io/en/stable/)
- Een e-mailaccount (bijvoorbeeld Gmail) met SMTP-toegang
- Vereiste Python packages (zie hieronder)

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
- Er is een voorbeeldbestand `.env.copy` aanwezig.
- Maak een kopie van dit bestand en hernoem het naar `.env`:
    ```
    copy .env.copy .env
    ```
- Open `.env` en vul de juiste waarden in:
    ```
    EMAIL_ACCOUNT=je.email@provider.com
    EMAIL_PASSWORD=je_app_wachtwoord_of_email_wachtwoord
    EMAIL_RECIPIENTS=ontvanger1@provider.com,ontvanger2@provider.com
    ```
    > **LET OP!** Gebruik een App Password als je 2FA hebt ingeschakeld.
4. **Voeg je Excel-bestanden toe**
Zorgt dat je je deelnemers bestand in de juiste [Deelnemers](http://_vscodecontentref_/0) map hebt staan

## Gebruik

1. **Voer het script uit**
Voer het volgende script uit om alles automatisch zijn werk te laten doen:
    ```
    python generate_all.py
    ```

## Opmerkingen

- Controleer altijd of je [.env](http://_vscodecontentref_/3) bestand niet wordt meegestuurd in versiebeheer (staat in [.gitignore](http://_vscodecontentref_/4)).
- Je kunt meerdere ontvangers opgeven door e-mailadressen te scheiden met een komma in de `EMAIL_RECIPIENTS` variabele.
- Logging van de scripts is zichtbaar in de terminal voor eenvoudige foutopsporing.

---

Voor vragen of problemen, neem contact op met de beheerder van dit project.