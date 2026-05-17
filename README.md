# Sollicitatielogger

Dit script verzamelt mails uit een specifieke map op je IMAP-mailaccount. Daarna schrijft hij de relevante informatie (datum, afzender, bedrijf, onderwerp) in een Excel-logbestand. Het is ontworpen om sollicitatie-antwoorden bij te houden in een georganiseerde manier.

## Versiegeschiedenis

- **Versie 1**: De info werd weggeschreven in een CSV-bestand.
- **Versie 2**: Het CSV-bestand werd vervangen door een Excel-bestand.
- **Versie 3**: Toegevoegd: automatische checks voor permissies (bijv. als Excel-bestand open is), duplicate prevention om dubbele e-mails te voorkomen, automatische aanmaak van Excel-structuur bij eerste gebruik en ondersteuning voor internationale karakters in e-mailheaders. Het script is nu klaar voor gebruik met alle IMAP-compatibele mailaccounts.

## Installatie

1. Zorg ervoor dat je Python 3 hebt geïnstalleerd (versie 3.6 of hoger aanbevolen).
2. Installeer de benodigde dependencies via pip:
   ```
   pip install openpyxl python-dotenv
   ```
3. (Aanbevolen) Maak een `requirements.txt` bestand met de volgende inhoud:
   ```
   openpyxl
   python-dotenv
   ```
   Installeer dan met: `pip install -r requirements.txt`

## Configuratie

Het script gebruikt een `.env` bestand voor gevoelige informatie zoals je mailgegevens. Maak een bestand genaamd `.env` in dezelfde map als `sollog.py` met de volgende inhoud:

```
IMAP_SERVER=imap.gmail.com
MAIL_ADRES=jouw.email@example.com
MAIL_WACHTWOORD=jouw_wachtwoord
IMAP_MAP=Sollicitaties
DATA_MAP=C:\Pad\Naar\Uitvoermap
```

- `IMAP_SERVER`: Het IMAP-serveradres van je mailprovider (bijv. `imap.gmail.com` voor Gmail).
- `MAIL_ADRES`: Je volledige e-mailadres.
- `MAIL_WACHTWOORD`: Je e-mailwachtwoord (of app-wachtwoord voor Gmail).
- `IMAP_MAP`: De naam van de map waar de sollicitatie-mails staan (bijv. "Inbox", "Sollicitaties").
- `DATA_MAP`: Het volledige pad naar de map waar het Excel-bestand opgeslagen moet worden.

**Belangrijk:** Voeg `.env` toe aan je `.gitignore` bestand om je wachtwoord niet te delen.

## Gebruik

1. Zorg dat je `.env` bestand correct is ingesteld.
2. Open een terminal in de map waar `sollog.py` staat.
3. Voer het script uit:
   ```
   python sollog.py
   ```
4. Het script zal verbinding maken met je mailaccount, e-mails ophalen, en nieuwe entries toevoegen aan `sollicitatielog.xlsx` in de opgegeven `DATA_MAP`.
5. Voorbeeld output:
   ```
   ✅ MAIL OK: 15-03-2024 | hr@example.com | Bedankt voor je sollicitatie
   ⚠️  Geen nieuwe mails gevonden.
   📁 Log opgeslagen in: C:\Pad\Naar\Uitvoermap\sollicitatielog.xlsx
   ```

## Ondersteunde IMAP Servers

Het script werkt met alle IMAP-compatibele mailproviders, zoals Gmail, Outlook, Yahoo, etc. Het is getest met Gmail, maar zou moeten werken met elke server die IMAP ondersteunt. Controleer de IMAP-instellingen van je provider voor het juiste serveradres en poort.

## Notities
Ontwikkeld met in VSCode met behulp van Claude Code.
