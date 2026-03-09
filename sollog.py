#Sollicitatielog
# Versie 1: Haalt info uit een Gmail map, en schrijft dit naar een CSV-bestand
# Versie 2: Zelfde info, maar naar een Excel-bestand

import imaplib
import email
from email.header import decode_header
import csv
import os
from datetime import datetime
from dotenv import load_dotenv
load_dotenv()
from openpyxl import load_workbook, Workbook

# Instelling laden via dotenv, zodat het paswoord niet openbaar wordt
# Zeker pip install python niet vergeten

# ─── INSTELLINGEN ───────────────────────────────────────────
GMAIL_ADRES    = os.getenv("GMAIL_ADRES")
APP_WACHTWOORD = os.getenv("APP_WACHTWOORD")
IMAP_MAP       = os.getenv("IMAP_MAP")
XLS_PAD        = os.getenv("XLSPAD")
# ────────────────────────────────────────────────────────────

def decodeer_header(waarde):
    if not waarde:
        return ""
    delen = decode_header(waarde)
    resultaat = []
    for deel, codering in delen:
        if isinstance(deel, bytes):
            resultaat.append(deel.decode(codering or "utf-8", errors="replace"))
        else:
            resultaat.append(deel)
    return " ".join(resultaat)

def haal_mails_op():
    print("Verbinding maken met Gmail...")
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_ADRES, APP_WACHTWOORD)

    status, _ = mail.select(f'"{IMAP_MAP}"')
    if status != "OK":
        print(f"❌ Map '{IMAP_MAP}' niet gevonden. Controleer de mapnaam.")
        mail.logout()
        return []

    status, berichten = mail.search(None, "ALL")
    if status != "OK" or not berichten[0]:
        print("Geen mails gevonden in deze map.")
        mail.logout()
        return []

    mail_ids = berichten[0].split()
    print(f"✅ {len(mail_ids)} mail(s) gevonden.")

    resultaten = []
    for mail_id in mail_ids:
        status, data = mail.fetch(mail_id, "(RFC822)")
        if status != "OK":
            continue

        bericht = email.message_from_bytes(data[0][1])
        onderwerp = decodeer_header(bericht.get("Subject", "(geen onderwerp)"))
        aan       = decodeer_header(bericht.get("To", "(onbekend)"))
        aan       = email.utils.parseaddr(aan)[1]
        domein = aan.split("@")[1] if "@" in aan else ""
        bedrijf = domein.split(".")[0].capitalize() if domein else "(onbekend)"
        datum_raw = bericht.get("Date", "")

        try:
            datum_obj = email.utils.parsedate_to_datetime(datum_raw)
            datum     = datum_obj.strftime("%d-%m-%Y")
        except Exception:
            datum = datum_raw

        resultaten.append({
            "Datum":     datum,
            "Aan":       aan,
            "Bedrijf":   bedrijf,
            "Onderwerp": onderwerp,
            
        })

    mail.logout()
    return resultaten

def sla_op_als_excel(mails):
    os.makedirs(os.path.dirname(XLS_PAD), exist_ok=True)

    # Bestand aanmaken als het nog niet bestaat
    if not os.path.exists(XLS_PAD):
        wb = Workbook()
        ws = wb.active
        ws.append(["Datum", "Aan", "Bedrijf", "Onderwerp"])
        wb.save(XLS_PAD)

    wb = load_workbook(XLS_PAD)
    ws = wb.active

    # Bestaande rijen inlezen om duplicaten te vermijden
    bestaande = set()
    for rij in ws.iter_rows(min_row=2, values_only=True):
        bestaande.add((rij[0], rij[3]))  # Datum + Onderwerp

    nieuw = [m for m in mails if (m["Datum"], m["Onderwerp"]) not in bestaande]

    if not nieuw:
        print("ℹ️  Geen nieuwe mails om toe te voegen.")
        return

    for m in nieuw:
        ws.append([m["Datum"], m["Aan"], m["Bedrijf"], m["Onderwerp"]])

    wb.save(XLS_PAD)
    print(f"✅ {len(nieuw)} nieuwe sollicitatie(s) toegevoegd aan {XLS_PAD}")

if __name__ == "__main__":
    mails = haal_mails_op()
    try:
        if mails:
            sla_op_als_excel(mails)
    except Exception as e:
        print(f"❌ Fout: {e}")
#Leesbaar maken van fouten        
    print("\nKlaar! Druk op Enter om af te sluiten.")
    input()