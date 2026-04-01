#!/usr/bin/env python3
"""
UPS Kraftstoffzuschlag Updater
===============================
Liest die aktuellen Kraftstoffzuschläge von der UPS-Website (statisches HTML)
und aktualisiert die Excel-Datei im KEP Frachtenrechner Repository.

Ausführung: python scripts/update_fuel_surcharges.py
"""

import re
import sys
import openpyxl
import requests
from bs4 import BeautifulSoup
from pathlib import Path

# ── Konfiguration ─────────────────────────────────────────────────────────────

EXCEL_FILENAME = "0493_Frachtenrechner_KEP_DATA.xlsx"
SHEET_NAME = "adds"
UPS_URL = "https://www.ups.com/de/de/support/shipping-support/shipping-costs-rates/fuel-surcharges"

# Tarif-Zuordnung: Welche Tarife bekommen den Standard-, welche den Express-Satz
STANDARD_TARIFFS = {
    "StandardSingleDE", "StandardSingleALL",
    "StandardMultiDE", "StandardMultiALL"
}
EXPRESS_TARIFFS = {
    "ExpressSaverALL_env", "ExpressSaverALL_doc", "ExpressSaverALL_pkg",
    "ExpressALL_env",      "ExpressALL_doc",      "ExpressALL_pkg",
    "ExpressNoon_env",     "ExpressNoon_pkg"
}

# Plausibilitätsbereiche (als Dezimalzahl)
STANDARD_RATE_MIN = 0.10   # 10 %
STANDARD_RATE_MAX = 0.50   # 50 %
EXPRESS_RATE_MIN  = 0.20   # 20 %
EXPRESS_RATE_MAX  = 0.70   # 70 %

# ── Hilfsfunktionen ───────────────────────────────────────────────────────────

def parse_percent(text: str) -> float | None:
    """
    Wandelt einen deutschen/englischen Prozentwert-String in eine Dezimalzahl um.
    z.B. '25,50 %' -> 0.255  |  '43.5%' -> 0.435
    """
    text = text.strip().replace("\xa0", "").replace(" ", "").replace("%", "")
    text = text.replace(",", ".")
    try:
        return round(float(text) / 100, 6)
    except ValueError:
        return None


# ── Haupt-Scraping-Funktion ───────────────────────────────────────────────────

def scrape_ups_fuel_surcharges() -> tuple[float, float]:
    """
    Ruft die aktuellen UPS Kraftstoffzuschläge per HTTP-Request ab.
    Die Daten sind im statischen HTML der Seite enthalten.
    Gibt (standard_rate, express_rate) als Dezimalzahlen zurück.
    """
    print(f"\n[1] Lade UPS-Seite: {UPS_URL}")

    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "de-DE,de;q=0.9,en-US;q=0.8,en;q=0.7",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
    }

    response = requests.get(UPS_URL, headers=headers, timeout=30)
    response.raise_for_status()
    print(f"  HTTP Status: {response.status_code} | Größe: {len(response.text)} Zeichen")

    soup = BeautifulSoup(response.text, "html.parser")

    # Alle Prozentwerte zur Übersicht ausgeben
    all_pct = re.findall(r"\d{1,3}[,.]\d{1,2}\s*%", response.text)
    print(f"  Alle Prozentwerte auf der Seite: {all_pct[:20]}")

    standard_rate = None
    express_rate  = None

    # ── Strategie 1: Tabellenzeilen mit Datum-Spalte (DD/MM/YYYY) ────────────
    print("\n[2] Suche in Tabellen nach aktuellster Zeile …")
    date_pattern = re.compile(r"^\d{2}/\d{2}/\d{4}$")

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        # Header-Zeile überspringen
        for row in rows:
            cells = [td.get_text(strip=True) for td in row.find_all("td")]
            if len(cells) < 3:
                continue
            # Prüfen ob erste Spalte ein Datum ist
            if not date_pattern.match(cells[0]):
                continue
            # Erste Daten-Zeile = aktuellste Woche
            s = parse_percent(cells[1])
            e = parse_percent(cells[2])
            if s and e:
                standard_rate = s
                express_rate  = e
                print(f"  Datum: {cells[0]} | Standard: {cells[1]} -> {s:.4f} | Express: {cells[2]} -> {e:.4f}")
                break
        if standard_rate and express_rate:
            break

    # ── Strategie 2: Alle Prozentwerte aus dem Text (Fallback) ───────────────
    if standard_rate is None or express_rate is None:
        print("\n[3] Fallback: Suche im Volltext …")
        text = soup.get_text()
        pct_pattern = re.compile(r"(\d{1,3}[,.]\d{1,2})\s*%")

        for line in text.splitlines():
            line_lower = line.lower()
            percents = pct_pattern.findall(line)
            if not percents:
                continue
            if "standard" in line_lower and standard_rate is None:
                rate = parse_percent(percents[0])
                if rate and STANDARD_RATE_MIN <= rate <= STANDARD_RATE_MAX:
                    standard_rate = rate
                    print(f"  Standard gefunden: {rate:.4f} in '{line.strip()}'")
            if "express" in line_lower and express_rate is None:
                rate = parse_percent(percents[0])
                if rate and EXPRESS_RATE_MIN <= rate <= EXPRESS_RATE_MAX:
                    express_rate = rate
                    print(f"  Express gefunden:  {rate:.4f} in '{line.strip()}'")

    # ── Ergebnis-Validierung ─────────────────────────────────────────────────
    if standard_rate is None or express_rate is None:
        # HTML-Auszug für Debugging ausgeben
        print("\n  --- HTML-Auszug (erste 3000 Zeichen) ---")
        print(response.text[:3000])
        print("  ---")
        raise ValueError(
            f"Kraftstoffzuschläge konnten nicht extrahiert werden! "
            f"Standard={standard_rate}, Express={express_rate}"
        )

    if not (STANDARD_RATE_MIN <= standard_rate <= STANDARD_RATE_MAX):
        raise ValueError(f"Standard-Satz außerhalb des Plausibilitätsbereichs: {standard_rate}")
    if not (EXPRESS_RATE_MIN <= express_rate <= EXPRESS_RATE_MAX):
        raise ValueError(f"Express-Satz außerhalb des Plausibilitätsbereichs: {express_rate}")

    return standard_rate, express_rate


# ── Excel-Update ──────────────────────────────────────────────────────────────

def update_excel(
    excel_path: Path,
    standard_rate: float,
    express_rate: float
) -> tuple[bool, dict]:
    """
    Aktualisiert die Fuel Surcharge-Werte in der Excel-Datei.
    Gibt (changed, old_values) zurück.
    """
    wb = openpyxl.load_workbook(excel_path)

    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"Tabellenblatt '{SHEET_NAME}' nicht in {excel_path} gefunden!\nVorhandene Blätter: {wb.sheetnames}")

    ws      = wb[SHEET_NAME]
    changed    = False
    old_values = {}

    for row in ws.iter_rows(min_row=2):
        tarif = row[0].value
        if tarif is None:
            continue

        old_val = row[1].value
        new_val = None

        if tarif in STANDARD_TARIFFS:
            new_val = standard_rate
        elif tarif in EXPRESS_TARIFFS:
            new_val = express_rate

        if new_val is not None and old_val != new_val:
            old_values[tarif] = old_val
            row[1].value = new_val
            changed = True

    if changed:
        wb.save(excel_path)

    return changed, old_values


# ── Einstiegspunkt ────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  UPS Kraftstoffzuschlag Updater")
    print("=" * 60)

    # Excel-Datei suchen
    excel_path = Path(EXCEL_FILENAME)
    if not excel_path.exists():
        candidates = list(Path(".").glob("*.xlsx"))
        print(f"  {EXCEL_FILENAME} nicht gefunden. Verfügbare xlsx: {candidates}")
        sys.exit(1)

    print(f"\nExcel-Datei: {excel_path.resolve()}")

    # Aktuelle Werte aus Excel lesen (für Vergleich)
    wb_check = openpyxl.load_workbook(excel_path)
    if SHEET_NAME not in wb_check.sheetnames:
        print(f"  FEHLER: Blatt '{SHEET_NAME}' nicht gefunden. Vorhandene Blätter: {wb_check.sheetnames}")
        sys.exit(1)
    ws_check = wb_check[SHEET_NAME]
    current_standard = next(
        (r[1].value for r in ws_check.iter_rows(min_row=2) if r[0].value in STANDARD_TARIFFS), None
    )
    current_express = next(
        (r[1].value for r in ws_check.iter_rows(min_row=2) if r[0].value in EXPRESS_TARIFFS), None
    )
    print(f"Aktuelle Werte -> Standard: {current_standard}, Express: {current_express}")

    # Neue Werte von UPS holen
    try:
        new_standard, new_express = scrape_ups_fuel_surcharges()
    except Exception as e:
        print(f"\n❌ Fehler beim Abrufen: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"\n✅ Neue Werte -> Standard: {new_standard:.4f} ({new_standard*100:.2f}%)"
          f"  |  Express: {new_express:.4f} ({new_express*100:.2f}%)")

    # Excel aktualisieren
    changed, old_vals = update_excel(excel_path, new_standard, new_express)

    if changed:
        print(f"\n📝 Excel aktualisiert ({len(old_vals)} Zellen geändert):")
        for tarif, old in old_vals.items():
            new = new_standard if tarif in STANDARD_TARIFFS else new_express
            print(f"   {tarif}: {old} -> {new}")
        print(f"\n::set-output name=changed::true")
        print(f"::set-output name=standard_rate::{new_standard}")
        print(f"::set-output name=express_rate::{new_express}")
    else:
        print("\n✅ Keine Änderungen – Werte sind bereits aktuell.")
        print("::set-output name=changed::false")


if __name__ == "__main__":
    main()
