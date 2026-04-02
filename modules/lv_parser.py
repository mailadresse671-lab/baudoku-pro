"""
LV-Parser: Liest alle LV-PDFs und extrahiert strukturierte Positionen.
Wird einmalig ausgeführt und speichert das Ergebnis als JSON.
"""

import os
import re
import json
import subprocess
import pypdf


# LV-Titel aus dem Dateinamen / Inhalt
LV_TITEL = {
    "LV 1": "Fahrzeugrückhaltesystem",
    "LV 2": "Perimeterschutz - Zaun- und Toranlagen",
    "LV 3": "Leerrohrtrasse / Kabelziehschächte"
}

# Regex für Positionsnummern (alle drei LV-Formate)
POS_PATTERN = re.compile(
    r'^('
    r'\d+\.\.[0-9]+\.'       # LV1: 1..10. / 2..20.
    r'|\d+\.[0-9]+\.[0-9]+\.'  # LV2: 1.1.10. / 2.4.55.
    r'|\d+\.[0-9]+\.'         # LV2 Gruppe: 1.1. / 2.3.
    r'|\d+\.[\.]+[0-9]+'      # LV3: 1....1 / 2....1
    r')'
)

# Diese Zeilen sind kein Beschreibungstext
SKIP_WORDS = [
    "EUR", "Seite", "Druckdatum", "Menge ME", "Einheitspreis",
    "Gesamtbetrag", "Kurt Motz", "Ulmer Straße", "TransnetBW",
    "Illertissen", "Stuttgart", "Kalkulator", "Angebotssumme"
]


def pdf_zu_text(pfad: str) -> str:
    """PDF zu Text konvertieren - versucht pdftotext, dann pypdf."""
    # Methode 1: pdftotext (schneller, bessere Formatierung)
    try:
        result = subprocess.run(
            ["pdftotext", "-layout", pfad, "-"],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0 and result.stdout:
            return result.stdout
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    # Methode 2: pypdf Fallback
    try:
        reader = pypdf.PdfReader(pfad)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as e:
        print(f"  Fehler beim Lesen: {e}")
        return ""


def _zeile_ist_skip(zeile: str) -> bool:
    return any(skip in zeile for skip in SKIP_WORDS)


def parse_lv_text(text: str, lv_name: str) -> list[dict]:
    """Extrahiert alle Positionen aus einem LV-Text."""
    positionen = []
    zeilen = text.split("\n")
    i = 0

    while i < len(zeilen):
        zeile = zeilen[i].strip()
        match = POS_PATTERN.match(zeile)

        if match:
            pos_nr = match.group(1).rstrip(".")
            # Beschreibung: Rest der Zeile oder nächste Zeile
            rest = zeile[len(match.group(1)):].strip()

            if not rest or _zeile_ist_skip(rest):
                # Beschreibung in nächster Zeile suchen
                j = i + 1
                while j < len(zeilen):
                    naechste = zeilen[j].strip()
                    if naechste and not _zeile_ist_skip(naechste) and not POS_PATTERN.match(naechste):
                        rest = naechste
                        break
                    elif POS_PATTERN.match(naechste):
                        break
                    j += 1

            if rest and not _zeile_ist_skip(rest) and len(rest) > 3:
                # Kurztext: erste 120 Zeichen der Beschreibung
                kurztext = rest[:120].replace("\n", " ").strip()
                positionen.append({
                    "lv": lv_name,
                    "lv_titel": LV_TITEL.get(lv_name, lv_name),
                    "nr": pos_nr,
                    "beschreibung": kurztext,
                    "anzeige": f"{lv_name} | {pos_nr} — {kurztext}"
                })

        i += 1

    return positionen


def lade_alle_lv(projekt_pfad: str) -> list[dict]:
    """Liest alle LV-PDFs eines Projekts und gibt strukturierte Positionen zurück."""
    lv_ordner = os.path.join(projekt_pfad, "Projekt_Infos")
    alle_positionen = []

    if not os.path.exists(lv_ordner):
        return alle_positionen

    for datei in sorted(os.listdir(lv_ordner)):
        if not datei.lower().endswith(".pdf"):
            continue

        pfad = os.path.join(lv_ordner, datei)
        # LV-Name aus Dateiname: "LV 1.pdf" → "LV 1"
        lv_name = os.path.splitext(datei)[0]

        print(f"  Lese {datei}...")
        text = pdf_zu_text(pfad)
        if not text:
            print(f"  Kein Text gefunden in {datei}")
            continue

        positionen = parse_lv_text(text, lv_name)
        alle_positionen.extend(positionen)
        print(f"  → {len(positionen)} Positionen gefunden")

    return alle_positionen


def speichere_lv_json(positionen: list[dict], ausgabe_pfad: str):
    """Speichert die extrahierten Positionen als JSON."""
    with open(ausgabe_pfad, "w", encoding="utf-8") as f:
        json.dump(positionen, f, indent=2, ensure_ascii=False)
    print(f"  Gespeichert: {ausgabe_pfad} ({len(positionen)} Positionen)")


def lade_lv_json(projekt_pfad: str) -> list[dict]:
    """Lädt die gecachten LV-Positionen aus JSON (falls vorhanden)."""
    json_pfad = os.path.join(projekt_pfad, "Projekt_Infos", "lv_positionen.json")
    if os.path.exists(json_pfad):
        with open(json_pfad, encoding="utf-8") as f:
            return json.load(f)
    return []


def lv_fuer_ki(positionen: list[dict], max_zeichen: int = 6000) -> str:
    """
    Erstellt einen kompakten LV-Text für den KI-Prompt.
    Nur Position-Nummern und Kurzbeschreibungen.
    """
    zeilen = []
    gesamt = 0
    for p in positionen:
        zeile = f"{p['lv']} {p['nr']}: {p['beschreibung']}"
        gesamt += len(zeile)
        if gesamt > max_zeichen:
            break
        zeilen.append(zeile)
    return "\n".join(zeilen)
