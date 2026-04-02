"""
lv_einlesen.py – Einmalig ausführen um LV-PDFs zu parsen.

Verwendung in Termux:
    cd ~/baudoku-pro
    python lv_einlesen.py

Erzeugt für jedes Projekt unter Bau_Projekte/<ProjektName>/Projekt_Infos/
die Datei lv_positionen.json mit allen gefundenen Positionen.
"""

import os
import sys
from modules import lv_parser
import config


def verarbeite_alle_projekte():
    if not os.path.exists(config.BASIS_PFAD):
        print(f"Kein Projekt-Ordner gefunden: {config.BASIS_PFAD}")
        sys.exit(1)

    projekte = [p for p in os.listdir(config.BASIS_PFAD)
                if os.path.isdir(os.path.join(config.BASIS_PFAD, p))]

    if not projekte:
        print("Keine Projekte gefunden.")
        sys.exit(0)

    gesamt = 0
    for projekt_name in sorted(projekte):
        p_pfad = os.path.join(config.BASIS_PFAD, projekt_name)
        lv_ordner = os.path.join(p_pfad, "Projekt_Infos")

        if not os.path.exists(lv_ordner):
            print(f"\n[{projekt_name}] Kein Projekt_Infos-Ordner – übersprungen")
            continue

        # Prüfen ob PDFs vorhanden
        pdfs = [f for f in os.listdir(lv_ordner) if f.lower().endswith(".pdf")]
        if not pdfs:
            print(f"\n[{projekt_name}] Keine PDFs gefunden – übersprungen")
            continue

        print(f"\n[{projekt_name}] {len(pdfs)} PDF(s) werden eingelesen...")
        positionen = lv_parser.lade_alle_lv(p_pfad)

        if not positionen:
            print(f"  → Keine Positionen gefunden (PDF-Struktur prüfen)")
            continue

        json_pfad = os.path.join(lv_ordner, "lv_positionen.json")
        lv_parser.speichere_lv_json(positionen, json_pfad)
        gesamt += len(positionen)

        # Kurze Vorschau
        print(f"  Vorschau (erste 5):")
        for p in positionen[:5]:
            print(f"    {p['anzeige'][:80]}")

    print(f"\nFertig! Gesamt {gesamt} Positionen eingelesen.")
    print("Du kannst jetzt die App starten: python app.py")


if __name__ == "__main__":
    verarbeite_alle_projekte()
