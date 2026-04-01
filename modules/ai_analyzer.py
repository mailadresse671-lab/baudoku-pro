import json
import re
import time
import google.generativeai as genai
import config

BAUTAGESBERICHT_FELDER = """
{
  "ist_baustelle": true,
  "wetter_vormittag": "z.B. Sonnig, leicht bewölkt",
  "wetter_nachmittag": "z.B. Bewölkt, Regen",
  "temp_min": 0,
  "temp_max": 0,
  "personal_aufsicht": 0,
  "personal_facharbeiter": 0,
  "personal_maschinist": 0,
  "beschreibung_arbeiten": ["Zeile 1", "Zeile 2"],
  "lv_positionen": ["Pos. 1.2.3 - Beschreibung (falls erkennbar)"],
  "geraete_liste": ["Gerät 1", "Gerät 2"],
  "material_liste": ["Material 1"],
  "sonstiges": [],
  "ki_hinweise": ["Was die KI nicht sicher erkennen konnte"]
}
"""

REGIEBERICHT_FELDER = """
{
  "ist_baustelle": true,
  "beschreibung_arbeiten": ["Zeile 1", "Zeile 2"],
  "grund_regie": "Warum ist das Regiearbeit (nicht im LV)",
  "personal": [
    {"funktion": "Polier", "anzahl": 1, "stunden": 0},
    {"funktion": "Facharbeiter", "anzahl": 0, "stunden": 0},
    {"funktion": "Maschinist", "anzahl": 0, "stunden": 0}
  ],
  "geraete": [
    {"bezeichnung": "Bagger Takeuchi 6t", "einheit": "h", "menge": 0}
  ],
  "material": [
    {"bezeichnung": "Material", "einheit": "m³", "menge": 0}
  ],
  "ki_hinweise": ["Was die KI nicht sicher erkennen konnte"]
}
"""


def _build_prompt(lv_text: str, report_type: str) -> str:
    felder = BAUTAGESBERICHT_FELDER if report_type == "bautagesbericht" else REGIEBERICHT_FELDER

    lv_kontext = f"""
LEISTUNGSVERZEICHNIS (für LV-Positionen Zuordnung):
{lv_text[:80000]}
""" if lv_text else "Kein LV vorhanden."

    if report_type == "bautagesbericht":
        aufgabe = """Erstelle einen Bautagesbericht für Straßen- und Tiefbau.
Analysiere die Fotos und erkenne welche Arbeiten durchgeführt wurden.
Ordne die Arbeiten den LV-Positionen zu wenn möglich.
Erfinde KEINE Mengen - diese werden manuell eingetragen."""
    else:
        aufgabe = """Erstelle einen Regiebericht für Straßen- und Tiefbau.
Das sind Arbeiten die NICHT im Leistungsverzeichnis stehen (Zusatzarbeiten, unvorhergesehenes).
Erkläre warum diese Arbeiten als Regie abgerechnet werden.
Stunden und Mengen werden manuell eingetragen - setze diese auf 0."""

    return f"""Du bist Polier im Straßen- und Tiefbau. {aufgabe}

WICHTIGE REGELN:
- Sei ehrlich: Wenn du etwas nicht sicher erkennen kannst, schreibe es in "ki_hinweise"
- Erfinde keine Mengen (Meter, m², Tonnen) - diese sieht man nicht auf Fotos
- Verwende Fachbegriffe aus dem Tiefbau (z.B. "Leitungsgraben", "Frostschutzschicht", "Verdichtung")
- Antworte AUSSCHLIESSLICH als gültiges JSON ohne Markdown-Formatierung

{lv_kontext}

JSON STRUKTUR:
{felder}

Antworte nur mit dem JSON-Objekt, kein Text davor oder danach."""


def analyze(image_paths: list[str], lv_text: str, report_type: str = "bautagesbericht") -> dict:
    if not config.GEMINI_API_KEY:
        raise ValueError("GEMINI_API_KEY fehlt in der .env Datei")

    genai.configure(api_key=config.GEMINI_API_KEY)

    # Bilder hochladen (max 10)
    uploads = []
    for path in image_paths[:config.MAX_BILDER_PRO_TAG]:
        try:
            uploads.append(genai.upload_file(path))
        except Exception as e:
            print(f"Bild konnte nicht hochgeladen werden {path}: {e}")

    if not uploads:
        raise ValueError("Keine Bilder konnten hochgeladen werden")

    prompt = _build_prompt(lv_text, report_type)

    # Modelle der Reihe nach versuchen
    modelle = ["gemini-2.0-flash", "gemini-2.0-flash-lite", "gemini-1.5-flash-latest", "gemini-1.5-flash"]
    response = None

    for modell_name in modelle:
        try:
            model = genai.GenerativeModel(
                modell_name,
                generation_config={"response_mime_type": "application/json"}
            )
            for versuch in range(3):
                try:
                    response = model.generate_content([prompt, *uploads])
                    break
                except Exception as e:
                    if "429" in str(e) and versuch < 2:
                        print(f"Quota Limit — warte 60 Sekunden...")
                        time.sleep(60)
                    else:
                        raise
            break
        except Exception as e:
            if "404" in str(e) or "not found" in str(e).lower():
                print(f"Modell {modell_name} nicht verfügbar, versuche nächstes...")
                continue
            raise

    if not response:
        raise ValueError("Kein Gemini-Modell verfügbar. Bitte API-Key prüfen.")

    raw = response.text.strip()

    # JSON extrahieren falls Markdown vorhanden
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if match:
        raw = match.group(0)

    data = json.loads(raw)

    # Sicherheitscheck: ist es wirklich ein Baustellen-Foto?
    if not data.get("ist_baustelle", True):
        raise ValueError("Kein Baustellen-Foto erkannt. Bitte nur Baustellen-Bilder hochladen.")

    return data
