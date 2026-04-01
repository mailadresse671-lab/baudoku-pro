import base64
import json
import re
from groq import Groq
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


def _encode_image(path: str) -> tuple[str, str]:
    ext = path.rsplit(".", 1)[-1].lower()
    media_type = "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"
    with open(path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode("utf-8")
    return data, media_type


def _build_prompt(lv_text: str, report_type: str) -> str:
    felder = BAUTAGESBERICHT_FELDER if report_type == "bautagesbericht" else REGIEBERICHT_FELDER

    lv_kontext = f"""
LEISTUNGSVERZEICHNIS (für LV-Positionen Zuordnung):
{lv_text[:60000]}
""" if lv_text else "Kein LV vorhanden."

    if report_type == "bautagesbericht":
        aufgabe = """Erstelle einen Bautagesbericht für Straßen- und Tiefbau.
Analysiere die Fotos und erkenne welche Arbeiten durchgeführt wurden.
Ordne die Arbeiten den LV-Positionen zu wenn möglich.
Erfinde KEINE Mengen - diese werden manuell eingetragen."""
    else:
        aufgabe = """Erstelle einen Regiebericht für Straßen- und Tiefbau.
Das sind Arbeiten die NICHT im Leistungsverzeichnis stehen.
Erkläre warum diese Arbeiten als Regie abgerechnet werden.
Stunden und Mengen auf 0 setzen."""

    return f"""Du bist Polier im Straßen- und Tiefbau. {aufgabe}

REGELN:
- Wenn du etwas nicht sicher erkennen kannst → in "ki_hinweise" schreiben
- Keine Mengen erfinden (Meter, m², Tonnen)
- Fachbegriffe Tiefbau verwenden
- Antwort NUR als JSON, kein Text davor oder danach

{lv_kontext}

JSON STRUKTUR:
{felder}"""


def analyze(image_paths: list[str], lv_text: str, report_type: str = "bautagesbericht") -> dict:
    if not config.GROQ_API_KEY:
        raise ValueError("GROQ_API_KEY fehlt in der .env Datei")

    client = Groq(api_key=config.GROQ_API_KEY)

    # Bilder als base64 vorbereiten (max 10, Groq empfiehlt max 5 für beste Ergebnisse)
    content = []
    for path in image_paths[:5]:
        try:
            data, media_type = _encode_image(path)
            content.append({
                "type": "image_url",
                "image_url": {"url": f"data:{media_type};base64,{data}"}
            })
        except Exception as e:
            print(f"Bild konnte nicht geladen werden {path}: {e}")

    if not content:
        raise ValueError("Keine Bilder konnten geladen werden")

    content.append({"type": "text", "text": _build_prompt(lv_text, report_type)})

    response = client.chat.completions.create(
        model="meta-llama/llama-4-scout-17b-16e-instruct",
        messages=[{"role": "user", "content": content}],
        response_format={"type": "json_object"},
        max_tokens=2000
    )

    raw = response.choices[0].message.content.strip()

    # JSON extrahieren falls nötig
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if match:
        raw = match.group(0)

    data = json.loads(raw)

    if not data.get("ist_baustelle", True):
        raise ValueError("Kein Baustellen-Foto erkannt.")

    return data
