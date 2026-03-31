import base64
import json
import re
import anthropic
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
    if not config.CLAUDE_API_KEY:
        raise ValueError("CLAUDE_API_KEY fehlt in der .env Datei")

    client = anthropic.Anthropic(api_key=config.CLAUDE_API_KEY)

    content = []

    # Bilder hinzufügen (max 10)
    for path in image_paths[:config.MAX_BILDER_PRO_TAG]:
        try:
            data, media_type = _encode_image(path)
            content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": media_type, "data": data}
            })
        except Exception as e:
            print(f"Bild konnte nicht geladen werden {path}: {e}")

    if not content:
        raise ValueError("Keine Bilder konnten geladen werden")

    content.append({"type": "text", "text": _build_prompt(lv_text, report_type)})

    response = client.messages.create(
        model=config.CLAUDE_MODEL,
        max_tokens=2000,
        messages=[{"role": "user", "content": content}]
    )

    raw = response.content[0].text.strip()

    # JSON extrahieren falls Markdown vorhanden
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if match:
        raw = match.group(0)

    data = json.loads(raw)

    # Sicherheitscheck: ist es wirklich ein Baustellen-Foto?
    if not data.get("ist_baustelle", True):
        raise ValueError("Kein Baustellen-Foto erkannt. Bitte nur Baustellen-Bilder hochladen.")

    return data
