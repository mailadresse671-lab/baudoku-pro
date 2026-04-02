import base64
import io
import json
import re
from PIL import Image
from groq import Groq
import config


def _encode_image(path: str) -> tuple[str, str]:
    """Bild auf max 800px verkleinern, 70% Qualität – gut für KI-Erkennung, klein genug für API."""
    img = Image.open(path)
    img.thumbnail((800, 800), Image.LANCZOS)
    buffer = io.BytesIO()
    img.convert("RGB").save(buffer, format="JPEG", quality=70)
    buffer.seek(0)
    return base64.standard_b64encode(buffer.read()).decode("utf-8"), "image/jpeg"


BAUTAGESBERICHT_FELDER = """
{
  "ist_baustelle": true,
  "wetter_vormittag": "Sonnig / Bewölkt / Regen / Schnee",
  "wetter_nachmittag": "Sonnig / Bewölkt / Regen / Schnee",
  "temp_min": 0,
  "temp_max": 0,
  "personal_aufsicht": 1,
  "personal_facharbeiter": 0,
  "personal_maschinist": 0,
  "beschreibung_arbeiten": [
    "Was genau zu sehen ist, z.B. Erdaushub, Rohrverlegung, Asphalteinbau, Betonarbeiten..."
  ],
  "lv_positionen": ["LV-Positionsnummer und Kurztext aus dem Leistungsverzeichnis"],
  "geraete_liste": ["Bagger, Radlader, Rüttelplatte, LKW – was erkennbar ist"],
  "material_liste": ["Rohre, Kies, Beton, Asphalt, Stahl – was erkennbar ist"],
  "sonstiges": [],
  "ki_hinweise": ["Was nicht sicher erkennbar war"]
}
"""

REGIEBERICHT_FELDER = """
{
  "ist_baustelle": true,
  "beschreibung_arbeiten": ["Was genau zu sehen ist"],
  "grund_regie": "Warum Regiearbeit: z.B. nicht im LV enthalten, Zusatzarbeit, Änderung",
  "personal": [
    {"funktion": "Polier", "anzahl": 1, "stunden": 0},
    {"funktion": "Facharbeiter", "anzahl": 0, "stunden": 0},
    {"funktion": "Maschinist", "anzahl": 0, "stunden": 0}
  ],
  "geraete": [
    {"bezeichnung": "Gerät was erkennbar ist", "einheit": "h", "menge": 0}
  ],
  "material": [
    {"bezeichnung": "Material was erkennbar ist", "einheit": "m²", "menge": 0}
  ],
  "ki_hinweise": ["Was nicht sicher erkennbar war"]
}
"""


def _build_prompt(lv_text: str, report_type: str) -> str:
    felder = BAUTAGESBERICHT_FELDER if report_type == "bautagesbericht" else REGIEBERICHT_FELDER

    lv_kontext = (
        f"LEISTUNGSVERZEICHNIS (für LV-Positionen):\n{lv_text[:6000]}"
        if lv_text else ""
    )

    if report_type == "bautagesbericht":
        aufgabe = (
            "Analysiere diese Baustellen-Fotos aus dem Straßen- und Tiefbau.\n"
            "Beschreibe GENAU was du siehst:\n"
            "- Welche Bauarbeiten werden durchgeführt? (Erdaushub, Rohrverlegung, "
            "Betonarbeiten, Asphalt, Kabelmontage, Leitplanken, Zaunbau, etc.)\n"
            "- Welche Maschinen und Geräte sind erkennbar?\n"
            "- Welche Materialien liegen vor oder werden verbaut?\n"
            "- Wie viele Personen sind sichtbar?\n"
            "- Wie ist das Wetter auf dem Foto?\n"
            "- Welche LV-Position passt am besten dazu?"
        )
    else:
        aufgabe = (
            "Analysiere diese Baustellen-Fotos für einen Regiebericht.\n"
            "Beschreibe GENAU was du siehst und erkläre warum diese Arbeit "
            "als Regiearbeit abgerechnet wird (nicht im Leistungsverzeichnis).\n"
            "Stunden und Mengen auf 0 setzen – werden manuell eingetragen."
        )

    return f"""Du bist erfahrener Polier im Straßen- und Tiefbau.

{aufgabe}

WICHTIGE REGELN:
- Beschreibe NUR was wirklich auf den Fotos zu sehen ist
- Keine Mengen erfinden (Meter, m², Tonnen, Stunden)
- Fachbegriffe aus dem Tiefbau verwenden
- Wenn etwas unklar ist → in "ki_hinweise" eintragen
- Antwort ausschließlich als JSON-Objekt

{lv_kontext}

JSON-Struktur (alle Felder ausfüllen):
{felder}"""


def analyze(image_paths: list[str], lv_text: str, report_type: str = "bautagesbericht") -> dict:
    if not config.GROQ_API_KEY:
        raise ValueError("GROQ_API_KEY fehlt in der .env Datei")

    client = Groq(api_key=config.GROQ_API_KEY)

    content = []
    for path in image_paths[:2]:
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

    # llama-4-maverick ist stärker bei Bilderkennung als scout
    models = [
        "meta-llama/llama-4-maverick-17b-128e-instruct",
        "meta-llama/llama-4-scout-17b-16e-instruct",
    ]

    last_error = None
    for model in models:
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": content}],
                response_format={"type": "json_object"},
                max_tokens=2000
            )
            break
        except Exception as e:
            last_error = e
            err_str = str(e)
            # Bei 413 (zu groß) oder 404 (Modell nicht gefunden) nächstes probieren
            if "413" in err_str or "404" in err_str or "model" in err_str.lower():
                print(f"Modell {model} fehlgeschlagen: {e} – nächstes probieren...")
                continue
            raise
    else:
        raise ValueError(f"Alle Modelle fehlgeschlagen: {last_error}")

    raw = response.choices[0].message.content.strip()

    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if match:
        raw = match.group(0)

    data = json.loads(raw)

    if not data.get("ist_baustelle", True):
        raise ValueError("Kein Baustellen-Foto erkannt.")

    return data
