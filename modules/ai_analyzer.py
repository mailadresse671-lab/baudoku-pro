import base64
import io
import json
import re
from PIL import Image
from groq import Groq
import config


def _encode_image(path: str) -> tuple[str, str]:
    """Bild auf max 800px verkleinern, 70% Qualität."""
    img = Image.open(path)
    img.thumbnail((800, 800), Image.LANCZOS)
    buffer = io.BytesIO()
    img.convert("RGB").save(buffer, format="JPEG", quality=70)
    buffer.seek(0)
    return base64.standard_b64encode(buffer.read()).decode("utf-8"), "image/jpeg"


BAUTAGESBERICHT_FELDER = """{
  "ist_baustelle": true,
  "wetter_vormittag": "Sonnig / Bewölkt / Regen / Schnee",
  "wetter_nachmittag": "Sonnig / Bewölkt / Regen / Schnee",
  "temp_min": 0,
  "temp_max": 0,
  "personal_aufsicht": 1,
  "personal_facharbeiter": 0,
  "personal_maschinist": 0,
  "beschreibung_arbeiten": [
    "Vollständiger Satz was gemacht wurde, z.B.:",
    "Herstellung Perimeterschutz / Untergrabschutz: Betonieren der Fundamente mittels Schalungsrohren DN 110.",
    "Aufstellen provisorischer Bauzaun zur Vorbereitung der Demontage des Bestandszauns."
  ],
  "lv_positionen": [],
  "geraete_liste": ["Bagger", "Betonmischer-LKW"],
  "material_liste": ["Beton C25/30", "Schalungsrohre DN 110"],
  "sonstiges": [],
  "ki_hinweise": []
}"""

REGIEBERICHT_FELDER = """{
  "ist_baustelle": true,
  "beschreibung_arbeiten": [
    "Vollständiger Satz was gemacht wurde"
  ],
  "grund_regie": "Warum Regiearbeit: nicht im LV enthalten / Änderung durch AG / unvorhergesehene Arbeit",
  "personal": [
    {"funktion": "Polier", "anzahl": 1, "stunden": 0},
    {"funktion": "Facharbeiter", "anzahl": 0, "stunden": 0},
    {"funktion": "Maschinist", "anzahl": 0, "stunden": 0}
  ],
  "geraete": [
    {"bezeichnung": "Gerät", "einheit": "h", "menge": 0}
  ],
  "material": [
    {"bezeichnung": "Material", "einheit": "m²", "menge": 0}
  ],
  "ki_hinweise": []
}"""


def _build_prompt(stichworte: str, lv_text: str, report_type: str, hat_bilder: bool) -> str:
    felder = BAUTAGESBERICHT_FELDER if report_type == "bautagesbericht" else REGIEBERICHT_FELDER

    lv_kontext = (
        f"\nLEISTUNGSVERZEICHNIS (als Referenz):\n{lv_text[:5000]}"
        if lv_text else ""
    )

    foto_hinweis = (
        "Die beigefügten Fotos dienen als Dokumentation und zur Ergänzung (Wetter, Geräte)."
        if hat_bilder else
        "Keine Fotos vorhanden."
    )

    if report_type == "bautagesbericht":
        aufgabe = f"""Der Polier hat heute folgende Arbeiten durchgeführt:

--- HEUTIGE ARBEITEN ---
{stichworte}
--- ENDE ---

Schreibe daraus einen professionellen Bautagesbericht im Straßen- und Tiefbau.

ANFORDERUNGEN an die Beschreibung:
- Vollständige Sätze in Fachsprache (Tiefbau / Straßenbau)
- Jeden Arbeitsschritt einzeln und präzise beschreiben
- Fachbegriffe verwenden (z.B. "Herstellung", "Einbau", "Verlegung", "Demontage")
- KEINE Mengen erfinden (Meter, m², Tonnen → auf 0 lassen)
- Stunden auf 0 lassen (wird manuell eingetragen)
{foto_hinweis}"""
    else:
        aufgabe = f"""Der Polier hat heute folgende Regiearbeiten durchgeführt:

--- HEUTIGE ARBEITEN ---
{stichworte}
--- ENDE ---

Schreibe daraus einen professionellen Regiebericht im Straßen- und Tiefbau.
Begründe warum diese Arbeiten als Regie abgerechnet werden.
Stunden und Mengen auf 0 setzen.
{foto_hinweis}"""

    return f"""Du bist erfahrener Polier im Straßen- und Tiefbau.

{aufgabe}
{lv_kontext}

Antworte NUR mit einem JSON-Objekt, kein Text davor oder danach:
{felder}"""


def analyze(image_paths: list[str], lv_text: str,
            report_type: str = "bautagesbericht",
            stichworte: str = "") -> dict:

    if not config.GROQ_API_KEY:
        raise ValueError("GROQ_API_KEY fehlt in der .env Datei")

    if not stichworte:
        raise ValueError("Bitte beschreibe kurz was heute gemacht wurde.")

    client = Groq(api_key=config.GROQ_API_KEY)

    # Bilder vorbereiten (max 2, für Wetter + Geräte-Erkennung)
    content = []
    for path in image_paths[:2]:
        try:
            data, media_type = _encode_image(path)
            content.append({
                "type": "image_url",
                "image_url": {"url": f"data:{media_type};base64,{data}"}
            })
        except Exception as e:
            print(f"Bild übersprungen {path}: {e}")

    hat_bilder = len(content) > 0
    content.append({
        "type": "text",
        "text": _build_prompt(stichworte, lv_text, report_type, hat_bilder)
    })

    models = [
        "meta-llama/llama-4-maverick-17b-128e-instruct",
        "meta-llama/llama-4-scout-17b-16e-instruct",
    ]

    last_error = None
    response = None
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
            if "413" in err_str or "404" in err_str or "model" in err_str.lower():
                print(f"Modell {model} fehlgeschlagen: {e}")
                continue
            raise

    if response is None:
        raise ValueError(f"API-Fehler: {last_error}")

    raw = response.choices[0].message.content.strip()
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if match:
        raw = match.group(0)

    data = json.loads(raw)

    if not data.get("ist_baustelle", True):
        raise ValueError("Keine Baustellendaten erkannt.")

    return data
