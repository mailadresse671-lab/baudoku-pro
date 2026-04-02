import os
import shutil
import textwrap
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment

ZEICHEN_PRO_ZEILE = 90
WOCHENTAGE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]


def _w(ws, cell_ref: str, wert):
    """Schreibt in eine Zelle, respektiert verbundene Zellen."""
    try:
        target = ws[cell_ref]
        for rng in ws.merged_cells.ranges:
            if cell_ref in rng:
                target = ws.cell(row=rng.min_row, column=rng.min_col)
                break
        target.value = wert
        target.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
    except Exception:
        pass


def _wl(ws, start_cell: str, items: list):
    """Schreibt eine Liste zeilenweise ab einer Startzelle."""
    if not items:
        return
    col = "".join(c for c in start_cell if c.isalpha())
    row = int("".join(c for c in start_cell if c.isdigit()))
    from openpyxl.utils import column_index_from_string
    col_idx = column_index_from_string(col)

    for item in items:
        for line in textwrap.wrap(str(item), width=ZEICHEN_PRO_ZEILE):
            target = ws.cell(row=row, column=col_idx)
            for rng in ws.merged_cells.ranges:
                if target.coordinate in rng:
                    target = ws.cell(row=rng.min_row, column=rng.min_col)
                    break
            target.value = line
            target.alignment = Alignment(wrap_text=False, vertical="bottom", horizontal="left")
            row += 1


def bautagesbericht(template_path: str, output_path: str, data: dict,
                    datum: str, bericht_nr: int, projekt_info: dict):
    """
    Füllt die Bautagesbericht-Vorlage.
    Montag:    Schreibt Kopfdaten + Tagesdaten
    Di-Sa:     Schreibt NUR Tagesdaten (Kopfdaten sind Formeln die auf Montag zeigen)
    """
    # Nur kopieren wenn noch nicht vorhanden (Wochendatei bleibt bestehen)
    if not os.path.exists(output_path):
        shutil.copy2(template_path, output_path)

    wb = openpyxl.load_workbook(output_path, keep_vba=True)

    dt = datetime.strptime(datum, "%d.%m.%Y")
    tag_name = WOCHENTAGE[dt.weekday()]
    ist_montag = dt.weekday() == 0

    if tag_name in wb.sheetnames:
        ws = wb[tag_name]
    else:
        ws = wb.active

    # ── Kopfdaten auf JEDEM Tab schreiben (nicht nur Montag) ──
    _w(ws, "C2", projekt_info.get("auftraggeber", ""))
    _w(ws, "C3", projekt_info.get("baustelle", ""))
    _w(ws, "C4", dt)
    _w(ws, "C6", projekt_info.get("bearbeiter", ""))

    # Bericht-Nr nur auf Montag (andere Tabs haben Formel +1)
    if ist_montag:
        _w(ws, "D1", bericht_nr)

    # ── Tagesdaten (alle Wochentage) ──
    _w(ws, "B8", data.get("wetter_vormittag", ""))
    _w(ws, "B9", data.get("wetter_nachmittag", ""))
    _w(ws, "H8", data.get("temp_min", ""))
    _w(ws, "H9", data.get("temp_max", ""))

    p1 = int(data.get("personal_aufsicht", 0))
    p2 = int(data.get("personal_facharbeiter", 0))
    p3 = int(data.get("personal_maschinist", 0))
    _w(ws, "C11", p1)
    _w(ws, "C12", p2)
    _w(ws, "C13", p3)

    # Arbeiten: LV-Positionen zuerst, dann Beschreibung
    arbeiten = list(data.get("lv_positionen", []))
    beschreibung = data.get("beschreibung_arbeiten", [])
    if arbeiten and beschreibung:
        arbeiten += ["---"] + beschreibung
    elif beschreibung:
        arbeiten = beschreibung

    # Regiebericht-Hinweis anhängen
    if data.get("regiebericht_erstellt"):
        arbeiten += [f"→ Regiebericht Nr. {data['regiebericht_nr']} wurde erstellt"]

    _wl(ws, "B17", arbeiten)
    _wl(ws, "B32", data.get("geraete_liste", []))
    _wl(ws, "B37", data.get("material_liste", []))

    sonstiges = list(data.get("sonstiges", []))
    hinweise = data.get("ki_hinweise", [])
    if hinweise:
        sonstiges += ["KI-Hinweise:"] + [f"- {h}" for h in hinweise]
    _wl(ws, "B42", sonstiges)

    wb.save(output_path)
    return output_path


def regiebericht(template_path: str, output_path: str, data: dict,
                 datum: str, bericht_nr: int, projekt_info: dict):
    """
    Füllt die Regiebericht-Vorlage mit den Feldern die die KI liefern kann.
    Personalstunden und Details muss der Polier manuell nachtragen.
    """
    shutil.copy2(template_path, output_path)
    wb = openpyxl.load_workbook(output_path, keep_vba=True)
    ws = wb["Regiebericht"] if "Regiebericht" in wb.sheetnames else wb.active

    dt = datetime.strptime(datum, "%d.%m.%Y")

    # Kopfdaten
    _w(ws, "H1", bericht_nr)
    _w(ws, "H2", projekt_info.get("baustelle", ""))
    _w(ws, "H5", projekt_info.get("bearbeiter", ""))
    _w(ws, "O6", datum)

    # Wochentag
    _w(ws, "E6", WOCHENTAGE[dt.weekday()])

    # Art der Leistung / Beschreibung (A8 bis ca. A14)
    beschreibung = data.get("beschreibung_arbeiten", [])
    grund = data.get("grund_regie", "")
    if grund:
        beschreibung = [grund] + ["---"] + beschreibung if beschreibung else [grund]
    _wl(ws, "A8", beschreibung)

    # Geräte (A27+)
    geraete = data.get("geraete_liste", [])
    _wl(ws, "A27", geraete)

    # Material (P27+)
    material = data.get("material_liste", [])
    _wl(ws, "P27", material)

    # Bemerkungen (A38+)
    bemerkungen = list(data.get("sonstiges", []))
    bemerkungen += ["→ Stunden und Mengen bitte manuell eintragen"]
    _wl(ws, "A38", bemerkungen)

    wb.save(output_path)
    return output_path


def get_output_path(projekt_pfad: str, datum: str, report_type: str, template_ext: str) -> str:
    """Gibt den Ausgabepfad für einen Bericht zurück (nach KW geordnet)."""
    dt = datetime.strptime(datum, "%d.%m.%Y")
    kw = dt.isocalendar()[1]
    year = dt.year
    typ = "Bautagesbericht" if report_type == "bautagesbericht" else "Regiebericht"
    kw_folder = os.path.join(projekt_pfad, "1.4 Berichte", "1.4.1 Tagesberichte",
                             "Fertig", f"KW{kw}_{year}")
    os.makedirs(kw_folder, exist_ok=True)

    if report_type == "bautagesbericht":
        # Wochendatei (eine Datei pro KW, alle Tage drin)
        return os.path.join(kw_folder, f"Bautagesbericht_KW{kw}_{year}{template_ext}")
    else:
        return os.path.join(kw_folder, f"Regiebericht_{datum.replace('.', '-')}{template_ext}")
