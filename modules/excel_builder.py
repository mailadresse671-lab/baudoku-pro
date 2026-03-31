import os
import shutil
import textwrap
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment

ZEICHEN_PRO_ZEILE = 85
WOCHENTAGE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]


def _schreibe_zelle(ws, cell_ref: str, wert):
    """Schreibt in eine Zelle, auch wenn sie Teil einer verbundenen Zelle ist."""
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


def _schreibe_liste(ws, start_cell: str, items: list):
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
    datum: "DD.MM.YYYY"
    projekt_info: {"auftraggeber": ..., "baustelle": ..., "bearbeiter": ...}
    """
    shutil.copy2(template_path, output_path)
    wb = openpyxl.load_workbook(output_path, keep_vba=True)

    dt = datetime.strptime(datum, "%d.%m.%Y")
    tag_name = WOCHENTAGE[dt.weekday()]

    if tag_name in wb.sheetnames:
        ws = wb[tag_name]
    else:
        ws = wb.active

    w = lambda cell, val: _schreibe_zelle(ws, cell, val)
    wl = lambda cell, items: _schreibe_liste(ws, cell, items)

    # Kopfdaten
    w("D1", bericht_nr)
    w("F1", bericht_nr)
    w("C2", projekt_info.get("auftraggeber", ""))
    w("C3", projekt_info.get("baustelle", ""))
    w("C4", datum)
    w("C6", projekt_info.get("bearbeiter", ""))

    # Wetter
    w("B8", data.get("wetter_vormittag", ""))
    w("B9", data.get("wetter_nachmittag", ""))
    w("H8", data.get("temp_min", ""))
    w("H9", data.get("temp_max", ""))

    # Personal
    p1 = int(data.get("personal_aufsicht", 0))
    p2 = int(data.get("personal_facharbeiter", 0))
    p3 = int(data.get("personal_maschinist", 0))
    w("C11", p1)
    w("C12", p2)
    w("C13", p3)
    w("C15", p1 + p2 + p3)

    # Arbeiten - LV-Positionen + Beschreibung zusammenführen
    arbeiten = list(data.get("beschreibung_arbeiten", []))
    lv_pos = data.get("lv_positionen", [])
    if lv_pos:
        arbeiten = lv_pos + ["---"] + arbeiten if arbeiten else lv_pos
    wl("B17", arbeiten)

    # Geräte + Material + Sonstiges
    wl("B32", data.get("geraete_liste", []))
    wl("B37", data.get("material_liste", []))

    sonstiges = list(data.get("sonstiges", []))
    hinweise = data.get("ki_hinweise", [])
    if hinweise:
        sonstiges += ["KI-Hinweise:"] + [f"- {h}" for h in hinweise]
    wl("B41", sonstiges)

    wb.save(output_path)
    return output_path


def regiebericht(template_path: str, output_path: str, data: dict,
                 datum: str, bericht_nr: int, projekt_info: dict):
    """
    Füllt die Regiebericht-Vorlage.
    """
    shutil.copy2(template_path, output_path)
    wb = openpyxl.load_workbook(output_path, keep_vba=True)
    ws = wb.active

    w = lambda cell, val: _schreibe_zelle(ws, cell, val)
    wl = lambda cell, items: _schreibe_liste(ws, cell, items)

    # Kopfdaten
    w("D1", bericht_nr)
    w("C2", projekt_info.get("auftraggeber", ""))
    w("C3", projekt_info.get("baustelle", ""))
    w("C4", datum)
    w("C6", projekt_info.get("bearbeiter", ""))

    # Beschreibung
    wl("B8", data.get("beschreibung_arbeiten", []))
    w("B15", data.get("grund_regie", ""))

    # Personal
    for i, p in enumerate(data.get("personal", [])[:5]):
        row = 20 + i
        _schreibe_zelle(ws, f"B{row}", p.get("funktion", ""))
        _schreibe_zelle(ws, f"D{row}", p.get("anzahl", 0))
        _schreibe_zelle(ws, f"F{row}", p.get("stunden", 0))

    # Geräte
    for i, g in enumerate(data.get("geraete", [])[:5]):
        row = 30 + i
        _schreibe_zelle(ws, f"B{row}", g.get("bezeichnung", ""))
        _schreibe_zelle(ws, f"E{row}", g.get("einheit", "h"))
        _schreibe_zelle(ws, f"F{row}", g.get("menge", 0))

    # Material
    for i, m in enumerate(data.get("material", [])[:5]):
        row = 38 + i
        _schreibe_zelle(ws, f"B{row}", m.get("bezeichnung", ""))
        _schreibe_zelle(ws, f"E{row}", m.get("einheit", ""))
        _schreibe_zelle(ws, f"F{row}", m.get("menge", 0))

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
    return os.path.join(kw_folder, f"{typ}_{datum.replace('.', '-')}{template_ext}")
