# -*- coding: utf-8 -*-
"""
VÖ-Tabellen – GUI (Tabelle 1/2/3/5)
- Ordnerauswahl für Monat/Quartal/Halbjahr/Jahr (optional)
- Ausgabe nach <Basis>\VÖ-Tabellen\<Eingangsordnername>
- Protokoll mit Datum/Uhrzeit nach Protokollordner (optional, sonst .\Protokolle)
- Layouts standardmäßig aus .\Layouts (optional wählbar)

Fix v012:
- Tabelle 1 / JJ / _g: Spalten J und K müssen aus der Eingangsdatei übernommen werden.
  => JJ/_g wird jetzt aus der Eingangsdatei erzeugt (nicht aus Layout), inkl. Markierung in Spalte G.
"""

import os
import re
import sys
import traceback
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# ============================================================
# Konfiguration / Konventionen
# ============================================================

APP_TITLE = "VÖ-Tabellen – GUI (Tabelle 1/2/3/5)"
DEFAULT_LAYOUT_DIR = os.path.join(".", "Layouts")
DEFAULT_PROTOCOL_DIR = os.path.join(".", "Protokolle")

RAW_SHEET_NAMES = {
    1: "Tabelle 1",
    2: "Tabelle 2",
    3: "Tabelle 3",
    5: "Tabelle 5",
}

TEMPLATES = {
    1: {"ext": "Tabelle-1-Layout_g.xlsx", "int": "Tabelle-1-Layout_INTERN.xlsx"},
    2: {"ext": "Tabelle-2-Layout_g.xlsx", "int": "Tabelle-2-Layout_INTERN.xlsx"},
    3: {"ext": "Tabelle-3-Layout_g.xlsx", "int": "Tabelle-3-Layout_INTERN.xlsx"},
    5: {"ext": "Tabelle-5-Layout_g.xlsx", "int": "Tabelle-5-Layout_INTERN.xlsx"},
}

RELEVANT_PREFIXES = [
    "Tabelle-1-Land",
    "Tabelle-2-Land",
    "Tabelle-3-Land",
    "Tabelle-5-Land",
    # später: Tab8/Tab9
]

# ============================================================
# Logging
# ============================================================

class Logger:
    def __init__(self, base_dir: str):
        os.makedirs(base_dir, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        self.path = os.path.join(base_dir, f"vo_tabellen_{ts}.log")
        with open(self.path, "w", encoding="utf-8") as f:
            f.write(f"=== Protokollstart: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n")

    def log(self, msg: str):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {msg}"
        with open(self.path, "a", encoding="utf-8") as f:
            f.write(line + "\n")

# ============================================================
# Hilfsfunktionen (Pfad / Periodenlogik / Merges)
# ============================================================

def normalize_path(p: str) -> str:
    if not p:
        return ""
    p = p.strip().strip('"')
    return os.path.normpath(p)

def parse_period_from_filename(filename: str) -> str:
    base = os.path.splitext(os.path.basename(filename))[0]
    m = re.search(r"_(\d{4}-(?:\d{2}|Q\d|H\d|JJ))$", base)
    if m:
        return m.group(1)
    m = re.search(r"_(\d{4}-(?:\d{2}|Q\d|H\d|JJ))_", base)
    if m:
        return m.group(1)
    # Fallback: irgendwas mit Jahr/Monat/Quartal/Halbjahr
    m = re.search(r"(20\d{2}-(?:\d{2}|Q\d|H\d|JJ))", base)
    return m.group(1) if m else ""

def period_is_jj(period: str) -> bool:
    return str(period).upper().endswith("JJ") or re.fullmatch(r"\d{4}", str(period).strip()) is not None

def year_from_period(period: str) -> str:
    # "2025-JJ" -> "2025", "2025" -> "2025"
    s = str(period).strip()
    if re.fullmatch(r"\d{4}", s):
        return s
    if s.endswith("-JJ"):
        return s.split("-")[0]
    if "-" in s:
        return s.split("-")[0]
    return s

def set_value_merge_safe(ws, row, col, value):
    """
    Setzt den Wert auch bei Merges: bei Merge-Bereich wird in die Top-Left-Zelle geschrieben.
    """
    cell = ws.cell(row=row, column=col)
    # Wenn Teil eines Merge-Bereichs: finde top-left
    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            tl = ws.cell(row=mr.min_row, column=mr.min_col)
            tl.value = value
            return
    cell.value = value

def copy_cell_value(src_ws, dst_ws, r, c):
    dst_ws.cell(row=r, column=c).value = src_ws.cell(row=r, column=c).value

def _extract_period_text(ws, row_guess: int) -> str:
    v = ws.cell(row=row_guess, column=1).value
    if v is None:
        # suche in den ersten 10 Zeilen die erste nichtleere Zelle in Spalte A
        for rr in range(1, 11):
            vv = ws.cell(row=rr, column=1).value
            if vv is not None and str(vv).strip():
                return str(vv).strip()
        return ""
    return str(v).strip()

def find_period_text(ws_raw) -> str:
    """
    Zeitbezug aus Eingangsdatei:
    - Monat/Q/H: steht in A3 (oder A4, etc.)
    - JJ: steht meist als " 2025" (nur Jahr). Dann daraus "Jahr 2025" machen.
    """
    t = _extract_period_text(ws_raw, 3)
    if re.fullmatch(r"\d{4}", str(t).strip()):
        return f"Jahr {str(t).strip()}"
    return str(t).strip()

# ============================================================
# Markierungen / Formatregeln
# ============================================================

def mark_cells_with_1_or_2(ws, col_index_g: int, fill: PatternFill):
    """
    Markiert Zeilen, bei denen in Spalte G (Index col_index_g) der Wert 1 oder 2 steht.
    Markierung: ganzer Zeilenbereich (A bis letzte Spalte) oder nur relevante Spalten? -> hier ganze Zeile.
    """
    max_row = ws.max_row
    max_col = ws.max_column
    for r in range(1, max_row + 1):
        v = ws.cell(row=r, column=col_index_g).value
        if v in (1, 2, "1", "2"):
            for c in range(1, max_col + 1):
                ws.cell(row=r, column=c).fill = fill

# ============================================================
# Tabelle 1
# ============================================================

def build_table1_workbook(raw_path, template_path, internal_layout: bool):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[1]] if RAW_SHEET_NAMES[1] in wb_raw.sheetnames else wb_raw.active

    # Zeitbezug ermitteln
    period_text = find_period_text(ws_raw)

    # JJ (extern): Eingangsdatei weitgehend übernehmen (inkl. Spalten J/K) -> wird in process_table1_file erledigt.
    # Hier bleibt build_table1_workbook als "Layout-Builder" bestehen.
    wb_out = openpyxl.load_workbook(template_path)
    ws_out = wb_out.active

    # Kopfzeilen: Bayern + Titel + Zeitbezug
    set_value_merge_safe(ws_out, 1, 1, ws_raw.cell(row=1, column=1).value)
    set_value_merge_safe(ws_out, 2, 1, ws_raw.cell(row=2, column=1).value)

    # INTERN hat Zeitbezug typischerweise ab Zeile 5, extern ab Zeile 3/5 je nach Layout.
    # Wir schreiben in die "klassische" A3/A5-Position; set_value_merge_safe kümmert sich um Merge-TopLeft.
    if internal_layout:
        set_value_merge_safe(ws_out, 5, 1, period_text)
    else:
        set_value_merge_safe(ws_out, 3, 1, period_text)

    # Datenbereich (Layout ist Grundlage):
    # Wir kopieren die "Tabelle" (Zahlen/Labels) aus RAW in OUT anhand des verwendeten Layouts.
    # Für Tabelle 1 gehen wir davon aus, dass Roh- und Layoutdatenbereich deckungsgleich sind (ab Zeile 6).
    # Kopiere alle Zellen, die im Layout existieren.
    max_row = min(ws_raw.max_row, ws_out.max_row)
    max_col = min(ws_raw.max_column, ws_out.max_column)

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            # Kopfbereich ist bereits gesetzt; wir kopieren ab Zeile 6
            if r >= 6:
                ws_out.cell(row=r, column=c).value = ws_raw.cell(row=r, column=c).value

    return wb_out

def process_table1_file(raw_path, output_dir, logger: Logger):
    base = os.path.splitext(os.path.basename(raw_path))[0]
    is_jj = "-JJ" in base

    layout_g = os.path.join(LAYOUT_DIR, TEMPLATES[1]["ext"])
    layout_i = os.path.join(LAYOUT_DIR, TEMPLATES[1]["int"])

    wb_i = build_table1_workbook(raw_path, layout_i, internal_layout=True)
    out_i = os.path.join(output_dir, base + "_INTERN.xlsx")
    wb_i.save(out_i)
    logger.log(f"[T1] INTERN -> {out_i}")

    if is_jj:
        # JJ (_g): soll der Eingangsdatei entsprechen (inkl. Spalten J/K) + Markierung.
        wb_g = openpyxl.load_workbook(raw_path)  # Styles/Spalten komplett behalten
        ws = wb_g[RAW_SHEET_NAMES[1]] if RAW_SHEET_NAMES[1] in wb_g.sheetnames else wb_g.active

        # Zeitbezug aus Eingangsdatei übernehmen, aber bei JJ als 'Jahr 2025' ausgeben
        wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
        ws_raw = wb_raw[RAW_SHEET_NAMES[1]] if RAW_SHEET_NAMES[1] in wb_raw.sheetnames else wb_raw.active
        period_text = find_period_text(ws_raw)  # liefert bei JJ bereits 'Jahr 2025'
        if period_text:
            set_value_merge_safe(ws, 3, 1, period_text)

        fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        mark_cells_with_1_or_2(ws, 7, fill)  # Tabelle 1: Spalte G markieren

        out_g = os.path.join(output_dir, base + "_g.xlsx")
        wb_g.save(out_g)
        logger.log(f"[T1] _g (JJ: Eingang + Markierung) -> {out_g}")
    else:
        wb_g = build_table1_workbook(raw_path, layout_g, internal_layout=False)
        out_g = os.path.join(output_dir, base + "_g.xlsx")
        wb_g.save(out_g)
        logger.log(f"[T1] _g -> {out_g}")

# ============================================================
# Tabelle 2
# ============================================================

def build_table2_workbook(raw_path, template_path, internal_layout: bool):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[2]] if RAW_SHEET_NAMES[2] in wb_raw.sheetnames else wb_raw.active

    period_text = find_period_text(ws_raw)

    wb_out = openpyxl.load_workbook(template_path)
    ws_out = wb_out.active

    # Kopf (T2: Titel ggf. mehrzeilig)
    set_value_merge_safe(ws_out, 1, 1, ws_raw.cell(row=1, column=1).value)
    # Titel (Zeile 2) + ggf. Fortsetzung (Zeile 3) in _g soll wie INTERN aussehen
    # Wir übernehmen Zeile 2 und 3 aus RAW (RAW hat die richtige zweizeilige Überschrift im Monat/Quartal/Halbjahr)
    set_value_merge_safe(ws_out, 2, 1, ws_raw.cell(row=2, column=1).value)
    # Falls RAW in A3 einen Titelteil hat (z.B. "… und Zahl der Arbeitnehmer/-innen"), dann übernehmen
    raw3 = ws_raw.cell(row=3, column=1).value
    if raw3 and str(raw3).strip() and not re.fullmatch(r"\d{4}.*", str(raw3).strip()):
        # Das ist vermutlich der Fortsetzungstitel
        set_value_merge_safe(ws_out, 3, 1, raw3)

    # Zeitbezug: bei RAW ist der Zeitbezug in der Regel eine Zeile unter dem Titelblock
    # Layout unterscheidet sich extern/intern – wir schreiben in die im Layout vorhandene Zeile:
    # INTERN: meist Zeile 4; _g: nach Titelblock eine Zeile darunter (oft Zeile 4/5)
    if internal_layout:
        # Zeitbezug üblicherweise in Zeile 4
        set_value_merge_safe(ws_out, 4, 1, period_text)
    else:
        # Zeitbezug unter dem Titelblock. Falls in Zeile 3 Fortsetzung steht, dann Zeitbezug in Zeile 4, sonst 3.
        if raw3 and str(raw3).strip() and not re.fullmatch(r"\d{4}.*", str(raw3).strip()):
            set_value_merge_safe(ws_out, 4, 1, period_text)
        else:
            set_value_merge_safe(ws_out, 3, 1, period_text)

    # Datenbereich kopieren (wie bei T1)
    max_row = min(ws_raw.max_row, ws_out.max_row)
    max_col = min(ws_raw.max_column, ws_out.max_column)
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            if r >= 6:  # Tabellenkörper
                ws_out.cell(row=r, column=c).value = ws_raw.cell(row=r, column=c).value

    return wb_out

def process_table2_file(raw_path, output_dir, logger: Logger):
    base = os.path.splitext(os.path.basename(raw_path))[0]
    layout_g = os.path.join(LAYOUT_DIR, TEMPLATES[2]["ext"])
    layout_i = os.path.join(LAYOUT_DIR, TEMPLATES[2]["int"])

    wb_g = build_table2_workbook(raw_path, layout_g, internal_layout=False)
    out_g = os.path.join(output_dir, base + "_g.xlsx")
    wb_g.save(out_g)
    logger.log(f"[T2] _g -> {out_g}")

    wb_i = build_table2_workbook(raw_path, layout_i, internal_layout=True)
    out_i = os.path.join(output_dir, base + "_INTERN.xlsx")
    wb_i.save(out_i)
    logger.log(f"[T2] INTERN -> {out_i}")

# ============================================================
# Tabelle 3
# ============================================================

def build_table3_workbook(raw_path, template_path, internal_layout: bool):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[3]] if RAW_SHEET_NAMES[3] in wb_raw.sheetnames else wb_raw.active

    period_text = find_period_text(ws_raw)

    wb_out = openpyxl.load_workbook(template_path)
    ws_out = wb_out.active

    # Kopf
    set_value_merge_safe(ws_out, 1, 1, ws_raw.cell(row=1, column=1).value)
    set_value_merge_safe(ws_out, 2, 1, ws_raw.cell(row=2, column=1).value)

    # Zeitbezug:
    if internal_layout:
        set_value_merge_safe(ws_out, 5, 1, period_text)
    else:
        # extern: Zeile 3 oder 5 je nach Layout – wir nutzen A3 (wie Muster)
        set_value_merge_safe(ws_out, 3, 1, period_text)

    # Datenbereich kopieren
    max_row = min(ws_raw.max_row, ws_out.max_row)
    max_col = min(ws_raw.max_column, ws_out.max_column)
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            if r >= 6:
                ws_out.cell(row=r, column=c).value = ws_raw.cell(row=r, column=c).value

    return wb_out

def process_table3_file(raw_path, output_dir, logger: Logger):
    base = os.path.splitext(os.path.basename(raw_path))[0]
    layout_g = os.path.join(LAYOUT_DIR, TEMPLATES[3]["ext"])
    layout_i = os.path.join(LAYOUT_DIR, TEMPLATES[3]["int"])

    wb_g = build_table3_workbook(raw_path, layout_g, internal_layout=False)
    out_g = os.path.join(output_dir, base + "_g.xlsx")
    wb_g.save(out_g)
    logger.log(f"[T3] _g -> {out_g}")

    wb_i = build_table3_workbook(raw_path, layout_i, internal_layout=True)
    out_i = os.path.join(output_dir, base + "_INTERN.xlsx")
    wb_i.save(out_i)
    logger.log(f"[T3] INTERN -> {out_i}")

# ============================================================
# Tabelle 5
# ============================================================

def find_last_nonempty_col_in_row(ws, row, max_search=50):
    last = None
    for c in range(1, max_search + 1):
        v = ws.cell(row=row, column=c).value
        if v is not None and str(v).strip() != "":
            last = c
    return last

def move_stand_to_last_col(ws, stand_row: int, from_col: int):
    """
    Verschiebt "Stand: xxxx" (oder ähnliches) von from_col in die letzte belegte Spalte der Zeile (unter Copyright).
    """
    v = ws.cell(row=stand_row, column=from_col).value
    if not v or "Stand" not in str(v):
        return
    last = find_last_nonempty_col_in_row(ws, stand_row, max_search=80)
    if not last:
        return
    # last ist unter Umständen die Copyright-Spalte oder die alte Stand-Spalte. Ziel: letzte belegte Spalte (ohne Stand)
    # Wir nehmen last, aber wenn last==from_col, nehmen wir die letzte belegte Spalte links davon.
    if last == from_col:
        for c in range(from_col - 1, 0, -1):
            vv = ws.cell(row=stand_row, column=c).value
            if vv is not None and str(vv).strip() != "":
                last = c
                break
    # Ziel ist "unter die letzte Tabellenspalte" -> das ist last (Copyright steht meist dort)
    ws.cell(row=stand_row, column=from_col).value = None
    ws.cell(row=stand_row, column=last).value = v

def build_table5_workbook(raw_path, template_path, internal_layout: bool):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)

    wb_out = openpyxl.load_workbook(template_path)

    # Tabelle 5 hat mehrere Blätter – wir kopieren pro Blatt die Wertebereiche in das entsprechende Layoutblatt
    # Annahme: Layout hat gleiche Blattanzahl/Blattnamen wie Rohdatei.
    for ws_raw in wb_raw.worksheets:
        name = ws_raw.title
        if name in wb_out.sheetnames:
            ws_out = wb_out[name]
        else:
            # Falls Layout andere Namen hat, fallback: nach Index
            ws_out = wb_out.worksheets[wb_raw.worksheets.index(ws_raw)]

        # Zeitbezug aus Eingangsdatei: bei Tabelle 5 steht es üblicherweise in A3 (JJ ggf. nur Jahr)
        period_text = find_period_text(ws_raw)

        # Kopf in A1/A2 und Zeitbezug in A3/A4 je nach Layout
        set_value_merge_safe(ws_out, 1, 1, ws_raw.cell(row=1, column=1).value)
        set_value_merge_safe(ws_out, 2, 1, ws_raw.cell(row=2, column=1).value)

        # Zeitbezug-Position: wir suchen in den ersten 10 Zeilen im Layout die erste leere Zeile nach Titelblock
        # oder setzen in A3 (extern) / A5 (intern) wie bei anderen Tabellen.
        if internal_layout:
            set_value_merge_safe(ws_out, 5, 1, period_text)
        else:
            set_value_merge_safe(ws_out, 3, 1, period_text)

        # Datenbereich kopieren (ab Zeile 6)
        max_row = min(ws_raw.max_row, ws_out.max_row)
        max_col = min(ws_raw.max_column, ws_out.max_column)
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                if r >= 6:
                    ws_out.cell(row=r, column=c).value = ws_raw.cell(row=r, column=c).value

        # Spezieller Fix: Blatt 5.5 (XML-Tab5-Land) "Stand:" Position unter letzte Spalte
        # In euren Beispielen: Stand ist in J143, soll in H143. Das ist: von col 10 nach col 8.
        if "5.5" in str(name) or "XML-Tab5-Land" in str(name):
            move_stand_to_last_col(ws_out, stand_row=143, from_col=10)

    return wb_out

def process_table5_file(raw_path, output_dir, logger: Logger):
    base = os.path.splitext(os.path.basename(raw_path))[0]
    layout_g = os.path.join(LAYOUT_DIR, TEMPLATES[5]["ext"])
    layout_i = os.path.join(LAYOUT_DIR, TEMPLATES[5]["int"])

    wb_g = build_table5_workbook(raw_path, layout_g, internal_layout=False)
    out_g = os.path.join(output_dir, base + "_g.xlsx")
    wb_g.save(out_g)
    logger.log(f"[T5] _g -> {out_g}")

    wb_i = build_table5_workbook(raw_path, layout_i, internal_layout=True)
    out_i = os.path.join(output_dir, base + "_INTERN.xlsx")
    wb_i.save(out_i)
    logger.log(f"[T5] INTERN -> {out_i}")

# ============================================================
# Scannen / Verarbeitung pro Eingangsordner
# ============================================================

def is_relevant_file(filename: str) -> bool:
    if not filename.lower().endswith(".xlsx"):
        return False
    base = os.path.basename(filename)
    # Keine bereits erzeugten Ausgaben erneut verarbeiten
    if base.endswith("_g.xlsx") or base.endswith("_INTERN.xlsx"):
        return False
    for p in RELEVANT_PREFIXES:
        if base.startswith(p):
            return True
    return False

def process_input_folder(input_dir: str, output_base: str, logger: Logger):
    input_dir = normalize_path(input_dir)
    output_base = normalize_path(output_base)

    if not input_dir or not os.path.isdir(input_dir):
        logger.log(f"[SKIP] Kein gültiger Eingangspfad: {input_dir}")
        return

    folder_name = os.path.basename(input_dir)
    out_dir = os.path.join(output_base, "VÖ-Tabellen", folder_name)
    os.makedirs(out_dir, exist_ok=True)

    logger.log(f"--- Eingang: {input_dir}")
    logger.log(f"--- Ausgabe: {out_dir}")

    files = sorted([os.path.join(input_dir, f) for f in os.listdir(input_dir) if is_relevant_file(f)])
    logger.log(f"[SCAN] {len(files)} Dateien gefunden (relevant).")

    for f in files:
        base = os.path.basename(f)
        try:
            logger.log(f"[START] {base}")
            if base.startswith("Tabelle-1-Land"):
                process_table1_file(f, out_dir, logger)
            elif base.startswith("Tabelle-2-Land"):
                process_table2_file(f, out_dir, logger)
            elif base.startswith("Tabelle-3-Land"):
                process_table3_file(f, out_dir, logger)
            elif base.startswith("Tabelle-5-Land"):
                process_table5_file(f, out_dir, logger)
            else:
                logger.log(f"[SKIP] Nicht implementiert: {base}")
        except Exception as e:
            logger.log(f"[FEHLER] {repr(e)}")
            logger.log(traceback.format_exc())
            raise

# ============================================================
# GUI
# ============================================================

def browse_dir(var: tk.StringVar):
    d = filedialog.askdirectory()
    if d:
        var.set(d)

def _run(month_dir, quarter_dir, half_dir, year_dir, out_base, protocol_dir, layout_dir, overwrite, root):
    global LAYOUT_DIR
    LAYOUT_DIR = normalize_path(layout_dir) if layout_dir else DEFAULT_LAYOUT_DIR

    out_base = normalize_path(out_base)
    protocol_dir = normalize_path(protocol_dir) if protocol_dir else DEFAULT_PROTOCOL_DIR

    logger = Logger(protocol_dir)
    logger.log("=== START GUI-Lauf ===")
    logger.log(f"Ausgabe-Basis: {out_base}")
    logger.log(f"Layouts:       {LAYOUT_DIR}")
    logger.log(f"Überschreiben: {'JA' if overwrite else 'NEIN'}")
    logger.log(f"Jobs: Monat={month_dir}, Quartal={quarter_dir}, Halbjahr={half_dir}, Jahr={year_dir}")

    # overwrite: wir löschen Zielordner nicht, aber Ausgaben werden beim Speichern überschrieben
    try:
        for label, idir in [("Monat", month_dir), ("Quartal", quarter_dir), ("Halbjahr", half_dir), ("Jahr", year_dir)]:
            idir = normalize_path(idir)
            if idir:
                logger.log(f"== JOB: {label} ==")
                process_input_folder(idir, out_base, logger)

        logger.log("[FERTIG] Verarbeitung abgeschlossen.")
        messagebox.showinfo("Fertig", f"Verarbeitung abgeschlossen.\n\nProtokoll:\n{logger.path}")
    except Exception as e:
        messagebox.showerror("Fehler", f"Es ist ein Fehler aufgetreten:\n{e}\n\nBitte Protokoll prüfen.\n\nProtokoll:\n{logger.path}")
        raise

def start_gui():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("880x360")

    month_var = tk.StringVar()
    quarter_var = tk.StringVar()
    half_var = tk.StringVar()
    year_var = tk.StringVar()
    out_var = tk.StringVar()
    protocol_var = tk.StringVar()
    layout_var = tk.StringVar()

    overwrite_var = tk.BooleanVar(value=True)

    # Default: Layouts relativ
    layout_var.set(DEFAULT_LAYOUT_DIR)

    row = 0
    def add_row(label, var, browse=True):
        nonlocal row
        tk.Label(root, text=label, anchor="w").grid(row=row, column=0, sticky="w", padx=8, pady=4)
        tk.Entry(root, textvariable=var, width=90).grid(row=row, column=1, sticky="w", padx=8, pady=4)
        if browse:
            tk.Button(root, text="Auswählen…", command=lambda: browse_dir(var)).grid(row=row, column=2, padx=8, pady=4)
        row += 1

    add_row("Monat – Eingangstabellen (optional):", month_var)
    add_row("Quartal – Eingangstabellen (optional):", quarter_var)
    add_row("Halbjahr – Eingangstabellen (optional):", half_var)
    add_row("Jahr – Eingangstabellen (optional):", year_var)
    add_row("Ausgabe-Basisordner (Pflicht):", out_var)
    add_row("Protokollordner (optional, leer = .\\Protokolle):", protocol_var)
    add_row("Layouts-Ordner (optional, leer = .\\Layouts):", layout_var)

    row += 1
    tk.Checkbutton(root, text="Vorhandene Ausgabedateien überschreiben", variable=overwrite_var).grid(row=row, column=1, sticky="w", padx=8)
    row += 2

    status = tk.Label(root, text="", fg="red")
    status.grid(row=row, column=0, columnspan=3, sticky="w", padx=8)

    def on_start():
        status.config(text="")
        out_base = normalize_path(out_var.get())
        if not out_base:
            status.config(text="Ausgabe-Basisordner ist Pflicht.")
            return

        try:
            _run(month_var.get(), quarter_var.get(), half_var.get(), year_var.get(),
                 out_base, protocol_var.get(), layout_var.get(), overwrite_var.get(),
                 root)
        except Exception:
            status.config(text="Fehler – siehe Protokoll.")
            # Exception wurde bereits geloggt

    tk.Button(root, text="Start", command=on_start, width=12).grid(row=row, column=1, sticky="e", padx=8, pady=10)
    tk.Button(root, text="Schließen", command=root.destroy, width=12).grid(row=row, column=2, sticky="w", padx=8, pady=10)

    # Infozeile
    row += 1
    info = tk.Label(root, text="Hinweis: Layouts werden aus .\\Layouts geladen. Ausgaben gehen nach <Basis>\\VÖ-Tabellen\\<Eingangsordnername>.", fg="gray")
    info.grid(row=row, column=0, columnspan=3, sticky="w", padx=8)

    row += 1
    prot = tk.Label(root, text="", fg="gray")
    prot.grid(row=row, column=0, columnspan=3, sticky="w", padx=8)

    def update_protocol_hint():
        p = normalize_path(protocol_var.get()) if protocol_var.get().strip() else DEFAULT_PROTOCOL_DIR
        # Logger-Datei wird erst beim Start erzeugt, daher nur Ordner anzeigen
        prot.config(text=f"Protokolle werden geschrieben nach: {p}")
        root.after(500, update_protocol_hint)

    update_protocol_hint()

    root.mainloop()

# ============================================================
# Main
# ============================================================

if __name__ == "__main__":
    # global for layout dir
    LAYOUT_DIR = DEFAULT_LAYOUT_DIR
    start_gui()
