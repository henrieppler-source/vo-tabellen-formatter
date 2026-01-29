import os
import glob
import re
from datetime import datetime
from copy import copy as copy_style

import openpyxl
from openpyxl.styles import Alignment, PatternFill

import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ============================================================
# FESTE VERZEICHNISSE (wie vorgegeben)
# ============================================================

PROTOKOLL_DIR = "Protokolle"  # Standard: relativ -> neben EXE/WorkingDir
LAYOUT_DIR = "Layouts"
INTERNAL_HEADER_TEXT = "NUR FÜR DEN INTERNEN DIENSTGEBRAUCH"

RAW_SHEET_NAMES = {
    1: "XML-Tab1-Land",
    2: "XML-Tab2-Land",
    3: "XML-Tab3-Land",
    5: "XML-Tab5-Land",
}

TEMPLATES = {
    1: {"ext": "Tabelle-1-Layout_g.xlsx", "int": "Tabelle-1-Layout_INTERN.xlsx"},
    2: {"ext": "Tabelle-2-Layout_g.xlsx", "int": "Tabelle-2-Layout_INTERN.xlsx"},
    3: {"ext": "Tabelle-3-Layout_g.xlsx", "int": "Tabelle-3-Layout_INTERN.xlsx"},
    5: {"ext": "Tabelle-5-Layout_g.xlsx", "int": "Tabelle-5-Layout_INTERN.xlsx"},
}

PREFIX_TO_TABLE = {
    "Tabelle-1-Land": 1,
    "Tabelle-2-Land": 2,
    "Tabelle-3-Land": 3,
    "Tabelle-5-Land": 5,
}

# später:
# TAB8_PREFIXES = ["25_Tab8_", "26_Tab8_", "27_Tab8_", "28_Tab8_"]
# TAB9_PREFIXES = ["29_Tab9_", "30_Tab9_", "31_Tab9_", "32_Tab9_"]


# ============================================================
# Logging
# ============================================================

class Logger:
    """Einfacher Logger: schreibt immer in eine Logdatei.
    Das Zielverzeichnis kann zur Laufzeit umgestellt werden (move_to)."""

    def __init__(self, base_dir: str | None = None):
        if base_dir is None or str(base_dir).strip() == "":
            base_dir = os.getcwd()
        self.base_dir = os.path.normpath(base_dir)
        if not os.path.isabs(self.base_dir):
            self.base_dir = os.path.normpath(os.path.join(os.getcwd(), self.base_dir))
        os.makedirs(self.base_dir, exist_ok=True)

        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        self.path = os.path.join(self.base_dir, f"vo_tabellen_{ts}.log")
        self._fh = open(self.path, "a", encoding="utf-8")
        self.log(f"=== Protokollstart: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

    def close(self):
        try:
            self._fh.close()
        except Exception:
            pass

    def log(self, msg: str):
        line = f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
        self._fh.write(line + "\n")
        self._fh.flush()

    def move_to(self, new_dir: str):
        """Wechselt das Protokollverzeichnis. Neuer Logfile-Name, weiter schreiben."""
        new_dir = os.path.normpath(new_dir)
        if not new_dir:
            return
        os.makedirs(new_dir, exist_ok=True)

        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        new_path = os.path.join(new_dir, f"vo_tabellen_{ts}.log")

        old_path = getattr(self, "path", None)

        try:
            self._fh.close()
        except Exception:
            pass

        self.base_dir = new_dir
        self.path = new_path
        self._fh = open(self.path, "a", encoding="utf-8")
        self.log(f"=== Protokoll fortgesetzt (vorher: {old_path}) ===")
def is_numeric_like(v):
    if v is None:
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        s = v.strip()
        if s in ("-", "X"):
            return True
        s2 = s.replace(".", "").replace(",", "").replace(" ", "")
        return s2.isdigit()
    return False


GER_MONTHS = (
    "Januar", "Februar", "März", "Maerz", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember"
)
PERIOD_TOKENS = ("Quartal", "Halbjahr", "Jahr", "Jahres", "Q1", "Q2", "Q3", "Q4", "H1", "H2", "JJ")


def find_period_text(ws, search_rows=30):
    hits = []
    for r in range(1, min(search_rows, ws.max_row) + 1):
        v = ws.cell(row=r, column=1).value
        if not isinstance(v, str):
            continue
        s = v.strip()
        if not re.search(r"(20\d{2})", s):
            continue
        if any(m in s for m in GER_MONTHS) or any(tok in s for tok in PERIOD_TOKENS):
            hits.append(s)
    return hits[-1] if hits else None


def extract_stand_from_raw(ws, max_search_rows=80):
    max_row = ws.max_row
    max_col = ws.max_column
    for r in range(max_row, max(max_row - max_search_rows, 1) - 1, -1):
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "Stand:" in v:
                return v.strip()
    return None


def get_merged_secondary_checker(ws):
    merged = list(ws.merged_cells.ranges)

    def is_secondary(row, col):
        for rg in merged:
            if rg.min_row <= row <= rg.max_row and rg.min_col <= col <= rg.max_col:
                return not (row == rg.min_row and col == rg.min_col)
        return False
    return is_secondary


def update_footer_with_stand_and_copyright(ws, stand_text):
    max_row = ws.max_row
    max_col = ws.max_column
    current_year = datetime.now().year


def ensure_copyright_footer_fixed_row(ws_target, ws_template, stand_text, fixed_row=65, stand_row_offset=0):
    """
    Ensures a footer at a fixed row.

    - Row `fixed_row`, Col A: '(C)opyright <current_year> Bayerisches Landesamt für Statistik'
    - Row `fixed_row + stand_row_offset`, last data column: 'Stand: dd.mm.yyyy' (right aligned)
      (If stand_row_offset==0 -> same row as copyright.)
    - Styles are copied from the template's copyright cell if available.
    """
    from datetime import datetime

    current_year = datetime.now().year

    # locate style source in template (copyright cell in col A)
    style_cell = None
    if ws_template is not None:
        for r in range(ws_template.max_row, 0, -1):
            v = ws_template.cell(row=r, column=1).value
            if isinstance(v, str) and "(C)opyright" in v:
                style_cell = ws_template.cell(row=r, column=1)
                break
    if style_cell is None:
        style_cell = ws_target.cell(row=1, column=1)

    stand_row = fixed_row + int(stand_row_offset or 0)

    # Remove other Stand: occurrences (keep only the chosen stand row)
    for r in range(1, ws_target.max_row + 1):
        if r == stand_row:
            continue
        for c in range(1, ws_target.max_column + 1):
            v = ws_target.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().startswith("Stand:"):
                ws_target.cell(row=r, column=c).value = ""

    # Write copyright safely even if merged
    set_value_merge_safe(ws_target, fixed_row, 1, f"(C)opyright {current_year} Bayerisches Landesamt für Statistik")
    cop_cell = ws_target.cell(row=fixed_row, column=1)
    cop_cell.font = copy_style(style_cell.font)
    cop_cell.border = copy_style(style_cell.border)
    cop_cell.fill = copy_style(style_cell.fill)
    cop_cell.number_format = style_cell.number_format
    cop_cell.protection = copy_style(style_cell.protection)
    cop_cell.alignment = copy_style(style_cell.alignment) if style_cell.alignment else Alignment(horizontal="left", vertical="center")

    # Stand in last data column under the table
    last_col = get_last_data_col(ws_target, end_row=max(1, fixed_row - 1))
    stand_cell = ws_target.cell(row=stand_row, column=last_col)
    stand_cell.value = stand_text or ""
    stand_cell.font = copy_style(style_cell.font)
    stand_cell.border = copy_style(style_cell.border)
    stand_cell.fill = copy_style(style_cell.fill)
    stand_cell.number_format = style_cell.number_format
    stand_cell.protection = copy_style(style_cell.protection)
    stand_cell.alignment = Alignment(horizontal="right", vertical=style_cell.alignment.vertical if style_cell.alignment else "center")

def mark_cells_with_1_or_2(ws, col_index, fill):
    for r in range(1, ws.max_row + 1):
        cell = ws.cell(row=r, column=col_index)
        v = cell.value
        if isinstance(v, (int, float)) and v in (1, 2):
            cell.fill = fill
        elif isinstance(v, str) and v.strip() in ("1", "2"):
            cell.fill = fill


def format_numeric_cells(ws, skip_cols=None):
    """
    Ganzzahlen: Tausendertrennzeichen = Leerzeichen, keine Dezimalstellen
    Negative: "- " vorangestellt
    """
    if skip_cols is None:
        skip_cols = set()

    pos = "#\\ ###\\ ###\\ ###\\ ###\\ ##0"
    neg = "-\\ " + pos
    fmt = f"{pos};{neg};0"

    for row in ws.iter_rows():
        for cell in row:
            if cell.column in skip_cols:
                continue
            v = cell.value
            if v is None or v in ("-", "X"):
                continue
            if isinstance(v, (int, float)):
                if isinstance(v, float):
                    cell.value = int(round(v))
                cell.number_format = fmt


def format_percent_column(ws, col_index: int):
    """
    Prozentspalten ohne %-Zeichen:
    - immer 1 Nachkommastelle
    - außer 0 -> 0
    """
    pct_fmt = "0.0;-0.0;0"
    for r in range(1, ws.max_row + 1):
        cell = ws.cell(row=r, column=col_index)
        v = cell.value
        if v is None or v in ("-", "X"):
            continue
        if isinstance(v, (int, float)):
            cell.number_format = pct_fmt


def detect_data_and_footer_tab1(sheet):
    """
    Tabelle 1: Datenstart = erste Zeile mit Zahl ab Spalte B irgendwo
    Fußnotenstart = erste Zeile in A mit '-'
    """
    max_row = sheet.max_row
    first_data = None
    for r in range(1, max_row + 1):
        for c in range(2, sheet.max_column + 1):
            if is_numeric_like(sheet.cell(row=r, column=c).value):
                first_data = r
                break
        if first_data:
            break
    if first_data is None:
        first_data = 1

    footnote_start = max_row + 1
    for r in range(1, max_row + 1):
        v = sheet.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip().startswith("-"):
            footnote_start = r
            break
    return first_data, footnote_start


def detect_data_and_footer_tab2_3(sheet):
    """
    Tabelle 2/3: Datenstart über numerische Werte ab Spalte C (B ist Text)
    """
    max_row = sheet.max_row
    first_data = None
    for r in range(1, max_row + 1):
        for c in range(3, sheet.max_column + 1):  # ab C
            if is_numeric_like(sheet.cell(row=r, column=c).value):
                first_data = r
                break
        if first_data:
            break
    if first_data is None:
        first_data = 1

    footnote_start = max_row + 1
    for r in range(1, max_row + 1):
        v = sheet.cell(row=r, column=1).value
        if isinstance(v, str) and v.strip().startswith("-"):
            footnote_start = r
            break
    return first_data, footnote_start


# ============================================================
# Verarbeitung Tabelle 1
# ============================================================

def build_table1_workbook(raw_path, layout_path, internal_layout: bool):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[1]]

    period_text = find_period_text(ws_raw)
    stand_text = extract_stand_from_raw(ws_raw)

    wb_out = openpyxl.load_workbook(layout_path)
    ws_out = wb_out[wb_out.sheetnames[0]]

    if internal_layout:
        ws_out.cell(row=1, column=1).value = INTERNAL_HEADER_TEXT
        ws_out.cell(row=5, column=1).value = period_text
    else:
        ws_out.cell(row=3, column=1).value = period_text

    is_sec = get_merged_secondary_checker(ws_out)
    fdr_raw, ft_raw = detect_data_and_footer_tab1(ws_raw)
    fdr_out, ft_out = detect_data_and_footer_tab1(ws_out)

    n_rows = min(ft_raw - fdr_raw, ft_out - fdr_out)
    max_col_out = ws_out.max_column

    for off in range(n_rows):
        r_raw = fdr_raw + off
        r_out = fdr_out + off
        for c in range(1, max_col_out + 1):
            if is_sec(r_out, c):
                continue
            ws_out.cell(row=r_out, column=c).value = ws_raw.cell(row=r_raw, column=c).value

    update_footer_with_stand_and_copyright(ws_out, stand_text)

    # Tabelle 1: Spalte I (9) = Prozentwerte
    format_percent_column(ws_out, 9)
    format_numeric_cells(ws_out, skip_cols={9})

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
        # JJ (_g): entspricht der Eingangsdatei (inkl. Spalten J/K), ohne INTERN-Kopfzeile,
        # mit Markierung in Spalte G (1/2) und Footer in Zeile 65.
        wb_g = openpyxl.load_workbook(raw_path)  # Formate/Merges/Spalten komplett behalten
        ws_g = wb_g[RAW_SHEET_NAMES[1]]

        # Berichtszeitraum aus Eingangsdatei -> bei JJ zu "Jahr 2025"
        wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
        ws_raw = wb_raw[RAW_SHEET_NAMES[1]]
        period_text = find_period_text(ws_raw)
        if period_text:
            pt = str(period_text).strip()
            if re.fullmatch(r"\s*\d{4}\s*", pt):
                pt = f"Jahr {pt}"
            set_value_merge_safe(ws_g, 3, 1, pt)

        # Stand aus Eingangsdatei
        stand_text = extract_stand_from_raw(ws_raw)

        # Markierung: Spalte G
        fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        mark_cells_with_1_or_2(ws_g, 7, fill)

        # Footer sicherstellen: Zeile 65 (Styles aus INTERN-Layout)
        ws_i = wb_i[wb_i.sheetnames[0]] if wb_i.sheetnames else None
        ensure_copyright_footer_fixed_row(ws_g, ws_i, stand_text, fixed_row=65, stand_row_offset=1)

        out_g = os.path.join(output_dir, base + "_g.xlsx")
        wb_g.save(out_g)
        logger.log(f"[T1] _g (JJ: Eingang inkl. J/K + Markierung + Footer) -> {out_g}")
    else:
        wb_g = build_table1_workbook(raw_path, layout_g, internal_layout=False)
        out_g = os.path.join(output_dir, base + "_g.xlsx")
        wb_g.save(out_g)
        logger.log(f"[T1] _g -> {out_g}")


# ============================================================
# Verarbeitung Tabelle 2 & 3
# ============================================================

def build_table2_3_workbook(table_no, raw_path, layout_path, internal_layout: bool):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[table_no]]

    period_text = find_period_text(ws_raw)
    stand_text = extract_stand_from_raw(ws_raw)

    wb_out = openpyxl.load_workbook(layout_path)
    ws_out = wb_out[wb_out.sheetnames[0]]

    # Kopf/Bezugszeitraum
    if internal_layout:
        ws_out.cell(row=1, column=1).value = INTERNAL_HEADER_TEXT
        ws_out.cell(row=6, column=1).value = period_text
    else:
        ws_out.cell(row=3, column=1).value = period_text
        if table_no == 2:
            # alte Monatszeile raus
            ws_out.cell(row=4, column=1).value = None

    # Daten kopieren: ab Spalte B überschreiben (damit B nicht aus Layout bleibt)
    START_COL_COPY = 2
    is_sec = get_merged_secondary_checker(ws_out)

    fdr_raw, ft_raw = detect_data_and_footer_tab2_3(ws_raw)
    fdr_out, ft_out = detect_data_and_footer_tab2_3(ws_out)

    n_rows = min(ft_raw - fdr_raw, ft_out - fdr_out)

    for off in range(n_rows):
        r_raw = fdr_raw + off
        r_out = fdr_out + off
        for c in range(START_COL_COPY, ws_out.max_column + 1):
            if is_sec(r_out, c):
                continue
            ws_out.cell(row=r_out, column=c).value = ws_raw.cell(row=r_raw, column=c).value

    update_footer_with_stand_and_copyright(ws_out, stand_text)

    # Tabelle 2/3: Spalte G (7) = Prozent
    format_percent_column(ws_out, 7)
    format_numeric_cells(ws_out, skip_cols={7})

    return wb_out


def process_table2_or_3_file(table_no, raw_path, output_dir, logger: Logger):
    base = os.path.splitext(os.path.basename(raw_path))[0]
    is_jj = "-JJ" in base

    layout_g = os.path.join(LAYOUT_DIR, TEMPLATES[table_no]["ext"])
    layout_i = os.path.join(LAYOUT_DIR, TEMPLATES[table_no]["int"])

    wb_i = build_table2_3_workbook(table_no, raw_path, layout_i, internal_layout=True)
    out_i = os.path.join(output_dir, base + "_INTERN.xlsx")
    wb_i.save(out_i)
    logger.log(f"[T{table_no}] INTERN -> {out_i}")

    if is_jj:
        # JJ (_g): entspricht der Eingangsdatei (inkl. Spalten J/K), ohne INTERN-Kopfzeile,
        # mit Markierung in Spalte G (1/2) und Footer in Zeile 65.
        wb_g = openpyxl.load_workbook(raw_path)  # Formate/Merges/Spalten komplett behalten
        ws_g = wb_g[RAW_SHEET_NAMES[1]]

        # Berichtszeitraum aus Eingangsdatei -> bei JJ zu "Jahr 2025"
        wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
        ws_raw = wb_raw[RAW_SHEET_NAMES[1]]
        period_text = find_period_text(ws_raw)
        if period_text:
            pt = str(period_text).strip()
            if re.fullmatch(r"\s*\d{4}\s*", pt):
                pt = f"Jahr {pt}"
            set_value_merge_safe(ws_g, 3, 1, pt)

        # Stand aus Eingangsdatei
        stand_text = extract_stand_from_raw(ws_raw)

        # Markierung: Spalte G
        fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        mark_cells_with_1_or_2(ws_g, 7, fill)

        # Footer sicherstellen: Zeile 65 (Styles aus INTERN-Layout)
        ws_i = wb_i[wb_i.sheetnames[0]] if wb_i.sheetnames else None
        ensure_copyright_footer_fixed_row(ws_g, ws_i, stand_text, fixed_row=65, stand_row_offset=1)

        out_g = os.path.join(output_dir, base + "_g.xlsx")
        wb_g.save(out_g)
        logger.log(f"[T1] _g (JJ: Eingang inkl. J/K + Markierung + Footer) -> {out_g}")
    else:
        wb_g = build_table2_3_workbook(table_no, raw_path, layout_g, internal_layout=False)
        out_g = os.path.join(output_dir, base + "_g.xlsx")
        wb_g.save(out_g)
        logger.log(f"[T{table_no}] _g -> {out_g}")


# ============================================================
# Verarbeitung Tabelle 5
# ============================================================

def build_table5_workbook(raw_path, layout_path, internal_layout: bool, is_jj: bool):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[5]]

    period_text = find_period_text(ws_raw)
    stand_text = extract_stand_from_raw(ws_raw)

    wb_out = openpyxl.load_workbook(layout_path)
    max_row_raw = ws_raw.max_row

    # Blockstart ("Bayern 1)"..)
    starts = []
    for r in range(1, max_row_raw + 1):
        v = ws_raw.cell(row=r, column=2).value
        if isinstance(v, str) and re.match(r"Bayern\s+\d\)", v.strip()):
            starts.append(r)

    block_ranges = []
    for i, start in enumerate(starts):
        end = (starts[i + 1] - 1) if i < len(starts) - 1 else max_row_raw
        last_nonempty = start
        for rr in range(start, end + 1):
            if any(ws_raw.cell(row=rr, column=c).value not in (None, "") for c in range(1, 25)):
                last_nonempty = rr
        block_ranges.append((start, last_nonempty))

    def fill_sheet_from_block(ws_out, start_row, end_row):
        is_sec = get_merged_secondary_checker(ws_out)

        # Datenstart im Layout: erste Zeile mit Zahl in Spalte C
        first_data_out = None
        for r in range(1, ws_out.max_row + 1):
            if is_numeric_like(ws_out.cell(row=r, column=3).value):
                first_data_out = r
                break
        if first_data_out is None:
            return

        # INTERN: immer C..J
        # _g:     JJ -> C..J, sonst C..H und I/J leer
        if internal_layout:
            cols = range(3, 11)  # C..J
        else:
            cols = range(3, 11) if is_jj else range(3, 9)  # JJ: C..J, sonst C..H

        raw_r = start_row
        out_r = first_data_out
        while raw_r <= end_row and out_r <= ws_out.max_row:
            for c in cols:
                if is_sec(out_r, c):
                    continue
                ws_out.cell(row=out_r, column=c).value = ws_raw.cell(row=raw_r, column=c).value
            raw_r += 1
            out_r += 1

        if (not internal_layout) and (not is_jj):
            # I/J sicher leer lassen
            for rr in range(first_data_out, out_r):
                for cc in (9, 10):
                    if not is_sec(rr, cc):
                        ws_out.cell(row=rr, column=cc).value = None

    for i, (start, end) in enumerate(block_ranges):
        if i >= len(wb_out.worksheets):
            break

        ws = wb_out.worksheets[i]

        if internal_layout:
            ws.cell(row=1, column=1).value = INTERNAL_HEADER_TEXT
            ws.cell(row=5, column=1).value = period_text
        else:
            ws.cell(row=3, column=1).value = period_text

        fill_sheet_from_block(ws, start, end)
        update_footer_with_stand_and_copyright(ws, stand_text)

        # Tabelle 5: Spalte H (8) = Prozent
        format_percent_column(ws, 8)
        format_numeric_cells(ws, skip_cols={8})

    return wb_out


def process_table5_file(raw_path, output_dir, logger: Logger):
    base = os.path.splitext(os.path.basename(raw_path))[0]
    is_jj = "-JJ" in base

    layout_g = os.path.join(LAYOUT_DIR, TEMPLATES[5]["ext"])
    layout_i = os.path.join(LAYOUT_DIR, TEMPLATES[5]["int"])

    wb_i = build_table5_workbook(raw_path, layout_i, internal_layout=True, is_jj=is_jj)
    out_i = os.path.join(output_dir, base + "_INTERN.xlsx")
    wb_i.save(out_i)
    logger.log(f"[T5] INTERN -> {out_i}")

    if is_jj:
        # JJ (_g): entspricht der Eingangsdatei (inkl. Spalten J/K), ohne INTERN-Kopfzeile,
        # mit Markierung in Spalte G (1/2) und Footer in Zeile 65.
        wb_g = openpyxl.load_workbook(raw_path)  # Formate/Merges/Spalten komplett behalten
        ws_g = wb_g[RAW_SHEET_NAMES[1]]

        # Berichtszeitraum aus Eingangsdatei -> bei JJ zu "Jahr 2025"
        wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
        ws_raw = wb_raw[RAW_SHEET_NAMES[1]]
        period_text = find_period_text(ws_raw)
        if period_text:
            pt = str(period_text).strip()
            if re.fullmatch(r"\s*\d{4}\s*", pt):
                pt = f"Jahr {pt}"
            set_value_merge_safe(ws_g, 3, 1, pt)

        # Stand aus Eingangsdatei
        stand_text = extract_stand_from_raw(ws_raw)

        # Markierung: Spalte G
        fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        mark_cells_with_1_or_2(ws_g, 7, fill)

        # Footer sicherstellen: Zeile 65 (Styles aus INTERN-Layout)
        ws_i = wb_i[wb_i.sheetnames[0]] if wb_i.sheetnames else None
        ensure_copyright_footer_fixed_row(ws_g, ws_i, stand_text, fixed_row=65, stand_row_offset=1)

        out_g = os.path.join(output_dir, base + "_g.xlsx")
        wb_g.save(out_g)
        logger.log(f"[T1] _g (JJ: Eingang inkl. J/K + Markierung + Footer) -> {out_g}")
    else:
        wb_g = build_table5_workbook(raw_path, layout_g, internal_layout=False, is_jj=is_jj)
        out_g = os.path.join(output_dir, base + "_g.xlsx")
        wb_g.save(out_g)
        logger.log(f"[T5] _g -> {out_g}")


# ============================================================
# Dateisuche & Ausgabeordner
# ============================================================

def find_raw_files(input_dir):
    """
    Nur Rohdateien, die mit Tabelle-1/2/3/5-Land beginnen.
    _g.xlsx und _INTERN.xlsx werden ignoriert.
    """
    files = []
    for prefix in PREFIX_TO_TABLE.keys():
        files.extend(glob.glob(os.path.join(input_dir, f"{prefix}_*.xlsx")))
    files = sorted(set(files))
    files = [f for f in files if not f.endswith("_g.xlsx") and not f.endswith("_INTERN.xlsx")]
    return files


def ensure_output_run_folder(base_out_dir: str, input_folder: str) -> str:
    """
    base_out_dir = z.B. ...\2025\2025-12
    Ausgabe     = base_out_dir\VÖ-Tabellen\<Eingangsordnername>
    """
    vo_base = os.path.join(base_out_dir, "VÖ-Tabellen")
    os.makedirs(vo_base, exist_ok=True)

    run_dir = os.path.join(vo_base, os.path.basename(input_folder.rstrip("\\/")))
    os.makedirs(run_dir, exist_ok=True)
    return run_dir


# ============================================================
# GUI
# ============================================================

def choose_dir(entry: ttk.Entry):
    d = filedialog.askdirectory()
    if d:
        entry.delete(0, tk.END)
        entry.insert(0, d)


def run_for_one_input_dir(input_dir: str, base_out_dir: str, logger: Logger, status_var: tk.StringVar):
    if not os.path.isdir(input_dir):
        logger.log(f"[SKIP] Eingangspfad existiert nicht: {input_dir}")
        return

    out_dir = ensure_output_run_folder(base_out_dir, input_dir)
    logger.log(f"--- Eingang: {input_dir}")
    logger.log(f"--- Ausgabe: {out_dir}")

    files = find_raw_files(input_dir)
    if not files:
        logger.log(f"[INFO] Keine passenden Rohdateien gefunden in: {input_dir}")
        return

    # Layout-Existenz prüfen (für 1/2/3/5)
    missing_layouts = []
    for t in (1, 2, 3, 5):
        for k in ("ext", "int"):
            p = os.path.join(LAYOUT_DIR, TEMPLATES[t][k])
            if not os.path.exists(p):
                missing_layouts.append(p)
    if missing_layouts:
        raise FileNotFoundError("Layout-Dateien fehlen:\n" + "\n".join(missing_layouts))

    for f in files:
        fname = os.path.basename(f)
        table_no = None
        for prefix, tno in PREFIX_TO_TABLE.items():
            if fname.startswith(prefix):
                table_no = tno
                break

        if table_no is None:
            logger.log(f"[SKIP] Unbekanntes Muster: {f}")
            continue

        status_var.set(f"Verarbeite {os.path.basename(input_dir)}: {fname}")
        logger.log(f"[START] {fname}")

        if table_no == 1:
            process_table1_file(f, out_dir, logger)
        elif table_no in (2, 3):
            process_table2_or_3_file(table_no, f, out_dir, logger)
        elif table_no == 5:
            process_table5_file(f, out_dir, logger)

        logger.log(f"[OK]    {fname}")


def run_processing(monat_dir, quartal_dir, halbjahr_dir, jahr_dir, base_out_dir, logger: Logger, status_var: tk.StringVar):
    if not base_out_dir:
        messagebox.showerror("Fehler", "Ausgabe-Basisordner muss angegeben werden.")
        return

    jobs = []
    if monat_dir.strip(): jobs.append(("Monat", monat_dir.strip()))
    if quartal_dir.strip(): jobs.append(("Quartal", quartal_dir.strip()))
    if halbjahr_dir.strip(): jobs.append(("Halbjahr", halbjahr_dir.strip()))
    if jahr_dir.strip(): jobs.append(("Jahr", jahr_dir.strip()))

    if not jobs:
        messagebox.showwarning("Hinweis", "Kein Eingangsverzeichnis gesetzt (Monat/Quartal/Halbjahr/Jahr).")
        return

    logger.log("=== START GUI-Lauf ===")
    logger.log(f"Ausgabe-Basis: {base_out_dir}")
    logger.log(f"Layouts:       {os.path.abspath(LAYOUT_DIR)}")
    logger.log(f"Überschreiben: JA (vorhandene Dateien werden ersetzt)")
    logger.log("Jobs: " + ", ".join([f"{lbl}={path}" for lbl, path in jobs]))

    try:
        for lbl, in_dir in jobs:
            status_var.set(f"Starte: {lbl}")
            logger.log(f"== JOB: {lbl} ==")
            run_for_one_input_dir(in_dir, base_out_dir, logger, status_var)

        status_var.set("Fertig.")
        logger.log("=== ENDE GUI-Lauf ===")
        messagebox.showinfo("Fertig", f"Verarbeitung abgeschlossen.\nProtokoll:\n{logger.path}")
    except Exception as e:
        status_var.set("Fehler – siehe Protokoll.")
        logger.log(f"[FEHLER] {repr(e)}")
        messagebox.showerror("Fehler", f"Es ist ein Fehler aufgetreten:\n{e}\n\nProtokoll:\n{logger.path}")


def start_gui():
    logger = Logger()

    root = tk.Tk()
    root.title("VÖ-Tabellen – GUI (Tabelle 1/2/3/5)")

    frm = ttk.Frame(root, padding=12)
    frm.grid(row=0, column=0, sticky="nsew")

    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    frm.columnconfigure(1, weight=1)

    entries = {}

    def add_row(r, label, key):
        ttk.Label(frm, text=label).grid(row=r, column=0, sticky="w", pady=4)
        e = ttk.Entry(frm, width=90)
        e.grid(row=r, column=1, sticky="we", padx=6)
        ttk.Button(frm, text="Auswählen…", command=lambda: choose_dir(e)).grid(row=r, column=2, sticky="e")
        entries[key] = e

    add_row(0, "Monat – Eingangstabellen (optional):", "monat")
    add_row(1, "Quartal – Eingangstabellen (optional):", "quartal")
    add_row(2, "Halbjahr – Eingangstabellen (optional):", "halbjahr")
    add_row(3, "Jahr – Eingangstabellen (optional):", "jahr")
    add_row(4, "Ausgabe-Basisordner (Pflicht):", "outbase")

    status_var = tk.StringVar(value="Bereit.")
    ttk.Label(frm, textvariable=status_var).grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))

    def on_run():
        run_processing(
            entries["monat"].get(),
            entries["quartal"].get(),
            entries["halbjahr"].get(),
            entries["jahr"].get(),
            entries["outbase"].get(),
            logger,
            status_var
        )

    btn_row = ttk.Frame(frm)
    btn_row.grid(row=6, column=0, columnspan=3, sticky="e", pady=10)

    ttk.Button(btn_row, text="Start", command=on_run).grid(row=0, column=0, padx=6)
    ttk.Button(btn_row, text="Schließen", command=root.destroy).grid(row=0, column=1, padx=6)

    ttk.Label(
        frm,
        text="Hinweis: Layouts werden aus .\\Layouts geladen. Ausgaben gehen nach <Basis>\\VÖ-Tabellen\\<Eingangsordnername>.",
    ).grid(row=7, column=0, columnspan=3, sticky="w")

    ttk.Label(
        frm,
        text=f"Protokolle werden geschrieben nach: {PROTOKOLL_DIR}",
    ).grid(row=8, column=0, columnspan=3, sticky="w")

    root.mainloop()


if __name__ == "__main__":
    start_gui()
