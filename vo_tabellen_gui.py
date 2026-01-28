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

PROTOKOLL_DIR = ""  # optional; wird in der GUI gewählt (leer = .\\Protokolle neben EXE)
LAYOUT_DIR = "Layouts"
INTERNAL_HEADER_TEXT = "NUR FÜR DEN INTERNEN DIENSTGEBRAUCH"

# ============================================================
# Hilfsfunktionen: Merge-sicher schreiben
# ============================================================

def _merged_top_left(ws, row, col):
    """Gibt (min_row, min_col) des Merge-Bereichs zurück, wenn (row,col) in einer Merge-Range liegt."""
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return rng.min_row, rng.min_col
    return None

def set_value_merge_safe(ws, row, col, value):
    """
    Setzt einen Wert auch dann, wenn die Zielzelle eine MergedCell ist.
    In dem Fall wird in die Top-Left-Zelle der Merge-Range geschrieben.
    """
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        tl = _merged_top_left(ws, row, col)
        if tl is None:
            return False
        row, col = tl
    ws.cell(row=row, column=col).value = value
    return True


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

GER_MONTHS = [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember"
]
PERIOD_TOKENS = ["Quartal", "Halbjahr"]


class Logger:
    def __init__(self):
        # PROTOKOLL_DIR kommt aus dem GUI (optional).
        # Wenn leer, loggen wir lokal neben der EXE (Unterordner 'Protokolle').
        log_dir = PROTOKOLL_DIR.strip() if isinstance(PROTOKOLL_DIR, str) else ""
        if not log_dir:
            log_dir = os.path.join(os.getcwd(), "Protokolle")
        os.makedirs(log_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        self.path = os.path.join(log_dir, f"vo_tabellen_{ts}.log")

        with open(self.path, "w", encoding="utf-8") as f:
            f.write(f"=== Protokollstart: {datetime.now().isoformat(sep=' ', timespec='seconds')} ===\n")

    def log(self, msg: str):
        line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
        print(line)
        with open(self.path, "a", encoding="utf-8") as f:
            f.write(line + "\n")


def find_period_text(ws, search_rows=40, search_cols=12):
    """
    Ermittelt den Berichtszeitraum aus der Eingangsdatei.
    Monat/Quartal/Halbjahr: enthält z.B. 'Dezember 2025', '4. Quartal 2025', '1. Halbjahr 2025'
    Jahr: steht oft nur '2025' -> wird zu 'Jahr 2025'

    Wir scannen die oberen Zeilen (typisch oberhalb des Tabellenkopfs) über mehrere Spalten,
    weil der Text häufig in zusammengeführten Zellen steht.
    """
    hits = []
    max_r = min(search_rows, ws.max_row)
    max_c = min(search_cols, ws.max_column)

    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            s = v.strip()
            if not s:
                continue

            m_year = re.search(r"(20\d{2})", s)
            if not m_year:
                continue
            year = m_year.group(1)

            if any(m in s for m in GER_MONTHS) or any(tok in s for tok in PERIOD_TOKENS):
                hits.append(s)
                continue

            # nur Jahreszahl (oder ' 2025')
            if re.fullmatch(r"\D*" + re.escape(year) + r"\D*", s):
                hits.append(f"Jahr {year}")

    return hits[-1] if hits else None


def extract_stand_from_raw(ws):
    for r in range(ws.max_row, 0, -1):
        for c in range(ws.max_column, 0, -1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().startswith("Stand:"):
                return v.strip()
    return None


def is_numeric_like(v):
    if v is None:
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return False
        try:
            float(s.replace(",", "."))
            return True
        except Exception:
            return False
    return False


def get_merged_secondary_checker(ws):
    merged_ranges = list(ws.merged_cells.ranges)

    def is_secondary_cell(r, c):
        cell = ws.cell(row=r, column=c)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            return True
        for rng in merged_ranges:
            if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
                return not (r == rng.min_row and c == rng.min_col)
        return False

    return is_secondary_cell


def detect_data_and_footer_tab1(ws):
    first_data = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=2).value
        if isinstance(v, str) and v.strip() == "Insgesamt":
            first_data = r
            break

    footer = ws.max_row
    for r in range(ws.max_row, 0, -1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and "(C)opyright" in v:
            footer = r
            break

    return first_data, footer


def detect_data_and_footer_tab2_3(ws):
    first_data = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=2).value
        if isinstance(v, str) and v.strip() == "Insgesamt":
            first_data = r
            break

    footer = ws.max_row
    for r in range(ws.max_row, 0, -1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and "(C)opyright" in v:
            footer = r
            break

    return first_data, footer


def detect_data_and_footer_tab5(ws):
    footer = ws.max_row
    for r in range(ws.max_row, 0, -1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and "(C)opyright" in v:
            footer = r
            break
    return footer


def get_last_data_col(ws, end_row, max_scan_col=30):
    """
    Ermittelt die letzte 'echte' Tabellenspalte anhand nicht-leerer Werte oberhalb des Footers.
    Verhindert, dass ws.max_column (Styling) z.B. J liefert, obwohl die Tabelle nur bis H geht.
    """
    end_row = max(1, end_row)
    last = 1
    max_c = min(max_scan_col, ws.max_column)
    for r in range(1, end_row + 1):
        for c in range(1, max_c + 1):
            v = ws.cell(row=r, column=c).value
            if v not in (None, ""):
                last = max(last, c)
    return last


def update_footer_with_stand_and_copyright(ws, stand_text):
    max_row = ws.max_row
    max_col = ws.max_column
    current_year = datetime.now().year

    copyright_row = None
    for r in range(max_row, 0, -1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and "(C)opyright" in v:
            new_text = re.sub(r"\(C\)opyright\s+\d{4}", f"(C)opyright {current_year}", v)
            ws.cell(row=r, column=1).value = new_text
            copyright_row = r
            break

    if not copyright_row:
        return

    # andere Stand:-Zeilen entfernen
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().startswith("Stand:") and r != copyright_row:
                ws.cell(row=r, column=c).value = ""

    if not stand_text:
        return

    # Stand-Spalte finden (oder letzte echte Datenspalte)
    stand_col = None
    for c in range(1, max_col + 1):
        v = ws.cell(row=copyright_row, column=c).value
        if isinstance(v, str) and "Stand:" in v:
            stand_col = c
            break
    if stand_col is None:
        stand_col = get_last_data_col(ws, end_row=copyright_row-1)

    cop_cell = ws.cell(row=copyright_row, column=1)
    tgt = ws.cell(row=copyright_row, column=stand_col)
    tgt.value = stand_text

    tgt.font = copy_style(cop_cell.font)
    tgt.border = copy_style(cop_cell.border)
    tgt.fill = copy_style(cop_cell.fill)
    tgt.number_format = cop_cell.number_format
    tgt.protection = copy_style(cop_cell.protection)
    tgt.alignment = Alignment(horizontal="right",
                             vertical=cop_cell.alignment.vertical if cop_cell.alignment else "center")


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


def format_percent_column(ws, col_index):
    """
    Prozentwerte: 1 Nachkommastelle, außer 0 (ohne Nachkommastelle).
    Enthält kein Prozentzeichen in der Zelle.
    """
    fmt = "0.0"
    fmt_zero = "0"

    for r in range(1, ws.max_row + 1):
        cell = ws.cell(row=r, column=col_index)
        v = cell.value
        if v is None or v in ("-", "X"):
            continue
        if isinstance(v, (int, float)):
            if float(v) == 0.0:
                cell.number_format = fmt_zero
            else:
                cell.number_format = fmt
        elif isinstance(v, str):
            s = v.strip().replace(",", ".")
            try:
                f = float(s)
                if f == 0.0:
                    cell.number_format = fmt_zero
                else:
                    cell.number_format = fmt
            except Exception:
                pass


# ============================================================
# Tabelle 1
# ============================================================

def build_table1_workbook(raw_path, layout_path, internal_layout: bool):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[1]]

    period_text = find_period_text(ws_raw)
    stand_text = extract_stand_from_raw(ws_raw)

    wb_out = openpyxl.load_workbook(layout_path)
    ws_out = wb_out[wb_out.sheetnames[0]]

    if internal_layout:
        set_value_merge_safe(ws_out, 1, 1, INTERNAL_HEADER_TEXT)
        set_value_merge_safe(ws_out, 5, 1, period_text)
    else:
        set_value_merge_safe(ws_out, 3, 1, period_text)

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
        # Für JJ soll _g inhaltlich der Eingangsdatei entsprechen (inkl. Spalten J/K),
        # daher aus dem _g-Layout erzeugen und anschließend markieren.
        wb_g = build_table1_workbook(raw_path, layout_g, internal_layout=False)
        ws = wb_g[wb_g.sheetnames[0]]

        # Zeitbezug explizit setzen (Zeile 3)
        wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
        ws_raw = wb_raw[RAW_SHEET_NAMES[1]]
        period_text = find_period_text(ws_raw)
        if period_text:
            set_value_merge_safe(ws, 3, 1, period_text)

        fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        mark_cells_with_1_or_2(ws, 7, fill)  # Tabelle 1: Spalte G markieren

        out_g = os.path.join(output_dir, base + "_g.xlsx")
        wb_g.save(out_g)
        logger.log(f"[T1] _g (JJ: wie Eingang + Markierung) -> {out_g}")
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

    # Kopf/Bezugszeitraum (Positionen unterscheiden sich je Tabelle/Typ)
    # Grundregeln aus den geprüften Mustern:
    #   Tabelle 2: INTERN -> Zeile 6, _g (Monat/Q/H) -> Zeile 4, _g (Jahr) -> Zeile 6
    #   Tabelle 3: INTERN -> Zeile 5, _g (Monat/Q/H) -> Zeile 3, _g (Jahr) -> Zeile 5
    is_year = isinstance(period_text, str) and period_text.strip().startswith("Jahr ")

    if internal_layout:
        set_value_merge_safe(ws_out, 1, 1, INTERNAL_HEADER_TEXT)
        if table_no == 2:
            set_value_merge_safe(ws_out, 6, 1, period_text)
        else:
            set_value_merge_safe(ws_out, 5, 1, period_text)
    else:
        if table_no == 2:
            set_value_merge_safe(ws_out, 6 if is_year else 4, 1, period_text)
        else:
            set_value_merge_safe(ws_out, 5 if is_year else 3, 1, period_text)

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
        ws = wb_i[wb_i.sheetnames[0]]
        ws.cell(row=1, column=1).value = None

        fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        mark_cells_with_1_or_2(ws, 5, fill)  # Tabelle 2/3: Spalte E markieren

        out_g = os.path.join(output_dir, base + "_g.xlsx")
        wb_i.save(out_g)
        logger.log(f"[T{table_no}] _g (JJ=INTERN ohne Kopf + Markierung) -> {out_g}")
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
            set_value_merge_safe(ws, 1, 1, INTERNAL_HEADER_TEXT)
            set_value_merge_safe(ws, 5, 1, period_text)
        else:
            set_value_merge_safe(ws, 3, 1, period_text)

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

    wb_g = build_table5_workbook(raw_path, layout_g, internal_layout=False, is_jj=is_jj)
    out_g = os.path.join(output_dir, base + "_g.xlsx")
    wb_g.save(out_g)
    logger.log(f"[T5] _g -> {out_g}")


# ============================================================
# Dateisuche / Ausgabepfade
# ============================================================

def find_raw_files(input_dir: str):
    hits = []
    for pref in PREFIX_TO_TABLE.keys():
        hits += glob.glob(os.path.join(input_dir, pref + "*.xlsx"))
    hits = [h for h in hits if not (h.endswith("_g.xlsx") or h.endswith("_INTERN.xlsx"))]
    return sorted(set(hits))


def ensure_output_run_folder(base_out_dir: str, input_dir: str):
    os.makedirs(base_out_dir, exist_ok=True)
    vo_root = os.path.join(base_out_dir, "VÖ-Tabellen")
    os.makedirs(vo_root, exist_ok=True)

    run_name = os.path.basename(os.path.normpath(input_dir))
    out_dir = os.path.join(vo_root, run_name)
    os.makedirs(out_dir, exist_ok=True)
    return out_dir


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


def choose_dir(entry: ttk.Entry):
    d = filedialog.askdirectory()
    if d:
        entry.delete(0, tk.END)
        entry.insert(0, d)


def start_gui():
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
    add_row(5, "Protokollordner (optional, leer = .\\Protokolle):", "protokoll")

    status_var = tk.StringVar(value="Bereit.")
    ttk.Label(frm, textvariable=status_var).grid(row=6, column=0, columnspan=3, sticky="w", pady=(10, 0))

    def on_run():
        global PROTOKOLL_DIR
        PROTOKOLL_DIR = entries["protokoll"].get().strip()
        logger = Logger()
        logpath_var.set(f"Protokoll wird geschrieben nach: {logger.path}")
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
    btn_row.grid(row=7, column=0, columnspan=3, sticky="e", pady=10)

    ttk.Button(btn_row, text="Start", command=on_run).grid(row=0, column=0, padx=6)
    ttk.Button(btn_row, text="Schließen", command=root.destroy).grid(row=0, column=1, padx=6)

    ttk.Label(
        frm,
        text="Hinweis: Layouts werden aus .\\Layouts geladen. Ausgaben gehen nach <Basis>\\VÖ-Tabellen\\<Eingangsordnername>.",
    ).grid(row=8, column=0, columnspan=3, sticky="w")

    logpath_var = tk.StringVar(value="Protokoll wird geschrieben nach: (noch nicht gestartet)")
    ttk.Label(frm, textvariable=logpath_var).grid(row=9, column=0, columnspan=3, sticky="w")

    root.mainloop()


if __name__ == "__main__":
    start_gui()
