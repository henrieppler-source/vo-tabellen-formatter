# vo_tabellen_gui_v005.py
# (kompletter Code)

from __future__ import annotations

import os
import re
import sys
import shutil
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, List, Dict

import tkinter as tk
from tkinter import filedialog, messagebox

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from copy import copy

APP_TITLE = "VÖ-Tabellen – GUI (Tabelle 1/2/3/5)"

# -----------------------------
# Defaults / Konventionen
# -----------------------------
DEFAULT_LAYOUTS_SUBDIR = "Layouts"
DEFAULT_VOE_SUBDIR = "VÖ-Tabellen"
DEFAULT_LOG_SUBDIR = "Protokolle"

RAW_PREFIXES_1TO5 = [
    "Tabelle-1-Land_",
    "Tabelle-2-Land_",
    "Tabelle-3-Land_",
    "Tabelle-5-Land_",
]

RAW_SHEET_NAMES = {
    1: "XML-Tab1-Land",
    2: "XML-Tab2-Land",
    3: "XML-Tab3-Land",
    5: "XML-Tab5-Land",
}

LAYOUT_FILES = {
    (1, True): "Tabelle-1-Layout_INTERN.xlsx",
    (1, False): "Tabelle-1-Layout_g.xlsx",
    (2, True): "Tabelle-2-Layout_INTERN.xlsx",
    (2, False): "Tabelle-2-Layout_g.xlsx",
    (3, True): "Tabelle-3-Layout_INTERN.xlsx",
    (3, False): "Tabelle-3-Layout_g.xlsx",
    (5, True): "Tabelle-5-Layout_INTERN.xlsx",
    (5, False): "Tabelle-5-Layout_g.xlsx",
}

# Zahlenformat: Tausender mit Leerzeichen (Excel-Formatcode)
FMT_INT_SPACE = "# ##0"
FMT_INT_SPACE_NEG = "# ##0;-# ##0"
# Prozent-Spalten: immer 1 Nachkommastelle, außer 0 -> 0 (ohne Nachkommastelle)
# (wir lösen das über zwei Formate + Wertprüfung im Code)
FMT_PCT_1_DEC = "0.0"
FMT_PCT_ZERO = "0"

# -------------------------------------
# Logging
# -------------------------------------
class Logger:
    def __init__(self, log_dir: Path, base_dir: Path):
        self.log_dir = log_dir
        self.base_dir = base_dir
        self.log_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        self.path = self.base_dir / f"vo_tabellen_{ts}.log"
        self._fh = open(self.path, "w", encoding="utf-8", newline="\n")
        self.log(f"=== Protokollstart: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

    def log(self, msg: str):
        line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
        print(line)
        try:
            self._fh.write(line + "\n")
            self._fh.flush()
        except Exception:
            pass

    def close(self):
        try:
            self._fh.close()
        except Exception:
            pass

# -------------------------------------
# Utilities
# -------------------------------------
def norm_path(p: str) -> str:
    return str(Path(p).resolve())

def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)

def is_relevant_file(name: str) -> bool:
    if not name.lower().endswith(".xlsx"):
        return False
    return any(name.startswith(pref) for pref in RAW_PREFIXES_1TO5)

def detect_period_from_filename(fn: str) -> str:
    # Erwartet z.B.: Tabelle-1-Land_2025-12.xlsx / 2025-Q4 / 2025-H2 / 2025-JJ
    m = re.search(r"_([0-9]{4}-(?:[0-9]{2}|Q[1-4]|H[1-2]|JJ))", fn)
    return m.group(1) if m else "UNBEKANNT"

def parse_table_no(fn: str) -> Optional[int]:
    m = re.match(r"Tabelle-(\d+)-Land_", fn)
    if not m:
        return None
    try:
        return int(m.group(1))
    except Exception:
        return None

def set_value_merge_safe(ws: openpyxl.worksheet.worksheet.Worksheet, row: int, col: int, value):
    """
    Setzt einen Wert in eine Zelle. Falls die Zelle in einem Merge liegt und nicht die
    Top-Left Zelle ist, wird die Top-Left Zelle gesetzt.
    """
    cell = ws.cell(row=row, column=col)
    for rng in ws.merged_cells.ranges:
        if (row, col) in rng:
            tl = ws.cell(rng.min_row, rng.min_col)
            tl.value = value
            return
    cell.value = value

def copy_sheet_to_workbook(src_ws, dst_wb, new_title: str):
    dst_ws = dst_wb.create_sheet(title=new_title)

    # Copy column widths
    for col_letter, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width = dim.width

    # Copy row heights
    for row_idx, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row_idx].height = dim.height

    # Copy merged cells
    for mr in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(mr))

    # Copy cells (values + styles)
    for row in src_ws.iter_rows():
        for cell in row:
            dcell = dst_ws.cell(row=cell.row, column=cell.col_idx, value=cell.value)
            if cell.has_style:
                dcell._style = copy(cell._style)
            dcell.number_format = cell.number_format
            dcell.font = copy(cell.font)
            dcell.border = copy(cell.border)
            dcell.fill = copy(cell.fill)
            dcell.alignment = copy(cell.alignment)
            dcell.protection = copy(cell.protection)
            dcell.comment = copy(cell.comment) if cell.comment else None

    return dst_ws

def find_period_text(ws) -> Optional[str]:
    """
    Zeitbezug steht in den Eingangstabellen "oberhalb des Tabellenkopfs".
    Für Tabellen 1/3: i.d.R. A3
    Für Tabelle 2: A4 oder A3, abhängig vom Umbruch (wir scannen A1..A8)
    Für Tabelle 5: je Blatt unterschiedlich, aber wir nutzen spezifische Regeln dort.
    """
    for r in range(1, 9):
        v = ws.cell(row=r, column=1).value
        if v is None:
            continue
        s = str(v).strip()
        if not s:
            continue
        # ausschließen: "Bayern", Überschriftzeilen, etc.
        if s.lower() == "bayern":
            continue
        # typischer Zeitbezug: "Dezember 2025", "4. Quartal 2025", "1. Halbjahr 2025", "2025"
        if re.search(r"(januar|februar|märz|maerz|april|mai|juni|juli|august|september|oktober|november|dezember)\s+\d{4}", s, re.I):
            return s
        if re.search(r"\d\.\s*quartal\s+\d{4}", s, re.I):
            return s
        if re.search(r"\d\.\s*halbjahr\s+\d{4}", s, re.I):
            return s
        if re.fullmatch(r"\d{4}", s):
            return s
    return None

def apply_int_format_with_space(ws, skip_cols: Optional[set] = None, allow_minus_space: bool = True):
    """
    Formatiert Ganzzahlen: Tausendertrennzeichen = Leerzeichen, keine Dezimalstellen.
    0 bleibt 0. "-" ignorieren. "X" ignorieren.
    Negative Werte sollen als "- 123" (Minus + Leerzeichen) dargestellt werden.
    => Wir lösen das, indem wir den Zellenwert als int belassen, aber das Anzeigeformat
       auf "# ##0;-# ##0" setzen und zusätzlich bei negativen Werten ein Custom-Format
       nutzen. Das "Minus + Leerzeichen" erzwingen wir via Zellformat nicht zuverlässig
       in allen Excel-Locales; daher setzen wir für negative Werte als Text "- {abs}".
       (Nur wenn Zelle wirklich numeric ist.)
    """
    if skip_cols is None:
        skip_cols = set()

    for row in ws.iter_rows():
        for cell in row:
            if cell.col_idx in skip_cols:
                continue

            v = cell.value
            if v is None:
                continue
            if isinstance(v, str):
                s = v.strip()
                if s in ("-", "X", ""):
                    continue
                # wenn es bereits Text ist -> nicht anfassen
                continue

            # numerisch
            if isinstance(v, (int, float)):
                # Prozentspalten werden woanders behandelt
                # Ganzzahl behandeln
                if isinstance(v, float):
                    # wenn float aber faktisch ganzzahlig
                    if abs(v - int(v)) < 1e-9:
                        v = int(v)
                        cell.value = v
                    else:
                        # echte Kommazahl -> nicht anfassen
                        continue

                if isinstance(v, int):
                    if v == 0:
                        cell.number_format = FMT_INT_SPACE
                        continue
                    if v < 0 and allow_minus_space:
                        cell.value = f"- {abs(v):,}".replace(",", " ")
                        cell.number_format = "@"  # Text
                    else:
                        cell.number_format = FMT_INT_SPACE

def apply_percent_format_one_decimal_except_zero(ws, col_idx: int):
    """
    Prozentwerte ohne % Zeichen: immer 1 Nachkommastelle, außer bei 0 => 0.
    Wenn Zelle "X" oder "-" enthält -> ignorieren.
    """
    for r in range(1, ws.max_row + 1):
        cell = ws.cell(row=r, column=col_idx)
        v = cell.value
        if v is None:
            continue
        if isinstance(v, str):
            s = v.strip()
            if s in ("-", "X", ""):
                continue
            # Text sonst ignorieren
            continue
        if isinstance(v, (int, float)):
            if abs(v) < 1e-12:
                cell.number_format = FMT_PCT_ZERO
            else:
                cell.number_format = FMT_PCT_1_DEC

# -------------------------------------
# Builders: Tabelle 1/2/3/5
# -------------------------------------
def build_table1_workbook(raw_path, layout_path, internal_layout: bool):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[1]] if RAW_SHEET_NAMES[1] in wb_raw.sheetnames else wb_raw[wb_raw.sheetnames[0]]

    wb_out = openpyxl.load_workbook(layout_path)
    ws_out = wb_out[wb_out.sheetnames[0]]

    period_text = find_period_text(ws_raw)

    # INTERNAL: Zeitbezug sitzt i.d.R. in Zeile 5 (A5)
    # _g: Zeitbezug sitzt i.d.R. in Zeile 3 (A3) außer Tabelle 2 Sonderfall
    if internal_layout:
        if period_text is not None:
            t = str(period_text).strip()
            if t.isdigit():
                t = f"Jahr {t}"
            set_value_merge_safe(ws_out, 5, 1, t)
    else:
        if period_text is not None:
            t = str(period_text).strip()
            if t.isdigit():
                t = f"Jahr {t}"
            set_value_merge_safe(ws_out, 3, 1, t)

    # Copy content block (großzügig)
    max_r = min(ws_out.max_row, ws_raw.max_row)
    max_c = min(ws_out.max_column, ws_raw.max_column)
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            # nicht Überschriften übersteuern: Layout bleibt, aber Datenbereich übernehmen
            # (hier pragmatisch: alles ab Zeile 14 und ab Spalte 4)
            if r >= 14 and c >= 4:
                set_value_merge_safe(ws_out, r, c, ws_raw.cell(row=r, column=c).value)

    # Zahlenformat (Spalte I ist Prozent? -> Tabelle 1: Spalte I = Prozentwerte)
    apply_int_format_with_space(ws_out, skip_cols={9})
    apply_percent_format_one_decimal_except_zero(ws_out, col_idx=9)

    return wb_out

def process_table1_file(raw_path: Path, layouts_dir: Path, out_dir: Path, logger: Logger):
    fn = raw_path.name
    period = detect_period_from_filename(fn)
    is_jj = "-JJ" in fn

    layout_int = layouts_dir / LAYOUT_FILES[(1, True)]
    layout_g = layouts_dir / LAYOUT_FILES[(1, False)]

    # INTERN
    wb_int = build_table1_workbook(str(raw_path), str(layout_int), internal_layout=True)
    out_int = out_dir / fn.replace(".xlsx", "_INTERN.xlsx")
    wb_int.save(out_int)
    logger.log(f"  -> Intern: {out_int}")

    # _g
    if is_jj:
        # Für JJ soll _g inhaltlich der Eingangsdatei entsprechen (inkl. Spalten J/K).
        # Am stabilsten: Eingangsdatei 1:1 übernehmen und nur (a) Zeitbezug als "Jahr XXXX"
        # setzen und (b) Markierung (Spalte G = 1/2) anwenden.
        wb_g = openpyxl.load_workbook(raw_path)  # Styles + Layout bleiben erhalten
        ws = wb_g[wb_g.sheetnames[0]]

        period_text = find_period_text(ws)  # steht in der Eingangsdatei oberhalb des Tabellenkopfs
        if period_text is not None:
            t = str(period_text).strip()
            period_out = f"Jahr {t}" if t.isdigit() else period_text
            set_value_merge_safe(ws, 3, 1, period_out)

        fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        # Markierung: alle Zeilen, deren Marker in Spalte G = 1 oder 2 ist
        for r in range(1, ws.max_row + 1):
            v = ws.cell(row=r, column=7).value
            if v in (1, 2, "1", "2"):
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).fill = fill

    else:
        wb_g = build_table1_workbook(str(raw_path), str(layout_g), internal_layout=False)
        # Markierung für Monat/Q/H? (Tabelle 1: Marker in Spalte G)
        fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ws = wb_g[wb_g.sheetnames[0]]
        for r in range(1, ws.max_row + 1):
            v = ws.cell(row=r, column=7).value
            if v in (1, 2, "1", "2"):
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).fill = fill

    out_g = out_dir / fn.replace(".xlsx", "_g.xlsx")
    wb_g.save(out_g)
    logger.log(f"  -> Extern: {out_g}")

def build_table2_or_3(raw_path, layout_path, internal_layout: bool, table_no: int):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    ws_raw = wb_raw[RAW_SHEET_NAMES[table_no]] if RAW_SHEET_NAMES[table_no] in wb_raw.sheetnames else wb_raw[wb_raw.sheetnames[0]]

    wb_out = openpyxl.load_workbook(layout_path)
    ws_out = wb_out[wb_out.sheetnames[0]]

    period_text = find_period_text(ws_raw)

    if table_no == 2:
        # Tabelle 2: Überschrift ist über 2 Zeilen (A2 + A3 im INTERN, A2 + A3 im _g)
        # Zeitbezug muss bei _g wie INTERN "eine Zeile tiefer" sitzen:
        # INTERN: Zeitbezug A4? In Mustern: A4 = "Dezember 2025"
        # _g: soll genauso sein wie INTERN, nur ohne "NUR FÜR DEN INTERNEN..."
        if internal_layout:
            if period_text is not None:
                t = str(period_text).strip()
                if t.isdigit():
                    t = f"Jahr {t}"
                set_value_merge_safe(ws_out, 4, 1, t)
        else:
            if period_text is not None:
                t = str(period_text).strip()
                if t.isdigit():
                    t = f"Jahr {t}"
                # _g: Zeitbezug in Zeile 4, nicht 3
                set_value_merge_safe(ws_out, 4, 1, t)
    else:
        # Tabelle 3: INTERN Zeitbezug i.d.R. Zeile 5 (A5), _g Zeile 3 (A3) oder 5? (im Muster _g Zeilen 1-3)
        if internal_layout:
            if period_text is not None:
                t = str(period_text).strip()
                if t.isdigit():
                    t = f"Jahr {t}"
                set_value_merge_safe(ws_out, 5, 1, t)
        else:
            if period_text is not None:
                t = str(period_text).strip()
                if t.isdigit():
                    t = f"Jahr {t}"
                set_value_merge_safe(ws_out, 3, 1, t)

    # Copy content block (Datenbereich)
    max_r = min(ws_out.max_row, ws_raw.max_row)
    max_c = min(ws_out.max_column, ws_raw.max_column)
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            if r >= 12 and c >= 4:
                set_value_merge_safe(ws_out, r, c, ws_raw.cell(row=r, column=c).value)

    # Zahlenformate
    # Tabelle 2: Spalte G = Prozent
    # Tabelle 3: Spalte G = Prozent
    if table_no == 2:
        apply_int_format_with_space(ws_out, skip_cols={7})
        apply_percent_format_one_decimal_except_zero(ws_out, col_idx=7)
    else:
        apply_int_format_with_space(ws_out, skip_cols={7})
        apply_percent_format_one_decimal_except_zero(ws_out, col_idx=7)

    return wb_out

def process_table2_or_3_file(raw_path: Path, layouts_dir: Path, out_dir: Path, logger: Logger, table_no: int):
    fn = raw_path.name
    is_jj = "-JJ" in fn

    layout_int = layouts_dir / LAYOUT_FILES[(table_no, True)]
    layout_g = layouts_dir / LAYOUT_FILES[(table_no, False)]

    wb_int = build_table2_or_3(str(raw_path), str(layout_int), internal_layout=True, table_no=table_no)
    out_int = out_dir / fn.replace(".xlsx", "_INTERN.xlsx")
    wb_int.save(out_int)
    logger.log(f"  -> Intern: {out_int}")

    wb_g = build_table2_or_3(str(raw_path), str(layout_g), internal_layout=False, table_no=table_no)
    out_g = out_dir / fn.replace(".xlsx", "_g.xlsx")
    wb_g.save(out_g)
    logger.log(f"  -> Extern: {out_g}")

def fix_table5_stand_position(ws, logger: Logger):
    """
    Blatt 5.5: "Stand: xxxx" muss in derselben Zeile wie Copyright
    unter der letzten Tabellenspalte stehen.
    In den _g Layouts war es manchmal zu weit rechts.
    Wir suchen die Zeile mit "Copyright" und setzen "Stand:" an die letzte Spalte der Tabelle.
    """
    target_row = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v and "copyright" in str(v).lower():
            target_row = r
            break

    if target_row is None:
        return

    # Stand-Text suchen (irgendwo in der Zeile)
    stand_text = None
    stand_col = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=target_row, column=c).value
        if v and "stand" in str(v).lower():
            stand_text = v
            stand_col = c
            break

    if stand_text is None:
        # vielleicht steht Stand in der Zeile drüber?
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=target_row - 0, column=c).value
            if v and "stand" in str(v).lower():
                stand_text = v
                stand_col = c
                break

    if stand_text is None:
        return

    # Letzte "Tabellenspalte" heuristisch: letzte nicht-leere Spalte in Kopfzeile 6..10?
    # Wir nehmen die letzte Spalte, in der in Zeile 6..10 irgendein Wert steht.
    last_col = 1
    for c in range(1, ws.max_column + 1):
        for rr in range(6, 11):
            if rr <= ws.max_row:
                if ws.cell(row=rr, column=c).value not in (None, ""):
                    last_col = max(last_col, c)

    if stand_col and stand_col != last_col:
        # Move stand text
        ws.cell(row=target_row, column=stand_col).value = None
        ws.cell(row=target_row, column=last_col).value = stand_text
        logger.log(f"    [FIX] Stand verschoben: Spalte {get_column_letter(stand_col)} -> {get_column_letter(last_col)} (Zeile {target_row})")

def build_table5_workbook(raw_path, template_path, internal_layout: bool, is_jj: bool):
    wb_raw = openpyxl.load_workbook(raw_path, data_only=True)
    wb_tpl = openpyxl.load_workbook(template_path)

    # Output = Template-Kopie, danach Werte aus raw eintragen
    wb_out = wb_tpl

    # Zeitbezug aus Eingangsdatei Blatt 1 (oder aus einem Blatt, wo er steht)
    ws_raw_first = wb_raw[wb_raw.sheetnames[0]]
    period_text = find_period_text(ws_raw_first)
    if period_text is not None:
        t = str(period_text).strip()
        if t.isdigit():
            t = f"Jahr {t}"

    # Tabelle 5: mehrere Blätter, alle 1:1 Wertebereiche übernehmen
    for idx, tpl_ws in enumerate(wb_out.worksheets):
        if idx >= len(wb_raw.worksheets):
            break
        raw_ws = wb_raw.worksheets[idx]

        # Zeitbezug in Tabelle 5 sitzt i.d.R. in A4 (bei Monat/Quartal/Halbjahr) und bei JJ als "Jahr 2025"
        # In der Layoutdatei sitzt er schon; wir überschreiben gezielt, wenn wir was gefunden haben.
        if period_text is not None:
            if is_jj:
                set_value_merge_safe(tpl_ws, 4, 1, f"Jahr {str(period_text).strip()}" if str(period_text).strip().isdigit() else t)
            else:
                set_value_merge_safe(tpl_ws, 4, 1, t)

        max_r = min(tpl_ws.max_row, raw_ws.max_row)
        max_c = min(tpl_ws.max_column, raw_ws.max_column)
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                # Datenbereich: ab Zeile 12 (bei Tab5 meist)
                if r >= 12 and c >= 4:
                    set_value_merge_safe(tpl_ws, r, c, raw_ws.cell(row=r, column=c).value)

        # Prozentspalte Tabelle 5: Spalte H ist Prozent (ohne % Zeichen)
        apply_int_format_with_space(tpl_ws, skip_cols={8})
        apply_percent_format_one_decimal_except_zero(tpl_ws, col_idx=8)

    return wb_out

def process_table5_file(raw_path: Path, layouts_dir: Path, out_dir: Path, logger: Logger):
    fn = raw_path.name
    is_jj = "-JJ" in fn

    layout_int = layouts_dir / LAYOUT_FILES[(5, True)]
    layout_g = layouts_dir / LAYOUT_FILES[(5, False)]

    wb_int = build_table5_workbook(str(raw_path), str(layout_int), internal_layout=True, is_jj=is_jj)
    out_int = out_dir / fn.replace(".xlsx", "_INTERN.xlsx")
    wb_int.save(out_int)
    logger.log(f"  -> Intern: {out_int}")

    wb_g = build_table5_workbook(str(raw_path), str(layout_g), internal_layout=False, is_jj=is_jj)

    # Fix: Stand-Position auf Blatt 5.5 bei _g (Monat/Q/H)
    try:
        if not is_jj and len(wb_g.worksheets) >= 5:
            fix_table5_stand_position(wb_g.worksheets[4], logger)
    except Exception:
        logger.log("    [WARN] Stand-Fix nicht anwendbar (ignoriert).")

    out_g = out_dir / fn.replace(".xlsx", "_g.xlsx")
    wb_g.save(out_g)
    logger.log(f"  -> Extern: {out_g}")

# -------------------------------------
# Runner per input folder
# -------------------------------------
def process_input_folder(input_dir: Path, layouts_dir: Path, out_base_dir: Path, logger: Logger):
    input_dir = Path(input_dir)
    layouts_dir = Path(layouts_dir)
    out_base_dir = Path(out_base_dir)

    out_root = out_base_dir / DEFAULT_VOE_SUBDIR / input_dir.name
    ensure_dir(out_root)

    logger.log(f"--- Eingang: {input_dir}")
    logger.log(f"--- Ausgabe: {out_root}")

    files = sorted([p for p in input_dir.iterdir() if p.is_file() and is_relevant_file(p.name)])
    if not files:
        logger.log("[SCAN] 0 relevante Dateien gefunden.")
        return

    logger.log(f"[SCAN] {len(files)} Dateien gefunden (relevant).")

    for f in files:
        tno = parse_table_no(f.name)
        if tno not in (1, 2, 3, 5):
            continue

        logger.log(f"[START] {f.name}")
        if tno == 1:
            process_table1_file(f, layouts_dir, out_root, logger)
        elif tno == 2:
            process_table2_or_3_file(f, layouts_dir, out_root, logger, table_no=2)
        elif tno == 3:
            process_table2_or_3_file(f, layouts_dir, out_root, logger, table_no=3)
        elif tno == 5:
            process_table5_file(f, layouts_dir, out_root, logger)

# -------------------------------------
# GUI
# -------------------------------------
@dataclass
class JobConfig:
    month_dir: str = ""
    quarter_dir: str = ""
    half_dir: str = ""
    year_dir: str = ""
    out_base: str = ""
    log_dir: str = ""  # optional
    layouts_dir: str = ""  # optional (falls leer -> <exe>\Layouts)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x360")

        self.cfg = JobConfig()

        self._build()

    def _build(self):
        pad = {"padx": 8, "pady": 4}

        def row(y, label, attr, required=False):
            tk.Label(self, text=label + (" (Pflicht)" if required else "")).grid(row=y, column=0, sticky="w", **pad)
            e = tk.Entry(self, width=90)
            e.grid(row=y, column=1, sticky="we", **pad)
            b = tk.Button(self, text="Auswählen...", command=lambda: self._pick_dir(e, attr))
            b.grid(row=y, column=2, sticky="e", **pad)
            return e

        self.e_month = row(0, "Monat – Eingangstabellen (optional):", "month_dir")
        self.e_quarter = row(1, "Quartal – Eingangstabellen (optional):", "quarter_dir")
        self.e_half = row(2, "Halbjahr – Eingangstabellen (optional):", "half_dir")
        self.e_year = row(3, "Jahr – Eingangstabellen (optional):", "year_dir")
        self.e_out = row(4, "Ausgabe-Basisordner:", "out_base", required=True)
        self.e_log = row(5, "Protokollordner (optional, leer = .\\Protokolle):", "log_dir")
        self.e_layouts = row(6, "Layouts-Ordner (optional, leer = .\\Layouts):", "layouts_dir")

        self.status = tk.Label(self, text="", fg="black")
        self.status.grid(row=7, column=0, columnspan=3, sticky="w", padx=8, pady=6)

        btn_frame = tk.Frame(self)
        btn_frame.grid(row=8, column=0, columnspan=3, sticky="e", padx=8, pady=8)

        tk.Button(btn_frame, text="Start", width=12, command=self._run).pack(side="right", padx=6)
        tk.Button(btn_frame, text="Schließen", width=12, command=self.destroy).pack(side="right", padx=6)

        note = tk.Label(
            self,
            text="Hinweis: Layouts werden aus .\\Layouts geladen. Ausgaben gehen nach <Basis>\\VÖ-Tabellen\\<Eingangsordnername>.",
            fg="gray"
        )
        note.grid(row=9, column=0, columnspan=3, sticky="w", padx=8, pady=2)

    def _pick_dir(self, entry: tk.Entry, attr: str):
        d = filedialog.askdirectory()
        if not d:
            return
        entry.delete(0, tk.END)
        entry.insert(0, d)
        setattr(self.cfg, attr, d)

    def _collect_cfg(self):
        self.cfg.month_dir = self.e_month.get().strip()
        self.cfg.quarter_dir = self.e_quarter.get().strip()
        self.cfg.half_dir = self.e_half.get().strip()
        self.cfg.year_dir = self.e_year.get().strip()
        self.cfg.out_base = self.e_out.get().strip()
        self.cfg.log_dir = self.e_log.get().strip()
        self.cfg.layouts_dir = self.e_layouts.get().strip()

    def _run(self):
        self._collect_cfg()
        try:
            if not self.cfg.out_base:
                messagebox.showerror("Fehler", "Ausgabe-Basisordner ist Pflicht.")
                return

            exe_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
            base_dir = Path(self.cfg.out_base).resolve()

            layouts_dir = Path(self.cfg.layouts_dir).resolve() if self.cfg.layouts_dir else (Path(__file__).resolve().parent / DEFAULT_LAYOUTS_SUBDIR)
            log_dir = Path(self.cfg.log_dir).resolve() if self.cfg.log_dir else (base_dir / DEFAULT_LOG_SUBDIR)

            logger = Logger(log_dir=log_dir, base_dir=base_dir)
            self.status.configure(text=f"Protokoll wird geschrieben nach: {logger.path}", fg="black")

            logger.log("=== START GUI-Lauf ===")
            logger.log(f"Ausgabe-Basis: {base_dir}")
            logger.log(f"Layouts:       {layouts_dir}")
            logger.log("Überschreiben: JA (vorhandene Dateien werden ersetzt)")

            jobs = [
                ("Monat", self.cfg.month_dir),
                ("Quartal", self.cfg.quarter_dir),
                ("Halbjahr", self.cfg.half_dir),
                ("Jahr", self.cfg.year_dir),
            ]

            logger.log("Jobs: " + ", ".join([f"{n}={p}" for n, p in jobs]))

            for name, p in jobs:
                if not p:
                    continue
                in_dir = Path(p)
                if not in_dir.exists():
                    logger.log(f"[WARN] {name}-Pfad existiert nicht: {in_dir}")
                    continue
                logger.log(f"== JOB: {name} ==")
                try:
                    process_input_folder(in_dir, Path(layouts_dir), base_dir, logger)
                except Exception as e:
                    logger.log(f"[FEHLER] Job {name} abgebrochen: {repr(e)}")
                    logger.log(traceback.format_exc())
                    raise

            logger.log("[FERTIG] Verarbeitung abgeschlossen.")
            logger.close()
            messagebox.showinfo("Fertig", "Verarbeitung abgeschlossen.\n\nProtokoll:\n" + str(logger.path))
            self.status.configure(text=f"Fertig. Protokoll: {logger.path}", fg="green")

        except Exception as e:
            try:
                self.status.configure(text="Fehler – siehe Protokoll.", fg="red")
            except Exception:
                pass
            messagebox.showerror("Fehler", f"Es ist ein Fehler aufgetreten:\n{e}\n\nBitte Protokoll prüfen.")

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
