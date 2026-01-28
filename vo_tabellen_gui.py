# vo_tabellen_gui.py
# GUI + Formatierung Tabellen 1/2/3/5 (INTERN & _g) inkl. Output-Struktur + Protokollierung
# Fixes:
# - kein fester L:-Pfad mehr (Protokolle: GUI oder Fallback ./Protokolle)
# - Zeitbezug robust aus Eingangsdatei (inkl. Jahr-only "2025")
# - Tabelle 2 (_g): Zeitbezug eine Zeile tiefer (Überschrift-Zeile 3 bleibt)
# - Tabelle 5 (_g): "Stand:" unter letzte sinnvolle Spalte (nicht max_column)

import os
import re
import sys
import traceback
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

import openpyxl
from openpyxl.utils import get_column_letter


# ============================================================
# Defaults / Konstanten
# ============================================================

# Wird im GUI gesetzt; wenn leer -> Logger nutzt ./Protokolle
PROTOKOLL_DIR = ""  # wird im GUI gewählt; Fallback: ./Protokolle

GER_MONTHS = [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember"
]
PERIOD_TOKENS = ["Quartal", "Halbjahr", "Jahr"]


# ============================================================
# Logging
# ============================================================

class Logger:
    def __init__(self):
        # PROTOKOLL_DIR kommt normalerweise aus dem GUI.
        # Wenn leer, loggen wir lokal neben der EXE (Unterordner 'Protokolle').
        log_dir = PROTOKOLL_DIR.strip() if isinstance(PROTOKOLL_DIR, str) else ""
        if not log_dir:
            log_dir = os.path.join(os.getcwd(), "Protokolle")
        os.makedirs(log_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        self.path = os.path.join(log_dir, f"vo_tabellen_{ts}.log")

    def log(self, msg: str):
        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{stamp}] {msg}"
        print(line)
        with open(self.path, "a", encoding="utf-8") as f:
            f.write(line + "\n")


# ============================================================
# Excel Helper
# ============================================================

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


def normalize_number_cell(v):
    """Ganzzahlen mit Leerzeichen als Tausendertrennzeichen, '-'/'X' ignorieren,
    Minuszeichen mit Leerzeichen: '- 25', ohne Dezimalstellen.
    Prozentspalten werden separat behandelt.
    """
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if s in ("-", "X", ""):
            return s  # 그대로
        # falls bereits mit Leerzeichen formatiert -> lassen
        # ansonsten versuchen als Zahl zu parsen
        s_clean = s.replace(" ", "").replace(".", "").replace(",", ".")
        try:
            num = float(s_clean)
        except Exception:
            return v
        # Ganzzahl
        if abs(num - int(num)) < 1e-9:
            n = int(num)
        else:
            n = int(round(num))
        if n < 0:
            return f"- {format(abs(n), ',').replace(',', ' ')}"
        return format(n, ",").replace(",", " ")
    if isinstance(v, (int, float)):
        n = int(round(v))
        if n < 0:
            return f"- {format(abs(n), ',').replace(',', ' ')}"
        return format(n, ",").replace(",", " ")
    return v


def normalize_percent_cell(v):
    """Prozentwerte: immer 1 Nachkommastelle, außer bei 0 => '0' (ohne Prozentzeichen).
    '-'/'X' bleiben wie sie sind.
    """
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if s in ("-", "X", ""):
            return s
        s_clean = s.replace(" ", "").replace(".", "").replace(",", ".")
        try:
            num = float(s_clean)
        except Exception:
            return v
    elif isinstance(v, (int, float)):
        num = float(v)
    else:
        return v

    # 0 -> "0"
    if abs(num) < 1e-12:
        return "0"

    # 1 Nachkommastelle, Minus mit Leerzeichen
    sign = "-" if num < 0 else ""
    num_abs = abs(num)
    txt = f"{num_abs:.1f}".replace(".", ",")
    if sign:
        return f"- {txt}"
    return txt


def copy_cell_style(src_cell, dst_cell):
    dst_cell._style = src_cell._style
    dst_cell.font = src_cell.font
    dst_cell.border = src_cell.border
    dst_cell.fill = src_cell.fill
    dst_cell.number_format = src_cell.number_format
    dst_cell.protection = src_cell.protection
    dst_cell.alignment = src_cell.alignment


def copy_sheet_values_and_styles(ws_src, ws_dst, min_row=1, max_row=None, min_col=1, max_col=None):
    if max_row is None:
        max_row = ws_src.max_row
    if max_col is None:
        max_col = ws_src.max_column

    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            sc = ws_src.cell(row=r, column=c)
            dc = ws_dst.cell(row=r, column=c)
            dc.value = sc.value
            copy_cell_style(sc, dc)


def find_period_text(ws, search_rows=30):
    """Liest den Berichtszeitraum aus der Eingangsdatei.
    Robust gegen unterschiedliche Tabellen (1/2/3/5) und Perioden (Monat/Quartal/Halbjahr/Jahr).
    Rückgabe z.B. 'Dezember 2025', '4. Quartal 2025', '1. Halbjahr 2025' oder '2025'.
    """
    hits = []
    max_r = min(search_rows, ws.max_row)
    # Zeitraum steht je nach Tabelle nicht immer in Spalte A -> wir scannen A..E
    for r in range(1, max_r + 1):
        for c in range(1, 6):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            s = v.strip()
            if not s:
                continue
            # Jahr-only (z.B. '2025')
            if re.fullmatch(r"20\d{2}", s):
                hits.append(s)
                continue
            # Sonst: muss ein Jahr enthalten und Monat/Token
            if not re.search(r"20\d{2}", s):
                continue
            if any(m in s for m in GER_MONTHS) or any(tok in s for tok in PERIOD_TOKENS):
                hits.append(s)
    return hits[-1] if hits else None


def last_used_col(ws, scan_rows=80):
    """Ermittelt die letzte 'sinnvolle' Spalte (gegen sehr große max_column-Werte durch Styles)."""
    max_r = min(ws.max_row, scan_rows)
    last = 1
    for r in range(1, max_r + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v not in (None, ""):
                if c > last:
                    last = c
    return last


def update_footer_with_stand_and_copyright(ws, stand_text: str, copyright_text: str):
    """Setzt Copyright + Stand in die letzte Zeile (oder vorhandene Zeile)"""
    # Suche nach Copyrightzeile
    max_r = min(ws.max_row, 200)
    copyright_row = None
    stand_col = None

    for r in range(1, max_r + 1):
        for c in range(1, min(ws.max_column, 50) + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "©" in v:
                copyright_row = r
            if isinstance(v, str) and "Stand:" in v:
                copyright_row = r
                stand_col = c

    if copyright_row is None:
        # fallback: letzte belegte Zeile + 1
        copyright_row = min(ws.max_row + 1, 500)

    # Copyright links
    ws.cell(row=copyright_row, column=1).value = copyright_text

    # Stand rechts -> aber "sinnvolle" letzte Spalte
    if stand_col is None:
        stand_col = last_used_col(ws)
    else:
        stand_col = min(stand_col, last_used_col(ws))

    ws.cell(row=copyright_row, column=stand_col).value = stand_text


# ============================================================
# Datei-/Perioden-Erkennung
# ============================================================

def detect_period_from_filename(fn: str):
    # Beispiele: Tabelle-1-Land_2025-12.xlsx / Tabelle-1-Land_2025-Q4.xlsx / Tabelle-1-Land_2025-H2.xlsx / Tabelle-1-Land_2025-JJ.xlsx
    m = re.search(r"_(20\d{2}-(?:\d{2}|Q\d|H\d|JJ))", fn)
    if m:
        return m.group(1)
    return None


def ensure_dir(p):
    os.makedirs(p, exist_ok=True)


# ============================================================
# Layout-Dateien: Name->Pfad
# ============================================================

def resolve_layout_paths(layout_dir: str):
    # erwartet dort z.B.:
    # Tabelle-1-Layout_INTERN.xlsx, Tabelle-1-Layout_g.xlsx, ...
    needed = {}
    for f in os.listdir(layout_dir):
        if not f.lower().endswith(".xlsx"):
            continue
        lf = f.lower()
        if "tabelle-1" in lf and "layout" in lf:
            needed[("1", "intern" if "intern" in lf else "g")] = os.path.join(layout_dir, f)
        if "tabelle-2" in lf and "layout" in lf:
            needed[("2", "intern" if "intern" in lf else "g")] = os.path.join(layout_dir, f)
        if "tabelle-3" in lf and "layout" in lf:
            needed[("3", "intern" if "intern" in lf else "g")] = os.path.join(layout_dir, f)
        if "tabelle-5" in lf and "layout" in lf:
            needed[("5", "intern" if "intern" in lf else "g")] = os.path.join(layout_dir, f)
    return needed


# ============================================================
# Verarbeitung Tabellen 1/2/3/5
# ============================================================

def process_table1(raw_path, tpl_path, out_path, internal_layout: bool, logger: Logger):
    logger.log(f"[T1] raw={raw_path}")
    wb_raw = openpyxl.load_workbook(raw_path)
    ws_raw = wb_raw.active

    wb_out = openpyxl.load_workbook(tpl_path)
    ws_out = wb_out.active

    title_text = ws_raw.cell(row=2, column=1).value
    period_text = find_period_text(ws_raw)

    # INTERN-Headerzeile
    if internal_layout:
        ws_out.cell(row=1, column=1).value = "Bayern"

    # Titel + Zeitraum (Zeilen wie Layout; wir überschreiben nur Inhalt)
    # Tabelle 1: Titel in Zeile 2, Zeitraum irgendwo in Kopfbereich -> wir setzen ihn in die Zeile,
    # in der bereits ein Zeitraum steht (aus dem Layout) – ansonsten auf Zeile 5.
    ws_out.cell(row=2, column=1).value = title_text

    if period_text:
        # finde alte Zeitraum-Zelle im Layout (typisch Zeile 3/5)
        placed = False
        for r in range(1, 15):
            v = ws_out.cell(row=r, column=1).value
            if isinstance(v, str) and (any(m in v for m in GER_MONTHS) or any(tok in v for tok in PERIOD_TOKENS) or re.fullmatch(r"20\d{2}", v.strip())):
                ws_out.cell(row=r, column=1).value = period_text if not re.fullmatch(r"20\d{2}", period_text.strip()) else f"Jahr {period_text.strip()}"
                placed = True
                break
        if not placed:
            ws_out.cell(row=5, column=1).value = period_text if not re.fullmatch(r"20\d{2}", period_text.strip()) else f"Jahr {period_text.strip()}"

    # Datenbereich: alles ab Zeile 13 (wie bisher) übernehmen
    # (die genaue Range stammt aus eurer bisherigen Routine; ggf. anpassen)
    for r in range(13, ws_raw.max_row + 1):
        for c in range(1, ws_raw.max_column + 1):
            v = ws_raw.cell(row=r, column=c).value
            if is_numeric_like(v):
                ws_out.cell(row=r, column=c).value = normalize_number_cell(v)
            else:
                ws_out.cell(row=r, column=c).value = v

    wb_out.save(out_path)
    logger.log(f"[T1] geschrieben: {out_path}")


def process_table2_or_3(raw_path, tpl_path, out_path, internal_layout: bool, table_no: int, logger: Logger):
    logger.log(f"[T{table_no}] raw={raw_path}")
    wb_raw = openpyxl.load_workbook(raw_path)
    ws_raw = wb_raw.active

    wb_out = openpyxl.load_workbook(tpl_path)
    ws_out = wb_out.active

    title_text = ws_raw.cell(row=2, column=1).value
    period_text = find_period_text(ws_raw)
    if period_text and re.fullmatch(r"20\d{2}", period_text.strip()):
        period_text = f"Jahr {period_text.strip()}"

    # INTERN-Kopf
    if internal_layout:
        ws_out.cell(row=1, column=1).value = "Bayern"
        # Titel aus raw: Zeile 2 + ggf. Fortsetzung (bei Tabelle 2)
        ws_out.cell(row=2, column=1).value = title_text
        if table_no == 2:
            ws_out.cell(row=3, column=1).value = ws_raw.cell(row=3, column=1).value
        # Zeitraum: im Layout vorhandene Zeitraum-Zelle überschreiben (sonst fallback)
        if period_text:
            placed = False
            for r in range(1, 20):
                v = ws_out.cell(row=r, column=1).value
                if isinstance(v, str) and (any(m in v for m in GER_MONTHS) or any(tok in v for tok in PERIOD_TOKENS) or re.fullmatch(r"Jahr\s+20\d{2}", v.strip())):
                    ws_out.cell(row=r, column=1).value = period_text
                    placed = True
                    break
            if not placed:
                ws_out.cell(row=6, column=1).value = period_text

    else:
        if table_no == 2:
            # Tabelle 2: zweizeilige Überschrift, Zeitraum eine Zeile tiefer (wie _INTERN)
            ws_out.cell(row=1, column=1).value = "Bayern"
            ws_out.cell(row=2, column=1).value = title_text
            ws_out.cell(row=3, column=1).value = ws_raw.cell(row=3, column=1).value
            ws_out.cell(row=4, column=1).value = period_text
        else:
            ws_out.cell(row=1, column=1).value = "Bayern"
            ws_out.cell(row=2, column=1).value = title_text
            ws_out.cell(row=3, column=1).value = period_text

    # Prozentspalten:
    # Tabelle 2: Spalte G
    # Tabelle 3: Spalte G
    percent_col = 7

    # Datenbereich: ab Zeile 12 (wie in euren bisherigen Routinen)
    for r in range(12, ws_raw.max_row + 1):
        for c in range(1, ws_raw.max_column + 1):
            v = ws_raw.cell(row=r, column=c).value
            if c == percent_col:
                ws_out.cell(row=r, column=c).value = normalize_percent_cell(v) if is_numeric_like(v) else v
            else:
                ws_out.cell(row=r, column=c).value = normalize_number_cell(v) if is_numeric_like(v) else v

    wb_out.save(out_path)
    logger.log(f"[T{table_no}] geschrieben: {out_path}")


def build_table5_workbook(raw_path, tpl_path, out_path, internal_layout: bool, logger: Logger):
    logger.log(f"[T5] raw={raw_path}")
    wb_raw = openpyxl.load_workbook(raw_path)
    wb_out = openpyxl.load_workbook(tpl_path)

    # Zeitraum aus Eingangsdatei (steht im Blatt 1 zuverlässig)
    period_text = find_period_text(wb_raw.worksheets[0])
    if period_text and re.fullmatch(r"20\d{2}", period_text.strip()):
        period_text = f"Jahr {period_text.strip()}"

    # Es gibt mehrere Blätter (5.1..5.5). Wir übernehmen Werte passend zur Layoutstruktur.
    for i, ws_out in enumerate(wb_out.worksheets):
        ws_raw = wb_raw.worksheets[i] if i < len(wb_raw.worksheets) else None
        if ws_raw is None:
            continue

        # Kopf Bayern + Titel
        if internal_layout:
            ws_out.cell(row=1, column=1).value = "Bayern"
        else:
            ws_out.cell(row=1, column=1).value = "Bayern"

        title = ws_raw.cell(row=2, column=1).value
        ws_out.cell(row=2, column=1).value = title

        # Zeitraum in vorhandene Zeitraum-Zelle im Layout schreiben (sonst fallback Zeile 3)
        if period_text:
            placed = False
            for r in range(1, 20):
                v = ws_out.cell(row=r, column=1).value
                if isinstance(v, str) and (any(m in v for m in GER_MONTHS) or any(tok in v for tok in PERIOD_TOKENS) or re.fullmatch(r"Jahr\s+20\d{2}", v.strip())):
                    ws_out.cell(row=r, column=1).value = period_text
                    placed = True
                    break
            if not placed:
                ws_out.cell(row=3, column=1).value = period_text

        # Daten: ab Zeile 12 (wie bei euch)
        for r in range(12, ws_raw.max_row + 1):
            for c in range(1, ws_raw.max_column + 1):
                v = ws_raw.cell(row=r, column=c).value
                ws_out.cell(row=r, column=c).value = normalize_number_cell(v) if is_numeric_like(v) else v

        # Footer/Stand:
        # "Stand:" aus raw (wenn vorhanden) -> sonst leer
        stand_text = None
        for r in range(1, min(ws_raw.max_row, 120) + 1):
            for c in range(1, min(ws_raw.max_column, 40) + 1):
                vv = ws_raw.cell(row=r, column=c).value
                if isinstance(vv, str) and "Stand:" in vv:
                    stand_text = vv.strip()
        if not stand_text:
            # fallback: Stand: <Zeitraum>
            if period_text:
                stand_text = f"Stand: {period_text}"
            else:
                stand_text = "Stand:"

        copyright_text = "© Statistisches Amt"

        update_footer_with_stand_and_copyright(ws_out, stand_text=stand_text, copyright_text=copyright_text)

    wb_out.save(out_path)
    logger.log(f"[T5] geschrieben: {out_path}")


# ============================================================
# Hauptlauf pro Eingangsordner (Monat/Quartal/Halbjahr/Jahr)
# ============================================================

def run_for_input_folder(input_dir, out_base_dir, layout_dir, period_label, logger: Logger):
    """
    input_dir: z.B. ...\\Arbeitstabellen\\1 Monat_12-2025
    out_base_dir: z.B. ...\\2025-12  (VÖ-Tabellen wird darunter angelegt)
    layout_dir: Layouts-Ordner
    period_label: Name des Eingangsordners (für Ausgabe-Unterordner)
    """
    if not input_dir or not os.path.isdir(input_dir):
        return

    layouts = resolve_layout_paths(layout_dir)

    vo_root = os.path.join(out_base_dir, "VÖ-Tabellen")
    ensure_dir(vo_root)
    out_dir = os.path.join(vo_root, period_label)
    ensure_dir(out_dir)

    logger.log(f"[IN] {period_label}: {input_dir}")
    logger.log(f"[OUT] {out_dir}")

    # Dateien suchen
    files = {1: None, 2: None, 3: None, 5: None}
    for f in os.listdir(input_dir):
        if not f.lower().endswith(".xlsx"):
            continue
        lf = f.lower()
        if lf.startswith("tabelle-1-land") and "_g" not in lf and "_intern" not in lf:
            files[1] = os.path.join(input_dir, f)
        elif lf.startswith("tabelle-2-land") and "_g" not in lf and "_intern" not in lf:
            files[2] = os.path.join(input_dir, f)
        elif lf.startswith("tabelle-3-land") and "_g" not in lf and "_intern" not in lf:
            files[3] = os.path.join(input_dir, f)
        elif lf.startswith("tabelle-5-land") and "_g" not in lf and "_intern" not in lf:
            files[5] = os.path.join(input_dir, f)

    # Periode aus Dateiname ableiten (anhand Tabelle 1 wenn möglich)
    period_token = None
    if files[1]:
        period_token = detect_period_from_filename(os.path.basename(files[1]))
    if not period_token:
        # fallback: irgendeine
        for k in (2, 3, 5):
            if files[k]:
                period_token = detect_period_from_filename(os.path.basename(files[k]))
                break
    if not period_token:
        period_token = "UNBEKANNT"

    logger.log(f"[PERIODE] {period_token}")

    # Ausgabe-Dateinamen
    def out_name(table_no, suffix):
        return os.path.join(out_dir, f"Tabelle-{table_no}-Land_{period_token}{suffix}.xlsx")

    # Tabelle 1
    if files[1]:
        tpl_intern = layouts.get(("1", "intern"))
        tpl_g = layouts.get(("1", "g"))
        if tpl_intern:
            process_table1(files[1], tpl_intern, out_name(1, "_INTERN"), True, logger)
        if tpl_g:
            process_table1(files[1], tpl_g, out_name(1, "_g"), False, logger)

    # Tabelle 2
    if files[2]:
        tpl_intern = layouts.get(("2", "intern"))
        tpl_g = layouts.get(("2", "g"))
        if tpl_intern:
            process_table2_or_3(files[2], tpl_intern, out_name(2, "_INTERN"), True, 2, logger)
        if tpl_g:
            process_table2_or_3(files[2], tpl_g, out_name(2, "_g"), False, 2, logger)

    # Tabelle 3
    if files[3]:
        tpl_intern = layouts.get(("3", "intern"))
        tpl_g = layouts.get(("3", "g"))
        if tpl_intern:
            process_table2_or_3(files[3], tpl_intern, out_name(3, "_INTERN"), True, 3, logger)
        if tpl_g:
            process_table2_or_3(files[3], tpl_g, out_name(3, "_g"), False, 3, logger)

    # Tabelle 5
    if files[5]:
        tpl_intern = layouts.get(("5", "intern"))
        tpl_g = layouts.get(("5", "g"))
        if tpl_intern:
            build_table5_workbook(files[5], tpl_intern, out_name(5, "_INTERN"), True, logger)
        if tpl_g:
            build_table5_workbook(files[5], tpl_g, out_name(5, "_g"), False, logger)

    logger.log("[FERTIG] Ordner verarbeitet.")


# ============================================================
# GUI
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("VO-Tabellen-Formatter")
        self.geometry("900x520")

        self.layout_dir = tk.StringVar()
        self.out_base_dir = tk.StringVar()
        self.in_month = tk.StringVar()
        self.in_quarter = tk.StringVar()
        self.in_half = tk.StringVar()
        self.in_year = tk.StringVar()
        self.log_dir = tk.StringVar()

        row = 0

        def add_path_row(label, var, choose_dir=True):
            nonlocal row
            tk.Label(self, text=label, anchor="w").grid(row=row, column=0, sticky="ew", padx=8, pady=4)
            tk.Entry(self, textvariable=var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
            btn = tk.Button(self, text="Auswählen...", command=lambda: self.pick_path(var, choose_dir))
            btn.grid(row=row, column=2, padx=8, pady=4)
            row += 1

        self.columnconfigure(1, weight=1)

        add_path_row("Layouts-Ordner (enthält Layouts/*.xlsx):", self.layout_dir, True)
        add_path_row("Ausgabe-Basisordner (VÖ-Tabellen wird darunter angelegt):", self.out_base_dir, True)

        tk.Label(self, text="Eingangsordner (nur gefüllte werden verarbeitet):", anchor="w").grid(row=row, column=0, sticky="ew", padx=8, pady=(12, 4))
        row += 1

        add_path_row("Monat:", self.in_month, True)
        add_path_row("Quartal:", self.in_quarter, True)
        add_path_row("Halbjahr:", self.in_half, True)
        add_path_row("Jahr:", self.in_year, True)

        add_path_row("Protokoll-Ordner:", self.log_dir, True)

        frm = tk.Frame(self)
        frm.grid(row=row, column=0, columnspan=3, sticky="ew", padx=8, pady=10)
        tk.Button(frm, text="Start", command=self.start_run).pack(side="left", padx=5)
        tk.Button(frm, text="Beenden", command=self.destroy).pack(side="left", padx=5)
        row += 1

        self.txt = tk.Text(self, height=12)
        self.txt.grid(row=row, column=0, columnspan=3, sticky="nsew", padx=8, pady=8)
        self.rowconfigure(row, weight=1)

        self.write("[INFO] GUI bereit.\n")

    def pick_path(self, var, choose_dir=True):
        if choose_dir:
            p = filedialog.askdirectory()
        else:
            p = filedialog.askopenfilename()
        if p:
            var.set(p)

    def write(self, s):
        self.txt.insert("end", s)
        self.txt.see("end")
        self.update_idletasks()

    def start_run(self):
        global PROTOKOLL_DIR
        PROTOKOLL_DIR = self.log_dir.get().strip()

        logger = Logger()

        try:
            layout_dir = self.layout_dir.get().strip()
            out_base = self.out_base_dir.get().strip()

            if not layout_dir or not os.path.isdir(layout_dir):
                messagebox.showerror("Fehler", "Bitte Layouts-Ordner auswählen.")
                return
            if not out_base or not os.path.isdir(out_base):
                messagebox.showerror("Fehler", "Bitte Ausgabe-Basisordner auswählen.")
                return

            # Protokollordner anlegen (wie gefordert)
            if PROTOKOLL_DIR:
                ensure_dir(PROTOKOLL_DIR)

            logger.log("Starte VO-Tabellen-Formatter...")

            # Monat
            if self.in_month.get().strip():
                run_for_input_folder(self.in_month.get().strip(), out_base, layout_dir,
                                     os.path.basename(self.in_month.get().strip()), logger)
            # Quartal
            if self.in_quarter.get().strip():
                run_for_input_folder(self.in_quarter.get().strip(), out_base, layout_dir,
                                     os.path.basename(self.in_quarter.get().strip()), logger)
            # Halbjahr
            if self.in_half.get().strip():
                run_for_input_folder(self.in_half.get().strip(), out_base, layout_dir,
                                     os.path.basename(self.in_half.get().strip()), logger)
            # Jahr
            if self.in_year.get().strip():
                run_for_input_folder(self.in_year.get().strip(), out_base, layout_dir,
                                     os.path.basename(self.in_year.get().strip()), logger)

            logger.log("[FERTIG] Verarbeitung abgeschlossen.")
            self.write("[FERTIG] Verarbeitung abgeschlossen.\n")

        except Exception as e:
            tb = traceback.format_exc()
            logger.log("FEHLER:\n" + tb)
            self.write("\n[FEHLER]\n" + tb + "\n")
            messagebox.showerror("Fehler", str(e))


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
