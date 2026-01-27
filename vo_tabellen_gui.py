import os
import glob
import re
from datetime import datetime
from copy import copy as copy_style

import openpyxl
from openpyxl.styles import Alignment, PatternFill

import tkinter as tk
from tkinter import filedialog, messagebox


APP_NAME = "VO-Tabellen-Formatter"
DEFAULT_PROTOKOLL_DIR = r"L:\Abteilung5\sg52\50_INSO\1_Erhebungen\11_Insolvenzstatistiken\110_52411_BEANTRAGTE\1101_Monatsabschluss\Protokolle"


# ----------------------------
# Logging
# ----------------------------
class Logger:
    def __init__(self, protokoll_dir: str):
        self.protokoll_dir = protokoll_dir
        self.log_path = None
        self._ensure_dir()

    def _ensure_dir(self):
        os.makedirs(self.protokoll_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        self.log_path = os.path.join(self.protokoll_dir, f"{APP_NAME}_{ts}.log")

    def write(self, msg: str):
        line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  {msg}"
        print(line)
        with open(self.log_path, "a", encoding="utf-8") as f:
            f.write(line + "\n")


# ----------------------------
# Helpers
# ----------------------------
def safe_makedirs(path: str, logger: Logger | None = None):
    if not path:
        return
    os.makedirs(path, exist_ok=True)
    if logger:
        logger.write(f"[DIR] OK: {path}")


def guess_period_tag(filename: str) -> str:
    """
    Erwartete Muster: 2025-12 / 2025-Q4 / 2025-H2 / 2025-JJ
    """
    m = re.search(r"_(\d{4}-(?:\d{2}|Q[1-4]|H[1-2]|JJ))", filename)
    if m:
        return m.group(1)
    # fallback: versuch ohne underscore
    m = re.search(r"(\d{4}-(?:\d{2}|Q[1-4]|H[1-2]|JJ))", filename)
    if m:
        return m.group(1)
    return "UNBEKANNT"


def period_kind(period_tag: str) -> str:
    if re.fullmatch(r"\d{4}-\d{2}", period_tag):
        return "Monat"
    if re.fullmatch(r"\d{4}-Q[1-4]", period_tag):
        return "Quartal"
    if re.fullmatch(r"\d{4}-H[1-2]", period_tag):
        return "Halbjahr"
    if re.fullmatch(r"\d{4}-JJ", period_tag):
        return "Jahr"
    return "Unbekannt"


def list_existing_tables(input_dir: str):
    """
    Sucht nur die Dateien die verarbeitet werden sollen.
    Aktuell: Tabelle 1/2/3/5.
    Später: Tab8/Tab9.
    """
    patterns = [
        "Tabelle-1-Land*.xlsx",
        "Tabelle-2-Land*.xlsx",
        "Tabelle-3-Land*.xlsx",
        "Tabelle-5-Land*.xlsx",
        "25_Tab8_*.xlsx",
        "26_Tab8_*.xlsx",
        "27_Tab8_*.xlsx",
        "28_Tab8_*.xlsx",
        "29_Tab9_*.xlsx",
        "30_Tab9_*.xlsx",
        "31_Tab9_*.xlsx",
        "32_Tab9_*.xlsx",
    ]
    hits = []
    for p in patterns:
        hits.extend(glob.glob(os.path.join(input_dir, p)))
    return sorted(set(hits))


def repo_dir() -> str:
    """
    Wenn PyInstaller: sys._MEIPASS / oder aktuelles Verzeichnis.
    Wir nutzen für relative Pfade (Layouts, etc.) IMMER das Exe-Verzeichnis.
    """
    try:
        import sys
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
    except Exception:
        pass
    return os.getcwd()


def default_layout_dir() -> str:
    return os.path.join(repo_dir(), "Layouts")


def load_layout_wb(path: str):
    return openpyxl.load_workbook(path)


def copy_sheet_to_workbook(src_ws, dest_wb, new_title: str):
    """
    Kopiert ein Worksheet in ein anderes Workbook (Werte + Styles + Merges + Dimensions).
    """
    dest_ws = dest_wb.create_sheet(title=new_title)

    # Row/col dimensions
    for col_letter, dim in src_ws.column_dimensions.items():
        dest_ws.column_dimensions[col_letter].width = dim.width
        dest_ws.column_dimensions[col_letter].hidden = dim.hidden
        dest_ws.column_dimensions[col_letter].outline_level = dim.outline_level

    for row_idx, dim in src_ws.row_dimensions.items():
        dest_ws.row_dimensions[row_idx].height = dim.height
        dest_ws.row_dimensions[row_idx].hidden = dim.hidden
        dest_ws.row_dimensions[row_idx].outline_level = dim.outline_level

    # Cells
    for row in src_ws.iter_rows():
        for cell in row:
            d = dest_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                d.font = copy_style(cell.font)
                d.border = copy_style(cell.border)
                d.fill = copy_style(cell.fill)
                d.number_format = cell.number_format
                d.protection = copy_style(cell.protection)
                d.alignment = copy_style(cell.alignment)

    # merges
    for mr in src_ws.merged_cells.ranges:
        dest_ws.merge_cells(str(mr))

    # print settings
    dest_ws.page_setup = copy_style(src_ws.page_setup)
    dest_ws.page_margins = copy_style(src_ws.page_margins)
    dest_ws.print_options = copy_style(src_ws.print_options)
    dest_ws.sheet_properties = copy_style(src_ws.sheet_properties)

    # view (IMPORTANT: sheet_view has no setter -> nur Attribute kopieren die existieren)
    try:
        # einige openpyxl Versionen erlauben nicht das direkte Setzen, daher defensiv:
        dest_ws.sheet_view.showGridLines = src_ws.sheet_view.showGridLines
        dest_ws.sheet_view.zoomScale = src_ws.sheet_view.zoomScale
        dest_ws.sheet_view.zoomScaleNormal = src_ws.sheet_view.zoomScaleNormal
        dest_ws.sheet_view.view = src_ws.sheet_view.view
        dest_ws.sheet_view.selection = src_ws.sheet_view.selection
    except Exception:
        pass

    return dest_ws


# ----------------------------
# TAB 8 Prototype (Monat)
# ----------------------------
def build_table8_combined(input_files: list[str], template_path: str, out_path: str,
                         internal: bool, logger: Logger):
    """
    Kombiniert 4 Einzeldateien (25..28_Tab8_YYYY-MM) in 1 Workbook.
    Sheet-Namen = Dateiname ohne .xlsx.
    Layout: basiert auf Template (INTERN oder _g).
    """
    logger.write(f"[TAB8] Template: {template_path}")
    tmpl_wb = openpyxl.load_workbook(template_path)
    # wir nehmen die erste Vorlage als Basis und leeren danach die Sheets, dann fügen wir neue ein
    out_wb = openpyxl.Workbook()
    # remove default sheet
    if out_wb.worksheets:
        out_wb.remove(out_wb.worksheets[0])

    # Wir kopieren Settings aus template: alle Sheets einzeln als "Basis" verwenden wir das erste Template-Sheet
    # (damit Seitenränder, Print Setup etc. vorhanden sind) – in der Praxis klappt's gut.
    tmpl_ws = tmpl_wb.worksheets[0]

    for f in input_files:
        base = os.path.splitext(os.path.basename(f))[0]
        logger.write(f"[TAB8] add sheet: {base} <- {os.path.basename(f)}")
        src_wb = openpyxl.load_workbook(f)
        src_ws = src_wb.worksheets[0]

        # Erzeuge Zielsheet aus Template (Kopie) und übertrage Werte aus Quelle
        tmp_wb2 = openpyxl.Workbook()
        tmp_wb2.remove(tmp_wb2.worksheets[0])
        # Template-Sheet in tmp kopieren, dann in out kopieren
        tmp_ws = copy_sheet_to_workbook(tmpl_ws, tmp_wb2, base)
        # Werte aus src auf tmp schreiben (nur values, formats bleiben vom Template)
        for row in src_ws.iter_rows():
            for cell in row:
                # nur wenn in src etwas steht, überschreiben
                if cell.value is not None and str(cell.value).strip() != "":
                    tmp_ws.cell(row=cell.row, column=cell.column).value = cell.value

        # optional: interne Kopfzeile setzen (falls Template nicht alles abdeckt)
        if internal:
            # typischerweise steht das in A1 oder in einer merge-Region. Wir setzen A1 falls leer.
            if (tmp_ws["A1"].value is None) or (str(tmp_ws["A1"].value).strip() == ""):
                tmp_ws["A1"].value = "NUR FÜR DEN INTERNEN DIENSTGEBRAUCH"
                tmp_ws["A1"].alignment = Alignment(horizontal="center")

        # tmp_ws in out_wb übernehmen
        copy_sheet_to_workbook(tmp_ws, out_wb, base)

    safe_makedirs(os.path.dirname(out_path), logger)
    out_wb.save(out_path)
    logger.write(f"[TAB8] OK: {out_path}")


# ----------------------------
# GUI
# ----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("900x520")

        self.protokoll_dir_var = tk.StringVar(value=DEFAULT_PROTOKOLL_DIR)
        self.out_base_var = tk.StringVar(value="")
        self.in_monat_var = tk.StringVar(value="")
        self.in_quartal_var = tk.StringVar(value="")
        self.in_halbjahr_var = tk.StringVar(value="")
        self.in_jahr_var = tk.StringVar(value="")

        self.layout_dir_var = tk.StringVar(value=default_layout_dir())

        self._build_ui()

    def _build_ui(self):
        frm = tk.Frame(self)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        # Layout dir
        tk.Label(frm, text="Layouts-Ordner (enthält Layouts/*.xlsx):").grid(row=0, column=0, sticky="w")
        tk.Entry(frm, textvariable=self.layout_dir_var, width=90).grid(row=1, column=0, sticky="we")
        tk.Button(frm, text="Auswählen...", command=self.pick_layout_dir).grid(row=1, column=1, padx=5)

        # Output base
        tk.Label(frm, text="Ausgabe-Basisordner (VÖ-Tabellen wird darunter angelegt):").grid(row=2, column=0, sticky="w", pady=(10, 0))
        tk.Entry(frm, textvariable=self.out_base_var, width=90).grid(row=3, column=0, sticky="we")
        tk.Button(frm, text="Auswählen...", command=self.pick_out_dir).grid(row=3, column=1, padx=5)

        # Input dirs
        tk.Label(frm, text="Eingangsordner (nur gefüllte werden verarbeitet):").grid(row=4, column=0, sticky="w", pady=(10, 0))

        self._input_row(frm, 5, "Monat:", self.in_monat_var, self.pick_in_monat)
        self._input_row(frm, 6, "Quartal:", self.in_quartal_var, self.pick_in_quartal)
        self._input_row(frm, 7, "Halbjahr:", self.in_halbjahr_var, self.pick_in_halbjahr)
        self._input_row(frm, 8, "Jahr:", self.in_jahr_var, self.pick_in_jahr)

        # Protokoll
        tk.Label(frm, text="Protokoll-Ordner:").grid(row=9, column=0, sticky="w", pady=(10, 0))
        tk.Entry(frm, textvariable=self.protokoll_dir_var, width=90).grid(row=10, column=0, sticky="we")
        tk.Button(frm, text="Auswählen...", command=self.pick_protokoll_dir).grid(row=10, column=1, padx=5)

        # Buttons
        btnfrm = tk.Frame(frm)
        btnfrm.grid(row=11, column=0, columnspan=2, sticky="w", pady=(15, 0))
        tk.Button(btnfrm, text="Start", command=self.run).pack(side="left")
        tk.Button(btnfrm, text="Beenden", command=self.destroy).pack(side="left", padx=10)

        # Text log
        self.txt = tk.Text(frm, height=12)
        self.txt.grid(row=12, column=0, columnspan=2, sticky="nsew", pady=(12, 0))

        frm.grid_columnconfigure(0, weight=1)
        frm.grid_rowconfigure(12, weight=1)

    def _input_row(self, parent, row, label, var, pick_cmd):
        tk.Label(parent, text=label).grid(row=row, column=0, sticky="w")
        tk.Entry(parent, textvariable=var, width=90).grid(row=row, column=0, sticky="we", padx=(70, 0))
        tk.Button(parent, text="Auswählen...", command=pick_cmd).grid(row=row, column=1, padx=5)

    def pick_layout_dir(self):
        d = filedialog.askdirectory(title="Layouts-Ordner auswählen")
        if d:
            self.layout_dir_var.set(d)

    def pick_out_dir(self):
        d = filedialog.askdirectory(title="Ausgabe-Basisordner auswählen")
        if d:
            self.out_base_var.set(d)

    def pick_protokoll_dir(self):
        d = filedialog.askdirectory(title="Protokoll-Ordner auswählen")
        if d:
            self.protokoll_dir_var.set(d)

    def pick_in_monat(self):
        d = filedialog.askdirectory(title="Eingangsordner Monat auswählen")
        if d:
            self.in_monat_var.set(d)

    def pick_in_quartal(self):
        d = filedialog.askdirectory(title="Eingangsordner Quartal auswählen")
        if d:
            self.in_quartal_var.set(d)

    def pick_in_halbjahr(self):
        d = filedialog.askdirectory(title="Eingangsordner Halbjahr auswählen")
        if d:
            self.in_halbjahr_var.set(d)

    def pick_in_jahr(self):
        d = filedialog.askdirectory(title="Eingangsordner Jahr auswählen")
        if d:
            self.in_jahr_var.set(d)

    def log_ui(self, msg: str):
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.update_idletasks()

    def run(self):
        # IMPORTANT: Protokoll-Dir darf lokal sein – nicht zwingend L:
        protokoll_dir = self.protokoll_dir_var.get().strip()
        if not protokoll_dir:
            messagebox.showerror("Fehler", "Bitte Protokoll-Ordner angeben.")
            return

        try:
            logger = Logger(protokoll_dir)
        except Exception as e:
            messagebox.showerror("Fehler", f"Protokoll-Ordner kann nicht angelegt werden:\n{e}")
            return

        def log(msg):
            logger.write(msg)
            self.log_ui(msg)

        out_base = self.out_base_var.get().strip()
        if not out_base:
            log("[ABBRUCH] Kein Ausgabe-Basisordner angegeben.")
            messagebox.showerror("Fehler", "Bitte Ausgabe-Basisordner angeben.")
            return

        layout_dir = self.layout_dir_var.get().strip()
        if not layout_dir or not os.path.isdir(layout_dir):
            log("[ABBRUCH] Layouts-Ordner ungültig.")
            messagebox.showerror("Fehler", "Bitte gültigen Layouts-Ordner angeben.")
            return

        # sammle aktive Eingänge
        inputs = [
            ("Monat", self.in_monat_var.get().strip()),
            ("Quartal", self.in_quartal_var.get().strip()),
            ("Halbjahr", self.in_halbjahr_var.get().strip()),
            ("Jahr", self.in_jahr_var.get().strip()),
        ]
        inputs = [(k, p) for (k, p) in inputs if p]

        if not inputs:
            log("[ABBRUCH] Keine Eingangsordner angegeben.")
            messagebox.showerror("Fehler", "Bitte mindestens einen Eingangsordner (Monat/Quartal/Halbjahr/Jahr) angeben.")
            return

        vo_root = os.path.join(out_base, "VÖ-Tabellen")
        safe_makedirs(vo_root, logger)
        log(f"[OUT] VÖ-Root: {vo_root}")

        # ---- PROTOTYP: TAB8 für Monat, wenn 25..28 vorhanden ----
        for kind, in_dir in inputs:
            in_dir = os.path.normpath(in_dir)
            if not os.path.isdir(in_dir):
                log(f"[SKIP] {kind}: Ordner existiert nicht: {in_dir}")
                continue

            folder_name = os.path.basename(in_dir)
            out_dir = os.path.join(vo_root, folder_name)
            safe_makedirs(out_dir, logger)

            log(f"[IN] {kind}: {in_dir}")
            files = list_existing_tables(in_dir)
            log(f"[SCAN] {len(files)} Dateien gefunden (relevant).")

            # Tab8 Monatsprototyp: 25..28_Tab8_*.xlsx im Ordner -> in 1 Datei
            tab8_files = [f for f in files if re.search(r"(25|26|27|28)_Tab8_", os.path.basename(f))]
            if tab8_files:
                # sort by prefix number
                tab8_files = sorted(tab8_files, key=lambda x: os.path.basename(x))
                # period ermitteln aus erstem File
                period = guess_period_tag(os.path.basename(tab8_files[0]))
                # output names (nur Vorschlag)
                out_g = os.path.join(out_dir, f"Tabelle-8-Land_{period}_g.xlsx")
                out_i = os.path.join(out_dir, f"Tabelle-8-Land_{period}_INTERN.xlsx")

                layout_g = os.path.join(layout_dir, "Tabelle-8-Layout_g.xlsx")
                layout_i = os.path.join(layout_dir, "Tabelle-8-Layout_INTERN.xlsx")

                if os.path.isfile(layout_g) and os.path.isfile(layout_i):
                    try:
                        log(f"[TAB8] baue (_g): {os.path.basename(out_g)}")
                        build_table8_combined(tab8_files, layout_g, out_g, internal=False, logger=logger)
                        log(f"[TAB8] baue (INTERN): {os.path.basename(out_i)}")
                        build_table8_combined(tab8_files, layout_i, out_i, internal=True, logger=logger)
                    except Exception as e:
                        log(f"[TAB8][FEHLER] {e}")
                else:
                    log("[TAB8][SKIP] Layoutdateien fehlen (Tabelle-8-Layout_g.xlsx / Tabelle-8-Layout_INTERN.xlsx).")

            # TODO: hier später: Tabelle 1/2/3/5 Verarbeitung + Tab9

        log("[FERTIG] Verarbeitung abgeschlossen.")
        messagebox.showinfo("Fertig", "Verarbeitung abgeschlossen.\nDetails siehe Protokoll-Datei.")

        # Fenster offen lassen – du wolltest Meldungen sehen
        # (kein sys.exit)

def start_gui():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    start_gui()
