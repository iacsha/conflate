"""
Conflate v1.0
Fuzzy-match deduplication and master mapping tool for Excel / CSV data.
"""

import customtkinter as ctk
import pandas as pd
from rapidfuzz import process, fuzz
import os
import json
import traceback
import logging
import sys
import time
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import defaultdict
import numpy as np
import scipy.sparse as sp_sparse
from sklearn.feature_extraction.text import TfidfVectorizer
from logging.handlers import RotatingFileHandler



# ===========================================================
# VERSION & PATHS
# ===========================================================
VERSION      = "1.0"
APP_NAME     = "Conflate"
APP_TITLE    = f"{APP_NAME} v{VERSION} — Data Deduplication & Master Mapper"
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
LOG_PATH     = os.path.join(SCRIPT_DIR, f"{APP_NAME}.log")

# ===========================================================
# LOGGING SETUP  (rotating, max 5 MB × 3 files)
# ===========================================================
_root_logger = logging.getLogger()
_root_logger.setLevel(logging.DEBUG)          # capture DEBUG+ internally

_file_handler = RotatingFileHandler(
    LOG_PATH, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
)
_file_handler.setLevel(logging.DEBUG)
_file_handler.setFormatter(logging.Formatter(
    "%(asctime)s [%(levelname)-8s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
))
_root_logger.addHandler(_file_handler)

_console_handler = logging.StreamHandler(sys.stdout)
_console_handler.setLevel(logging.INFO)
_console_handler.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
_root_logger.addHandler(_console_handler)

def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logging.critical("UNCAUGHT EXCEPTION", exc_info=(exc_type, exc_value, exc_traceback))

sys.excepthook = handle_exception
logging.info(f"{'='*60}")
logging.info(f"  {APP_NAME} v{VERSION} — SESSION STARTED")
logging.info(f"  Log file: {LOG_PATH}")
logging.info(f"{'='*60}")

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


def ts() -> str:
    """Current timestamp string for filenames."""
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


# ===========================================================
# IN-APP LOG VIEWER
# ===========================================================
class LogViewerWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title(f"{APP_NAME} — Log Viewer")
        self.geometry("900x600")
        self.log_path = LOG_PATH

        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", padx=10, pady=(10, 0))

        ctk.CTkLabel(toolbar, text="Application Log",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        ctk.CTkButton(toolbar, text="⟳ Refresh", width=90,
                      command=self._load).pack(side="right", padx=5)
        ctk.CTkButton(toolbar, text="📂 Open Folder", width=110,
                      command=self._open_folder).pack(side="right", padx=5)

        # Level filter
        self.filter_var = ctk.StringVar(value="ALL")
        for lvl in ["ALL", "DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]:
            ctk.CTkRadioButton(toolbar, text=lvl, variable=self.filter_var,
                               value=lvl, command=self._load,
                               width=60).pack(side="left", padx=3)

        self.textbox = ctk.CTkTextbox(self, font=("Courier New", 11))
        self.textbox.pack(fill="both", expand=True, padx=10, pady=10)

        self._load()

    def _load(self):
        level_filter = self.filter_var.get()
        try:
            with open(self.log_path, "r", encoding="utf-8", errors="replace") as f:
                lines = f.readlines()
        except FileNotFoundError:
            lines = ["Log file not found yet.\n"]

        if level_filter != "ALL":
            lines = [l for l in lines if f"[{level_filter}" in l or not l.startswith("20")]

        self.textbox.configure(state="normal")
        self.textbox.delete("0.0", "end")
        self.textbox.insert("0.0", "".join(lines))
        self.textbox.configure(state="disabled")
        self.textbox.see("end")  # scroll to bottom (most recent)

    def _open_folder(self):
        import subprocess
        folder = os.path.dirname(self.log_path)
        if sys.platform == "win32":
            subprocess.Popen(f'explorer "{folder}"')
        elif sys.platform == "darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])


# ===========================================================
# HISTOGRAM WINDOW
# ===========================================================
class HistogramWindow(ctk.CTkToplevel):
    BUCKET_COLORS = {
        "50–60": "#cc3333", "60–70": "#cc6633", "70–80": "#ccaa33",
        "80–90": "#88aa33", "90–100": "#33aa33", "100": "#00cc88",
    }

    def __init__(self, parent, matches, on_proceed):
        super().__init__(parent)
        self.title("Match Score Distribution")
        self.geometry("640x430")
        self.resizable(False, False)
        self.on_proceed = on_proceed
        self.grab_set()

        ctk.CTkLabel(self, text="Match Score Distribution",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        ctk.CTkLabel(self, text=f"Total matches found: {len(matches):,}",
                     text_color="gray").pack()

        buckets = {"50–60": 0, "60–70": 0, "70–80": 0,
                   "80–90": 0, "90–100": 0, "100": 0}
        for m in matches:
            s = m["Score"]
            if s >= 100:        buckets["100"]    += 1
            elif s >= 90:       buckets["90–100"] += 1
            elif s >= 80:       buckets["80–90"]  += 1
            elif s >= 70:       buckets["70–80"]  += 1
            elif s >= 60:       buckets["60–70"]  += 1
            else:               buckets["50–60"]  += 1

        max_count  = max(buckets.values()) or 1
        canvas     = tk.Canvas(self, height=270, bg="#2b2b2b", highlightthickness=0)
        canvas.pack(fill="x", padx=20, pady=10)

        labels, counts = list(buckets.keys()), list(buckets.values())
        n, canvas_w, gap, bar_area_h = len(labels), 600, 8, 200
        bar_w = (canvas_w - gap * (n + 1)) // n

        for i, (label, count) in enumerate(zip(labels, counts)):
            x0    = gap + i * (bar_w + gap)
            bar_h = int((count / max_count) * bar_area_h)
            y0    = bar_area_h + 10 - bar_h
            color = self.BUCKET_COLORS.get(label, "#4488ff")
            canvas.create_rectangle(x0, y0, x0 + bar_w, bar_area_h + 10,
                                    fill=color, outline="")
            canvas.create_text(x0 + bar_w // 2, bar_area_h + 25,
                               text=label, fill="white", font=("Arial", 9))
            if count > 0:
                canvas.create_text(x0 + bar_w // 2, y0 - 12,
                                   text=f"{count:,}", fill="white",
                                   font=("Arial", 9, "bold"))

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="Proceed to Review →",
                      command=self._proceed).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Cancel",
                      fg_color="gray", hover_color="darkgray",
                      command=self.destroy).pack(side="left", padx=10)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _proceed(self):
        self.destroy()
        self.on_proceed()


# ===========================================================
# COLUMN MAPPING WIDGET  (Master Mode only)
# ===========================================================
class ColumnMappingWidget(ctk.CTkFrame):
    def __init__(self, parent, primary_cols, master_cols, **kwargs):
        super().__init__(parent, **kwargs)
        self.primary_cols = primary_cols
        self.master_cols  = master_cols
        self.pairs        = []

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=5, pady=(5, 0))
        ctk.CTkLabel(
            header,
            text="3b. Column Mappings — Primary → Master  (many-to-many supported):",
            font=ctk.CTkFont(weight="bold")
        ).pack(side="left")
        ctk.CTkButton(header, text="+ Add Pair", width=100,
                      command=self.add_pair).pack(side="right", padx=5)

        self.pairs_frame = ctk.CTkScrollableFrame(self, height=110)
        self.pairs_frame.pack(fill="both", expand=True, padx=5, pady=5)
        ctk.CTkLabel(
            self,
            text="Tip: pairs define what is compared for scoring. Add one pair per logical field.",
            text_color="gray", font=ctk.CTkFont(size=11)
        ).pack(pady=(0, 5))
        self.add_pair()

    def add_pair(self):
        row_frame = ctk.CTkFrame(self.pairs_frame, fg_color="transparent")
        row_frame.pack(fill="x", pady=3)
        p_var = ctk.StringVar(value=self.primary_cols[0] if self.primary_cols else "")
        m_var = ctk.StringVar(value=self.master_cols[0]  if self.master_cols  else "")
        ctk.CTkOptionMenu(row_frame, values=self.primary_cols,
                          variable=p_var, width=200).pack(side="left", padx=5)
        ctk.CTkLabel(row_frame, text="→",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        ctk.CTkOptionMenu(row_frame, values=self.master_cols,
                          variable=m_var, width=200).pack(side="left", padx=5)
        pair = (p_var, m_var, row_frame)
        self.pairs.append(pair)
        ctk.CTkButton(
            row_frame, text="✕", width=32,
            fg_color="#aa2222", hover_color="#cc0000",
            command=lambda p=pair: self.remove_pair(p)
        ).pack(side="left", padx=5)

    def remove_pair(self, pair):
        if len(self.pairs) <= 1:
            return
        _, _, row_frame = pair
        row_frame.destroy()
        self.pairs.remove(pair)

    def get_mappings(self):
        return [(p.get(), m.get()) for p, m, _ in self.pairs if p.get() and m.get()]

    def update_columns(self, primary_cols, master_cols):
        self.primary_cols, self.master_cols = primary_cols, master_cols
        for _, _, f in self.pairs: f.destroy()
        self.pairs.clear()
        self.add_pair()


# ===========================================================
# MAIN APP
# ===========================================================
class DataMatchApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("950x1060")

        # --- STATE ---
        self.file_path        = ""
        self.master_file_path = ""
        self.df               = None
        self.df_master        = None
        self.matches          = []
        self.current_index    = 0
        self.approved_merges  = []
        self.flagged_merges   = []
        self.skipped_count    = 0
        self.decision_history = []   # "A" | "B" | "S" | "F" per step
        self.progress_file    = ""
        self.target_cols      = []
        self.target_cols_master = []
        self.col_mappings     = []
        self.checkbox_vars    = {}
        self.master_checkbox_vars = {}
        self.cancel_scan      = False
        self.review_mode_active = False
        self.is_master_mode   = False
        self.primary_lookup   = {}
        self.master_lookup    = {}
        self.review_start_time = None
        self.column_mapping_widget = None
        self.primary_id_col   = ""
        self.master_id_col    = ""
        self._raw_export_path = None
        self._scan_settings   = {}   # captured for crash context
        self._note_has_focus  = False  # suppresses hotkeys while typing a note

        # --- KEY BINDINGS ---
        self.bind("<Left>",      lambda e: self._safe_key_decision("A"))
        self.bind("1",           lambda e: self._safe_key_decision("A"))
        self.bind("<Right>",     lambda e: self._safe_key_decision("B"))
        self.bind("2",           lambda e: self._safe_key_decision("B"))
        self.bind("<space>",     lambda e: self._safe_key_decision("S"))
        self.bind("3",           lambda e: self._safe_key_decision("S"))
        self.bind("f",           lambda e: self._safe_key_decision("F"))
        self.bind("4",           lambda e: self._safe_key_decision("F"))
        self.bind("<Control-z>", lambda e: self._safe_undo())

        self._build_setup_ui()
        self._build_review_ui()

    # =========================================================
    # UI — SETUP SCREEN
    # =========================================================
    def _build_setup_ui(self):
        self.setup_frame = ctk.CTkFrame(self)
        self.setup_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Title + log viewer button
        title_row = ctk.CTkFrame(self.setup_frame, fg_color="transparent")
        title_row.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkLabel(
            title_row, text=f"{APP_NAME} — Data Deduplication & Master Mapper",
            font=ctk.CTkFont(size=22, weight="bold")
        ).pack(side="left")
        ctk.CTkButton(
            title_row, text="📋 View Log", width=100,
            fg_color="#444", hover_color="#666",
            command=self._open_log_viewer
        ).pack(side="right", padx=5)

        # --- File loaders ---
        loader_frame = ctk.CTkFrame(self.setup_frame, fg_color="transparent")
        loader_frame.pack(pady=5, fill="x")
        loader_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(loader_frame, text="1. Select Primary Data",
                      command=self.load_file).grid(row=0, column=0, padx=10, pady=5)
        self.lbl_file = ctk.CTkLabel(loader_frame, text="No primary file", text_color="gray")
        self.lbl_file.grid(row=1, column=0)

        ctk.CTkButton(loader_frame, text="1b. Select Master List (Optional)",
                      command=self.load_master_file,
                      fg_color="teal", hover_color="darkcyan"
                      ).grid(row=0, column=1, padx=10, pady=5)
        self.lbl_master_file = ctk.CTkLabel(loader_frame,
                                             text="No master file (Dedupe Mode)",
                                             text_color="gray")
        self.lbl_master_file.grid(row=1, column=1)

        # --- Settings ---
        settings_frame = ctk.CTkFrame(self.setup_frame, fg_color="transparent")
        settings_frame.pack(pady=10)

        self.lbl_slider = ctk.CTkLabel(settings_frame, text="Match Strictness: 85%")
        self.lbl_slider.grid(row=0, column=0, padx=15)
        self.slider_score = ctk.CTkSlider(settings_frame, from_=50, to=100,
                                          number_of_steps=50,
                                          command=self._update_slider_label)
        self.slider_score.set(85)
        self.slider_score.grid(row=1, column=0, padx=15, pady=5)

        ctk.CTkLabel(settings_frame, text="Max Matches per Item:").grid(row=0, column=1, padx=15)
        self.combo_max = ctk.CTkOptionMenu(settings_frame, values=["1", "3", "5", "10", "20"])
        self.combo_max.set("5")
        self.combo_max.grid(row=1, column=1, padx=15, pady=5)

        self.var_test_mode = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(settings_frame, text="Test Mode\n(First 100 Rows)",
                        variable=self.var_test_mode
                        ).grid(row=0, column=2, rowspan=2, padx=15)

        # --- Engine toggle ---
        engine_frame = ctk.CTkFrame(self.setup_frame)
        engine_frame.pack(pady=5, padx=20, fill="x")
        ctk.CTkLabel(engine_frame, text="Processing Engine:",
                     font=ctk.CTkFont(weight="bold")).pack(pady=(5, 0))
        self.engine_var = ctk.StringVar(value="RapidFuzz")
        ctk.CTkRadioButton(engine_frame,
                           text="Standard (RapidFuzz) — Best for < 10,000 rows",
                           variable=self.engine_var, value="RapidFuzz").pack(pady=5)
        ctk.CTkRadioButton(engine_frame,
                           text="Heavy Duty (TF-IDF) — Best for > 10,000 rows",
                           variable=self.engine_var, value="TFIDF").pack(pady=(0, 10))

        # --- Column selectors ---
        ctk.CTkLabel(self.setup_frame, text="2. Select Columns to Search:",
                     font=ctk.CTkFont(weight="bold")).pack(pady=(12, 4))

        col_frame = ctk.CTkFrame(self.setup_frame, fg_color="transparent")
        col_frame.pack(fill="x", padx=20)
        col_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(col_frame, text="Primary File",
                     font=ctk.CTkFont(weight="bold")).grid(row=0, column=0)
        ctk.CTkLabel(col_frame, text="Master File",
                     font=ctk.CTkFont(weight="bold")).grid(row=0, column=1)

        self.scroll_cols = ctk.CTkScrollableFrame(col_frame, height=95)
        self.scroll_cols.grid(row=1, column=0, padx=10, sticky="nsew")
        ctk.CTkLabel(self.scroll_cols, text="Load a primary file…",
                     text_color="gray").pack(pady=20)

        self.scroll_cols_master = ctk.CTkScrollableFrame(col_frame, height=95)
        self.scroll_cols_master.grid(row=1, column=1, padx=10, sticky="nsew")
        ctk.CTkLabel(self.scroll_cols_master, text="Load a master file…",
                     text_color="gray").pack(pady=20)

        # --- Unique ID selectors ---
        id_section = ctk.CTkFrame(self.setup_frame)
        id_section.pack(pady=(8, 0), padx=20, fill="x")

        ctk.CTkLabel(
            id_section,
            text="2b. Select Unique ID Column(s) to retain in output:",
            font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, columnspan=2, pady=(6, 2))

        ctk.CTkLabel(id_section, text="Primary File ID:").grid(
            row=1, column=0, padx=(20, 5), sticky="e")
        self.combo_primary_id = ctk.CTkOptionMenu(
            id_section, values=["— none —"], width=220,
            command=lambda v: setattr(self, "primary_id_col",
                                      "" if v == "— none —" else v)
        )
        self.combo_primary_id.set("— none —")
        self.combo_primary_id.grid(row=1, column=1, padx=(0, 20), pady=4, sticky="w")

        ctk.CTkLabel(id_section, text="Master File ID:").grid(
            row=2, column=0, padx=(20, 5), sticky="e")
        self.combo_master_id = ctk.CTkOptionMenu(
            id_section, values=["— none —"], width=220,
            command=lambda v: setattr(self, "master_id_col",
                                      "" if v == "— none —" else v)
        )
        self.combo_master_id.set("— none —")
        self.combo_master_id.grid(row=2, column=1, padx=(0, 20), pady=(4, 8), sticky="w")

        ctk.CTkLabel(
            id_section,
            text="Both IDs appear in every output row so you can VLOOKUP back to source data.",
            text_color="gray", font=ctk.CTkFont(size=11)
        ).grid(row=3, column=0, columnspan=2, pady=(0, 6))

        # --- Mapping section ---
        self.mapping_section = ctk.CTkFrame(self.setup_frame)
        self.mapping_section.pack(pady=5, padx=20, fill="x")
        ctk.CTkLabel(
            self.mapping_section,
            text="Load both files to configure column mappings (Master Mode only).",
            text_color="gray"
        ).pack(pady=8)

        # --- Actions ---
        action_frame = ctk.CTkFrame(self.setup_frame, fg_color="transparent")
        action_frame.pack(pady=(12, 5))

        self.btn_start = ctk.CTkButton(action_frame, text="3. Start Scan",
                                       command=self.start_processing)
        self.btn_start.grid(row=0, column=0, padx=10)

        self.btn_stop = ctk.CTkButton(action_frame, text="Stop Scan",
                                      fg_color="red", hover_color="darkred",
                                      state="disabled", command=self._trigger_stop)
        self.btn_stop.grid(row=0, column=1, padx=10)

        self.progress_bar = ctk.CTkProgressBar(self.setup_frame, width=340)
        self.progress_bar.set(0)
        self.status_label = ctk.CTkLabel(self.setup_frame, text="",
                                         text_color="gray", font=ctk.CTkFont(size=12))

    # =========================================================
    # UI — REVIEW SCREEN
    # =========================================================
    def _build_review_ui(self):
        self.review_frame = ctk.CTkFrame(self)

        self.lbl_progress = ctk.CTkLabel(
            self.review_frame, text="Reviewing Match X of Y",
            font=ctk.CTkFont(size=18, weight="bold"))
        self.lbl_progress.pack(pady=8)

        self.lbl_score = ctk.CTkLabel(
            self.review_frame, text="Match Score: —",
            text_color="orange", font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_score.pack(pady=4)

        # --- Session stats bar ---
        stats_frame = ctk.CTkFrame(self.review_frame)
        stats_frame.pack(fill="x", padx=20, pady=(0, 6))
        stats_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)

        headers = ["✅ Approved", "🚩 Flagged", "⏭ Skipped", "🔲 Remaining", "⏱ Est. Left"]
        for col, h in enumerate(headers):
            ctk.CTkLabel(stats_frame, text=h,
                         font=ctk.CTkFont(weight="bold")).grid(
                row=0, column=col, padx=4, pady=(6, 0))

        self.lbl_stat_approved  = ctk.CTkLabel(stats_frame, text="0", text_color="#00cc66")
        self.lbl_stat_flagged   = ctk.CTkLabel(stats_frame, text="0", text_color="#ffaa00")
        self.lbl_stat_skipped   = ctk.CTkLabel(stats_frame, text="0", text_color="gray")
        self.lbl_stat_remaining = ctk.CTkLabel(stats_frame, text="0", text_color="orange")
        self.lbl_stat_time      = ctk.CTkLabel(stats_frame, text="–",  text_color="lightblue")

        for col, lbl in enumerate([self.lbl_stat_approved, self.lbl_stat_flagged,
                                    self.lbl_stat_skipped, self.lbl_stat_remaining,
                                    self.lbl_stat_time]):
            lbl.grid(row=1, column=col, padx=4, pady=(0, 6))

        # --- Match display ---
        self.lbl_a = ctk.CTkLabel(self.review_frame, text="Item A (Your Data):",
                                  font=ctk.CTkFont(weight="bold"))
        self.lbl_a.pack(anchor="w", padx=20)
        self.textbox_a = ctk.CTkTextbox(self.review_frame, height=50, width=850)
        self.textbox_a.pack(pady=3, padx=20)

        self.lbl_b = ctk.CTkLabel(self.review_frame, text="Item B:",
                                  font=ctk.CTkFont(weight="bold"))
        self.lbl_b.pack(anchor="w", padx=20)
        self.textbox_b = ctk.CTkTextbox(self.review_frame, height=50, width=850)
        self.textbox_b.pack(pady=3, padx=20)

        # --- Decision buttons ---
        btn_frame = ctk.CTkFrame(self.review_frame, fg_color="transparent")
        btn_frame.pack(pady=10)

        self.btn_keep_a = ctk.CTkButton(
            btn_frame, text="Keep Yours\n[← / 1]",
            fg_color="green", hover_color="darkgreen",
            command=lambda: self.make_decision("A"))
        self.btn_keep_a.grid(row=0, column=0, padx=6)

        self.btn_keep_b = ctk.CTkButton(
            btn_frame, text="Override with Match\n[→ / 2]",
            fg_color="teal", hover_color="darkcyan",
            command=lambda: self.make_decision("B"))
        self.btn_keep_b.grid(row=0, column=1, padx=6)

        self.btn_skip = ctk.CTkButton(
            btn_frame, text="Skip\n[Space / 3]",
            fg_color="gray", hover_color="darkgray",
            command=lambda: self.make_decision("S"))
        self.btn_skip.grid(row=0, column=2, padx=6)

        self.btn_flag = ctk.CTkButton(
            btn_frame, text="Flag for Review\n[F / 4]",
            fg_color="#b05800", hover_color="#d07000",
            command=lambda: self.make_decision("F"))
        self.btn_flag.grid(row=0, column=3, padx=6)

        self.btn_undo = ctk.CTkButton(
            btn_frame, text="Undo Last\n[Ctrl+Z]",
            fg_color="#555", hover_color="#777",
            command=self.undo_decision)
        self.btn_undo.grid(row=0, column=4, padx=6)

        # --- Notes field ---
        note_frame = ctk.CTkFrame(self.review_frame, fg_color="transparent")
        note_frame.pack(fill="x", padx=20, pady=(4, 0))
        ctk.CTkLabel(note_frame, text="Decision Note — optional, saved with output  (hotkeys pause while typing):",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.entry_note = ctk.CTkEntry(note_frame, placeholder_text="e.g. 'Vendor confirmed same supplier'",
                                       width=850)
        self.entry_note.pack(fill="x", pady=(3, 0))

        # Suspend hotkeys while the user is typing a note
        self.entry_note.bind("<FocusIn>",  lambda e: setattr(self, "_note_has_focus", True))
        self.entry_note.bind("<FocusOut>", lambda e: setattr(self, "_note_has_focus", False))
        ctk.CTkLabel(self.review_frame, text="Full Row Context:",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=(10, 0))
        self.textbox_context = ctk.CTkTextbox(self.review_frame, height=185, width=850)
        self.textbox_context.pack(pady=5, padx=20)

        self.btn_save_exit = ctk.CTkButton(
            self.review_frame, text="Save Progress & Exit",
            fg_color="red", hover_color="darkred",
            command=self.save_and_exit)
        self.btn_save_exit.pack(pady=8)

    # =========================================================
    # SETUP HANDLERS
    # =========================================================
    def _update_slider_label(self, value):
        self.lbl_slider.configure(text=f"Match Strictness: {int(value)}%")

    def _safe_key_decision(self, choice):
        if self._note_has_focus:
            return   # let the keystroke reach the entry widget instead
        if self.review_mode_active and self.current_index < len(self.matches):
            self.make_decision(choice)

    def _safe_undo(self):
        if self.review_mode_active:
            self.undo_decision()

    def _trigger_stop(self):
        self.cancel_scan = True
        self.btn_stop.configure(text="Stopping…", state="disabled")
        logging.info("User triggered Stop Scan.")

    def _open_log_viewer(self):
        LogViewerWindow(self)

    def load_file(self):
        fp = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
        if not fp:
            return
        self.file_path = fp
        self.lbl_file.configure(text=os.path.basename(fp))
        stem = os.path.splitext(os.path.basename(fp))[0]
        self.progress_file = os.path.join(
            os.path.dirname(fp), f"DataMatch_Progress_{stem}.json")
        logging.info(f"Primary file selected: {fp}")
        try:
            cols = self._read_columns(fp)
            self._populate_checkboxes(cols, is_master=False)
            self._refresh_mapping_section()
        except Exception as e:
            messagebox.showerror("Error", f"Could not read columns.\n\n{e}")

    def load_master_file(self):
        fp = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
        if not fp:
            return
        self.master_file_path = fp
        self.lbl_master_file.configure(
            text=f"Master: {os.path.basename(fp)}", text_color="teal")
        logging.info(f"Master file selected: {fp}")
        try:
            cols = self._read_columns(fp)
            self._populate_checkboxes(cols, is_master=True)
            self._refresh_mapping_section()
        except Exception as e:
            messagebox.showerror("Error", f"Could not read columns.\n\n{e}")

    def _read_columns(self, fp):
        return (pd.read_csv(fp, nrows=0) if fp.endswith(".csv")
                else pd.read_excel(fp, nrows=0)).columns.tolist()

    def _populate_checkboxes(self, columns, is_master=False):
        frame = self.scroll_cols_master if is_master else self.scroll_cols
        d     = self.master_checkbox_vars if is_master else self.checkbox_vars
        for w in frame.winfo_children():
            w.destroy()
        d.clear()
        for col in columns:
            var = ctk.BooleanVar(value=False)
            ctk.CTkCheckBox(frame, text=str(col), variable=var
                            ).pack(pady=2, padx=10, anchor="w")
            d[col] = var
        opts = ["— none —"] + columns
        if is_master:
            self.combo_master_id.configure(values=opts)
            self.combo_master_id.set("— none —")
            self.master_id_col = ""
        else:
            self.combo_primary_id.configure(values=opts)
            self.combo_primary_id.set("— none —")
            self.primary_id_col = ""

    def _refresh_mapping_section(self):
        primary_cols = list(self.checkbox_vars.keys())
        master_cols  = list(self.master_checkbox_vars.keys())
        for w in self.mapping_section.winfo_children():
            w.destroy()
        if not primary_cols or not master_cols:
            ctk.CTkLabel(
                self.mapping_section,
                text="Load both files to configure column mappings (Master Mode only).",
                text_color="gray"
            ).pack(pady=8)
            self.column_mapping_widget = None
            return
        self.column_mapping_widget = ColumnMappingWidget(
            self.mapping_section, primary_cols, master_cols)
        self.column_mapping_widget.pack(fill="both", expand=True, padx=5, pady=5)

    # =========================================================
    # DATA HELPERS
    # =========================================================
    def _load_dataframe(self, path):
        return (pd.read_csv(path) if path.endswith(".csv")
                else pd.read_excel(path))

    def _clean_df(self, df, cols):
        df = df.copy()
        df["Combined_Search"] = ""
        for col in cols:
            df["Combined_Search"] += df[col].astype(str).replace("nan", "") + " "
        df["Cleaned"] = (
            df["Combined_Search"].str.strip().str.lower()
            .str.replace("[.,-]", "", regex=True)
        )
        return df[df["Cleaned"].astype(bool)].reset_index(drop=True)

    def _build_lookups(self):
        skip = {"Combined_Search", "Cleaned"}
        t0 = time.perf_counter()
        self.primary_lookup = {
            row["Cleaned"]: {k: v for k, v in row.items() if k not in skip}
            for _, row in self.df.iterrows()
        }
        if self.is_master_mode and self.df_master is not None:
            self.master_lookup = {
                row["Cleaned"]: {k: v for k, v in row.items() if k not in skip}
                for _, row in self.df_master.iterrows()
            }
        logging.debug(f"Lookup dicts built in {time.perf_counter()-t0:.2f}s "
                      f"({len(self.primary_lookup):,} primary, "
                      f"{len(self.master_lookup):,} master)")

    def _add_match(self, item_a_cleaned, item_b_cleaned, score):
        row_a = self.primary_lookup.get(item_a_cleaned)
        row_b = (self.master_lookup if self.is_master_mode
                 else self.primary_lookup).get(item_b_cleaned)
        if row_a is None or row_b is None:
            return
        val_a  = " | ".join(str(row_a.get(c, "")) for c in self.target_cols)
        cols_b = self.target_cols_master if self.is_master_mode else self.target_cols
        val_b  = " | ".join(str(row_b.get(c, "")) for c in cols_b)
        self.matches.append({
            "Score":   round(score, 1),
            "Match_A": val_a,
            "Match_B": val_b,
            "Row_A":   row_a,
            "Row_B":   row_b,
        })

    # =========================================================
    # ENGINE: RAPIDFUZZ
    # =========================================================
    def _run_rapidfuzz(self, unique_primary, unique_master, score_cutoff, max_limit):
        total = len(unique_primary)
        logging.info(f"RapidFuzz: scanning {total:,} primary items…")
        t_engine = time.perf_counter()

        for i, item_p in enumerate(unique_primary):
            if self.cancel_scan:
                break
            search_list = unique_master if self.is_master_mode else unique_primary[i + 1:]
            if not search_list:
                break
            results = process.extract(
                item_p, search_list,
                scorer=fuzz.token_sort_ratio,
                limit=max_limit,
                score_cutoff=score_cutoff
            )
            for match_str, score, _ in results:
                self._add_match(item_p, match_str, score)

            if i % max(1, total // 200) == 0:
                self.progress_bar.set((i + 1) / total)
                self.status_label.configure(
                    text=f"Scanning {i + 1:,} / {total:,} items…")
                self.update()

        elapsed = time.perf_counter() - t_engine
        logging.info(f"RapidFuzz complete: {len(self.matches):,} matches "
                     f"in {elapsed:.1f}s ({elapsed/max(total,1)*1000:.1f}ms/item)")

    # =========================================================
    # ENGINE: TF-IDF  (sklearn — best for > 10,000 rows)
    # =========================================================
    def _run_tfidf(self, unique_primary, unique_master, score_cutoff, max_limit):
        score_decimal = score_cutoff / 100.0
        batch_size    = 2000
        vectorizer    = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
        t_engine      = time.perf_counter()

        self.status_label.configure(text="TF-IDF: Fitting vocabulary…")
        self.update()

        if self.is_master_mode:
            logging.info("TF-IDF Master Mode: fitting combined vocab…")
            t0 = time.perf_counter()
            vectorizer.fit(unique_primary + unique_master)
            tfidf_p = vectorizer.transform(unique_primary)
            tfidf_m = vectorizer.transform(unique_master)
            logging.debug(f"TF-IDF transform in {time.perf_counter()-t0:.2f}s "
                          f"P={tfidf_p.shape} M={tfidf_m.shape}")
            n_rows = tfidf_p.shape[0]

            for start_row in range(0, n_rows, batch_size):
                if self.cancel_scan:
                    break
                end_row = min(start_row + batch_size, n_rows)
                t0 = time.perf_counter()
                self.status_label.configure(
                    text=f"TF-IDF Batch {start_row:,}–{end_row:,} of {n_rows:,}…")
                self.progress_bar.set(start_row / n_rows)
                self.update()

                batch_sim = tfidf_p[start_row:end_row].dot(tfidf_m.T).tocoo()
                row_hits  = defaultdict(list)
                for local_i, j, s in zip(batch_sim.row, batch_sim.col, batch_sim.data):
                    if s >= score_decimal:
                        row_hits[local_i].append((s, j))
                for local_i, hits in row_hits.items():
                    hits.sort(reverse=True)
                    for s, j in hits[:max_limit]:
                        self._add_match(unique_primary[start_row + local_i],
                                        unique_master[j], s * 100)
                logging.debug(f"Batch {start_row}–{end_row} in {time.perf_counter()-t0:.2f}s")

        else:
            logging.info("TF-IDF Dedupe Mode: fitting & transforming…")
            t0 = time.perf_counter()
            tfidf_matrix = vectorizer.fit_transform(unique_primary)
            logging.debug(f"TF-IDF fit_transform in {time.perf_counter()-t0:.2f}s "
                          f"matrix={tfidf_matrix.shape}")
            n_rows = tfidf_matrix.shape[0]

            for start_row in range(0, n_rows, batch_size):
                if self.cancel_scan:
                    break
                end_row = min(start_row + batch_size, n_rows)
                t0 = time.perf_counter()
                self.status_label.configure(
                    text=f"TF-IDF Batch {start_row:,}–{end_row:,} of {n_rows:,}…")
                self.progress_bar.set(start_row / n_rows)
                self.update()

                batch_sim = tfidf_matrix[start_row:end_row].dot(tfidf_matrix.T).tocoo()
                row_hits  = defaultdict(list)
                for local_i, col_idx, s in zip(batch_sim.row, batch_sim.col, batch_sim.data):
                    real_row = start_row + local_i
                    if col_idx <= real_row:
                        continue
                    if s >= score_decimal:
                        row_hits[local_i].append((s, col_idx))
                for local_i, hits in row_hits.items():
                    hits.sort(reverse=True)
                    for s, col_idx in hits[:max_limit]:
                        self._add_match(unique_primary[start_row + local_i],
                                        unique_primary[col_idx], s * 100)
                logging.debug(f"Batch {start_row}–{end_row} in {time.perf_counter()-t0:.2f}s")

        elapsed = time.perf_counter() - t_engine
        logging.info(f"TF-IDF complete: {len(self.matches):,} matches in {elapsed:.1f}s")

    # =========================================================
    # MAIN START
    # =========================================================
    def start_processing(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select your Primary Data file first.")
            return

        self.target_cols = [c for c, v in self.checkbox_vars.items() if v.get()]
        if not self.target_cols:
            messagebox.showerror("Error",
                                 "Please check at least one column in the Primary list.")
            return

        self.is_master_mode = bool(self.master_file_path)
        if self.is_master_mode:
            if self.column_mapping_widget:
                self.col_mappings = self.column_mapping_widget.get_mappings()
            else:
                self.col_mappings = []

            if self.col_mappings:
                self.target_cols        = [p for p, _ in self.col_mappings]
                self.target_cols_master = [m for _, m in self.col_mappings]
            else:
                self.target_cols_master = [c for c, v in
                                           self.master_checkbox_vars.items() if v.get()]
                if not self.target_cols_master:
                    messagebox.showerror(
                        "Error",
                        "Master file loaded but no master columns selected.\n"
                        "Add a Column Mapping pair or check master columns.")
                    return

        if os.path.exists(self.progress_file):
            choice = messagebox.askyesnocancel(
                "Saved Progress",
                "Found saved progress.\n\nYES = Resume\nNO = Start fresh\nCANCEL = Abort")
            if choice is True:
                self.load_progress()
                self.switch_to_review()
                return
            elif choice is False:
                os.remove(self.progress_file)
            else:
                return

        # Capture settings for crash context
        self._scan_settings = {
            "file":         self.file_path,
            "master_file":  self.master_file_path,
            "engine":       self.engine_var.get(),
            "cutoff":       int(self.slider_score.get()),
            "max_limit":    int(self.combo_max.get()),
            "mode":         "master" if self.is_master_mode else "dedupe",
            "test_mode":    self.var_test_mode.get(),
            "primary_cols": self.target_cols,
            "master_cols":  self.target_cols_master if self.is_master_mode else [],
        }
        logging.info(f"Scan settings: {json.dumps(self._scan_settings, default=str)}")

        # Reset state
        self.cancel_scan      = False
        self.matches          = []
        self.approved_merges  = []
        self.flagged_merges   = []
        self.skipped_count    = 0
        self.decision_history = []

        self.btn_start.configure(text="Analyzing…", state="disabled")
        self.btn_stop.configure(text="Stop Scan", state="normal")
        self.progress_bar.pack(pady=5)
        self.status_label.pack(pady=2)
        self.progress_bar.set(0)
        self.update()

        score_cutoff   = int(self.slider_score.get())
        max_limit      = int(self.combo_max.get())
        engine         = self.engine_var.get()
        unique_primary = []
        unique_master  = []

        try:
            t_total = time.perf_counter()

            logging.info("Loading primary data…")
            t0 = time.perf_counter()
            self.df = self._load_dataframe(self.file_path)
            if self.var_test_mode.get():
                self.df = self.df.head(100)
            self.df    = self._clean_df(self.df, self.target_cols)
            unique_primary = self.df["Cleaned"].unique().tolist()
            logging.info(f"Primary loaded: {len(unique_primary):,} unique items "
                         f"in {time.perf_counter()-t0:.2f}s")

            if self.is_master_mode:
                logging.info("Loading master data…")
                t0 = time.perf_counter()
                self.df_master = self._load_dataframe(self.master_file_path)
                self.df_master = self._clean_df(self.df_master, self.target_cols_master)
                unique_master  = self.df_master["Cleaned"].unique().tolist()
                logging.info(f"Master loaded: {len(unique_master):,} unique items "
                             f"in {time.perf_counter()-t0:.2f}s")

            self._build_lookups()

            if engine == "RapidFuzz":
                self._run_rapidfuzz(unique_primary, unique_master, score_cutoff, max_limit)
            else:
                self._run_tfidf(unique_primary, unique_master, score_cutoff, max_limit)

            logging.info(f"Total scan time: {time.perf_counter()-t_total:.1f}s")

        except Exception:
            tb = traceback.format_exc()
            # Log crash WITH full settings context
            logging.error(
                f"CRASH in start_processing\n"
                f"Settings: {json.dumps(self._scan_settings, default=str)}\n"
                f"Primary rows: {len(unique_primary):,} | "
                f"Master rows: {len(unique_master):,}\n"
                f"Matches so far: {len(self.matches):,}\n"
                f"{tb}"
            )
            messagebox.showerror("Crash", tb)

        finally:
            self.btn_start.configure(text="3. Start Scan", state="normal")
            self.btn_stop.configure(text="Stop Scan", state="disabled")
            self.progress_bar.pack_forget()
            self.status_label.pack_forget()

        if self.cancel_scan and not self.matches:
            return
        if not self.matches:
            messagebox.showinfo("Done",
                                "No matches found. Try lowering the Match Strictness slider.")
            return

        self.matches = sorted(self.matches, key=lambda x: x["Score"], reverse=True)
        self.current_index    = 0
        self._raw_export_path = None
        self._export_raw_matches()
        self.save_progress()

        logging.info(f"Opening histogram for {len(self.matches):,} matches…")
        HistogramWindow(self, self.matches, self.switch_to_review)

    # =========================================================
    # RAW MATCH EXPORT
    # =========================================================
    def _export_raw_matches(self):
        if not self.matches:
            return
        stem     = os.path.splitext(os.path.basename(self.file_path))[0]
        prefix   = "Master_Scan" if self.is_master_mode else "Dedupe_Scan"
        out_path = os.path.join(
            os.path.dirname(self.file_path),
            f"{prefix}_{stem}_{ts()}.xlsx"
        )
        rows = []
        for m in self.matches:
            row = {"Score (%)": m["Score"],
                   "Primary Item": m["Match_A"],
                   "Matched Item": m["Match_B"]}
            for k, v in m.get("Row_A", {}).items():
                row[f"A_{k}"] = v
            for k, v in m.get("Row_B", {}).items():
                row[f"B_{k}"] = v
            rows.append(row)
        try:
            pd.DataFrame(rows).to_excel(out_path, index=False)
            logging.info(f"Raw match export saved: {out_path}")
            self._raw_export_path = out_path
        except Exception as e:
            logging.warning(f"Raw export failed: {e}")
            self._raw_export_path = None

    # =========================================================
    # REVIEW SCREEN
    # =========================================================
    def switch_to_review(self):
        self.setup_frame.pack_forget()
        self.review_frame.pack(pady=20, padx=20, fill="both", expand=True)
        self.review_mode_active = True
        self.review_start_time  = time.time()

        if self.is_master_mode:
            self.lbl_b.configure(text="Item B (Master Data)  [→ / 2]:")
            self.btn_keep_b.configure(text="Override with Master\n[→ / 2]",
                                      fg_color="teal", hover_color="darkcyan")
        else:
            self.lbl_b.configure(text="Item B (Duplicate)  [→ / 2]:")
            self.btn_keep_b.configure(text="Merge Duplicate\n[→ / 2]",
                                      fg_color="green", hover_color="darkgreen")
        self.load_current_match()

    def _update_stats(self):
        remaining = len(self.matches) - self.current_index
        reviewed  = self.current_index

        self.lbl_stat_approved.configure(text=f"{len(self.approved_merges):,}")
        self.lbl_stat_flagged.configure(text=f"{len(self.flagged_merges):,}")
        self.lbl_stat_skipped.configure(text=f"{self.skipped_count:,}")
        self.lbl_stat_remaining.configure(text=f"{remaining:,}")

        if reviewed > 0 and self.review_start_time:
            elapsed  = time.time() - self.review_start_time
            est_secs = (elapsed / reviewed) * remaining
            if est_secs < 60:
                self.lbl_stat_time.configure(text=f"{int(est_secs)}s")
            else:
                self.lbl_stat_time.configure(
                    text=f"{int(est_secs // 60)}m {int(est_secs % 60)}s")
        else:
            self.lbl_stat_time.configure(text="–")

    def load_current_match(self):
        if self.current_index >= len(self.matches):
            self.finish_review()
            return

        match     = self.matches[self.current_index]
        score_val = match["Score"]

        self.lbl_progress.configure(
            text=f"Reviewing Match {self.current_index + 1} of {len(self.matches):,}")
        color = "#00b300" if score_val >= 95 else ("orange" if score_val >= 85 else "red")
        self.lbl_score.configure(text=f"Match Score: {score_val}%", text_color=color)

        for box, key in [(self.textbox_a, "Match_A"), (self.textbox_b, "Match_B")]:
            box.configure(state="normal")
            box.delete("0.0", "end")
            box.insert("0.0", match[key])
            box.configure(state="disabled")

        # Clear note field for next decision
        self.entry_note.delete(0, "end")

        row_a  = match.get("Row_A", {})
        row_b  = match.get("Row_B", {})
        master_header = "Master Supplier Data" if self.is_master_mode else "Internal Duplicate"

        context  = "--- ITEM A (Your Data) ---\n"
        context += "  |  ".join(f"{k}: {v}" for k, v in row_a.items())

        if self.col_mappings:
            context += "\n\n--- COLUMN-PAIR COMPARISON ---\n"
            for p_col, m_col in self.col_mappings:
                val_a      = row_a.get(p_col, "—")
                val_b      = row_b.get(m_col, "—")
                pair_score = fuzz.token_sort_ratio(str(val_a), str(val_b))
                context += (f"  {p_col} → {m_col}:  "
                            f"'{val_a}'  vs  '{val_b}'  [{pair_score}%]\n")

        context += f"\n\n--- ITEM B ({master_header}) ---\n"
        context += "  |  ".join(f"{k}: {v}" for k, v in row_b.items())

        self.textbox_context.configure(state="normal")
        self.textbox_context.delete("0.0", "end")
        self.textbox_context.insert("0.0", context)
        self.textbox_context.configure(state="disabled")

        self.btn_undo.configure(
            state="disabled" if not self.decision_history else "normal")
        self._update_stats()

    def make_decision(self, choice):
        if self.current_index >= len(self.matches):
            return

        match   = self.matches[self.current_index]
        row_a   = match.get("Row_A", {})
        row_b   = match.get("Row_B", {})
        id_a    = row_a.get(self.primary_id_col, "") if self.primary_id_col else ""
        id_b    = row_b.get(self.master_id_col,  "") if self.master_id_col  else ""
        id_b_label = "Master_ID" if self.is_master_mode else "Duplicate_ID"
        note    = self.entry_note.get().strip()

        base = {
            "Primary_ID":    id_a,
            id_b_label:      id_b,
            "Primary Item":  match["Match_A"],
            "Matched Item":  match["Match_B"],
            "Score":         match["Score"],
            "Note":          note,
        }

        if choice == "A":
            self.approved_merges.append(
                {**base, "Final Selection": match["Match_A"], "Action": "Kept Original"})
            logging.info(
                f"DECISION [Keep A] #{self.current_index+1} | "
                f"Score={match['Score']} | "
                f"A='{match['Match_A']}' | B='{match['Match_B']}' | Note='{note}'")

        elif choice == "B":
            action = "Mapped to Master" if self.is_master_mode else "Merged Duplicate"
            self.approved_merges.append(
                {**base, "Final Selection": match["Match_B"], "Action": action})
            logging.info(
                f"DECISION [Override B] #{self.current_index+1} | "
                f"Score={match['Score']} | "
                f"A='{match['Match_A']}' | B='{match['Match_B']}' | Note='{note}'")

        elif choice == "F":
            self.flagged_merges.append(
                {**base, "Final Selection": "", "Action": "Flagged for Review"})
            logging.info(
                f"DECISION [Flag] #{self.current_index+1} | "
                f"Score={match['Score']} | "
                f"A='{match['Match_A']}' | B='{match['Match_B']}' | Note='{note}'")

        elif choice == "S":
            self.skipped_count += 1
            logging.debug(
                f"DECISION [Skip] #{self.current_index+1} | "
                f"A='{match['Match_A']}' | B='{match['Match_B']}'")

        self.decision_history.append(choice)
        self.current_index += 1

        if self.current_index % 10 == 0:
            self.save_progress()

        self.load_current_match()

    def undo_decision(self):
        if self.current_index > 0 and self.decision_history:
            self.current_index -= 1
            last = self.decision_history.pop()
            if last == "A" or last == "B":
                self.approved_merges.pop()
            elif last == "F":
                self.flagged_merges.pop()
            elif last == "S":
                self.skipped_count = max(0, self.skipped_count - 1)
            logging.info(f"UNDO: reverted decision [{last}] at index {self.current_index}")
            self.save_progress()
            self.load_current_match()

    # =========================================================
    # PERSISTENCE
    # =========================================================
    def save_progress(self):
        data = {
            "matches":           self.matches,
            "current_index":     self.current_index,
            "approved_merges":   self.approved_merges,
            "flagged_merges":    self.flagged_merges,
            "skipped_count":     self.skipped_count,
            "decision_history":  self.decision_history,
            "is_master_mode":    self.is_master_mode,
            "col_mappings":      self.col_mappings,
            "primary_id_col":    self.primary_id_col,
            "master_id_col":     self.master_id_col,
            "raw_export_path":   self._raw_export_path,
        }
        try:
            with open(self.progress_file, "w") as f:
                json.dump(data, f)
        except Exception as e:
            logging.warning(f"Could not save progress: {e}")

    def load_progress(self):
        with open(self.progress_file, "r") as f:
            data = json.load(f)
        self.matches           = data["matches"]
        self.current_index     = data["current_index"]
        self.approved_merges   = data["approved_merges"]
        self.flagged_merges    = data.get("flagged_merges", [])
        self.skipped_count     = data.get("skipped_count", 0)
        self.decision_history  = data.get("decision_history", [])
        self.is_master_mode    = data.get("is_master_mode", False)
        self.col_mappings      = data.get("col_mappings", [])
        self.primary_id_col    = data.get("primary_id_col", "")
        self.master_id_col     = data.get("master_id_col", "")
        self._raw_export_path  = data.get("raw_export_path")
        logging.info(f"Progress resumed: index={self.current_index}, "
                     f"approved={len(self.approved_merges)}, "
                     f"flagged={len(self.flagged_merges)}")

    def save_and_exit(self):
        self.save_progress()
        messagebox.showinfo("Saved", "Progress saved! You can close and resume anytime.")
        self.destroy()

    # =========================================================
    # FINISH
    # =========================================================
    def finish_review(self):
        self.review_mode_active = False
        prefix   = "Master_Mapping" if self.is_master_mode else "Internal_Dedupe"
        stem     = os.path.splitext(os.path.basename(self.file_path))[0]
        out_dir  = os.path.dirname(self.file_path)
        out_path = os.path.join(out_dir, f"Final_{prefix}_{stem}_{ts()}.xlsx")

        approved_df = pd.DataFrame(self.approved_merges)
        flagged_df  = pd.DataFrame(self.flagged_merges)

        files_saved = []
        if not approved_df.empty or not flagged_df.empty:
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                if not approved_df.empty:
                    approved_df.to_excel(writer, sheet_name="Decisions", index=False)
                if not flagged_df.empty:
                    flagged_df.to_excel(writer, sheet_name="Flagged for Review", index=False)
            files_saved.append(f"Decisions file:\n  {out_path}")
            logging.info(f"Output saved: {out_path} "
                         f"({len(self.approved_merges)} approved, "
                         f"{len(self.flagged_merges)} flagged)")

        if self._raw_export_path:
            files_saved.append(f"Raw scan file:\n  {self._raw_export_path}")

        summary = (
            f"Review complete!\n\n"
            f"  ✅ Approved:  {len(self.approved_merges):,}\n"
            f"  🚩 Flagged:   {len(self.flagged_merges):,}\n"
            f"  ⏭ Skipped:   {self.skipped_count:,}\n\n"
        )
        if files_saved:
            summary += "\n\n".join(files_saved)
        else:
            summary += "No decisions were recorded."

        logging.info(
            f"Session complete — approved={len(self.approved_merges)}, "
            f"flagged={len(self.flagged_merges)}, skipped={self.skipped_count}")

        messagebox.showinfo("Complete!", summary)

        if os.path.exists(self.progress_file):
            os.remove(self.progress_file)

        self.destroy()


if __name__ == "__main__":
    app = DataMatchApp()
    app.mainloop()
