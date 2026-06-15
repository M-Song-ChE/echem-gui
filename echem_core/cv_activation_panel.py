"""CV Activation panel.

Tracks activation CV convergence: plots current at a target potential
vs cycle number, and evaluates whether the last N cycles show < threshold%
change (pass criterion for activation).

Layout
------
Left (scrollable):
    Files · column selectors · IR/RHE corrections ·
    activation settings (E_target, scan direction, convergence window/threshold) ·
    buttons · per-file results table.

Right:
    Upper – CV figure (all cycles of active file, gradient-colored)
    Lower – Cycle vs J@E_target (all loaded files overlaid)

Interactions (both plots):
    Scroll          – zoom centred on cursor
    Left-drag       – pan
    Left-click      – annotate nearest data point
    Right-click     – clear annotation / set E_target (CV plot only, if no annotation)
"""

import os
from collections import OrderedDict

import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.cm as mpl_cm
import matplotlib.colors as mpl_colors

from .file_manager import _read_mpr
from .checklist import CheckableListbox
from .plotting import copy_figure_to_clipboard

# ── Constants ────────────────────────────────────────────────────────────────
_DEF = dict(e_target="0.70", window="10", threshold="2.0", r_sol="0", e_ref="0")
_DIRECTIONS = ["Anodic", "Cathodic", "Average"]
_TRACE_COLORS = [
    "#1f77b4", "#d62728", "#2ca02c", "#9467bd", "#ff7f0e",
    "#8c564b", "#17becf", "#e377c2", "#bcbd22", "#7f7f7f",
]
_CLICK_PX = 8   # pixel radius for repeated-click cycling through candidates


# ── Module-level helpers ─────────────────────────────────────────────────────
def _read_one(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".mpr":
        return _read_mpr(path)
    df = pd.read_csv(path, sep="\t", encoding="latin-1", on_bad_lines="skip")
    df.columns = [c.strip().replace("<", "").replace(">", "") for c in df.columns]
    return df.reset_index(drop=True)


def _split_scans(E: np.ndarray, I: np.ndarray):
    """Split E/I into anodic and cathodic halves (each sorted ascending in E)."""
    if len(E) < 4:
        return E, I, E, I
    n = len(E)
    min_idx = int(np.argmin(E))
    max_idx = int(np.argmax(E))
    if abs(max_idx / n - 0.5) <= abs(min_idx / n - 0.5):
        E_an  = E[:max_idx + 1].copy(); I_an  = I[:max_idx + 1].copy()
        E_cat = E[max_idx:].copy();     I_cat = I[max_idx:].copy()
    else:
        E_cat = E[:min_idx + 1].copy(); I_cat = I[:min_idx + 1].copy()
        E_an  = E[min_idx:].copy();     I_an  = I[min_idx:].copy()
    o = np.argsort(E_an);  E_an,  I_an  = E_an[o],  I_an[o]
    o = np.argsort(E_cat); E_cat, I_cat = E_cat[o], I_cat[o]
    return E_an, I_an, E_cat, I_cat


def _interp_at_e(E: np.ndarray, I: np.ndarray, e_target: float):
    E = np.asarray(E, dtype=float); I = np.asarray(I, dtype=float)
    if len(E) < 2: return None
    if e_target < E.min() or e_target > E.max(): return None
    return float(np.interp(e_target, E, I))


# ══════════════════════════════════════════════════════════════════════════════
class CvActivationPanel(ttk.Frame):
    """Self-contained CV Activation convergence panel."""

    def __init__(self, master):
        ttk.Frame.__init__(self, master)
        self.files       = OrderedDict()
        self.active_file = None
        self._loading    = False
        self._debounce_id = None

        # Annotation state (shared across both plots)
        self._ann          = None
        self._ann_dot      = None
        self._ann_ax       = None
        self._cand_idx     = 0
        self._last_click_pos = None
        # Pan state
        self._panning   = False
        self._pan_ax    = None
        self._pan_start = None
        self._pan_moved = False

        self._build_panel()

    # ════════════════════════════════════════════════════════════════
    # Panel construction
    # ════════════════════════════════════════════════════════════════
    def _build_panel(self):
        body = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True)

        # ── Scrollable left ───────────────────────────────────────
        left_outer = ttk.Frame(body, width=290)
        body.add(left_outer, weight=0)
        _lc = tk.Canvas(left_outer, highlightthickness=0)
        _ls = ttk.Scrollbar(left_outer, orient=tk.VERTICAL, command=_lc.yview)
        _lc.configure(yscrollcommand=_ls.set)
        _ls.pack(side=tk.RIGHT, fill=tk.Y)
        _lc.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        left = tk.Frame(_lc)
        _lwin = _lc.create_window((0, 0), window=left, anchor=tk.NW)
        left.bind("<Configure>", lambda e: _lc.configure(scrollregion=_lc.bbox("all")))
        _lc.bind("<Configure>", lambda e: _lc.itemconfig(_lwin, width=e.width))
        _lc.bind("<MouseWheel>", lambda e: _lc.yview_scroll(-1*(e.delta//120), "units"))

        # ── Files ─────────────────────────────────────────────────
        ttk.Label(left, text="Files", font=("", 9, "bold")).pack(
            anchor=tk.W, padx=4, pady=(6, 0))
        ttk.Label(left, text="Load one file per sample (activation CV series)",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)
        _fb = ttk.Frame(left); _fb.pack(fill=tk.X, padx=4, pady=2)
        ttk.Button(_fb, text="Load File(s)", command=self._load_files).pack(
            side=tk.LEFT, padx=(0, 4))
        ttk.Button(_fb, text="Remove", command=self._remove_file).pack(
            side=tk.LEFT, padx=(0, 4))
        ttk.Button(_fb, text="Merge…", command=self._open_merge_dialog).pack(side=tk.LEFT)
        _flf = ttk.Frame(left); _flf.pack(fill=tk.X, padx=4, pady=2)
        self.file_listbox = CheckableListbox(
            _flf, height=5, show_checkboxes=False, on_reorder=self._on_file_reorder)
        self.file_listbox.pack(fill=tk.X, expand=True)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        # ── Column selectors ──────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Column Mapping", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        ttk.Label(left, text="X-axis (potential):").pack(anchor=tk.W, padx=4, pady=(4, 0))
        self.x_var = tk.StringVar()
        self.x_combo = ttk.Combobox(left, textvariable=self.x_var, state="readonly", width=22)
        self.x_combo.pack(anchor=tk.W, padx=4, pady=2)
        self.x_combo.bind("<<ComboboxSelected>>", lambda e: self._schedule())

        ttk.Label(left, text="Y-axis (current):").pack(anchor=tk.W, padx=4)
        self.y_var = tk.StringVar()
        self.y_combo = ttk.Combobox(left, textvariable=self.y_var, state="readonly", width=22)
        self.y_combo.pack(anchor=tk.W, padx=4, pady=2)
        self.y_combo.bind("<<ComboboxSelected>>", lambda e: self._schedule())

        ttk.Label(left, text="Cycle column:").pack(anchor=tk.W, padx=4)
        self.cyc_var = tk.StringVar()
        self.cyc_combo = ttk.Combobox(left, textvariable=self.cyc_var, state="readonly", width=22)
        self.cyc_combo.pack(anchor=tk.W, padx=4, pady=2)
        self.cyc_combo.bind("<<ComboboxSelected>>", lambda e: self._schedule())

        # ── Corrections ───────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Corrections", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        _ir_row = ttk.Frame(left); _ir_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_ir_row, text="R_sol (Ω):", width=12, anchor=tk.W).pack(side=tk.LEFT)
        self.r_sol_var = tk.StringVar(value=_DEF["r_sol"])
        _ir_e = ttk.Entry(_ir_row, textvariable=self.r_sol_var, width=8)
        _ir_e.pack(side=tk.LEFT, padx=(2, 0))
        _ir_e.bind("<Return>",   lambda e: self._save_corr_and_schedule())
        _ir_e.bind("<FocusOut>", lambda e: self._save_corr_and_schedule())

        _rhe_row = ttk.Frame(left); _rhe_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_rhe_row, text="E_ref offset (V):", width=14, anchor=tk.W).pack(side=tk.LEFT)
        self.e_ref_var = tk.StringVar(value=_DEF["e_ref"])
        _rhe_e = ttk.Entry(_rhe_row, textvariable=self.e_ref_var, width=8)
        _rhe_e.pack(side=tk.LEFT, padx=(2, 0))
        _rhe_e.bind("<Return>",   lambda e: self._save_corr_and_schedule())
        _rhe_e.bind("<FocusOut>", lambda e: self._save_corr_and_schedule())
        ttk.Label(left, text="E_corr = E − I·R_sol + E_ref",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)

        # ── Activation settings ────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Activation Settings", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        _et_row = ttk.Frame(left); _et_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_et_row, text="E target (V):", width=16, anchor=tk.W).pack(side=tk.LEFT)
        self.e_target_var = tk.StringVar(value=_DEF["e_target"])
        _et_e = ttk.Entry(_et_row, textvariable=self.e_target_var, width=8)
        _et_e.pack(side=tk.LEFT, padx=(2, 0))
        _et_e.bind("<Return>",   lambda e: self._schedule())
        _et_e.bind("<FocusOut>", lambda e: self._schedule())

        _dir_row = ttk.Frame(left); _dir_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_dir_row, text="Scan direction:", width=16, anchor=tk.W).pack(side=tk.LEFT)
        self.direction_var = tk.StringVar(value="Anodic")
        ttk.Combobox(_dir_row, textvariable=self.direction_var,
                     values=_DIRECTIONS, state="readonly", width=10).pack(
                         side=tk.LEFT, padx=(2, 0))
        self.direction_var.trace_add("write", lambda *_: self._schedule())

        _win_row = ttk.Frame(left); _win_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_win_row, text="Conv. window (cycles):", width=20, anchor=tk.W).pack(side=tk.LEFT)
        self.window_var = tk.StringVar(value=_DEF["window"])
        _win_e = ttk.Entry(_win_row, textvariable=self.window_var, width=5)
        _win_e.pack(side=tk.LEFT, padx=(2, 0))
        _win_e.bind("<Return>",   lambda e: self._schedule())
        _win_e.bind("<FocusOut>", lambda e: self._schedule())

        _thr_row = ttk.Frame(left); _thr_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_thr_row, text="Threshold (%):", width=16, anchor=tk.W).pack(side=tk.LEFT)
        self.threshold_var = tk.StringVar(value=_DEF["threshold"])
        _thr_e = ttk.Entry(_thr_row, textvariable=self.threshold_var, width=6)
        _thr_e.pack(side=tk.LEFT, padx=(2, 0))
        _thr_e.bind("<Return>",   lambda e: self._schedule())
        _thr_e.bind("<FocusOut>", lambda e: self._schedule())

        ttk.Label(left, text="Pass: |ΔJ over last N cycles| / |J| < threshold%",
                  foreground="gray", font=("", 8), wraplength=270,
                  justify=tk.LEFT).pack(anchor=tk.W, padx=4, pady=(0, 4))

        _ol_row = ttk.Frame(left); _ol_row.pack(fill=tk.X, padx=4, pady=2)
        self.overlay_all_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(_ol_row, text="Show all samples in cycle plot",
                        variable=self.overlay_all_var,
                        command=self._schedule).pack(side=tk.LEFT)

        # ── Display Settings ──────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Display", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        # Layout selector (1x2 = side-by-side, 2x1 = stacked)
        _lay_row = ttk.Frame(left); _lay_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_lay_row, text="Layout:", width=12, anchor=tk.W).pack(side=tk.LEFT)
        self.layout_var = tk.StringVar(value="1x2")
        _lay_cb = ttk.Combobox(_lay_row, textvariable=self.layout_var,
                                values=["1x2", "2x1"], state="readonly", width=6)
        _lay_cb.pack(side=tk.LEFT, padx=(2, 0))
        _lay_cb.bind("<<ComboboxSelected>>", lambda e: self._apply_layout())

        # Plot size (W × H inches, per figure)
        _ps_row = ttk.Frame(left); _ps_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_ps_row, text="Plot size (in):", width=12, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Label(_ps_row, text="W").pack(side=tk.LEFT)
        self.plot_w_var = tk.StringVar(value="6.0")
        _pw_e = ttk.Entry(_ps_row, textvariable=self.plot_w_var, width=5)
        _pw_e.pack(side=tk.LEFT, padx=(1, 6))
        ttk.Label(_ps_row, text="H").pack(side=tk.LEFT)
        self.plot_h_var = tk.StringVar(value="5.0")
        _ph_e = ttk.Entry(_ps_row, textvariable=self.plot_h_var, width=5)
        _ph_e.pack(side=tk.LEFT, padx=(1, 0))
        for _e in (_pw_e, _ph_e):
            _e.bind("<Return>",   lambda ev: self._apply_plot_size())
            _e.bind("<FocusOut>", lambda ev: self._apply_plot_size())

        # Font sizes — separate per element type
        def _fs_field(row_parent, label, var_name, default):
            row = ttk.Frame(row_parent); row.pack(fill=tk.X, padx=4, pady=1)
            ttk.Label(row, text=label, width=14, anchor=tk.W).pack(side=tk.LEFT)
            sv = tk.StringVar(value=str(default))
            setattr(self, var_name, sv)
            e = ttk.Entry(row, textvariable=sv, width=5)
            e.pack(side=tk.LEFT, padx=(2, 0))
            e.bind("<Return>",   lambda ev: self._schedule())
            e.bind("<FocusOut>", lambda ev: self._schedule())

        _fs_field(left, "Title fs:",       "fs_title_var",  11)
        _fs_field(left, "Axis label fs:",  "fs_axis_var",    9)
        _fs_field(left, "Tick fs:",        "fs_tick_var",    8)
        _fs_field(left, "Legend fs:",      "fs_legend_var",  7)
        _fs_field(left, "Annot. fs:",      "fs_annot_var",  10)

        # Custom plot titles
        _tit1_row = ttk.Frame(left); _tit1_row.pack(fill=tk.X, padx=4, pady=(4, 1))
        ttk.Label(_tit1_row, text="CV title:", width=12, anchor=tk.W).pack(side=tk.LEFT)
        self.cv_title_var = tk.StringVar(value="")
        _ct1 = ttk.Entry(_tit1_row, textvariable=self.cv_title_var)
        _ct1.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))
        _ct1.bind("<Return>",   lambda e: self._save_titles_and_schedule())
        _ct1.bind("<FocusOut>", lambda e: self._save_titles_and_schedule())

        _tit2_row = ttk.Frame(left); _tit2_row.pack(fill=tk.X, padx=4, pady=(1, 2))
        ttk.Label(_tit2_row, text="Cycle title:", width=12, anchor=tk.W).pack(side=tk.LEFT)
        self.cyc_title_var = tk.StringVar(value="")
        _ct2 = ttk.Entry(_tit2_row, textvariable=self.cyc_title_var)
        _ct2.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))
        _ct2.bind("<Return>",   lambda e: self._schedule())
        _ct2.bind("<FocusOut>", lambda e: self._schedule())

        ttk.Label(left, text="(blank = auto; CV title is per-file)",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)

        # Axis range — per plot (blank = auto-fit).
        # CV range is per-file (saved into entry on FocusOut/Return);
        # Cycle range is shared across files (panel-level).
        def _range_row(parent, label, var_lo_name, var_hi_name, per_file=False):
            row = ttk.Frame(parent); row.pack(fill=tk.X, padx=4, pady=1)
            ttk.Label(row, text=label, width=12, anchor=tk.W).pack(side=tk.LEFT)
            ttk.Label(row, text="min").pack(side=tk.LEFT)
            lo = tk.StringVar(value="");  setattr(self, var_lo_name, lo)
            e1 = ttk.Entry(row, textvariable=lo, width=7)
            e1.pack(side=tk.LEFT, padx=(2, 4))
            ttk.Label(row, text="max").pack(side=tk.LEFT)
            hi = tk.StringVar(value="");  setattr(self, var_hi_name, hi)
            e2 = ttk.Entry(row, textvariable=hi, width=7)
            e2.pack(side=tk.LEFT, padx=(2, 0))
            if per_file:
                _cb = lambda ev: self._save_corr_and_schedule()
            else:
                _cb = lambda ev: self._schedule()
            for _e in (e1, e2):
                _e.bind("<Return>",   _cb)
                _e.bind("<FocusOut>", _cb)

        ttk.Label(left, text="CV plot range (per-file):",
                  font=("", 8, "italic")).pack(anchor=tk.W, padx=4, pady=(4, 0))
        _range_row(left, "X (E):", "cv_xmin_var", "cv_xmax_var", per_file=True)
        _range_row(left, "Y (I):", "cv_ymin_var", "cv_ymax_var", per_file=True)

        ttk.Label(left, text="Cycle plot range:",
                  font=("", 8, "italic")).pack(anchor=tk.W, padx=4, pady=(4, 0))
        _range_row(left, "X (cycle):", "cyc_xmin_var", "cyc_xmax_var")
        _range_row(left, "Y (J):",     "cyc_ymin_var", "cyc_ymax_var")

        # ── Buttons ───────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        _btn_row = ttk.Frame(left); _btn_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Button(_btn_row, text="Analyze",
                   command=self._update_all).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(_btn_row, text="Clear",
                   command=self._clear_all).pack(side=tk.LEFT)
        _btn_row2 = ttk.Frame(left); _btn_row2.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Button(_btn_row2, text="Apply Settings to All Files",
                   command=self._apply_settings_to_all).pack(side=tk.LEFT)
        ttk.Label(_btn_row2, text="(E target / direction / window / threshold)",
                  foreground="gray", font=("", 7)).pack(side=tk.LEFT, padx=(6, 0))

        # ── Results table ─────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        _rtv_lf = ttk.LabelFrame(left, text="Results")
        _rtv_lf.pack(fill=tk.BOTH, padx=4, pady=4, expand=False)
        _rtv_cols = ("file", "e_tgt", "dir", "w_thr", "j_final", "delta_pct", "status")
        self._results_tv = ttk.Treeview(
            _rtv_lf, columns=_rtv_cols, show="headings", height=6, selectmode="browse")
        self._results_tv.heading("file",      text="Sample")
        self._results_tv.heading("e_tgt",     text="E tgt (V)")
        self._results_tv.heading("dir",       text="Dir")
        self._results_tv.heading("w_thr",     text="W / Thr%")
        self._results_tv.heading("j_final",   text="J_final")
        self._results_tv.heading("delta_pct", text="Δ%")
        self._results_tv.heading("status",    text="Status")
        self._results_tv.column("file",      width=100, anchor=tk.W,      stretch=True)
        self._results_tv.column("e_tgt",     width=62,  anchor=tk.CENTER, stretch=False)
        self._results_tv.column("dir",       width=52,  anchor=tk.CENTER, stretch=False)
        self._results_tv.column("w_thr",     width=65,  anchor=tk.CENTER, stretch=False)
        self._results_tv.column("j_final",   width=62,  anchor=tk.CENTER, stretch=False)
        self._results_tv.column("delta_pct", width=52,  anchor=tk.CENTER, stretch=False)
        self._results_tv.column("status",    width=90,  anchor=tk.CENTER, stretch=False)
        _rtv_sb = ttk.Scrollbar(_rtv_lf, orient=tk.VERTICAL,
                                 command=self._results_tv.yview)
        self._results_tv.configure(yscrollcommand=_rtv_sb.set)
        _rtv_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._results_tv.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self._results_tv.tag_configure("pass", background="#c8e6c9")
        self._results_tv.tag_configure("fail", background="#ffcdd2")

        # ── Right: scrollable area with two figures (layout-driven) ─
        right = ttk.Frame(body)
        body.add(right, weight=1)
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)

        _plot_sc = tk.Canvas(right, highlightthickness=0)
        _right_vs = ttk.Scrollbar(right, orient=tk.VERTICAL,   command=_plot_sc.yview)
        _right_hs = ttk.Scrollbar(right, orient=tk.HORIZONTAL, command=_plot_sc.xview)
        _plot_sc.configure(yscrollcommand=_right_vs.set, xscrollcommand=_right_hs.set)
        _right_vs.grid(row=0, column=1, sticky="ns")
        _right_hs.grid(row=1, column=0, sticky="ew")
        _plot_sc.grid(row=0, column=0, sticky="nsew")
        _plot_sc.bind("<MouseWheel>",
                       lambda e: _plot_sc.yview_scroll(-1*(e.delta//120), "units"))
        _plot_sc.bind("<Shift-MouseWheel>",
                       lambda e: _plot_sc.xview_scroll(-1*(e.delta//120), "units"))
        self._plot_sc = _plot_sc

        _plots_frame = ttk.Frame(_plot_sc)
        _plot_sc.create_window((0, 0), window=_plots_frame, anchor=tk.NW)
        _plots_frame.bind("<Configure>",
                          lambda e: _plot_sc.configure(scrollregion=_plot_sc.bbox("all")))
        self._plots_frame = _plots_frame

        _fw = float(self.plot_w_var.get())
        _fh = float(self.plot_h_var.get())

        # CV figure
        self._cv_frame = ttk.Frame(_plots_frame)
        self._cv_fig = Figure(figsize=(_fw, _fh), dpi=100, constrained_layout=True)
        self._cv_ax  = self._cv_fig.add_subplot(111)
        self._cv_cv  = FigureCanvasTkAgg(self._cv_fig, master=self._cv_frame)
        self._cv_cv.get_tk_widget().pack()
        self._cv_cv.get_tk_widget().config(width=int(_fw * 100), height=int(_fh * 100))
        _cv_tb_row = ttk.Frame(self._cv_frame)
        _cv_tb_row.pack(fill=tk.X)
        self._cv_tb = NavigationToolbar2Tk(self._cv_cv, _cv_tb_row, pack_toolbar=False)
        self._cv_tb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._cv_tb.update()
        tk.Button(_cv_tb_row, text="Copy",
                  command=lambda: copy_figure_to_clipboard(self._cv_fig),
                  relief=tk.RAISED, borderwidth=1, padx=6).pack(
                      side=tk.LEFT, padx=(4, 2), pady=1)

        # Cycle figure
        self._cyc_frame = ttk.Frame(_plots_frame)
        self._cyc_fig = Figure(figsize=(_fw, _fh), dpi=100, constrained_layout=True)
        self._cyc_ax  = self._cyc_fig.add_subplot(111)
        self._cyc_cv  = FigureCanvasTkAgg(self._cyc_fig, master=self._cyc_frame)
        self._cyc_cv.get_tk_widget().pack()
        self._cyc_cv.get_tk_widget().config(width=int(_fw * 100), height=int(_fh * 100))
        _cyc_tb_row = ttk.Frame(self._cyc_frame)
        _cyc_tb_row.pack(fill=tk.X)
        self._cyc_tb = NavigationToolbar2Tk(self._cyc_cv, _cyc_tb_row, pack_toolbar=False)
        self._cyc_tb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._cyc_tb.update()
        tk.Button(_cyc_tb_row, text="Copy",
                  command=lambda: copy_figure_to_clipboard(self._cyc_fig),
                  relief=tk.RAISED, borderwidth=1, padx=6).pack(
                      side=tk.LEFT, padx=(4, 2), pady=1)

        # Initial layout: 1x2
        self._apply_layout()

        # ── Connect events ────────────────────────────────────────
        # Legend tracking — for hit testing in mouse handlers
        self._cv_legend  = None
        self._cyc_legend = None
        for cv in (self._cv_cv, self._cyc_cv):
            cv.mpl_connect("scroll_event",        self._on_scroll)
            cv.mpl_connect("button_press_event",  self._on_press)
            cv.mpl_connect("button_release_event",self._on_release)
            cv.mpl_connect("motion_notify_event", self._on_motion)

    # ════════════════════════════════════════════════════════════════
    # File management
    # ════════════════════════════════════════════════════════════════
    def _load_files(self):
        paths = filedialog.askopenfilenames(
            title="Load Activation CV File(s)",
            filetypes=[("Data files", "*.txt *.csv *.mpr *.mpt"), ("All files", "*.*")])
        if not paths:
            return
        for path in paths:
            short = os.path.basename(path)
            if short in self.files:
                continue
            try:
                df = _read_one(path)
            except Exception as exc:
                messagebox.showerror("Load error", f"{short}:\n{exc}")
                continue
            color_idx = len(self.files) % len(_TRACE_COLORS)
            entry = {
                "path": path, "df": df,
                "r_sol": 0.0, "e_ref": 0.0,
                "x_col": "", "y_col": "", "cyc_col": "",
                "e_target": "", "direction": "Anodic",
                "window": "10", "threshold": "2.0",
                "color": _TRACE_COLORS[color_idx],
                "result": None,
                "custom_cv_title": "",
                "cv_xmin": "", "cv_xmax": "",
                "cv_ymin": "", "cv_ymax": "",
            }
            self._init_entry_defaults(entry)   # auto-detect cols + e_target
            self.files[short] = entry
            self._loading = True
            self.file_listbox.insert(tk.END, short)
            self._loading = False

        if self.files:
            self._switch_file(list(self.files.keys())[-1])

    def _remove_file(self):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        idx  = sel[0]
        keys = list(self.files.keys())
        if idx >= len(keys):
            return
        del self.files[keys[idx]]
        self.file_listbox.delete(idx)
        keys = list(self.files.keys())
        if keys:
            self._switch_file(keys[min(idx, len(keys) - 1)])
        else:
            self.active_file = None
            self._clear_plots()
        self._rebuild_results_tv()

    def _on_file_reorder(self, new_order):
        new_files = OrderedDict()
        for name in new_order:
            if name in self.files:
                new_files[name] = self.files[name]
        for k, v in self.files.items():
            if k not in new_files:
                new_files[k] = v
        self.files = new_files
        self._schedule()

    def _on_file_select(self, _=None):
        if self._loading:
            return
        sel = self.file_listbox.curselection()
        if not sel:
            return
        keys = list(self.files.keys())
        if sel[0] < len(keys):
            self._switch_file(keys[sel[0]])

    def _switch_file(self, short):
        self._save_active_state()          # save before leaving current file
        self.active_file = short
        keys = list(self.files.keys())
        if short in keys:
            self._loading = True
            self.file_listbox.selection_clear(0, tk.END)
            self.file_listbox.selection_set(keys.index(short))
            self._loading = False
        self._update_column_combos()
        self._restore_state()
        self._schedule()

    def _apply_settings_to_all(self):
        """Broadcast current activation params to every file (corrections excluded)."""
        self._save_active_state()
        src = self.files.get(self.active_file)
        if not src:
            return
        for entry in self.files.values():
            entry["e_target"]  = src["e_target"]
            entry["direction"] = src["direction"]
            entry["window"]    = src["window"]
            entry["threshold"] = src["threshold"]
        self._update_all()

    def _save_active_state(self):
        """Save current UI values into the active file's entry."""
        entry = self.files.get(self.active_file)
        if not entry:
            return
        try: entry["r_sol"] = float(self.r_sol_var.get())
        except ValueError: pass
        try: entry["e_ref"] = float(self.e_ref_var.get())
        except ValueError: pass
        entry["e_target"]  = self.e_target_var.get()
        entry["direction"] = self.direction_var.get()
        entry["window"]    = self.window_var.get()
        entry["threshold"] = self.threshold_var.get()
        entry["x_col"]     = self.x_var.get()
        entry["y_col"]     = self.y_var.get()
        entry["cyc_col"]   = self.cyc_var.get()
        if hasattr(self, "cv_title_var"):
            entry["custom_cv_title"] = self.cv_title_var.get()
        # Per-file CV plot axis range (blank = auto)
        if hasattr(self, "cv_xmin_var"):
            entry["cv_xmin"] = self.cv_xmin_var.get()
            entry["cv_xmax"] = self.cv_xmax_var.get()
            entry["cv_ymin"] = self.cv_ymin_var.get()
            entry["cv_ymax"] = self.cv_ymax_var.get()

    def _init_entry_defaults(self, entry):
        """Auto-detect columns and e_target for a freshly created entry (no UI touch)."""
        df   = entry["df"]
        cols = list(df.columns)
        # x column
        xcol = ""
        for c in ["Ewe/V", "Ewe/V ", "E/V", "E (V)"]:
            if c in cols: xcol = c; break
        if not xcol:
            for c in cols:
                if "/V" in c or "(V)" in c: xcol = c; break
        # y column
        ycol = ""
        for c in ["I/mA", "<I>/mA", "I (mA)"]:
            if c in cols: ycol = c; break
        if not ycol:
            for c in cols:
                if "/mA" in c or "(mA)" in c or "/A" in c: ycol = c; break
        # cycle column
        ccol = "cycle number" if "cycle number" in cols else None
        entry["x_col"]   = xcol
        entry["y_col"]   = ycol
        entry["cyc_col"] = ccol if ccol else "(none)"
        # e_target
        if not xcol or not ycol:
            return
        sub = df
        if ccol:
            last_cn = sorted(df[ccol].unique())[-1]
            sub = df[df[ccol] == last_cn]
        E = sub[xcol].dropna().values
        I = sub[ycol].dropna().values
        n = min(len(E), len(I))
        if n < 4:
            return
        E_an, I_an, _, _ = _split_scans(E[:n], I[:n])
        if len(I_an) == 0:
            return
        entry["e_target"] = f"{float(E_an[np.argmax(I_an)]):.4f}"

    def _auto_set_e_target(self):
        """Refresh e_target for the active file using current UI column selection."""
        entry = self.files.get(self.active_file)
        if not entry:
            return
        xcol = self.x_var.get()
        ycol = self.y_var.get()
        if not xcol or not ycol or xcol not in entry["df"].columns or ycol not in entry["df"].columns:
            return
        df   = entry["df"]
        ccol = self._get_cycle_col(df)
        if ccol:
            last_cn = sorted(df[ccol].unique())[-1]
            sub = df[df[ccol] == last_cn]
        else:
            sub = df
        E = sub[xcol].dropna().values
        I = sub[ycol].dropna().values
        n = min(len(E), len(I))
        if n < 4:
            return
        E, I = E[:n], I[:n]
        E_an, I_an, _, _ = _split_scans(E, I)
        if len(I_an) == 0:
            return
        e_at_max_anodic = float(E_an[np.argmax(I_an)])
        self.e_target_var.set(f"{e_at_max_anodic:.4f}")
        entry["e_target"] = self.e_target_var.get()

    # ════════════════════════════════════════════════════════════════
    # Column management
    # ════════════════════════════════════════════════════════════════
    def _update_column_combos(self):
        entry = self.files.get(self.active_file)
        if not entry:
            return
        cols = list(entry["df"].columns)
        self.x_combo["values"] = cols
        self.y_combo["values"] = cols
        self.cyc_combo["values"] = ["(none)"] + cols
        # Restore saved column or auto-detect
        sx = entry.get("x_col", "")
        if sx and sx in cols:
            self.x_var.set(sx)
        elif self.x_var.get() not in cols:
            for c in ["Ewe/V", "Ewe/V ", "E/V", "E (V)"]:
                if c in cols: self.x_var.set(c); break
            else:
                for c in cols:
                    if "/V" in c or "(V)" in c: self.x_var.set(c); break
        sy = entry.get("y_col", "")
        if sy and sy in cols:
            self.y_var.set(sy)
        elif self.y_var.get() not in cols:
            for c in ["I/mA", "<I>/mA", "I (mA)"]:
                if c in cols: self.y_var.set(c); break
            else:
                for c in cols:
                    if "/mA" in c or "(mA)" in c or "/A" in c: self.y_var.set(c); break
        sc = entry.get("cyc_col", "")
        if sc:   # previously saved for this file — restore exactly
            self.cyc_var.set(sc if sc in ["(none)"] + cols else "(none)")
        else:    # never saved → always auto-detect so "(none)" from another file doesn't stick
            self.cyc_var.set("cycle number" if "cycle number" in cols else "(none)")

    def _get_cycle_col(self, df):
        c = self.cyc_var.get()
        return c if (c and c != "(none)" and c in df.columns) else None

    # ════════════════════════════════════════════════════════════════
    # Corrections
    # ════════════════════════════════════════════════════════════════
    def _restore_state(self):
        """Restore all per-file UI values from entry."""
        entry = self.files.get(self.active_file)
        if not entry:
            return
        self.r_sol_var.set(str(entry.get("r_sol", 0.0)))
        self.e_ref_var.set(str(entry.get("e_ref", 0.0)))
        # Always restore e_target — even if empty. If we skip when empty,
        # the previous file's value stays in the UI and gets saved to this file.
        self.e_target_var.set(entry.get("e_target", ""))
        self.direction_var.set(entry.get("direction", "Anodic"))
        self.window_var.set(entry.get("window", "10"))
        self.threshold_var.set(entry.get("threshold", "2.0"))
        if hasattr(self, "cv_title_var"):
            self.cv_title_var.set(entry.get("custom_cv_title", "") or "")
        # Per-file CV plot axis range (blank = auto)
        if hasattr(self, "cv_xmin_var"):
            self.cv_xmin_var.set(entry.get("cv_xmin", "") or "")
            self.cv_xmax_var.set(entry.get("cv_xmax", "") or "")
            self.cv_ymin_var.set(entry.get("cv_ymin", "") or "")
            self.cv_ymax_var.set(entry.get("cv_ymax", "") or "")

    def _save_corr_and_schedule(self):
        self._save_active_state()
        self._schedule()

    def _apply_correction(self, df, r_sol, e_ref, xcol=None, ycol=None):
        xcol = xcol or self.x_var.get()
        ycol = ycol or self.y_var.get()
        if not xcol or not ycol or xcol not in df.columns or ycol not in df.columns:
            return df
        df = df.copy()
        if r_sol != 0:
            df[xcol] = df[xcol].values - df[ycol].values * 1e-3 * r_sol
        if e_ref != 0:
            df[xcol] = df[xcol].values + e_ref
        return df

    # ════════════════════════════════════════════════════════════════
    # Debounced real-time update
    # ════════════════════════════════════════════════════════════════
    def _schedule(self, *_):
        if self._debounce_id is not None:
            try: self.after_cancel(self._debounce_id)
            except Exception: pass
        self._debounce_id = self.after(300, self._update_all)

    def _update_all(self):
        self._debounce_id = None
        self._replot_cv()
        self._run_analysis()
        self._replot_cycle()

    # ════════════════════════════════════════════════════════════════
    # Core extraction + analysis
    # ════════════════════════════════════════════════════════════════
    def _extract_cycle_j(self, df_c, e_target: float, direction: str,
                         xcol=None, ycol=None, ccol=None):
        xcol = xcol or self.x_var.get()
        ycol = ycol or self.y_var.get()
        if ccol is None:
            ccol = self._get_cycle_col(df_c)
        if not xcol or not ycol:
            return []
        raw_cycles = sorted(df_c[ccol].unique()) if ccol else [None]
        result = []
        for cn in raw_cycles:
            sub = df_c[df_c[ccol] == cn] if cn is not None else df_c
            E = sub[xcol].dropna().values; I = sub[ycol].dropna().values
            n = min(len(E), len(I))
            if n < 4: continue
            E, I = E[:n], I[:n]
            if direction == "Anodic":
                E_u, I_u, _, _ = _split_scans(E, I)
            elif direction == "Cathodic":
                _, _, E_u, I_u = _split_scans(E, I)
            else:
                E_an, I_an, E_cat, I_cat = _split_scans(E, I)
                vals = [v for v in [_interp_at_e(E_an, I_an, e_target),
                                    _interp_at_e(E_cat, I_cat, e_target)]
                        if v is not None]
                if vals:
                    result.append((int(cn) if cn is not None else 0,
                                   float(np.mean(vals))))
                continue
            j = _interp_at_e(E_u, I_u, e_target)
            if j is not None:
                result.append((int(cn) if cn is not None else 0, j))
        return result

    def _check_convergence(self, cycle_j: list, window: int, threshold: float):
        out = []
        for i, (cn, j) in enumerate(cycle_j):
            if i < window:
                out.append((cn, j, None, None))
            else:
                j_prev = cycle_j[i - window][1]
                delta  = abs(j - j_prev) / max(abs(j_prev), 1e-12) * 100.0
                out.append((cn, j, delta, delta < threshold))
        return out

    def _run_analysis(self):
        """Run extraction + convergence for all files using per-file settings."""
        # Save current UI first so active file's settings are up to date
        self._save_active_state()
        for short, entry in self.files.items():
            try:
                e_target  = float(entry.get("e_target") or 0)
                window    = int(entry.get("window", "10"))
                threshold = float(entry.get("threshold", "2.0"))
            except (ValueError, TypeError):
                entry["result"] = None; continue
            direction = entry.get("direction", "Anodic")
            xcol = entry.get("x_col") or self.x_var.get()
            ycol = entry.get("y_col") or self.y_var.get()
            ccol_name = entry.get("cyc_col", "")
            ccol = ccol_name if (ccol_name and ccol_name != "(none)"
                                 and ccol_name in entry["df"].columns) else None
            df_c    = self._apply_correction(entry["df"], entry["r_sol"], entry["e_ref"],
                                             xcol=xcol, ycol=ycol)
            cycle_j = self._extract_cycle_j(df_c, e_target, direction,
                                            xcol=xcol, ycol=ycol, ccol=ccol)
            if not cycle_j:
                entry["result"] = None
                continue
            conv = self._check_convergence(cycle_j, window, threshold)
            with_delta = [(cn, j, dp, p) for cn, j, dp, p in conv if dp is not None]
            if with_delta:
                cn_f, j_f, dp_f, passed_f = with_delta[-1]
            else:
                cn_f, j_f, dp_f, passed_f = conv[-1][0], conv[-1][1], None, None
            entry["result"] = {
                "cycle_j": cycle_j, "conv": conv,
                "j_final": j_f, "delta_pct": dp_f, "passed": passed_f,
            }
        self._rebuild_results_tv()

    def _rebuild_results_tv(self):
        tv = self._results_tv
        tv.delete(*tv.get_children())
        for short, entry in self.files.items():
            et  = entry.get("e_target", "") or "—"
            dr  = entry.get("direction", "Anodic")[:2]   # "An" / "Ca" / "Av"
            w   = entry.get("window",    "10")
            thr = entry.get("threshold", "2.0")
            w_thr = f"{w} / {thr}%"
            res = entry.get("result")
            if res is None:
                tv.insert("", tk.END, values=(short, et, dr, w_thr, "—", "—", "—"))
                continue
            j_f = f"{res['j_final']:.4f}" if res["j_final"] is not None else "—"
            dp  = f"{res['delta_pct']:.2f}%" if res["delta_pct"] is not None else "—"
            if res["passed"] is True:
                status, tag = "Activated ✓", "pass"
            elif res["passed"] is False:
                status, tag = "Not activated ✗", "fail"
            else:
                status, tag = "—", ""
            tv.insert("", tk.END, values=(short, et, dr, w_thr, j_f, dp, status),
                      tags=(tag,))

    # ════════════════════════════════════════════════════════════════
    # Display helpers: layout, plot-size, fonts
    # ════════════════════════════════════════════════════════════════
    def _fs(self, var, default):
        """Read a font-size StringVar with fallback to default."""
        try:
            return max(4, int(float(var.get())))
        except (ValueError, TypeError, AttributeError):
            return default

    def _apply_layout(self):
        """Arrange CV/cycle frames inside _plots_frame per layout_var."""
        for w in (self._cv_frame, self._cyc_frame):
            w.grid_forget()
        if self.layout_var.get() == "2x1":
            self._cv_frame.grid(row=0,  column=0, padx=4, pady=4, sticky="nw")
            self._cyc_frame.grid(row=1, column=0, padx=4, pady=4, sticky="nw")
        else:  # 1x2
            self._cv_frame.grid(row=0,  column=0, padx=4, pady=4, sticky="nw")
            self._cyc_frame.grid(row=0, column=1, padx=4, pady=4, sticky="nw")
        self._plot_sc.after(
            50, lambda: self._plot_sc.configure(
                scrollregion=self._plot_sc.bbox("all")))

    def _apply_plot_size(self, event=None):
        """Resize both figures to plot_w_var × plot_h_var (inches)."""
        try:
            w = float(self.plot_w_var.get())
            h = float(self.plot_h_var.get())
        except ValueError:
            return
        w = max(2.0, min(50.0, w))
        h = max(1.5, min(50.0, h))
        dpi = 100
        for fig, canvas in ((self._cv_fig, self._cv_cv),
                            (self._cyc_fig, self._cyc_cv)):
            fig.set_size_inches(w, h)
            canvas.get_tk_widget().config(width=int(w * dpi), height=int(h * dpi))
            canvas.draw_idle()
        self._plot_sc.after(
            50, lambda: self._plot_sc.configure(
                scrollregion=self._plot_sc.bbox("all")))

    def _save_titles_and_schedule(self):
        """Save the CV title into the active file's entry, then replot."""
        entry = self.files.get(self.active_file)
        if entry is not None:
            entry["custom_cv_title"] = self.cv_title_var.get()
        self._schedule()

    def _apply_axis_range(self, ax, xmin_var, xmax_var, ymin_var, ymax_var):
        """Apply user-entered axis ranges. Blank fields keep matplotlib's
        auto-computed limits."""
        try:
            xlo = float(xmin_var.get())
            xhi = float(xmax_var.get())
            if xlo < xhi: ax.set_xlim(xlo, xhi)
        except (ValueError, AttributeError):
            pass
        try:
            ylo = float(ymin_var.get())
            yhi = float(ymax_var.get())
            if ylo < yhi: ax.set_ylim(ylo, yhi)
        except (ValueError, AttributeError):
            pass

    # ════════════════════════════════════════════════════════════════
    # Plotting
    # ════════════════════════════════════════════════════════════════
    def _replot_cv(self):
        """Redraw upper CV — clf() + re-add subplot avoids colorbar accumulation."""
        self._clear_ann(redraw=False)
        self._cv_fig.clf()
        self._cv_ax = self._cv_fig.add_subplot(111)
        ax = self._cv_ax

        entry = self.files.get(self.active_file)
        if not entry:
            ax.set_title("CV  (no file loaded)",
                         fontsize=self._fs(self.fs_title_var, 11))
            self._cv_cv.draw()
            self._cv_tb.update(); self._cv_tb.push_current()
            return

        xcol = self.x_var.get(); ycol = self.y_var.get()
        if not xcol or not ycol:
            self._cv_cv.draw()
            self._cv_tb.update(); self._cv_tb.push_current()
            return

        try: r_sol = float(self.r_sol_var.get())
        except ValueError: r_sol = 0.0
        try: e_ref = float(self.e_ref_var.get())
        except ValueError: e_ref = 0.0

        df_c  = self._apply_correction(entry["df"], r_sol, e_ref)
        ccol  = self._get_cycle_col(df_c)
        raw_cycles = sorted(df_c[ccol].unique()) if ccol else [None]
        n_cyc = len(raw_cycles)
        cmap  = mpl_cm.get_cmap("viridis", max(n_cyc, 2))

        for i, cn in enumerate(raw_cycles):
            sub   = df_c[df_c[ccol] == cn] if cn is not None else df_c
            # n_cyc==1 means no cycle column → use file's assigned colour, not viridis purple
            color = entry["color"] if n_cyc == 1 else cmap(i / max(n_cyc - 1, 1))
            lbl   = f"C{int(cn)}" if cn is not None else "data"
            ax.plot(sub[xcol].values, sub[ycol].values,
                    lw=1.0, color=color, label=lbl, alpha=0.85)

        try:
            e_t = float(self.e_target_var.get())
            ax.axvline(e_t, color="red", lw=1.2, ls="--", alpha=0.7,
                       label=f"E={e_t:.3f} V")
        except ValueError:
            pass

        ref_suffix = " vs RHE" if e_ref != 0 else ""
        fs_axis  = self._fs(self.fs_axis_var,    9)
        fs_title = self._fs(self.fs_title_var,  11)
        fs_tick  = self._fs(self.fs_tick_var,    8)
        fs_lgd   = self._fs(self.fs_legend_var,  7)
        ax.set_xlabel(f"{xcol}{ref_suffix}", fontsize=fs_axis)
        ax.set_ylabel(ycol, fontsize=fs_axis)
        custom_cv = entry.get("custom_cv_title", "") or ""
        default_cv = f"Activation CV — {self.active_file}  ({n_cyc} cycles)"
        ax.set_title(custom_cv if custom_cv.strip() else default_cv, fontsize=fs_title)
        ax.tick_params(labelsize=fs_tick)

        self._cv_legend = None
        if n_cyc <= 20:
            leg = ax.legend(fontsize=fs_lgd,
                            ncol=max(1, n_cyc // 8), frameon=True, loc="best")
            if leg is not None:
                try: leg.set_in_layout(False)
                except Exception: pass
                leg.set_draggable(True)
                self._cv_legend = leg
        else:
            cb = self._cv_fig.colorbar(
                mpl_cm.ScalarMappable(
                    norm=mpl_colors.Normalize(1, n_cyc), cmap="viridis"),
                ax=ax, shrink=0.8, pad=0.01, fraction=0.04)
            cb.set_label("Cycle #", fontsize=fs_axis)
            cb.ax.tick_params(labelsize=fs_tick)

        # Apply user-set axis range (blank = auto)
        self._apply_axis_range(ax,
                               self.cv_xmin_var, self.cv_xmax_var,
                               self.cv_ymin_var, self.cv_ymax_var)
        self._cv_cv.draw()           # synchronous so limits are set before toolbar sees them
        self._cv_tb.update()         # clear stale nav stack
        self._cv_tb.push_current()   # register current limits as Home view

    def _replot_cycle(self):
        """Redraw lower cycle-vs-J figure."""
        self._clear_ann(redraw=False)
        self._cyc_fig.clf()
        self._cyc_ax = self._cyc_fig.add_subplot(111)
        ax = self._cyc_ax

        fs_axis  = self._fs(self.fs_axis_var,    9)
        fs_title = self._fs(self.fs_title_var,  11)
        fs_tick  = self._fs(self.fs_tick_var,    8)
        fs_lgd   = self._fs(self.fs_legend_var,  7)
        fs_ann   = self._fs(self.fs_annot_var,  10)

        show_all  = self.overlay_all_var.get()
        files_to_show = (list(self.files.keys()) if show_all
                         else ([self.active_file] if self.active_file else []))

        any_data = False
        all_js = []
        for short in files_to_show:
            entry = self.files.get(short)
            if not entry:
                continue
            try:
                e_target  = float(entry.get("e_target") or 0)
                window    = int(entry.get("window", "10"))
                threshold = float(entry.get("threshold", "2.0"))
            except (ValueError, TypeError):
                continue
            direction = entry.get("direction", "Anodic")
            xcol = entry.get("x_col") or self.x_var.get()
            ycol = entry.get("y_col") or self.y_var.get()
            ccol_name = entry.get("cyc_col", "")
            ccol = ccol_name if (ccol_name and ccol_name != "(none)"
                                 and ccol_name in entry["df"].columns) else None
            df_c    = self._apply_correction(entry["df"], entry["r_sol"], entry["e_ref"],
                                             xcol=xcol, ycol=ycol)
            cycle_j = self._extract_cycle_j(df_c, e_target, direction,
                                            xcol=xcol, ycol=ycol, ccol=ccol)
            if not cycle_j:
                continue
            conv  = self._check_convergence(cycle_j, window, threshold)
            cns   = [x[0] for x in conv]
            js    = [x[1] for x in conv]
            all_js.extend(js)
            lbl = f"{short}  E={e_target:.3f}V" if len(files_to_show) > 1 else short
            ax.plot(cns, js, "o-", color=entry["color"], lw=1.6, ms=4, label=lbl)
            # Mark last convergence delta — draggable so user can move out-of-bounds text
            if len(conv) > window:
                cn_l, j_l, dp, passed = conv[-1]
                mk  = "✓" if passed else "✗"
                col = "#2e7d32" if passed else "#c62828"
                ann = ax.annotate(f"{mk} {dp:.1f}%", xy=(cn_l, j_l),
                                  xytext=(6, 4), textcoords="offset points",
                                  fontsize=fs_ann, fontweight="bold", color=col)
                try:
                    ann.draggable()
                except Exception:
                    pass
            any_data = True

        # Y-axis label: use active file's E_target if single file, else generic
        try:
            act_et = float(self.files.get(self.active_file, {}).get("e_target") or 0)
            y_lbl = f"J at E_target  (mA)"
        except (ValueError, TypeError):
            y_lbl = "J at E_target  (mA)"

        self._cyc_legend = None
        if any_data:
            ax.set_xlabel("Cycle number", fontsize=fs_axis)
            ax.set_ylabel(y_lbl, fontsize=fs_axis)
            try:
                act_w = int(self.files.get(self.active_file, {}).get("window", "10"))
                act_th = float(self.files.get(self.active_file, {}).get("threshold", "2.0"))
            except (ValueError, TypeError):
                act_w, act_th = 10, 2.0
            custom_cyc = (self.cyc_title_var.get() if hasattr(self, "cyc_title_var") else "") or ""
            default_cyc = f"Convergence check  (window={act_w} cyc, threshold={act_th}%)"
            ax.set_title(custom_cyc if custom_cyc.strip() else default_cyc, fontsize=fs_title)
            ax.tick_params(labelsize=fs_tick)
            if len(files_to_show) > 1:
                leg = ax.legend(fontsize=fs_lgd, frameon=True)
                if leg is not None:
                    try: leg.set_in_layout(False)
                    except Exception: pass
                    leg.set_draggable(True)
                    self._cyc_legend = leg
            ax.grid(True, alpha=0.3)
            # Y-axis: fit data (don't force 0 into range)
            if all_js:
                j_lo = min(all_js); j_hi = max(all_js)
                margin = max(abs(j_hi - j_lo) * 0.08, abs(j_lo) * 0.02, 0.01)
                ax.set_ylim(j_lo - margin, j_hi + margin)
        else:
            ax.set_title("Cycle vs J  (no data — check E_target or column mapping)",
                         fontsize=fs_title)

        # Apply user-set axis range (overrides auto-fit ylim above when set)
        self._apply_axis_range(ax,
                               self.cyc_xmin_var, self.cyc_xmax_var,
                               self.cyc_ymin_var, self.cyc_ymax_var)
        self._cyc_cv.draw()
        self._cyc_tb.update()
        self._cyc_tb.push_current()

    # ════════════════════════════════════════════════════════════════
    # Mouse interactions — scroll / pan / annotate
    # ════════════════════════════════════════════════════════════════
    def _get_canvas(self, ax):
        return self._cv_cv if ax is self._cv_ax else self._cyc_cv

    def _on_scroll(self, event):
        ax = event.inaxes
        if ax is None:
            return
        factor = 0.85 if event.button == "up" else 1.0 / 0.85
        xl, yl = ax.get_xlim(), ax.get_ylim()
        xd, yd = event.xdata, event.ydata
        ax.set_xlim(xd + (xl[0] - xd) * factor, xd + (xl[1] - xd) * factor)
        ax.set_ylim(yd + (yl[0] - yd) * factor, yd + (yl[1] - yd) * factor)
        self._get_canvas(ax).draw_idle()

    def _legend_for_ax(self, ax):
        if ax is self._cv_ax:  return self._cv_legend
        if ax is self._cyc_ax: return self._cyc_legend
        return None

    def _event_on_legend(self, event):
        """True if click event is inside the visible legend bbox for its axes."""
        if event.inaxes is None:
            return False
        leg = self._legend_for_ax(event.inaxes)
        if leg is None or not leg.get_visible():
            return False
        try:
            renderer = self._get_canvas(event.inaxes).get_renderer()
            bbox = leg.get_window_extent(renderer)
            return bbox.contains(event.x, event.y)
        except Exception:
            return False

    def _on_press(self, event):
        if event.button == 1 and event.inaxes:
            # Don't start panning when click is on the (draggable) legend —
            # let matplotlib's set_draggable handle that.
            if self._event_on_legend(event):
                return
            self._panning   = True
            self._pan_moved = False
            self._pan_ax    = event.inaxes
            self._pan_start = (event.xdata, event.ydata)

    def _on_motion(self, event):
        if not self._panning or self._pan_ax is None:
            return
        if event.xdata is None or event.ydata is None:
            return
        ax = self._pan_ax
        dx = self._pan_start[0] - event.xdata
        dy = self._pan_start[1] - event.ydata
        if abs(dx) > 1e-12 or abs(dy) > 1e-12:
            self._pan_moved = True
        xl = ax.get_xlim(); yl = ax.get_ylim()
        ax.set_xlim(xl[0] + dx, xl[1] + dx)
        ax.set_ylim(yl[0] + dy, yl[1] + dy)
        self._get_canvas(ax).draw_idle()

    def _on_release(self, event):
        was_panning = self._panning
        pan_moved   = self._pan_moved
        self._panning = False
        self._pan_ax  = None

        if event.button == 1 and not pan_moved and event.inaxes:
            self._annotate(event)
        elif event.button == 3 and event.inaxes:
            if self._ann is not None:
                # Right-click clears annotation
                self._clear_ann()
            elif event.inaxes is self._cv_ax and event.xdata is not None:
                # Right-click on CV (no annotation active) → set E_target
                self.e_target_var.set(f"{event.xdata:.4f}")
                self._schedule()

    def _annotate(self, event):
        ax = event.inaxes
        lines = [ln for ln in ax.lines
                 if len(ln.get_xdata()) > 0 and ln.get_visible()
                 and not ln.get_label().startswith("_")]
        if not lines:
            return
        candidates = []
        for ln in lines:
            xd   = np.asarray(ln.get_xdata(), dtype=float)
            yd   = np.asarray(ln.get_ydata(), dtype=float)
            mask = np.isfinite(xd) & np.isfinite(yd)
            if not mask.any():
                continue
            disp  = ax.transData.transform(np.column_stack([xd[mask], yd[mask]]))
            dists = np.hypot(disp[:, 0] - event.x, disp[:, 1] - event.y)
            best  = int(np.argmin(dists))
            candidates.append((float(dists[best]), ln,
                                float(xd[mask][best]), float(yd[mask][best])))
        if not candidates:
            return
        candidates.sort(key=lambda t: t[0])
        if (self._last_click_pos is not None
                and abs(event.x - self._last_click_pos[0]) <= _CLICK_PX
                and abs(event.y - self._last_click_pos[1]) <= _CLICK_PX):
            self._cand_idx = (self._cand_idx + 1) % len(candidates)
        else:
            self._cand_idx = 0
        self._last_click_pos = (event.x, event.y)
        n                    = len(candidates)
        _, ln, x, y          = candidates[self._cand_idx]
        label                = ln.get_label() or "?"
        xl, yl = ax.get_xlim(), ax.get_ylim()
        xf = (x - xl[0]) / (xl[1] - xl[0]) if xl[1] != xl[0] else 0.5
        yf = (y - yl[0]) / (yl[1] - yl[0]) if yl[1] != yl[0] else 0.5
        xoff = -95 if xf > 0.65 else 15
        yoff = -60 if yf > 0.65 else 15
        hint = f"  [{self._cand_idx + 1}/{n}]" if n > 1 else ""
        text = f"x = {x:.4g}\ny = {y:.4g}\n{label}{hint}"
        if n > 1 and self._cand_idx == 0:
            text += "\n↻ click again to cycle"
        self._clear_ann(redraw=False)
        self._ann_ax = ax
        self._ann = ax.annotate(
            text, xy=(x, y), xytext=(xoff, yoff), textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.4", fc="lightyellow", ec="gray", alpha=0.92),
            arrowprops=dict(arrowstyle="->", color="gray", lw=1.2),
            fontsize=self._fs(self.fs_annot_var, 10), zorder=10)
        self._ann.set_in_layout(False)
        self._ann_dot, = ax.plot(x, y, "o", color=ln.get_color(),
                                 markersize=7, zorder=11, label="_ann_dot")
        self._get_canvas(ax).draw_idle()

    def _clear_ann(self, redraw=True):
        canvas = self._get_canvas(self._ann_ax) if self._ann_ax is not None else None
        for artist in (self._ann, self._ann_dot):
            if artist is not None:
                try: artist.remove()
                except Exception: pass
        self._ann = self._ann_dot = None
        self._last_click_pos = None
        self._cand_idx       = 0
        self._ann_ax         = None
        if redraw and canvas is not None:
            canvas.draw_idle()

    # ════════════════════════════════════════════════════════════════
    # Misc
    # ════════════════════════════════════════════════════════════════
    def _clear_all(self):
        self._clear_ann(redraw=False)
        self._cv_fig.clf();  self._cv_ax  = self._cv_fig.add_subplot(111)
        self._cyc_fig.clf(); self._cyc_ax = self._cyc_fig.add_subplot(111)
        self._cv_cv.draw_idle()
        self._cyc_cv.draw_idle()
        for entry in self.files.values():
            entry["result"] = None
        self._rebuild_results_tv()

    def _clear_plots(self):
        self._cv_fig.clf();  self._cv_ax  = self._cv_fig.add_subplot(111)
        self._cyc_fig.clf(); self._cyc_ax = self._cyc_fig.add_subplot(111)
        self._cv_cv.draw_idle()
        self._cyc_cv.draw_idle()

    # ════════════════════════════════════════════════════════════════
    # File merge
    # ════════════════════════════════════════════════════════════════
    def _open_merge_dialog(self):
        if len(self.files) < 2:
            messagebox.showinfo("Merge", "Load at least 2 files to merge.")
            return

        dlg = tk.Toplevel(self)
        dlg.title("Merge Files")
        dlg.resizable(False, False)
        dlg.grab_set()

        ttk.Label(dlg, text="Click files to set merge order:",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=8, pady=(8, 2))
        ttk.Label(dlg, text="File 1 cycles kept as-is; each subsequent file's cycles\n"
                            "continue from where the previous file ended.",
                  foreground="gray", font=("", 8), justify=tk.LEFT).pack(
                      anchor=tk.W, padx=8)

        _btns_frame = ttk.Frame(dlg)
        _btns_frame.pack(fill=tk.X, padx=8, pady=4)

        order   = []    # short names in selected order
        btn_map = {}    # short → Button widget

        def _refresh():
            for sh, btn in btn_map.items():
                if sh in order:
                    idx = order.index(sh) + 1
                    btn.config(text=f"[{idx}] {sh}", relief=tk.SUNKEN, bg="#c8e6c9")
                else:
                    btn.config(text=sh, relief=tk.RAISED, bg="SystemButtonFace")

        def _toggle(sh):
            if sh in order:
                order.remove(sh)
            else:
                order.append(sh)
            _refresh()

        for sh in self.files:
            btn = tk.Button(_btns_frame, text=sh, anchor=tk.W,
                            command=lambda s=sh: _toggle(s))
            btn.pack(fill=tk.X, pady=1)
            btn_map[sh] = btn

        _name_row = ttk.Frame(dlg)
        _name_row.pack(fill=tk.X, padx=8, pady=(8, 4))
        ttk.Label(_name_row, text="Merged file name:", width=18,
                  anchor=tk.W).pack(side=tk.LEFT)
        _name_var = tk.StringVar(value="merged.txt")
        ttk.Entry(_name_row, textvariable=_name_var, width=26).pack(
            side=tk.LEFT, padx=(4, 0))

        def _confirm():
            if len(order) < 2:
                messagebox.showwarning("Merge", "Select at least 2 files.", parent=dlg)
                return
            name = _name_var.get().strip()
            if not name:
                messagebox.showwarning("Merge", "Enter a name for the merged file.", parent=dlg)
                return
            if name in self.files:
                messagebox.showwarning("Merge", f"'{name}' already exists.", parent=dlg)
                return
            dlg.destroy()
            self._do_merge(list(order), name)

        _act_row = ttk.Frame(dlg)
        _act_row.pack(fill=tk.X, padx=8, pady=(4, 10))
        ttk.Button(_act_row, text="Merge",  command=_confirm).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(_act_row, text="Cancel", command=dlg.destroy).pack(side=tk.LEFT)

    def _do_merge(self, order: list, merged_name: str):
        """Concatenate files, renumbering cycles so they run continuously."""
        pieces    = []
        max_cycle = 0
        first_entry = self.files[order[0]]

        for i, short in enumerate(order):
            entry = self.files[short]
            df    = entry["df"].copy()

            saved_cc = entry.get("cyc_col", "")
            if saved_cc and saved_cc != "(none)" and saved_cc in df.columns:
                ccol = saved_cc
            elif "cycle number" in df.columns:
                ccol = "cycle number"
            else:
                ccol = None

            if ccol:
                cyc_vals = df[ccol].values.copy().astype(float)
                min_cyc  = int(cyc_vals.min())
                max_cyc  = int(cyc_vals.max())
                if i == 0:
                    df["cycle number"] = cyc_vals.astype(int)
                    max_cycle = max_cyc
                else:
                    offset = max_cycle - min_cyc + 1
                    df["cycle number"] = (cyc_vals + offset).astype(int)
                    max_cycle = max_cycle + (max_cyc - min_cyc + 1)
            else:
                # No cycle info — treat whole file as one next cycle
                max_cycle += 1
                df["cycle number"] = max_cycle

            pieces.append(df)

        merged_df = pd.concat(pieces, ignore_index=True)

        color_idx = len(self.files) % len(_TRACE_COLORS)
        self.files[merged_name] = {
            "path":      "",
            "df":        merged_df,
            "r_sol":     first_entry.get("r_sol", 0.0),
            "e_ref":     first_entry.get("e_ref", 0.0),
            "x_col":     first_entry.get("x_col", ""),
            "y_col":     first_entry.get("y_col", ""),
            "cyc_col":   "cycle number",
            "e_target":  first_entry.get("e_target", ""),
            "direction": first_entry.get("direction", "Anodic"),
            "window":    first_entry.get("window", "10"),
            "threshold": first_entry.get("threshold", "2.0"),
            "color":     _TRACE_COLORS[color_idx],
            "result":    None,
            "custom_cv_title": "",
            "cv_xmin": "", "cv_xmax": "",
            "cv_ymin": "", "cv_ymax": "",
        }
        self._loading = True
        self.file_listbox.insert(tk.END, merged_name)
        self._loading = False
        self._switch_file(merged_name)
        self._auto_set_e_target()
