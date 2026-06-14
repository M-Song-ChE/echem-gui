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

from .file_manager import _read_mpr, _COLOR_NAMES, _COLOR_HEX
from .checklist import CheckableListbox
from .plotting import copy_figure_to_clipboard

# ── Defaults ────────────────────────────────────────────────────────────────
_DEF = dict(
    e_target="0.70",
    window="10",
    threshold="2.0",
    r_sol="0",
    e_ref="0",
)

_DIRECTIONS = ["Anodic", "Cathodic", "Average"]

# Multi-file cycle-vs-J plot colors
_TRACE_COLORS = [
    "#1f77b4", "#d62728", "#2ca02c", "#9467bd", "#ff7f0e",
    "#8c564b", "#17becf", "#e377c2", "#bcbd22", "#7f7f7f",
]

_trapz = getattr(np, "trapezoid", getattr(np, "trapz", None))


# ── File reading (mirrors hupd_panel) ───────────────────────────────────────
def _read_one(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".mpr":
        return _read_mpr(path)
    df = pd.read_csv(path, sep="\t", encoding="latin-1", on_bad_lines="skip")
    df.columns = [c.strip().replace("<", "").replace(">", "") for c in df.columns]
    return df.reset_index(drop=True)


def _get_cycles(df: pd.DataFrame):
    """Sorted unique cycle numbers, or [None] if no cycle column."""
    if "cycle number" in df.columns and len(df):
        return sorted(df["cycle number"].unique().tolist())
    return [None]


def _split_scans(E: np.ndarray, I: np.ndarray):
    """Split E/I arrays into anodic and cathodic halves (each sorted ascending in E)."""
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
    """Return interpolated I at e_target, or None if out of range."""
    E = np.asarray(E, dtype=float)
    I = np.asarray(I, dtype=float)
    if len(E) < 2:
        return None
    e_lo, e_hi = E.min(), E.max()
    if e_target < e_lo or e_target > e_hi:
        return None
    return float(np.interp(e_target, E, I))


# ══════════════════════════════════════════════════════════════════════════════
class CvActivationPanel(ttk.Frame):
    """Self-contained CV Activation convergence panel."""

    def __init__(self, master):
        ttk.Frame.__init__(self, master)
        self.files       = OrderedDict()   # short_name → entry dict
        self.active_file = None
        self._loading    = False
        self._cv_cbar    = None            # colorbar handle — removed before each redraw
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
        ttk.Button(_fb, text="Remove", command=self._remove_file).pack(side=tk.LEFT)
        _flf = ttk.Frame(left); _flf.pack(fill=tk.X, padx=4, pady=2)
        self.file_listbox = CheckableListbox(
            _flf, height=5, show_checkboxes=False,
            on_reorder=self._on_file_reorder)
        self.file_listbox.pack(fill=tk.X, expand=True)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        # ── Column selectors ──────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Column Mapping", font=("", 9, "bold")).pack(
            anchor=tk.W, padx=4)

        ttk.Label(left, text="X-axis (potential):").pack(anchor=tk.W, padx=4, pady=(4, 0))
        self.x_var = tk.StringVar()
        self.x_combo = ttk.Combobox(left, textvariable=self.x_var,
                                    state="readonly", width=22)
        self.x_combo.pack(anchor=tk.W, padx=4, pady=2)
        self.x_combo.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())

        ttk.Label(left, text="Y-axis (current):").pack(anchor=tk.W, padx=4)
        self.y_var = tk.StringVar()
        self.y_combo = ttk.Combobox(left, textvariable=self.y_var,
                                    state="readonly", width=22)
        self.y_combo.pack(anchor=tk.W, padx=4, pady=2)
        self.y_combo.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())

        ttk.Label(left, text="Cycle column:").pack(anchor=tk.W, padx=4)
        self.cyc_var = tk.StringVar()
        self.cyc_combo = ttk.Combobox(left, textvariable=self.cyc_var,
                                      state="readonly", width=22)
        self.cyc_combo.pack(anchor=tk.W, padx=4, pady=2)
        self.cyc_combo.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())

        # ── Corrections ───────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Corrections", font=("", 9, "bold")).pack(
            anchor=tk.W, padx=4)

        _ir_row = ttk.Frame(left); _ir_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_ir_row, text="R_sol (Ω):", width=12, anchor=tk.W).pack(side=tk.LEFT)
        self.r_sol_var = tk.StringVar(value=_DEF["r_sol"])
        _ir_e = ttk.Entry(_ir_row, textvariable=self.r_sol_var, width=8)
        _ir_e.pack(side=tk.LEFT, padx=(2, 0))
        _ir_e.bind("<Return>",   lambda e: self._save_and_replot())
        _ir_e.bind("<FocusOut>", lambda e: self._save_and_replot())

        _rhe_row = ttk.Frame(left); _rhe_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_rhe_row, text="E_ref offset (V):", width=14, anchor=tk.W).pack(side=tk.LEFT)
        self.e_ref_var = tk.StringVar(value=_DEF["e_ref"])
        _rhe_e = ttk.Entry(_rhe_row, textvariable=self.e_ref_var, width=8)
        _rhe_e.pack(side=tk.LEFT, padx=(2, 0))
        _rhe_e.bind("<Return>",   lambda e: self._save_and_replot())
        _rhe_e.bind("<FocusOut>", lambda e: self._save_and_replot())
        ttk.Label(left, text="E_corr = E − I·R_sol + E_ref",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)

        # ── Activation settings ────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Activation Settings", font=("", 9, "bold")).pack(
            anchor=tk.W, padx=4)

        _et_row = ttk.Frame(left); _et_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_et_row, text="E target (V):", width=16, anchor=tk.W).pack(side=tk.LEFT)
        self.e_target_var = tk.StringVar(value=_DEF["e_target"])
        _et_e = ttk.Entry(_et_row, textvariable=self.e_target_var, width=8)
        _et_e.pack(side=tk.LEFT, padx=(2, 0))
        _et_e.bind("<Return>",   lambda e: self._auto_replot())
        _et_e.bind("<FocusOut>", lambda e: self._auto_replot())

        _dir_row = ttk.Frame(left); _dir_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_dir_row, text="Scan direction:", width=16, anchor=tk.W).pack(side=tk.LEFT)
        self.direction_var = tk.StringVar(value="Anodic")
        ttk.Combobox(_dir_row, textvariable=self.direction_var,
                     values=_DIRECTIONS, state="readonly", width=10).pack(side=tk.LEFT, padx=(2, 0))
        self.direction_var.trace_add("write", lambda *_: self._auto_replot())

        _win_row = ttk.Frame(left); _win_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_win_row, text="Conv. window (N cycles):", width=22, anchor=tk.W).pack(side=tk.LEFT)
        self.window_var = tk.StringVar(value=_DEF["window"])
        _win_e = ttk.Entry(_win_row, textvariable=self.window_var, width=5)
        _win_e.pack(side=tk.LEFT, padx=(2, 0))
        _win_e.bind("<Return>",   lambda e: self._auto_replot())
        _win_e.bind("<FocusOut>", lambda e: self._auto_replot())

        _thr_row = ttk.Frame(left); _thr_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_thr_row, text="Threshold (%):", width=16, anchor=tk.W).pack(side=tk.LEFT)
        self.threshold_var = tk.StringVar(value=_DEF["threshold"])
        _thr_e = ttk.Entry(_thr_row, textvariable=self.threshold_var, width=6)
        _thr_e.pack(side=tk.LEFT, padx=(2, 0))
        _thr_e.bind("<Return>",   lambda e: self._auto_replot())
        _thr_e.bind("<FocusOut>", lambda e: self._auto_replot())

        ttk.Label(left,
                  text="Pass: |ΔJ over last N cycles| / |J| < threshold%",
                  foreground="gray", font=("", 8), wraplength=270,
                  justify=tk.LEFT).pack(anchor=tk.W, padx=4, pady=(0, 4))

        # ── Cycle-vs-J overlay option ──────────────────────────────
        _ol_row = ttk.Frame(left); _ol_row.pack(fill=tk.X, padx=4, pady=2)
        self.overlay_all_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(_ol_row, text="Show all samples in cycle plot",
                        variable=self.overlay_all_var,
                        command=self._replot_cycle).pack(side=tk.LEFT)

        # ── Buttons ───────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        _btn_row = ttk.Frame(left); _btn_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Button(_btn_row, text="Analyze",
                   command=self._analyze).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(_btn_row, text="Clear",
                   command=self._clear_all).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(_btn_row, text="Copy CV",
                   command=lambda: copy_figure_to_clipboard(self._cv_fig)).pack(
                       side=tk.LEFT)

        # ── Results table ─────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        _rtv_lf = ttk.LabelFrame(left, text="Results")
        _rtv_lf.pack(fill=tk.BOTH, padx=4, pady=4, expand=False)
        _rtv_cols = ("file", "j_final", "delta_pct", "status")
        self._results_tv = ttk.Treeview(
            _rtv_lf, columns=_rtv_cols, show="headings",
            height=6, selectmode="browse")
        self._results_tv.heading("file",      text="Sample")
        self._results_tv.heading("j_final",   text="J_final (mA)")
        self._results_tv.heading("delta_pct", text="Δ% / N cyc")
        self._results_tv.heading("status",    text="Status")
        self._results_tv.column("file",      width=110, anchor=tk.W,      stretch=True)
        self._results_tv.column("j_final",   width=80,  anchor=tk.CENTER, stretch=False)
        self._results_tv.column("delta_pct", width=70,  anchor=tk.CENTER, stretch=False)
        self._results_tv.column("status",    width=90,  anchor=tk.CENTER, stretch=False)
        _rtv_sb = ttk.Scrollbar(_rtv_lf, orient=tk.VERTICAL,
                                 command=self._results_tv.yview)
        self._results_tv.configure(yscrollcommand=_rtv_sb.set)
        _rtv_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._results_tv.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        # Color-code rows
        self._results_tv.tag_configure("pass", background="#c8e6c9")
        self._results_tv.tag_configure("fail", background="#ffcdd2")

        # ── Right panel: two figures stacked ─────────────────────
        right = ttk.Frame(body)
        body.add(right, weight=1)

        _rpw = ttk.PanedWindow(right, orient=tk.VERTICAL)
        _rpw.pack(fill=tk.BOTH, expand=True)

        # Upper: CV figure
        _cv_frame = ttk.Frame(_rpw)
        _rpw.add(_cv_frame, weight=3)
        self._cv_fig = Figure(figsize=(8, 4), dpi=100)
        self._cv_ax  = self._cv_fig.add_subplot(111)
        self._cv_cv  = FigureCanvasTkAgg(self._cv_fig, master=_cv_frame)
        _cv_tb = NavigationToolbar2Tk(self._cv_cv, _cv_frame, pack_toolbar=False)
        _cv_tb.pack(side=tk.BOTTOM, fill=tk.X)
        self._cv_cv.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self._cv_cv.mpl_connect("button_press_event", self._on_cv_click)

        # Lower: cycle-vs-J figure
        _cyc_frame = ttk.Frame(_rpw)
        _rpw.add(_cyc_frame, weight=2)
        self._cyc_fig = Figure(figsize=(8, 3), dpi=100)
        self._cyc_ax  = self._cyc_fig.add_subplot(111)
        self._cyc_cv  = FigureCanvasTkAgg(self._cyc_fig, master=_cyc_frame)
        _cyc_tb = NavigationToolbar2Tk(self._cyc_cv, _cyc_frame, pack_toolbar=False)
        _cyc_tb.pack(side=tk.BOTTOM, fill=tk.X)
        self._cyc_cv.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Vertical line for E_target on CV plot (set on click)
        self._cv_vline = None

    # ════════════════════════════════════════════════════════════════
    # File management
    # ════════════════════════════════════════════════════════════════
    def _load_files(self):
        paths = filedialog.askopenfilenames(
            title="Load Activation CV File(s)",
            filetypes=[("Data files", "*.txt *.csv *.mpr *.mpt"),
                       ("All files", "*.*")])
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
                "path":      path,
                "df":        df,
                "r_sol":     0.0,
                "e_ref":     0.0,
                "color":     _TRACE_COLORS[color_idx],
                "result":    None,   # filled by _analyze
            }
            self.files[short] = entry
            self._loading = True
            self.file_listbox.insert(tk.END, short)
            self._loading = False

        if self.files:
            last = list(self.files.keys())[-1]
            self._switch_file(last)

    def _remove_file(self):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        idx  = sel[0]
        keys = list(self.files.keys())
        if idx >= len(keys):
            return
        short = keys[idx]
        del self.files[short]
        self.file_listbox.delete(idx)
        keys = list(self.files.keys())
        if keys:
            new_idx = min(idx, len(keys) - 1)
            self._switch_file(keys[new_idx])
        else:
            self.active_file = None
            self._clear_plots()
        self._rebuild_results_tv()

    def _on_file_reorder(self, new_order):
        old_keys = list(self.files.keys())
        new_files = OrderedDict()
        for name in new_order:
            if name in self.files:
                new_files[name] = self.files[name]
        for k, v in self.files.items():
            if k not in new_files:
                new_files[k] = v
        self.files = new_files
        self._replot_cycle()

    def _on_file_select(self, _event=None):
        if self._loading:
            return
        sel = self.file_listbox.curselection()
        if not sel:
            return
        idx  = sel[0]
        keys = list(self.files.keys())
        if idx < len(keys):
            self._switch_file(keys[idx])

    def _switch_file(self, short):
        self.active_file = short
        # Sync listbox selection
        keys = list(self.files.keys())
        if short in keys:
            self._loading = True
            self.file_listbox.selection_clear(0, tk.END)
            self.file_listbox.selection_set(keys.index(short))
            self._loading = False
        self._update_column_combos()
        self._restore_corrections()
        self._auto_replot()

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

        # Auto-detect if not already set or column no longer valid
        if self.x_var.get() not in cols:
            for c in ["Ewe/V", "Ewe/V ", "E/V", "E (V)"]:
                if c in cols:
                    self.x_var.set(c); break
            else:
                # pick first voltage-like column
                for c in cols:
                    if "/V" in c or "(V)" in c:
                        self.x_var.set(c); break
        if self.y_var.get() not in cols:
            for c in ["I/mA", "<I>/mA", "I (mA)"]:
                if c in cols:
                    self.y_var.set(c); break
            else:
                for c in cols:
                    if "/mA" in c or "(mA)" in c or "/A" in c:
                        self.y_var.set(c); break
        if self.cyc_var.get() not in ["(none)"] + cols:
            if "cycle number" in cols:
                self.cyc_var.set("cycle number")
            else:
                self.cyc_var.set("(none)")

    def _get_cycle_col(self, df):
        c = self.cyc_var.get()
        if c and c != "(none)" and c in df.columns:
            return c
        return None

    # ════════════════════════════════════════════════════════════════
    # Corrections
    # ════════════════════════════════════════════════════════════════
    def _restore_corrections(self):
        entry = self.files.get(self.active_file)
        if not entry:
            return
        self.r_sol_var.set(str(entry["r_sol"]))
        self.e_ref_var.set(str(entry["e_ref"]))

    def _save_and_replot(self):
        entry = self.files.get(self.active_file)
        if entry:
            try: entry["r_sol"] = float(self.r_sol_var.get())
            except ValueError: pass
            try: entry["e_ref"] = float(self.e_ref_var.get())
            except ValueError: pass
        self._auto_replot()

    def _apply_correction(self, df, r_sol, e_ref):
        """IR compensation + RHE offset; returns corrected copy."""
        xcol = self.x_var.get()
        ycol = self.y_var.get()
        if not xcol or not ycol or xcol not in df.columns or ycol not in df.columns:
            return df
        df = df.copy()
        if r_sol != 0:
            # Assume current column is in mA → convert to A
            I_A = df[ycol].values * 1e-3
            df[xcol] = df[xcol].values - I_A * r_sol
        if e_ref != 0:
            df[xcol] = df[xcol].values + e_ref
        return df

    # ════════════════════════════════════════════════════════════════
    # Core extraction
    # ════════════════════════════════════════════════════════════════
    def _extract_cycle_j(self, df_c, e_target: float, direction: str):
        """Return list of (cycle_num_int, J_at_E) for each cycle."""
        xcol = self.x_var.get()
        ycol = self.y_var.get()
        ccol = self._get_cycle_col(df_c)
        if not xcol or not ycol:
            return []

        cycles = []
        if ccol:
            raw_cycles = sorted(df_c[ccol].unique())
        else:
            raw_cycles = [None]

        for cn in raw_cycles:
            if cn is not None:
                sub = df_c[df_c[ccol] == cn]
            else:
                sub = df_c
            E = sub[xcol].dropna().values
            I = sub[ycol].dropna().values
            # Trim to matching lengths
            n = min(len(E), len(I))
            if n < 4:
                continue
            E, I = E[:n], I[:n]

            if direction == "Anodic":
                E_use, I_use, _, _ = _split_scans(E, I)
            elif direction == "Cathodic":
                _, _, E_use, I_use = _split_scans(E, I)
            else:  # Average
                E_an, I_an, E_cat, I_cat = _split_scans(E, I)
                j_an  = _interp_at_e(E_an,  I_an,  e_target)
                j_cat = _interp_at_e(E_cat, I_cat, e_target)
                vals  = [v for v in [j_an, j_cat] if v is not None]
                if vals:
                    cn_int = int(cn) if cn is not None else 0
                    cycles.append((cn_int, float(np.mean(vals))))
                continue

            j = _interp_at_e(E_use, I_use, e_target)
            if j is not None:
                cn_int = int(cn) if cn is not None else 0
                cycles.append((cn_int, j))

        return cycles  # [(cycle_num, J)]

    def _check_convergence(self, cycle_j: list, window: int, threshold: float):
        """
        For each cycle i where i >= window:
          delta% = |J[i] - J[i-window]| / max(|J[i-window]|, 1e-12) * 100
        Returns list of (cn, J, delta_pct_or_None, pass_bool_or_None).
        """
        out = []
        for i, (cn, j) in enumerate(cycle_j):
            if i < window:
                out.append((cn, j, None, None))
            else:
                j_prev = cycle_j[i - window][1]
                delta  = abs(j - j_prev) / max(abs(j_prev), 1e-12) * 100.0
                passed = delta < threshold
                out.append((cn, j, delta, passed))
        return out

    # ════════════════════════════════════════════════════════════════
    # Plotting
    # ════════════════════════════════════════════════════════════════
    def _auto_replot(self):
        self._replot_cv()
        self._replot_cycle()

    def _replot_cv(self):
        """Redraw upper CV figure for active file — all cycles, gradient-colored."""
        # Remove stale colorbar before clearing axes (colorbar lives on the figure)
        if self._cv_cbar is not None:
            try:
                self._cv_cbar.remove()
            except Exception:
                pass
            self._cv_cbar = None

        ax = self._cv_ax
        ax.clear()
        entry = self.files.get(self.active_file)
        if not entry:
            ax.set_title("CV  (no file loaded)")
            self._cv_fig.tight_layout(pad=0.8)
            self._cv_fig.set_layout_engine("none")
            self._cv_cv.draw_idle()
            return

        xcol = self.x_var.get()
        ycol = self.y_var.get()
        if not xcol or not ycol:
            self._cv_cv.draw_idle()
            return

        try:
            r_sol = float(self.r_sol_var.get())
        except ValueError:
            r_sol = 0.0
        try:
            e_ref = float(self.e_ref_var.get())
        except ValueError:
            e_ref = 0.0

        df_c  = self._apply_correction(entry["df"], r_sol, e_ref)
        ccol  = self._get_cycle_col(df_c)

        if ccol:
            raw_cycles = sorted(df_c[ccol].unique())
        else:
            raw_cycles = [None]

        n_cyc = len(raw_cycles)
        cmap  = mpl_cm.get_cmap("viridis", max(n_cyc, 2))

        for i, cn in enumerate(raw_cycles):
            if cn is not None:
                sub = df_c[df_c[ccol] == cn]
            else:
                sub = df_c
            E = sub[xcol].values
            I = sub[ycol].values
            color = cmap(i / max(n_cyc - 1, 1))
            lbl = f"C{int(cn)}" if cn is not None else "data"
            ax.plot(E, I, lw=1.0, color=color, label=lbl, alpha=0.85)

        # E_target marker
        try:
            e_t = float(self.e_target_var.get())
            ax.axvline(e_t, color="red", lw=1.2, ls="--", alpha=0.7, label=f"E={e_t:.3f} V")
        except ValueError:
            pass

        x_lbl = xcol
        y_lbl = ycol
        ref_lbl = ""
        if e_ref != 0:
            ref_lbl = " vs RHE"
        ax.set_xlabel(f"{x_lbl}{ref_lbl}", fontsize=9)
        ax.set_ylabel(y_lbl, fontsize=9)
        ax.set_title(f"Activation CV — {self.active_file}  ({n_cyc} cycles)",
                     fontsize=9)
        ax.tick_params(labelsize=8)

        if n_cyc <= 20:
            ax.legend(fontsize=6, ncol=max(1, n_cyc // 8), frameon=True,
                      loc="best")
        else:
            self._cv_cbar = self._cv_fig.colorbar(
                mpl_cm.ScalarMappable(
                    norm=mpl_colors.Normalize(1, n_cyc), cmap="viridis"),
                ax=ax, shrink=0.8, pad=0.02)
            self._cv_cbar.set_label("Cycle #", fontsize=8)

        self._cv_fig.tight_layout(pad=0.8)
        self._cv_fig.set_layout_engine("none")
        self._cv_cv.draw_idle()

    def _replot_cycle(self):
        """Redraw lower cycle-vs-J figure."""
        ax = self._cyc_ax
        ax.clear()

        try:
            e_target = float(self.e_target_var.get())
        except ValueError:
            self._cyc_cv.draw_idle()
            return
        try:
            window = int(self.window_var.get())
        except ValueError:
            window = 10
        try:
            threshold = float(self.threshold_var.get())
        except ValueError:
            threshold = 2.0

        direction = self.direction_var.get()
        show_all  = self.overlay_all_var.get()

        files_to_show = (list(self.files.keys())
                         if show_all else
                         ([self.active_file] if self.active_file else []))

        any_data = False
        for fi, short in enumerate(files_to_show):
            entry = self.files.get(short)
            if not entry:
                continue
            try:
                r_sol = entry["r_sol"]
                e_ref = entry["e_ref"]
            except Exception:
                r_sol = e_ref = 0.0

            df_c     = self._apply_correction(entry["df"], r_sol, e_ref)
            cycle_j  = self._extract_cycle_j(df_c, e_target, direction)
            if not cycle_j:
                continue

            conv     = self._check_convergence(cycle_j, window, threshold)
            cns      = [x[0] for x in conv]
            js       = [x[1] for x in conv]
            color    = entry["color"]

            ax.plot(cns, js, "o-", color=color, lw=1.6, ms=4, label=short)

            # Mark the last window delta
            if len(conv) > window:
                *_, last = conv
                cn_last, j_last, dp, passed = last
                mk = "✓" if passed else "✗"
                col_mk = "#2e7d32" if passed else "#c62828"
                ax.annotate(
                    f"{mk} {dp:.1f}%",
                    xy=(cn_last, j_last),
                    xytext=(6, 4), textcoords="offset points",
                    fontsize=7, color=col_mk)
            any_data = True

        if any_data:
            ax.set_xlabel("Cycle number", fontsize=9)
            ax.set_ylabel(f"J at E={e_target:.3f} V  (mA)", fontsize=9)
            ax.set_title(f"Convergence check  (window={window} cyc, threshold={threshold}%)",
                         fontsize=9)
            ax.tick_params(labelsize=8)
            ax.axhline(0, color="k", lw=0.5, ls=":")
            if len(files_to_show) > 1:
                ax.legend(fontsize=7, frameon=True)
            ax.grid(True, alpha=0.3)
        else:
            ax.set_title("Cycle vs J  (no data — check E_target or column mapping)",
                         fontsize=9)

        self._cyc_fig.tight_layout(pad=0.8)
        self._cyc_fig.set_layout_engine("none")
        self._cyc_cv.draw_idle()

    # ════════════════════════════════════════════════════════════════
    # Analyze button → fill results table
    # ════════════════════════════════════════════════════════════════
    def _analyze(self):
        try:
            e_target  = float(self.e_target_var.get())
            window    = int(self.window_var.get())
            threshold = float(self.threshold_var.get())
        except ValueError:
            messagebox.showerror("Input error",
                                 "Check E target, window, and threshold values.")
            return

        direction = self.direction_var.get()

        for short, entry in self.files.items():
            try:
                r_sol = entry["r_sol"]
                e_ref = entry["e_ref"]
            except Exception:
                r_sol = e_ref = 0.0
            df_c    = self._apply_correction(entry["df"], r_sol, e_ref)
            cycle_j = self._extract_cycle_j(df_c, e_target, direction)
            if not cycle_j:
                entry["result"] = None
                continue
            conv = self._check_convergence(cycle_j, window, threshold)
            # Summary: use last entry that has a delta
            last_with_delta = [(cn, j, dp, p) for cn, j, dp, p in conv
                               if dp is not None]
            if last_with_delta:
                cn_f, j_f, dp_f, passed_f = last_with_delta[-1]
            else:
                cn_f, j_f, dp_f, passed_f = conv[-1][0], conv[-1][1], None, None

            entry["result"] = {
                "cycle_j":   cycle_j,
                "conv":      conv,
                "j_final":   j_f,
                "delta_pct": dp_f,
                "passed":    passed_f,
            }

        self._rebuild_results_tv()
        self._replot_cycle()

    def _rebuild_results_tv(self):
        tv = self._results_tv
        tv.delete(*tv.get_children())
        for short, entry in self.files.items():
            res = entry.get("result")
            if res is None:
                tv.insert("", tk.END, values=(short, "—", "—", "—"))
                continue
            j_f  = f"{res['j_final']:.4f}" if res["j_final"] is not None else "—"
            dp   = f"{res['delta_pct']:.2f}%" if res["delta_pct"] is not None else "—"
            if res["passed"] is True:
                status = "Activated ✓"
                tag    = "pass"
            elif res["passed"] is False:
                status = "Not activated ✗"
                tag    = "fail"
            else:
                status = "—"
                tag    = ""
            tv.insert("", tk.END, values=(short, j_f, dp, status), tags=(tag,))

    # ════════════════════════════════════════════════════════════════
    # Misc
    # ════════════════════════════════════════════════════════════════
    def _on_cv_click(self, event):
        """Right-click on CV sets E_target to clicked X value."""
        if event.button == 3 and event.inaxes:
            self.e_target_var.set(f"{event.xdata:.4f}")
            self._auto_replot()

    def _clear_all(self):
        if self._cv_cbar is not None:
            try: self._cv_cbar.remove()
            except Exception: pass
            self._cv_cbar = None
        self._cv_ax.clear()
        self._cyc_ax.clear()
        for fig in (self._cv_fig, self._cyc_fig):
            fig.tight_layout(pad=0.8)
            fig.set_layout_engine("none")
        self._cv_cv.draw_idle()
        self._cyc_cv.draw_idle()
        for entry in self.files.values():
            entry["result"] = None
        self._rebuild_results_tv()

    def _clear_plots(self):
        self._cv_ax.clear()
        self._cyc_ax.clear()
        self._cv_cv.draw_idle()
        self._cyc_cv.draw_idle()
