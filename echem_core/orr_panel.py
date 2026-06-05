"""ORR Analysis panel — background-subtracted rotating disk electrode polarization curves.

Each SAMPLE groups N2 (background) and O2 (signal) CV file pairs by RPM.

Processing per pair:
  1. Extract the last cycle from each file.
  2. Apply IR correction:  E = Ewe/V − (I/mA / 1000) × R_sol
     (separate R_sol_N2 and R_sol_O2 — measured in independent sessions).
  3. Apply RHE conversion: E += E_ref  (shared per sample).
  4. Extract the anodic scan direction (from minimum-E vertex upward).
  5. Interpolate N2 onto O2's E grid and subtract: I_net = I_O2 − I_N2_interp.
  6. Optionally normalise to electrode area → J (mA/cm²).

Multiple samples are displayed in a scrollable grid of subplots (configurable columns),
mirroring the layout of the Multi E.Chem 2 tab.
"""

from collections import OrderedDict
import re
import os
import math

import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from .file_manager import _read_mpr, _PALETTE, _COLOR_NAMES, _COLOR_HEX
from .plotting import (apply_grid, draw_reflines, copy_figure_to_clipboard,
                        _cycle_colors, _scale_legend_spacing)
from .checklist import CheckableListbox
from . import session_manager as _sm
from .legend_editor import open_legend_editor

# ── UI constants ────────────────────────────────────────────────────────
_SAMPLE_HDR_BG     = "#d1c4e9"   # light purple — distinct from ME1 blue / ME2 green
_SAMPLE_HDR_ACTIVE = "#ffd54f"   # gold  (matches other tabs)

# Extracts catalyst label from parentheses plus optional dataset suffix,
# e.g. "Sample_03(Pt) vs …" → "Pt", "Sample_03(Pt)_2 vs …" → "Pt_2"
_CATALYST_PAT = re.compile(r'\((\w+)\)(_\d+)?')

# Runtime-only keys stripped before JSON serialisation
_SAMPLE_RUNTIME = frozenset({
    "fig", "ax", "canvas", "toolbar", "plot_frame",
    "hdr_frame", "hdr_label", "legend",
    # ephemeral interaction state
    "panning", "pan_ax", "pan_start", "pan_moved",
    "leg_resize", "leg_resize_start_y", "leg_resize_start_sz",
    "ann", "ann_dot", "ann_last", "ann_idx", "_plot_data",
    "auto_xlim", "auto_ylim", "leg_size_live",
})
_PAIR_RUNTIME = frozenset({"df_n2", "df_o2"})

# Matches  _NN_CV_   or   _NNN_CV_   in a filename
_RPM_PAT = re.compile(r'_(\d{2,4})_CV_', re.IGNORECASE)

_GRID_STYLE_MAP = {
    "dashed": "--", "dotted": ":", "solid": "-", "dash-dot": "-."}

_ANN_DOT_LABEL = "_orr_dot"


# ── Module-level helpers ─────────────────────────────────────────────────

def _read_one_df(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".mpr":
        return _read_mpr(path)
    df = pd.read_csv(path, sep="\t", encoding="latin-1", on_bad_lines="skip")
    df.columns = [c.strip().replace("<", "").replace(">", "") for c in df.columns]
    return df.reset_index(drop=True)


def _detect_gas(stem: str) -> str:
    lo = stem.lower()
    # Match n2 / o2 as word-like fragments (not inside longer numbers)
    if re.search(r'(?<!\d)n2(?!\d)', lo):
        return "n2"
    if re.search(r'(?<!\d)o2(?!\d)', lo):
        return "o2"
    return ""


def _extract_rpm_id(stem: str) -> str:
    m = _RPM_PAT.search(stem)
    return m.group(1) if m else ""


def _detect_catalyst(stem: str) -> str:
    """Extract catalyst label including optional dataset suffix.

    'LTS-BDRDE_22(Pt) vs …'   → 'Pt'
    'LTS-BDRDE_22(Pt)_2 vs …' → 'Pt_2'
    """
    m = _CATALYST_PAT.search(stem)
    if not m:
        return ""
    return m.group(1) + (m.group(2) or "")


def _extract_anodic(E: np.ndarray, I: np.ndarray):
    """Return (E_sorted, I_sorted) for the anodic (E-increasing) half of the CV.

    Finds the cathodic vertex (minimum E), takes everything from that point
    onward, then sorts by E ascending so np.interp works correctly.
    """
    if len(E) < 4:
        return E, I
    min_idx = int(np.argmin(E))
    E_an = E[min_idx:]
    I_an = I[min_idx:]
    order = np.argsort(E_an)
    return E_an[order], I_an[order]


def _find_half_wave(E: np.ndarray, J: np.ndarray):
    """Return (E_half, J_half) for an ORR polarization curve (anodic scan, E ascending).

    J_lim = most negative J (diffusion-limited plateau at low E).
    E½    = E where J crosses J_lim / 2.
    Returns (None, None) when data is insufficient or has no cathodic current.
    """
    if len(E) < 4:
        return None, None
    j_lim = float(np.min(J))
    if j_lim >= 0:
        return None, None
    j_half = j_lim / 2.0
    diff = J - j_half
    sign_ch = np.where(np.diff(np.sign(diff)))[0]
    if len(sign_ch) == 0:
        return None, None
    idx = sign_ch[len(sign_ch) // 2]
    e0, e1 = E[idx], E[idx + 1]
    j0, j1 = J[idx], J[idx + 1]
    e_half = (e0 + (j_half - j0) / (j1 - j0) * (e1 - e0)
              if j1 != j0 else (e0 + e1) / 2.0)
    return float(e_half), float(j_half)


def _process_pair(pair: dict, r_sol_n2: float, r_sol_o2: float,
                  e_ref: float, area: float):
    """Return (E_plot, Y_plot) for one ORR pair, or None on failure."""
    df_n2 = pair.get("df_n2")
    df_o2 = pair.get("df_o2")
    if df_n2 is None or df_o2 is None:
        return None

    for df in (df_n2, df_o2):
        if "Ewe/V" not in df.columns or "I/mA" not in df.columns:
            return None

    def _last_cycle(df):
        if "cycle number" in df.columns and len(df["cycle number"].unique()) > 0:
            last = df["cycle number"].max()
            return df[df["cycle number"] == last].copy()
        return df.copy()

    lc_n2 = _last_cycle(df_n2)
    lc_o2 = _last_cycle(df_o2)
    if len(lc_n2) < 10 or len(lc_o2) < 10:
        return None

    E_n2 = lc_n2["Ewe/V"].values.astype(float)
    I_n2 = lc_n2["I/mA"].values.astype(float)
    E_o2 = lc_o2["Ewe/V"].values.astype(float)
    I_o2 = lc_o2["I/mA"].values.astype(float)

    # IR correction
    if r_sol_n2 != 0.0:
        E_n2 = E_n2 - (I_n2 / 1000.0) * r_sol_n2
    if r_sol_o2 != 0.0:
        E_o2 = E_o2 - (I_o2 / 1000.0) * r_sol_o2

    # RHE conversion
    if e_ref != 0.0:
        E_n2 = E_n2 + e_ref
        E_o2 = E_o2 + e_ref

    # Anodic scan
    E_n2_an, I_n2_an = _extract_anodic(E_n2, I_n2)
    E_o2_an, I_o2_an = _extract_anodic(E_o2, I_o2)
    if len(E_n2_an) < 5 or len(E_o2_an) < 5:
        return None

    # Restrict O2 to overlapping E range
    E_lo = max(E_n2_an[0],  E_o2_an[0])
    E_hi = min(E_n2_an[-1], E_o2_an[-1])
    if E_lo >= E_hi:
        return None

    mask    = (E_o2_an >= E_lo) & (E_o2_an <= E_hi)
    E_plot  = E_o2_an[mask]
    I_o2_cl = I_o2_an[mask]
    if len(E_plot) < 3:
        return None

    I_n2_interp = np.interp(E_plot, E_n2_an, I_n2_an)
    I_net       = I_o2_cl - I_n2_interp

    Y_plot = I_net / area if area > 0 else I_net
    return E_plot, Y_plot


# ════════════════════════════════════════════════════════════════════════
# ORRPanel
# ════════════════════════════════════════════════════════════════════════

class ORRPanel(ttk.Frame):
    """ORR analysis panel — group N2+O2 CV pairs into samples; plot background-subtracted
    ORR polarization curves (anodic scan, last cycle, per RPM)."""

    def __init__(self, master):
        ttk.Frame.__init__(self, master)
        self.loaded_files   = OrderedDict()  # short_name → {path, gas, rpm_id, df}
        self._loaded_keys   = []             # ordered list of short names (mirrors listbox)
        self.samples        = OrderedDict()  # sample_name → sample_entry
        self.active_sample  = None
        self._suppress_replot = False
        self._loading        = False
        self._drag                = None
        self._zoom_sample         = None
        self._copied_sample_fmt   = None  # clipboard for Copy/Paste format
        self._build_panel()
        self.after(500, self._auto_set_initial_size)

    # ════════════════════════════════════════════════════════════════
    # Panel construction
    # ════════════════════════════════════════════════════════════════
    def _build_panel(self):
        body = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True)

        # ── Scrollable left panel ─────────────────────────────────
        left_outer = ttk.Frame(body, width=360)
        body.add(left_outer, weight=0)

        _lc = tk.Canvas(left_outer, highlightthickness=0)
        _ls = ttk.Scrollbar(left_outer, orient=tk.VERTICAL, command=_lc.yview)
        _lc.configure(yscrollcommand=_ls.set)
        _ls.pack(side=tk.RIGHT, fill=tk.Y)
        _lc.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        left = tk.Frame(_lc)
        self._left_lc    = _lc    # outer scrollable canvas
        self._left_inner = left   # inner content frame
        _lwin = _lc.create_window((0, 0), window=left, anchor=tk.NW)
        left.bind("<Configure>", lambda e: _lc.configure(scrollregion=_lc.bbox("all")))
        _lc.bind("<Configure>", lambda e: _lc.itemconfig(_lwin, width=e.width))
        _lc.bind("<MouseWheel>", lambda e: _lc.yview_scroll(-1 * (e.delta // 120), "units"))

        # ══ LOADED FILES ════════════════════════════════════════════
        _lf_hdr = ttk.Frame(left)
        _lf_hdr.pack(fill=tk.X, padx=4, pady=(6, 0))
        ttk.Label(_lf_hdr, text="Loaded Files:", font=("", 9, "bold")).pack(side=tk.LEFT)

        ttk.Label(left, text="(N2/O2 and catalyst auto-detected from filename)",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)
        _fb = ttk.Frame(left)
        _fb.pack(fill=tk.X, padx=4)
        ttk.Button(_fb, text="Load Files", command=self._load_files).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(_fb, text="Remove",     command=self._remove_loaded_file).pack(side=tk.LEFT)

        _gb = ttk.Frame(left)
        _gb.pack(fill=tk.X, padx=4, pady=(1, 0))
        ttk.Button(_gb, text="Sel N2", width=7,
                   command=self._select_n2).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(_gb, text="Sel O2", width=7,
                   command=self._select_o2).pack(side=tk.LEFT)

        _lf = ttk.Frame(left)
        _lf.pack(fill=tk.X, padx=4, pady=2)
        self.loaded_tv = ttk.Treeview(_lf, height=5, selectmode="extended", show="tree")
        self.loaded_tv.column("#0", minwidth=50, stretch=True)
        self.loaded_tv.tag_configure("n2",  background="#dbeafe", foreground="#1e40af")
        self.loaded_tv.tag_configure("o2",  background="#ffedd5", foreground="#9a3412")
        self.loaded_tv.tag_configure("unk", background="white")
        self.loaded_tv.tag_configure("hdr", background="#e8e8e8", foreground="#333333",
                                     font=("", 8, "bold"))
        self.loaded_tv.bind("<Double-1>", self._on_loaded_tv_dblclick)
        _lf_sb = ttk.Scrollbar(_lf, orient=tk.VERTICAL, command=self.loaded_tv.yview)
        self.loaded_tv.configure(yscrollcommand=_lf_sb.set)
        _lf_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.loaded_tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Drag handle to resize the loaded-files treeview
        _lf_handle = tk.Frame(left, height=5, bg="#c8c8c8", cursor="sb_v_double_arrow")
        _lf_handle.pack(fill=tk.X, padx=4)
        _lf_handle.bind("<ButtonPress-1>", self._on_loaded_resize_start)
        _lf_handle.bind("<B1-Motion>",     self._on_loaded_resize_drag)

        # ══ SAMPLES ═════════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        _sh_hdr = ttk.Frame(left)
        _sh_hdr.pack(fill=tk.X, padx=4)
        ttk.Label(_sh_hdr, text="ORR Samples:", font=("", 9, "bold")).pack(side=tk.LEFT)

        _sb = ttk.Frame(left)
        _sb.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Button(_sb, text="New Sample", command=self._new_sample).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(_sb, text="Rename",     command=self._rename_sample).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(_sb, text="Delete",     command=self._delete_sample).pack(side=tk.LEFT)

        _slf = ttk.Frame(left)
        _slf.pack(fill=tk.X, padx=4, pady=2)
        self.sample_lb = CheckableListbox(
            _slf, height=3,
            on_check=self._on_sample_visibility_change,
            on_reorder=self._on_sample_reorder)
        self.sample_lb.pack(fill=tk.X, expand=True)
        self.sample_lb.bind("<<ListboxSelect>>", self._on_sample_select)

        # Drag handle to resize the samples listbox
        _slf_handle = tk.Frame(left, height=5, bg="#c8c8c8", cursor="sb_v_double_arrow")
        _slf_handle.pack(fill=tk.X, padx=4)
        _slf_handle.bind("<ButtonPress-1>", self._on_sample_resize_start)
        _slf_handle.bind("<B1-Motion>",     self._on_sample_resize_drag)

        ttk.Button(left, text="↓ Add Selected Files to Sample",
                   command=self._add_files_to_sample).pack(fill=tk.X, padx=4, pady=(2, 0))

        _cp_row = ttk.Frame(left)
        _cp_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Button(_cp_row, text="Copy Format",
                   command=self._copy_sample_format).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(_cp_row, text="Paste Format",
                   command=self._paste_sample_format).pack(side=tk.LEFT)
        ttk.Label(_cp_row, text="(grid/font/legend/reflines)",
                  foreground="gray", font=("", 7)).pack(side=tk.LEFT, padx=(6, 0))

        # ══ PAIR TABLE ══════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(left, text="RPM Pairs  (active sample)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        ttk.Label(left,
                  text="Edit RPM values; N2/O2 paired by filename RPM index",
                  foreground="gray", font=("", 8), wraplength=320).pack(
            anchor=tk.W, padx=4)

        _ptf = ttk.Frame(left)
        _ptf.pack(fill=tk.X, padx=4, pady=2)
        self._pair_tbl_canvas = tk.Canvas(_ptf, background="#f5f5f5",
                                          highlightthickness=1, highlightbackground="#cccccc",
                                          height=100)
        _pt_vs = ttk.Scrollbar(_ptf, orient=tk.VERTICAL, command=self._pair_tbl_canvas.yview)
        self._pair_tbl_canvas.configure(yscrollcommand=_pt_vs.set)
        _pt_vs.pack(side=tk.RIGHT, fill=tk.Y)
        self._pair_tbl_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._pair_tbl_inner = tk.Frame(self._pair_tbl_canvas, background="#f5f5f5")
        _pt_win = self._pair_tbl_canvas.create_window((0, 0), window=self._pair_tbl_inner, anchor=tk.NW)
        self._pair_tbl_inner.bind(
            "<Configure>",
            lambda e: self._pair_tbl_canvas.configure(scrollregion=self._pair_tbl_canvas.bbox("all")))
        self._pair_tbl_canvas.bind(
            "<Configure>", lambda e: self._pair_tbl_canvas.itemconfig(_pt_win, width=e.width))

        def _pt_wheel(e):
            self._pair_tbl_canvas.yview_scroll(-1 * (e.delta // 120), "units")
            return "break"
        self._pair_tbl_canvas.bind("<MouseWheel>", _pt_wheel)
        self._pair_tbl_inner.bind("<MouseWheel>", _pt_wheel)

        # Drag handle to resize the RPM pairs table
        _pt_handle = tk.Frame(left, height=5, bg="#c8c8c8", cursor="sb_v_double_arrow")
        _pt_handle.pack(fill=tk.X, padx=4)
        _pt_handle.bind("<ButtonPress-1>", self._on_pair_resize_start)
        _pt_handle.bind("<B1-Motion>",     self._on_pair_resize_drag)

        # ══ CORRECTION ══════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(left, text="IR / RHE Correction  (active sample)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        _cr0 = ttk.Frame(left)
        _cr0.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Label(_cr0, text="Catalyst:").pack(side=tk.LEFT)
        self._corr_catalyst_var = tk.StringVar(value="")
        self._corr_cat_cb = ttk.Combobox(
            _cr0, textvariable=self._corr_catalyst_var,
            values=[], state="readonly", width=14)
        self._corr_cat_cb.pack(side=tk.LEFT, padx=(4, 0))
        self._corr_cat_cb.bind("<<ComboboxSelected>>", self._on_corr_catalyst_select)

        _cst = ttk.Frame(left)
        _cst.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(_cst, text="Color:").pack(side=tk.LEFT)
        self._cat_color_var = tk.StringVar(value="")
        _cat_color_e = ttk.Entry(_cst, textvariable=self._cat_color_var, width=9)
        _cat_color_e.pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(_cst, text="Style:").pack(side=tk.LEFT)
        self._cat_ls_var = tk.StringVar(value="solid")
        _cat_ls_cb = ttk.Combobox(_cst, textvariable=self._cat_ls_var,
                                   values=["solid", "dashed", "dotted", "dash-dot"],
                                   state="readonly", width=8)
        _cat_ls_cb.pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(_cst, text="W:").pack(side=tk.LEFT)
        self._cat_lw_var = tk.StringVar(value="1.5")
        _cat_lw_e = ttk.Entry(_cst, textvariable=self._cat_lw_var, width=4)
        _cat_lw_e.pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(_cst, text="Marker:").pack(side=tk.LEFT)
        self._cat_mk_var = tk.StringVar(value="none")
        _cat_mk_cb = ttk.Combobox(_cst, textvariable=self._cat_mk_var,
                                   values=["none", "o", "s", "^", "D", "v", "x", "+"],
                                   state="readonly", width=6)
        _cat_mk_cb.pack(side=tk.LEFT, padx=(2, 0))

        for _w in (_cat_color_e, _cat_lw_e):
            _w.bind("<Return>",   lambda e: self._on_cat_style_change())
            _w.bind("<FocusOut>", lambda e: self._on_cat_style_change())
        _cat_ls_cb.bind("<<ComboboxSelected>>", lambda e: self._on_cat_style_change())
        _cat_mk_cb.bind("<<ComboboxSelected>>", lambda e: self._on_cat_style_change())

        _cr1 = ttk.Frame(left)
        _cr1.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_cr1, text="R_sol N2 (Ω):").pack(side=tk.LEFT)
        self.r_sol_n2_var = tk.StringVar(value="0")
        _n2_e = ttk.Entry(_cr1, textvariable=self.r_sol_n2_var, width=8)
        _n2_e.pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(_cr1, text="R_sol O2 (Ω):").pack(side=tk.LEFT)
        self.r_sol_o2_var = tk.StringVar(value="0")
        _o2_e = ttk.Entry(_cr1, textvariable=self.r_sol_o2_var, width=8)
        _o2_e.pack(side=tk.LEFT, padx=(4, 0))
        for _e in (_n2_e, _o2_e):
            _e.bind("<Return>",   lambda ev: self._on_correction_change())
            _e.bind("<FocusOut>", lambda ev: self._on_correction_change())

        _cr2 = ttk.Frame(left)
        _cr2.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(_cr2, text="E_ref (V vs RHE):").pack(side=tk.LEFT)
        self.e_ref_var = tk.StringVar(value="0")
        _eref_e = ttk.Entry(_cr2, textvariable=self.e_ref_var, width=8)
        _eref_e.pack(side=tk.LEFT, padx=(4, 12))
        _eref_e.bind("<Return>",   lambda ev: self._on_correction_change())
        _eref_e.bind("<FocusOut>", lambda ev: self._on_correction_change())
        ttk.Label(_cr2, text="Ref:").pack(side=tk.LEFT)
        self.ref_electrode_var = tk.StringVar(value="RHE")
        _ref_cb = ttk.Combobox(
            _cr2, textvariable=self.ref_electrode_var,
            values=["RHE", "Ag/AgCl", "SCE", "NHE", "SHE", "Hg/HgO", "Fc/Fc+"],
            state="readonly", width=10)
        _ref_cb.pack(side=tk.LEFT, padx=(2, 0))
        _ref_cb.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())

        _cr3 = ttk.Frame(left)
        _cr3.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(_cr3, text="Area (cm²):").pack(side=tk.LEFT)
        self.area_var = tk.StringVar(value="")
        _area_e = ttk.Entry(_cr3, textvariable=self.area_var, width=8)
        _area_e.pack(side=tk.LEFT, padx=(4, 6))
        ttk.Label(_cr3, text="(leave blank for I/mA; enter for J/mA/cm²)",
                  foreground="gray", font=("", 8)).pack(side=tk.LEFT)
        _area_e.bind("<Return>",   lambda ev: self._on_correction_change())
        _area_e.bind("<FocusOut>", lambda ev: self._on_correction_change())

        # Traces: immediately write any change to the active sample's dict so that
        # FocusOut events firing after a sample switch never overwrite the wrong sample.
        for _ck, _cv in [("r_sol_n2", self.r_sol_n2_var),
                         ("r_sol_o2", self.r_sol_o2_var),
                         ("e_ref",    self.e_ref_var)]:
            _cv.trace_add("write",
                          lambda *_a, k=_ck, v=_cv: self._on_corr_var_trace(k, v))
        self.area_var.trace_add("write",
                                lambda *_a: self._on_corr_var_trace("area", self.area_var))

        ttk.Label(left, text="(auto-applied on Enter / focus change)",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)

        # ══ AXIS / RANGE ════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Plot Range  (active sample)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        _xr = ttk.Frame(left)
        _xr.pack(fill=tk.X, padx=4, pady=(1, 0))
        ttk.Label(_xr, text="X min:").pack(side=tk.LEFT)
        self.x_min_var = tk.StringVar()
        _xmin = ttk.Entry(_xr, textvariable=self.x_min_var, width=7)
        _xmin.pack(side=tk.LEFT, padx=(2, 4))
        ttk.Label(_xr, text="X max:").pack(side=tk.LEFT)
        self.x_max_var = tk.StringVar()
        _xmax = ttk.Entry(_xr, textvariable=self.x_max_var, width=7)
        _xmax.pack(side=tk.LEFT, padx=(2, 4))
        ttk.Label(_xr, text="Int:").pack(side=tk.LEFT)
        self.x_grid_int_var = tk.StringVar(value="0")
        _xgi = ttk.Entry(_xr, textvariable=self.x_grid_int_var, width=5)
        _xgi.pack(side=tk.LEFT, padx=(2, 0))

        _yr = ttk.Frame(left)
        _yr.pack(fill=tk.X, padx=4, pady=(1, 0))
        ttk.Label(_yr, text="Y min:").pack(side=tk.LEFT)
        self.y_min_var = tk.StringVar()
        _ymin = ttk.Entry(_yr, textvariable=self.y_min_var, width=7)
        _ymin.pack(side=tk.LEFT, padx=(2, 4))
        ttk.Label(_yr, text="Y max:").pack(side=tk.LEFT)
        self.y_max_var = tk.StringVar()
        _ymax = ttk.Entry(_yr, textvariable=self.y_max_var, width=7)
        _ymax.pack(side=tk.LEFT, padx=(2, 4))
        ttk.Label(_yr, text="Int:").pack(side=tk.LEFT)
        self.y_grid_int_var = tk.StringVar(value="0")
        _ygi = ttk.Entry(_yr, textvariable=self.y_grid_int_var, width=5)
        _ygi.pack(side=tk.LEFT, padx=(2, 0))

        ttk.Label(left, text="(blank = auto)", foreground="gray",
                  font=("", 8)).pack(anchor=tk.W, padx=4)
        for _re in (_xmin, _xmax, _ymin, _ymax, _xgi, _ygi):
            _re.bind("<Return>",   lambda e: self._auto_replot())
            _re.bind("<FocusOut>", lambda e: self._auto_replot())

        _flip_row = ttk.Frame(left)
        _flip_row.pack(fill=tk.X, padx=4, pady=(2, 2))
        self.x_flip_var = tk.BooleanVar(value=False)
        self.y_flip_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(_flip_row, text="Flip X", variable=self.x_flip_var,
                        command=self._auto_replot).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Checkbutton(_flip_row, text="Flip Y", variable=self.y_flip_var,
                        command=self._auto_replot).pack(side=tk.LEFT)

        _ehalf_row = ttk.Frame(left)
        _ehalf_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        self.show_half_wave_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(_ehalf_row, text="Show E½ markers",
                        variable=self.show_half_wave_var,
                        command=self._auto_replot).pack(side=tk.LEFT)

        # ══ TITLE ═══════════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        _title_row = ttk.Frame(left)
        _title_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(_title_row, text="Title:").pack(side=tk.LEFT)
        self.plot_title_var = tk.StringVar(value="")
        _title_e = ttk.Entry(_title_row, textvariable=self.plot_title_var)
        _title_e.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
        _title_e.bind("<Return>",   lambda e: self._auto_replot())
        _title_e.bind("<FocusOut>", lambda e: self._auto_replot())

        # ══ LEGEND ══════════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Legend  (active sample)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        _lr1 = ttk.Frame(left)
        _lr1.pack(fill=tk.X, padx=4, pady=2)
        self.legend_show_var  = tk.BooleanVar(value=True)
        self.legend_frame_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(_lr1, text="Show Legend", variable=self.legend_show_var,
                        command=self._auto_replot).pack(side=tk.LEFT)
        ttk.Checkbutton(_lr1, text="Frame", variable=self.legend_frame_var,
                        command=self._auto_replot).pack(side=tk.LEFT, padx=(8, 0))
        _lr2 = ttk.Frame(left)
        _lr2.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_lr2, text="Size:").pack(side=tk.LEFT)
        self.legend_size_var = tk.StringVar(value="8")
        _lsz_e = ttk.Entry(_lr2, textvariable=self.legend_size_var, width=4)
        _lsz_e.pack(side=tk.LEFT, padx=(2, 8))
        _lsz_e.bind("<Return>",   lambda e: self._auto_replot())
        _lsz_e.bind("<FocusOut>", lambda e: self._auto_replot())
        ttk.Label(_lr2, text="Loc:").pack(side=tk.LEFT)
        self.legend_loc_var = tk.StringVar(value="best")
        _leg_loc_cb = ttk.Combobox(
            _lr2, textvariable=self.legend_loc_var,
            values=["best", "upper right", "upper left", "lower left", "lower right",
                    "right", "center left", "center right", "lower center",
                    "upper center", "center"],
            state="readonly", width=11)
        _leg_loc_cb.pack(side=tk.LEFT, padx=2)

        def _on_leg_loc_select(e=None):
            if self.active_sample and self.active_sample in self.samples:
                self.samples[self.active_sample].pop("legend_manual_pos", None)
                _leg = self.samples[self.active_sample].get("legend")
                if _leg is not None:
                    _leg._loc = 0
            self._auto_replot()
        _leg_loc_cb.bind("<<ComboboxSelected>>", _on_leg_loc_select)
        ttk.Label(left, text="(left-drag to move)",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)

        # ══ GRID ════════════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Grid  (active sample)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        _gxy = ttk.Frame(left)
        _gxy.pack(fill=tk.X, padx=4, pady=2)
        self.x_grid_var = tk.BooleanVar(value=False)
        self.y_grid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(_gxy, text="X", variable=self.x_grid_var,
                        command=self._auto_replot).pack(side=tk.LEFT)
        ttk.Checkbutton(_gxy, text="Y", variable=self.y_grid_var,
                        command=self._auto_replot).pack(side=tk.LEFT, padx=(10, 0))
        _gst = ttk.Frame(left)
        _gst.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(_gst, text="Style:").pack(side=tk.LEFT)
        self.grid_style_var = tk.StringVar(value="dashed")
        _gscb = ttk.Combobox(_gst, textvariable=self.grid_style_var,
                              values=["dashed", "dotted", "solid", "dash-dot"],
                              state="readonly", width=9)
        _gscb.pack(side=tk.LEFT, padx=(2, 6))
        _gscb.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())
        ttk.Label(_gst, text="Color:").pack(side=tk.LEFT)
        self.grid_color_var = tk.StringVar(value="gray")
        _gccb = ttk.Combobox(_gst, textvariable=self.grid_color_var,
                              values=["gray", "black", "red", "blue", "green",
                                      "orange", "purple", "crimson", "royalblue",
                                      "darkorange", "teal"],
                              state="readonly", width=9)
        _gccb.pack(side=tk.LEFT, padx=(2, 6))
        _gccb.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())
        ttk.Label(_gst, text="Width:").pack(side=tk.LEFT)
        self.grid_linewidth_var = tk.StringVar(value="0.8")
        _glw = ttk.Entry(_gst, textvariable=self.grid_linewidth_var, width=4)
        _glw.pack(side=tk.LEFT, padx=(2, 0))
        _glw.bind("<Return>",   lambda e: self._auto_replot())
        _glw.bind("<FocusOut>", lambda e: self._auto_replot())

        # ══ FONT ════════════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Font  (active sample)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        self.font_title_size_var = tk.StringVar(value="10")
        self.font_title_bold_var = tk.BooleanVar(value=False)
        self.font_label_size_var = tk.StringVar(value="10")
        self.font_label_bold_var = tk.BooleanVar(value=False)
        self.font_tick_size_var  = tk.StringVar(value="8")
        self.font_tick_bold_var  = tk.BooleanVar(value=False)
        for lbl, sz_v, bd_v in (
            ("Title:    Size", self.font_title_size_var, self.font_title_bold_var),
            ("Axis Lbl: Size", self.font_label_size_var, self.font_label_bold_var),
            ("Tick Nos: Size", self.font_tick_size_var,  self.font_tick_bold_var),
        ):
            _fr = ttk.Frame(left)
            _fr.pack(fill=tk.X, padx=4, pady=(2, 0))
            ttk.Label(_fr, text=lbl).pack(side=tk.LEFT)
            _fe = ttk.Entry(_fr, textvariable=sz_v, width=4)
            _fe.pack(side=tk.LEFT, padx=(2, 4))
            _fe.bind("<Return>",   lambda e: self._auto_replot())
            _fe.bind("<FocusOut>", lambda e: self._auto_replot())
            ttk.Checkbutton(_fr, text="Bold", variable=bd_v,
                            command=self._auto_replot).pack(side=tk.LEFT)
        _spc = ttk.Frame(left)
        _spc.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Label(_spc, text="Spacing: Title").pack(side=tk.LEFT)
        self.title_pad_var = tk.StringVar(value="6")
        _tpe = ttk.Entry(_spc, textvariable=self.title_pad_var, width=4)
        _tpe.pack(side=tk.LEFT, padx=(2, 6))
        _tpe.bind("<Return>",   lambda e: self._auto_replot())
        _tpe.bind("<FocusOut>", lambda e: self._auto_replot())
        ttk.Label(_spc, text="Label").pack(side=tk.LEFT)
        self.label_pad_var = tk.StringVar(value="4")
        _lpe = ttk.Entry(_spc, textvariable=self.label_pad_var, width=4)
        _lpe.pack(side=tk.LEFT, padx=(2, 0))
        _lpe.bind("<Return>",   lambda e: self._auto_replot())
        _lpe.bind("<FocusOut>", lambda e: self._auto_replot())

        # ══ REFERENCE LINES ══════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Reference Lines  (active sample)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        _ra = ttk.Frame(left)
        _ra.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_ra, text="X:").pack(side=tk.LEFT)
        self._ref_x_var = tk.StringVar()
        _rx_e = ttk.Entry(_ra, textvariable=self._ref_x_var, width=7)
        _rx_e.pack(side=tk.LEFT, padx=(2, 0))
        ttk.Button(_ra, text="+X", width=3, command=self._add_xrefline).pack(
            side=tk.LEFT, padx=(2, 8))
        ttk.Label(_ra, text="Y:").pack(side=tk.LEFT)
        self._ref_y_var = tk.StringVar()
        _ry_e = ttk.Entry(_ra, textvariable=self._ref_y_var, width=7)
        _ry_e.pack(side=tk.LEFT, padx=(2, 0))
        ttk.Button(_ra, text="+Y", width=3, command=self._add_yrefline).pack(
            side=tk.LEFT, padx=2)
        _rx_e.bind("<Return>", lambda e: self._add_xrefline())
        _ry_e.bind("<Return>", lambda e: self._add_yrefline())

        _rl_row = ttk.Frame(left)
        _rl_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        self._reflines_lb = tk.Listbox(_rl_row, height=3,
                                       selectmode=tk.SINGLE, exportselection=False)
        self._reflines_lb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._reflines_lb.bind("<<ListboxSelect>>", lambda e: self._on_refline_select())
        ttk.Button(_rl_row, text="Remove",
                   command=self._remove_refline).pack(side=tk.RIGHT, padx=(4, 0))
        _ro = ttk.Frame(left)
        _ro.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(_ro, text="Style:").pack(side=tk.LEFT)
        self._refline_style_var = tk.StringVar(value="dashed")
        _rls_cb = ttk.Combobox(_ro, textvariable=self._refline_style_var,
                                values=["dashed", "dotted", "solid", "dash-dot"],
                                state="readonly", width=9)
        _rls_cb.pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(_ro, text="Color:").pack(side=tk.LEFT)
        self._refline_color_var = tk.StringVar(value="dimgray")
        _rlc_cb = ttk.Combobox(_ro, textvariable=self._refline_color_var,
                                values=["dimgray", "black", "red", "blue", "green",
                                        "orange", "purple", "crimson", "royalblue",
                                        "darkorange", "teal", "saddlebrown"],
                                state="readonly", width=9)
        _rlc_cb.pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(_ro, text="Width:").pack(side=tk.LEFT)
        self._refline_lw_var = tk.StringVar(value="1.0")
        _rllw = ttk.Entry(_ro, textvariable=self._refline_lw_var, width=4)
        _rllw.pack(side=tk.LEFT, padx=(2, 0))
        _rls_cb.bind("<<ComboboxSelected>>", lambda e: self._on_refline_style_change())
        _rlc_cb.bind("<<ComboboxSelected>>", lambda e: self._on_refline_style_change())
        _rllw.bind("<Return>",   lambda e: self._on_refline_style_change())
        _rllw.bind("<FocusOut>", lambda e: self._on_refline_style_change())

        # ══ PLOT BUTTON + LOG ═══════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Button(left, text="Plot Active Sample",
                   command=self._auto_replot).pack(anchor=tk.W, padx=4, pady=(0, 4))

        # ══ ANALYSIS ════════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Analysis  (active sample)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        _an_row = ttk.Frame(left)
        _an_row.pack(fill=tk.X, padx=4, pady=(2, 2))
        ttk.Button(_an_row, text="Tafel Analysis",
                   command=self._open_tafel_window).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(_an_row, text="KL Analysis",
                   command=self._open_kl_window).pack(side=tk.LEFT)

        # ══ EXPORT ══════════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Button(left, text="Export Active Sample to Excel",
                   command=self._export_sample_excel).pack(anchor=tk.W, padx=4, pady=(0, 4))

        ttk.Label(left, text="Log", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        _logf = ttk.Frame(left)
        _logf.pack(fill=tk.X, padx=4, pady=2)
        self.log_text = tk.Text(_logf, height=4, state=tk.DISABLED,
                                wrap=tk.WORD, relief=tk.SUNKEN, borderwidth=1)
        _log_sc = ttk.Scrollbar(_logf, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=_log_sc.set)
        _log_sc.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # ── Right panel ────────────────────────────────────────────
        right_outer = ttk.Frame(body)
        body.add(right_outer, weight=1)
        right_outer.rowconfigure(0, weight=0)
        right_outer.rowconfigure(1, weight=0)
        right_outer.rowconfigure(2, weight=1)
        right_outer.columnconfigure(0, weight=1)

        self.plot_w_var     = tk.StringVar(value="10.5")
        self.plot_h_var     = tk.StringVar(value="5.5")
        self._grid_cols_var = tk.StringVar(value="2")
        _size_bar = ttk.Frame(right_outer)
        _size_bar.grid(row=0, column=0, sticky="ew", padx=4, pady=2)
        ttk.Label(_size_bar, text="Plot size (in):").pack(side=tk.LEFT, padx=(4, 2))
        ttk.Label(_size_bar, text="W").pack(side=tk.LEFT)
        _pw_e = ttk.Entry(_size_bar, textvariable=self.plot_w_var, width=5)
        _pw_e.pack(side=tk.LEFT, padx=(1, 6))
        ttk.Label(_size_bar, text="H").pack(side=tk.LEFT)
        _ph_e = ttk.Entry(_size_bar, textvariable=self.plot_h_var, width=5)
        _ph_e.pack(side=tk.LEFT, padx=(1, 0))
        ttk.Label(_size_bar, text="Cols:").pack(side=tk.LEFT, padx=(10, 2))
        _gc_e = ttk.Entry(_size_bar, textvariable=self._grid_cols_var, width=3)
        _gc_e.pack(side=tk.LEFT, padx=(1, 0))
        for _e in (_pw_e, _ph_e):
            _e.bind("<Return>",   lambda ev: self._apply_plot_size())
            _e.bind("<FocusOut>", lambda ev: self._apply_plot_size())
        _gc_e.bind("<Return>",   lambda ev: self._on_grid_cols_change())
        _gc_e.bind("<FocusOut>", lambda ev: self._on_grid_cols_change())

        self._zoom_bar = ttk.Frame(right_outer)
        ttk.Button(self._zoom_bar, text="← Back to Grid",
                   command=self._unzoom_sample_view).pack(side=tk.LEFT, padx=6, pady=3)
        self._zoom_bar.grid(row=1, column=0, sticky="ew")
        self._zoom_bar.grid_remove()

        _ri = ttk.Frame(right_outer)
        _ri.grid(row=2, column=0, sticky="nsew")
        _ri.rowconfigure(0, weight=1)
        _ri.columnconfigure(0, weight=1)

        self._right_canvas = tk.Canvas(_ri, highlightthickness=0)
        _rvs = ttk.Scrollbar(_ri, orient=tk.VERTICAL,   command=self._right_canvas.yview)
        _rhs = ttk.Scrollbar(_ri, orient=tk.HORIZONTAL, command=self._right_canvas.xview)
        self._right_canvas.configure(yscrollcommand=_rvs.set, xscrollcommand=_rhs.set)
        _rvs.grid(row=0, column=1, sticky="ns")
        _rhs.grid(row=1, column=0, sticky="ew")
        self._right_canvas.grid(row=0, column=0, sticky="nsew")

        self._plots_frame = ttk.Frame(self._right_canvas)
        self._plots_win   = self._right_canvas.create_window(
            (0, 0), window=self._plots_frame, anchor=tk.NW)
        self._plots_frame.bind(
            "<Configure>",
            lambda e: self._right_canvas.configure(
                scrollregion=self._right_canvas.bbox("all")))
        self._right_canvas.bind("<Configure>", self._on_right_canvas_configure)
        self._right_canvas.bind(
            "<MouseWheel>",
            lambda e: self._right_canvas.yview_scroll(-1 * (e.delta // 120), "units"))
        self._right_canvas.bind(
            "<Shift-MouseWheel>",
            lambda e: self._right_canvas.xview_scroll(-1 * (e.delta // 120), "units"))

        self._drop_line = tk.Frame(self._plots_frame, bg="#1a73e8", height=3)

        self._placeholder = ttk.Label(
            self._plots_frame,
            text="Create samples and load N2+O2 files to display ORR polarization curves.",
            foreground="gray", font=("", 10))
        self._placeholder.grid(row=0, column=0, columnspan=2, pady=60)

    # ════════════════════════════════════════════════════════════════
    # File loading (no auto-merge)
    # ════════════════════════════════════════════════════════════════
    def _load_files(self):
        paths = filedialog.askopenfilenames(
            title="Load ORR Files (N2 and O2)",
            filetypes=[("EC-Lab files", "*.mpr *.txt"), ("All files", "*.*")])
        if not paths:
            return
        added, errors = [], []
        for path in paths:
            stem  = os.path.splitext(os.path.basename(path))[0]
            short = os.path.basename(path)
            if short in self.loaded_files:
                continue
            try:
                df = _read_one_df(path)
            except Exception as exc:
                errors.append(f"{short}: {exc}")
                continue
            gas      = _detect_gas(stem)
            rpm_id   = _extract_rpm_id(stem)
            catalyst = _detect_catalyst(stem)
            self.loaded_files[short] = {
                "path": path, "gas": gas, "rpm_id": rpm_id,
                "catalyst": catalyst, "df": df}
            self._loaded_keys.append(short)
            cat_key   = catalyst if catalyst else ""
            cat_iid   = f"_cat_:{cat_key}"
            cat_label = catalyst if catalyst else "(no catalyst)"
            if not self.loaded_tv.exists(cat_iid):
                self.loaded_tv.insert("", tk.END, iid=cat_iid,
                                      text=f"  ▸ {cat_label}",
                                      tags=("hdr",), open=True)
            gas_tag = gas.upper() if gas else "??"
            self.loaded_tv.insert(cat_iid, tk.END, iid=short,
                                  text=f"    ({gas_tag})  {short}",
                                  tags=(gas or "unk",))
            added.append(short)

        if errors:
            messagebox.showerror("Load Error", "\n".join(errors[:5]))
        if added:
            self._log(f"Loaded {len(added)} file(s).")

    def _remove_loaded_file(self):
        sel = [iid for iid in self.loaded_tv.selection()
               if not iid.startswith("_cat_:")]
        if not sel:
            return
        for short in sel:
            if short not in self.loaded_files:
                continue
            parent = self.loaded_tv.parent(short)
            self.loaded_tv.delete(short)
            self.loaded_files.pop(short, None)
            if short in self._loaded_keys:
                self._loaded_keys.remove(short)
            if parent and not self.loaded_tv.get_children(parent):
                self.loaded_tv.delete(parent)

    def _on_loaded_tv_dblclick(self, event):
        """Double-click on a catalyst header → rename it."""
        iid = self.loaded_tv.identify_row(event.y)
        if iid and iid.startswith("_cat_:"):
            self._set_catalyst_name(iid)

    def _set_catalyst_name(self, cat_iid=None):
        """Rename the catalyst group and propagate to pairs."""
        if cat_iid is None:
            sel = self.loaded_tv.selection()
            if not sel:
                return
            iid = sel[0]
            cat_iid = iid if iid.startswith("_cat_:") else self.loaded_tv.parent(iid)
        if not cat_iid or not cat_iid.startswith("_cat_:"):
            return

        old_cat = cat_iid[len("_cat_:"):]
        if not cat_iid or not cat_iid.startswith("_cat_:"):
            messagebox.showwarning("ORR", "Could not determine catalyst group.")
            return

        old_cat = cat_iid[len("_cat_:"):]

        from tkinter.simpledialog import askstring
        new_cat = askstring(
            "Rename Catalyst",
            "Catalyst name:",
            initialvalue=old_cat,
            parent=self,
        )
        if new_cat is None:
            return
        new_cat = new_cat.strip()
        if new_cat == old_cat:
            return

        file_shorts = set(self.loaded_tv.get_children(cat_iid))

        # Update loaded_files
        for short in file_shorts:
            if short in self.loaded_files:
                self.loaded_files[short]["catalyst"] = new_cat

        # Update treeview header text
        self.loaded_tv.item(cat_iid,
                            text=f"  ▸ {new_cat if new_cat else '(no catalyst)'}")

        # Propagate to all sample pairs that reference these files
        affected_samples = set()
        for sname, sentry in self.samples.items():
            cc = sentry.setdefault("catalyst_corrections", {})
            cs = sentry.setdefault("catalyst_styles", {})
            renamed: dict = {}  # old_id → new_id

            for pair in sentry["pairs"]:
                if (pair.get("n2_short") not in file_shorts
                        and pair.get("o2_short") not in file_shorts):
                    continue
                old_base = pair.get("catalyst_base", "")
                old_id   = pair.get("catalyst_id",   "")
                suffix   = old_id[len(old_base):] if old_id.startswith(old_base) else ""
                new_id   = new_cat + suffix
                renamed[old_id] = new_id
                pair["catalyst_base"] = new_cat
                pair["catalyst_id"]   = new_id
                affected_samples.add(sname)

            for old_id, new_id in renamed.items():
                if old_id != new_id:
                    if old_id in cc:
                        cc[new_id] = cc.pop(old_id)
                    if old_id in cs:
                        cs[new_id] = cs.pop(old_id)

        for sname in affected_samples:
            self._plot_sample(sname)
        if self.active_sample in affected_samples:
            self._rebuild_pair_table(self.active_sample)
            self._update_catalyst_selector(self.active_sample)

    def _select_n2(self):
        shorts = [k for k in self._loaded_keys
                  if self.loaded_files.get(k, {}).get("gas") == "n2"]
        self.loaded_tv.selection_set(shorts)

    def _select_o2(self):
        shorts = [k for k in self._loaded_keys
                  if self.loaded_files.get(k, {}).get("gas") == "o2"]
        self.loaded_tv.selection_set(shorts)

    # ════════════════════════════════════════════════════════════════
    # Sample management
    # ════════════════════════════════════════════════════════════════
    def _new_sample(self):
        name = simpledialog.askstring("New ORR Sample", "Sample name:", parent=self)
        if not name:
            return
        name = name.strip()
        if not name or name in self.samples:
            return
        self.samples[name] = {"pairs": [], "reflines": []}
        self._rebuild_sample_listbox()
        self._create_sample_figure(name)
        self._save_active_sample_state()
        self._switch_active_sample(name)

    def _rename_sample(self):
        sel = self.sample_lb.curselection()
        if not sel:
            return
        old = list(self.samples.keys())[sel[0]]
        new = simpledialog.askstring("Rename Sample", "New name:",
                                     initialvalue=old, parent=self)
        if not new:
            return
        new = new.strip()
        if not new or new == old or new in self.samples:
            return
        new_samples = OrderedDict()
        for k, v in self.samples.items():
            new_samples[new if k == old else k] = v
        self.samples = new_samples
        if self.active_sample == old:
            self.active_sample = new
        self._rebuild_sample_listbox()
        gentry = self.samples.get(new, {})
        if "ax" in gentry:
            ax = gentry["ax"]
            ax.set_title(gentry.get("custom_title", "") or new, fontsize=9)
            gentry["canvas"].draw_idle()

    def _delete_sample(self):
        sel = self.sample_lb.curselection()
        if not sel:
            return
        name = list(self.samples.keys())[sel[0]]
        self._destroy_sample_figure(name)
        del self.samples[name]
        if self.active_sample == name:
            self.active_sample = None
            self._rebuild_pair_table(None)
            self._refresh_reflines_lb()
        if not self.samples:
            self._rebuild_sample_listbox()
            self._placeholder.grid(row=0, column=0, columnspan=2, pady=60)
        else:
            self._rebuild_sample_listbox()
            self._relayout_figures()
            new_idx = min(sel[0], self.sample_lb.size() - 1)
            if new_idx >= 0:
                self._switch_active_sample(list(self.samples.keys())[new_idx])

    def _rebuild_sample_listbox(self):
        self.sample_lb.clear()
        for sn, sentry in self.samples.items():
            vis = not sentry.get("hidden", False)
            self.sample_lb.insert(tk.END, sn, checked=vis)
        if self.active_sample and self.active_sample in self.samples:
            idx = list(self.samples.keys()).index(self.active_sample)
            self._loading = True
            try:
                self.sample_lb.selection_clear(0, tk.END)
                self.sample_lb.selection_set(idx)
            finally:
                self._loading = False

    def _on_sample_select(self, event):
        if self._loading:
            return
        sel = self.sample_lb.curselection()
        if not sel:
            return
        keys = list(self.samples.keys())
        if sel[0] >= len(keys):
            return
        name = keys[sel[0]]
        if name != self.active_sample:
            self._save_active_sample_state()
            self._switch_active_sample(name)

    def _on_sample_visibility_change(self, sample_name, visible):
        sentry = self.samples.get(sample_name)
        if sentry is None:
            return
        sentry["hidden"] = not visible
        self._relayout_figures()
        self._right_canvas.after(50, lambda: self._right_canvas.configure(
            scrollregion=self._right_canvas.bbox("all")))

    def _on_sample_reorder(self, new_names):
        self.samples = OrderedDict(
            (n, self.samples[n]) for n in new_names if n in self.samples)
        self._relayout_figures()

    # ════════════════════════════════════════════════════════════════
    # File → Sample assignment + pairing
    # ════════════════════════════════════════════════════════════════
    def _add_files_to_sample(self):
        if not self.active_sample:
            messagebox.showwarning("ORR", "Create or select a sample first.")
            return
        sel_iids = self.loaded_tv.selection()
        if not sel_iids:
            return
        sentry = self.samples[self.active_sample]

        expanded = []
        for iid in sel_iids:
            if iid.startswith("_cat_:"):
                expanded.extend(self.loaded_tv.get_children(iid))
            else:
                expanded.append(iid)
        selected_shorts = [iid for iid in expanded if iid in self.loaded_files]

        unknown = []
        file_entries = []
        for short in selected_shorts:
            fe = self.loaded_files.get(short)
            if fe is None:
                continue
            if fe["gas"] not in ("n2", "o2"):
                unknown.append(short)
            else:
                file_entries.append((short, fe))

        if unknown:
            ans = messagebox.askyesno(
                "Undetected Gas",
                f"{len(unknown)} file(s) have no 'n2'/'o2' in their name.\n"
                "Skip them?  (Yes = skip, No = cancel)")
            if not ans:
                return

        # Track all paths already committed to any pair slot (exact-dup guard).
        all_existing_paths = set()
        for p in sentry["pairs"]:
            if p.get("n2_path"):
                all_existing_paths.add(p["n2_path"])
            if p.get("o2_path"):
                all_existing_paths.add(p["o2_path"])

        added = 0

        # ── Batch-isolation: group this batch by (catalyst_base, rpm_id) ──
        # First-seen gas wins each slot; duplicates fall through to lone-file path.
        batch_groups: dict = {}
        for short, fe in file_entries:
            key = (fe.get("catalyst", ""), fe["rpm_id"])
            gas = fe["gas"]
            batch_groups.setdefault(key, {})
            if gas not in batch_groups[key]:
                batch_groups[key][gas] = (short, fe)

        # Step 1: create complete pairs from within this batch.
        # Batch-internal pairs are never merged into existing incomplete pairs —
        # this prevents cross-contamination when the same catalyst/rpm appears in
        # multiple experiments loaded in separate batches.
        batch_paired: set = set()
        for (cat_base, rpm_id), gas_map in batch_groups.items():
            if "n2" not in gas_map or "o2" not in gas_map:
                continue
            n2_short, n2_fe = gas_map["n2"]
            o2_short, o2_fe = gas_map["o2"]
            if n2_fe["path"] in all_existing_paths or o2_fe["path"] in all_existing_paths:
                batch_paired.add(n2_short)
                batch_paired.add(o2_short)
                continue
            used_cat_rpm = {(p.get("catalyst_id", ""), p.get("rpm_id", ""))
                            for p in sentry["pairs"]}
            cat_label = cat_base
            suffix = 2
            while (cat_label, rpm_id) in used_cat_rpm:
                cat_label = f"{cat_base}_{suffix}"
                suffix += 1
            new_pair = {
                "catalyst_id":   cat_label,
                "catalyst_base": cat_base,
                "rpm_id":        rpm_id,
                "rpm_val":       rpm_id,
                "n2_short":      n2_short,
                "o2_short":      o2_short,
                "n2_path":       n2_fe["path"],
                "o2_path":       o2_fe["path"],
                "df_n2":         n2_fe["df"],
                "df_o2":         o2_fe["df"],
            }
            sentry["pairs"].append(new_pair)
            all_existing_paths.add(n2_fe["path"])
            all_existing_paths.add(o2_fe["path"])
            batch_paired.add(n2_short)
            batch_paired.add(o2_short)
            cc = sentry.setdefault("catalyst_corrections", {})
            cc.setdefault(cat_label, {"r_sol_n2": 0.0, "r_sol_o2": 0.0,
                                      "e_ref": 0.0, "area": ""})
            cs = sentry.setdefault("catalyst_styles", {})
            cs.setdefault(cat_label, {"color": "", "linestyle": "solid",
                                      "linewidth": "1.5", "marker": "none"})
            added += 2

        # Step 2: process lone files not handled by batch-internal pairing.
        for short, fe in file_entries:
            if short in batch_paired:
                continue
            gas      = fe["gas"]
            rpm_id   = fe["rpm_id"]
            catalyst = fe.get("catalyst", "")
            path     = fe["path"]

            if path in all_existing_paths:
                continue

            # Try to merge into an existing incomplete pair that matches
            # (catalyst_base, rpm_id) and has this gas slot empty.
            merged = False
            for pair in sentry["pairs"]:
                pair_base = pair.get("catalyst_base", pair.get("catalyst_id", ""))
                if (pair_base == catalyst
                        and pair.get("rpm_id") == rpm_id
                        and not pair.get(f"{gas}_path")):
                    pair[f"{gas}_path"]  = path
                    pair[f"{gas}_short"] = short
                    pair[f"df_{gas}"]    = fe["df"]
                    merged = True
                    break

            if merged:
                all_existing_paths.add(path)
                added += 1
                continue

            # No mergeable slot — create a new (possibly incomplete) pair.
            # Auto-suffix catalyst label if (label, rpm_id) is already taken.
            used_cat_rpm = {(p.get("catalyst_id", ""), p.get("rpm_id", ""))
                            for p in sentry["pairs"]}
            cat_label = catalyst
            suffix = 2
            while (cat_label, rpm_id) in used_cat_rpm:
                cat_label = f"{catalyst}_{suffix}"
                suffix += 1
            new_pair = {
                "catalyst_id":   cat_label,
                "catalyst_base": catalyst,
                "rpm_id":        rpm_id,
                "rpm_val":       rpm_id,
                "n2_short":      short if gas == "n2" else "",
                "o2_short":      short if gas == "o2" else "",
                "n2_path":       path  if gas == "n2" else "",
                "o2_path":       path  if gas == "o2" else "",
                "df_n2":         fe["df"] if gas == "n2" else None,
                "df_o2":         fe["df"] if gas == "o2" else None,
            }
            sentry["pairs"].append(new_pair)
            all_existing_paths.add(path)
            cc = sentry.setdefault("catalyst_corrections", {})
            cc.setdefault(cat_label, {"r_sol_n2": 0.0, "r_sol_o2": 0.0,
                                      "e_ref": 0.0, "area": ""})
            cs = sentry.setdefault("catalyst_styles", {})
            cs.setdefault(cat_label, {"color": "", "linestyle": "solid",
                                      "linewidth": "1.5", "marker": "none"})
            added += 1

        sentry["pairs"].sort(key=lambda p: (p.get("catalyst_id", ""), p.get("rpm_id", "")))
        if added:
            self._rebuild_pair_table(self.active_sample)
            self._auto_replot()
            self._log(f"Added {added} pair(s) to sample '{self.active_sample}'.")
            self._update_catalyst_selector(self.active_sample)
        else:
            self._log("No new pairs added (already exist or no matching IDs).")

    # ════════════════════════════════════════════════════════════════
    # Pair table
    # ════════════════════════════════════════════════════════════════
    def _rebuild_pair_table(self, sample_name):
        for w in self._pair_tbl_inner.winfo_children():
            w.destroy()

        sentry = self.samples.get(sample_name) if sample_name else None
        pairs  = sentry["pairs"] if sentry else []

        if not pairs:
            tk.Label(self._pair_tbl_inner,
                     text="No pairs yet. Load files and click\n"
                          "'↓ Add Selected Files to Sample'.",
                     bg="#f5f5f5", fg="gray", font=("", 8),
                     justify=tk.LEFT).pack(padx=6, pady=6, anchor=tk.W)
            return

        # Header — checkbox col + catalyst + rpm fixed; N2/O2 expand; remove btn right
        hdr = tk.Frame(self._pair_tbl_inner, bg="#f5f5f5")
        hdr.pack(fill=tk.X, padx=2, pady=(2, 0))
        tk.Label(hdr, text="Plot", width=4, bg="#f5f5f5",
                 font=("", 8, "bold"), anchor=tk.W).pack(side=tk.LEFT)
        tk.Label(hdr, text="Catalyst", width=7, bg="#f5f5f5",
                 font=("", 8, "bold"), anchor=tk.W).pack(side=tk.LEFT)
        tk.Label(hdr, text="RPM", width=5, bg="#f5f5f5",
                 font=("", 8, "bold"), anchor=tk.W).pack(side=tk.LEFT)
        tk.Label(hdr, text="", width=2, bg="#f5f5f5").pack(side=tk.RIGHT)
        tk.Label(hdr, text="O2 file", bg="#f5f5f5",
                 font=("", 8, "bold"), anchor=tk.W).pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(2, 0))
        tk.Label(hdr, text="N2 file", bg="#f5f5f5",
                 font=("", 8, "bold"), anchor=tk.W).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 0))

        prev_cat = None
        row_idx  = 0
        for pair in pairs:
            cat = pair.get("catalyst_id", "")
            if cat != prev_cat:
                _sep_bg = "#e8d5f5"
                _sep = tk.Frame(self._pair_tbl_inner, bg=_sep_bg)
                _sep.pack(fill=tk.X, padx=2, pady=(6 if prev_cat is not None else 2, 0))
                tk.Label(_sep, text=f"── {cat} ──", bg=_sep_bg,
                         font=("", 8, "bold"), anchor=tk.W, padx=6).pack(fill=tk.X)
                prev_cat = cat
                row_idx  = 0

            i = row_idx
            row_idx += 1
            row_bg = "#ffffff" if i % 2 == 0 else "#eeeeee"
            row = tk.Frame(self._pair_tbl_inner, bg=row_bg)
            row.pack(fill=tk.X, padx=2, pady=1)

            cat_var     = tk.StringVar(value=pair.get("catalyst_id", ""))
            rpm_var     = tk.StringVar(value=pair.get("rpm_val", pair.get("rpm_id", "")))
            enabled_var = tk.BooleanVar(value=pair.get("enabled", True))

            def _toggle_enabled(p=pair, ev=enabled_var):
                p["enabled"] = ev.get()
                self._auto_replot()

            def _save_pair(cv=cat_var, rv=rpm_var, p=pair, sn=sample_name):
                new_cat = cv.get().strip()
                old_cat = p.get("catalyst_id", "")
                if new_cat and new_cat != old_cat:
                    sentry_ref = self.samples.get(sn, {})
                    for _p in sentry_ref.get("pairs", []):
                        if _p.get("catalyst_id") == old_cat:
                            _p["catalyst_id"] = new_cat
                    for _store_key in ("catalyst_corrections", "catalyst_styles"):
                        _store = sentry_ref.get(_store_key, {})
                        if old_cat in _store:
                            _store[new_cat] = _store.pop(old_cat)
                    if getattr(self, "_active_catalyst", None) == old_cat:
                        self._active_catalyst = new_cat
                    self._rebuild_pair_table(sn)
                    self._update_catalyst_selector(sn)
                else:
                    p["catalyst_id"] = new_cat or old_cat
                p["rpm_val"] = rv.get()
                self._auto_replot()

            # Checkbox — left edge
            tk.Checkbutton(row, variable=enabled_var, command=_toggle_enabled,
                           bg=row_bg, activebackground=row_bg,
                           relief=tk.FLAT).pack(side=tk.LEFT, padx=(2, 0))

            cat_e = tk.Entry(row, textvariable=cat_var, width=7, bg=row_bg,
                             relief=tk.GROOVE)
            cat_e.pack(side=tk.LEFT, padx=(1, 1))
            cat_e.bind("<Return>",   lambda e, fn=_save_pair: fn())
            cat_e.bind("<FocusOut>", lambda e, fn=_save_pair: fn())

            rpm_e = tk.Entry(row, textvariable=rpm_var, width=5, bg=row_bg,
                             relief=tk.GROOVE)
            rpm_e.pack(side=tk.LEFT, padx=(0, 2))
            rpm_e.bind("<Return>",   lambda e, fn=_save_pair: fn())
            rpm_e.bind("<FocusOut>", lambda e, fn=_save_pair: fn())

            n2_s = pair.get("n2_short", "") or "—"
            o2_s = pair.get("o2_short", "") or "—"
            n2_bg = "#d4edda" if pair.get("n2_short") else "#f8d7da"
            o2_bg = "#d1ecf1" if pair.get("o2_short") else "#f8d7da"

            def _remove(p=pair, sn=sample_name):
                if sn and sn in self.samples:
                    try:
                        self.samples[sn]["pairs"].remove(p)
                    except ValueError:
                        pass
                    self._rebuild_pair_table(sn)
                    self._auto_replot()

            # Remove button right-anchored so filename labels get all remaining space
            tk.Button(row, text="✕", width=2, command=_remove,
                      bg=row_bg, relief=tk.FLAT, font=("", 8)).pack(side=tk.RIGHT, padx=(1, 2))
            tk.Label(row, text=o2_s, bg=o2_bg, font=("", 7), anchor=tk.W,
                     relief=tk.GROOVE).pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(0, 1))
            tk.Label(row, text=n2_s, bg=n2_bg, font=("", 7), anchor=tk.W,
                     relief=tk.GROOVE).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 1))

    # ════════════════════════════════════════════════════════════════
    # Copy / Paste sample display format
    # ════════════════════════════════════════════════════════════════
    _FMT_KEYS = (
        "x_grid", "y_grid", "x_grid_int", "y_grid_int",
        "grid_style", "grid_color", "grid_linewidth",
        "x_flip", "y_flip",
        "legend_show", "legend_frame", "leg_size", "legend_loc",
        "font_title_size", "font_title_bold",
        "font_label_size", "font_label_bold",
        "font_tick_size",  "font_tick_bold",
        "title_pad", "label_pad",
        "ref_electrode",
    )

    def _copy_sample_format(self):
        if not self.active_sample or self.active_sample not in self.samples:
            return
        self._save_active_sample_state()
        g = self.samples[self.active_sample]
        self._copied_sample_fmt = {k: g.get(k) for k in self._FMT_KEYS}
        self._copied_sample_fmt["reflines"] = list(g.get("reflines", []))

    def _paste_sample_format(self):
        if not self._copied_sample_fmt:
            return
        if not self.active_sample or self.active_sample not in self.samples:
            return
        g = self.samples[self.active_sample]
        p = self._copied_sample_fmt
        for k in self._FMT_KEYS:
            if k in p and p[k] is not None:
                g[k] = p[k]
        if "reflines" in p:
            g["reflines"] = list(p["reflines"])
        self._switch_active_sample(self.active_sample)
        self._plot_sample(self.active_sample)

    # ════════════════════════════════════════════════════════════════
    # Active sample state save / restore
    # ════════════════════════════════════════════════════════════════
    def _save_active_sample_state(self):
        if not self.active_sample or self.active_sample not in self.samples:
            return
        g = self.samples[self.active_sample]

        def _fv(var, default=0.0):
            try:
                return float(var.get())
            except (ValueError, tk.TclError):
                return default

        g["ref_electrode"]   = self.ref_electrode_var.get()
        g["x_min"]           = self.x_min_var.get()
        g["x_max"]           = self.x_max_var.get()
        g["y_min"]           = self.y_min_var.get()
        g["y_max"]           = self.y_max_var.get()
        g["x_grid_int"]      = self.x_grid_int_var.get()
        g["y_grid_int"]      = self.y_grid_int_var.get()
        g["x_flip"]          = self.x_flip_var.get()
        g["y_flip"]          = self.y_flip_var.get()
        g["legend_show"]     = self.legend_show_var.get()
        g["legend_frame"]    = self.legend_frame_var.get()
        try:
            g["leg_size"]    = float(self.legend_size_var.get())
        except (ValueError, tk.TclError):
            pass
        g["legend_loc"]      = self.legend_loc_var.get()
        g["x_grid"]          = self.x_grid_var.get()
        g["y_grid"]          = self.y_grid_var.get()
        g["grid_style"]      = self.grid_style_var.get()
        g["grid_color"]      = self.grid_color_var.get()
        g["grid_linewidth"]  = self.grid_linewidth_var.get()
        g["font_title_size"] = self.font_title_size_var.get()
        g["font_title_bold"] = self.font_title_bold_var.get()
        g["font_label_size"] = self.font_label_size_var.get()
        g["font_label_bold"] = self.font_label_bold_var.get()
        g["font_tick_size"]  = self.font_tick_size_var.get()
        g["font_tick_bold"]  = self.font_tick_bold_var.get()
        g["title_pad"]       = self.title_pad_var.get()
        g["label_pad"]       = self.label_pad_var.get()
        g["custom_title"]    = self.plot_title_var.get()
        g["show_half_wave"]  = self.show_half_wave_var.get()

    def _switch_active_sample(self, sample_name):
        self._switching_sample = True
        self.active_sample = sample_name
        g = self.samples.get(sample_name, {})

        g.setdefault("ref_electrode",   "RHE")
        g.setdefault("x_min",           "")
        g.setdefault("x_max",           "")
        g.setdefault("y_min",           "")
        g.setdefault("y_max",           "")
        g.setdefault("x_grid_int",      "0")
        g.setdefault("y_grid_int",      "0")
        g.setdefault("x_flip",          False)
        g.setdefault("y_flip",          False)
        g.setdefault("legend_show",     True)
        g.setdefault("legend_frame",    True)
        g.setdefault("leg_size",        8.0)
        g.setdefault("legend_loc",      "best")
        g.setdefault("x_grid",          False)
        g.setdefault("y_grid",          False)
        g.setdefault("grid_style",      "dashed")
        g.setdefault("grid_color",      "gray")
        g.setdefault("grid_linewidth",  "0.8")
        g.setdefault("font_title_size", "10")
        g.setdefault("font_title_bold", False)
        g.setdefault("font_label_size", "10")
        g.setdefault("font_label_bold", False)
        g.setdefault("font_tick_size",  "8")
        g.setdefault("font_tick_bold",  False)
        g.setdefault("title_pad",       "6")
        g.setdefault("label_pad",       "4")
        g.setdefault("custom_title",    "")
        g.setdefault("show_half_wave",  True)
        g.setdefault("reflines",        [])
        g.setdefault("custom_xlabel",   None)
        g.setdefault("custom_ylabel",   None)

        old = self._suppress_replot
        self._suppress_replot = True
        try:
            def _sv(var, key, default=""):
                v = g.get(key, default)
                var.set("" if v == 0.0 and default == "" else str(v))

            self._update_catalyst_selector(sample_name)
            self.ref_electrode_var.set(g["ref_electrode"])
            self.x_min_var.set(g["x_min"])
            self.x_max_var.set(g["x_max"])
            self.y_min_var.set(g["y_min"])
            self.y_max_var.set(g["y_max"])
            self.x_grid_int_var.set(g["x_grid_int"])
            self.y_grid_int_var.set(g["y_grid_int"])
            self.x_flip_var.set(g["x_flip"])
            self.y_flip_var.set(g["y_flip"])
            self.legend_show_var.set(g["legend_show"])
            self.legend_frame_var.set(g["legend_frame"])
            _ls = g["leg_size"]
            self.legend_size_var.set(str(int(_ls) if float(_ls) == int(_ls) else _ls))
            self.legend_loc_var.set(g["legend_loc"])
            self.x_grid_var.set(g["x_grid"])
            self.y_grid_var.set(g["y_grid"])
            self.grid_style_var.set(g["grid_style"])
            self.grid_color_var.set(g["grid_color"])
            self.grid_linewidth_var.set(g["grid_linewidth"])
            self.font_title_size_var.set(g["font_title_size"])
            self.font_title_bold_var.set(g["font_title_bold"])
            self.font_label_size_var.set(g["font_label_size"])
            self.font_label_bold_var.set(g["font_label_bold"])
            self.font_tick_size_var.set(g["font_tick_size"])
            self.font_tick_bold_var.set(g["font_tick_bold"])
            self.title_pad_var.set(g["title_pad"])
            self.label_pad_var.set(g["label_pad"])
            self.plot_title_var.set(g.get("custom_title", ""))
            self.show_half_wave_var.set(g.get("show_half_wave", True))
            self._rebuild_pair_table(sample_name)
            self._refresh_reflines_lb()
        finally:
            self._suppress_replot = old
            self._switching_sample = False

        self._highlight_active_headers()
        self._auto_replot()

    # ════════════════════════════════════════════════════════════════
    # Per-catalyst correction UI helpers
    # ════════════════════════════════════════════════════════════════
    def _update_catalyst_selector(self, sample_name):
        """Rebuild the catalyst combobox and load the first (or previously selected) catalyst."""
        sentry = self.samples.get(sample_name, {})
        cats = list(sentry.get("catalyst_corrections", {}).keys())
        # Ensure any catalyst that has pairs but no entry yet is included
        for p in sentry.get("pairs", []):
            c = p.get("catalyst_id", "")
            if c and c not in cats:
                cats.append(c)
        self._corr_cat_cb["values"] = cats
        if cats:
            # Try to keep the previously selected catalyst if it still exists
            prev = getattr(self, "_active_catalyst", None)
            sel = prev if prev in cats else cats[0]
            self._corr_catalyst_var.set(sel)
            self._load_catalyst_corrections(sample_name, sel)
        else:
            self._corr_catalyst_var.set("")
            self._active_catalyst = None
            self._switching_sample = True
            try:
                self.r_sol_n2_var.set("0")
                self.r_sol_o2_var.set("0")
                self.e_ref_var.set("0")
                self.area_var.set("")
                self._cat_color_var.set("")
                self._cat_ls_var.set("solid")
                self._cat_lw_var.set("1.5")
                self._cat_mk_var.set("none")
            finally:
                self._switching_sample = False

    def _load_catalyst_corrections(self, sample_name, catalyst_id):
        """Load a catalyst's stored correction values into the shared UI vars."""
        sentry = self.samples.get(sample_name, {})
        cc = sentry.get("catalyst_corrections", {}).get(catalyst_id, {})
        self._switching_sample = True
        try:
            self.r_sol_n2_var.set(str(cc.get("r_sol_n2", 0.0)))
            self.r_sol_o2_var.set(str(cc.get("r_sol_o2", 0.0)))
            self.e_ref_var.set(str(cc.get("e_ref", 0.0)))
            self.area_var.set(cc.get("area", ""))
        finally:
            self._switching_sample = False
        self._active_catalyst = catalyst_id
        sentry = self.samples.get(sample_name, {})
        cs = sentry.get("catalyst_styles", {}).get(catalyst_id, {})
        self._switching_sample = True
        try:
            self._cat_color_var.set(cs.get("color", ""))
            self._cat_ls_var.set(cs.get("linestyle", "solid"))
            self._cat_lw_var.set(cs.get("linewidth", "1.5"))
            self._cat_mk_var.set(cs.get("marker", "none"))
        finally:
            self._switching_sample = False

    def _on_corr_catalyst_select(self, event=None):
        """User picked a different catalyst in the combobox — load its corrections."""
        cat = self._corr_catalyst_var.get()
        if cat and self.active_sample and self.active_sample in self.samples:
            self._load_catalyst_corrections(self.active_sample, cat)

    def _on_cat_style_change(self):
        if getattr(self, "_switching_sample", False):
            return
        if not self.active_sample or self.active_sample not in self.samples:
            return
        cat = getattr(self, "_active_catalyst", None)
        if not cat:
            return
        sentry = self.samples[self.active_sample]
        cs = sentry.setdefault("catalyst_styles", {})
        cat_cs = cs.setdefault(cat, {"color": "", "linestyle": "solid",
                                      "linewidth": "1.5", "marker": "none"})
        cat_cs["color"]     = self._cat_color_var.get().strip()
        cat_cs["linestyle"] = self._cat_ls_var.get()
        cat_cs["linewidth"] = self._cat_lw_var.get()
        cat_cs["marker"]    = self._cat_mk_var.get()
        self._auto_replot()

    # ════════════════════════════════════════════════════════════════
    # Correction trigger + per-sample immediate write
    # ════════════════════════════════════════════════════════════════
    def _on_corr_var_trace(self, key: str, var):
        if getattr(self, "_switching_sample", False):
            return
        if not self.active_sample or self.active_sample not in self.samples:
            return
        cat = getattr(self, "_active_catalyst", None)
        if not cat:
            return
        sentry = self.samples[self.active_sample]
        cc = sentry.setdefault("catalyst_corrections", {})
        cat_cc = cc.setdefault(cat, {"r_sol_n2": 0.0, "r_sol_o2": 0.0,
                                      "e_ref": 0.0, "area": ""})
        try:
            raw = var.get()
        except tk.TclError:
            return
        if key == "area":
            cat_cc["area"] = raw
        else:
            try:
                cat_cc[key] = float(raw or 0)
            except ValueError:
                pass

    def _on_correction_change(self):
        if getattr(self, "_switching_sample", False):
            return
        self._save_active_sample_state()
        self._auto_replot()

    # ════════════════════════════════════════════════════════════════
    # Plotting
    # ════════════════════════════════════════════════════════════════
    def _auto_replot(self):
        if self._suppress_replot or not self.active_sample:
            return
        self._plot_sample(self.active_sample)

    def _plot_sample(self, sample_name):
        sentry = self.samples.get(sample_name)
        if sentry is None or "ax" not in sentry:
            return

        is_active = (sample_name == self.active_sample)

        def _gv(key, default=""):
            var = getattr(self, key + "_var", None)
            return (var.get() if (is_active and var is not None)
                    else sentry.get(key, default))

        ref      = _gv("ref_electrode", "RHE")
        leg_show = sentry.get("legend_show", True)  if not is_active else self.legend_show_var.get()
        leg_frm  = sentry.get("legend_frame", True) if not is_active else self.legend_frame_var.get()
        leg_loc  = sentry.get("legend_loc", "best") if not is_active else self.legend_loc_var.get()
        try:
            leg_size = float(sentry.get("leg_size", 8.0) if not is_active
                             else self.legend_size_var.get())
        except (ValueError, TypeError):
            leg_size = 8.0

        ax     = sentry["ax"]
        canvas = sentry["canvas"]

        # Preserve zoom
        _prev_view = ((ax.get_xlim(), ax.get_ylim())
                      if sentry.get("auto_xlim") is not None else None)

        # Save legend manual position
        _old_leg = sentry.get("legend")
        if _old_leg is not None:
            _loc = getattr(_old_leg, "_loc", None)
            if isinstance(_loc, (tuple, list)):
                sentry["legend_manual_pos"] = tuple(_loc)
        sentry["legend"] = None
        self._clear_ann(sample_name, redraw=False)
        ax.clear()

        pairs = sentry.get("pairs", [])
        plot_data = []    # (E_arr, Y_arr, label, color) — cached for annotation

        # Group pairs by catalyst_id to assign per-catalyst base color + RPM gradient
        catalyst_order = []
        pairs_by_cat   = {}
        for pair in pairs:
            if not pair.get("enabled", True):
                continue
            cat = pair.get("catalyst_id", "")
            if cat not in catalyst_order:
                catalyst_order.append(cat)
                pairs_by_cat[cat] = []
            pairs_by_cat[cat].append(pair)

        for ci, cat in enumerate(catalyst_order):
            base_col  = _PALETTE[ci % len(_PALETTE)]
            cat_pairs = sorted(pairs_by_cat[cat],
                               key=lambda p: p.get("rpm_id", ""))
            n_cat = len(cat_pairs)
            cat_colors = (_cycle_colors(base_col, n_cat, step=0.10, reverse=False)
                          if n_cat > 1 else [base_col])
            for j, pair in enumerate(cat_pairs):
                if not pair.get("n2_short") or not pair.get("o2_short"):
                    continue
                _cat_cc = sentry.get("catalyst_corrections", {}).get(cat, {})
                _r_n2 = float(_cat_cc.get("r_sol_n2", 0) or 0)
                _r_o2 = float(_cat_cc.get("r_sol_o2", 0) or 0)
                _eref  = float(_cat_cc.get("e_ref", 0) or 0)
                _area  = float(_cat_cc.get("area", "") or 0)
                result = _process_pair(pair, _r_n2, _r_o2, _eref, _area)
                if result is None:
                    continue
                E_plot, Y_plot = result
                rpm_val = pair.get("rpm_val") or pair.get("rpm_id") or f"#{j+1}"
                prefix  = f"[{cat}] " if cat else ""
                label   = f"{prefix}{rpm_val} rpm"
                _cat_st = sentry.get("catalyst_styles", {}).get(cat, {})
                _col = _cat_st.get("color", "").strip() or cat_colors[j]
                try:
                    _lw = float(_cat_st.get("linewidth", "1.5") or 1.5)
                except (ValueError, TypeError):
                    _lw = 1.5
                _ls = {"solid": "-", "dashed": "--", "dotted": ":",
                       "dash-dot": "-."}.get(_cat_st.get("linestyle", "solid"), "-")
                _mk = _cat_st.get("marker", "none")
                _mk = None if _mk == "none" else _mk
                ax.plot(E_plot, Y_plot, color=_col, linewidth=_lw, linestyle=_ls,
                        marker=_mk, markersize=4, label=label, zorder=2)
                plot_data.append((E_plot, Y_plot, label, _col))

        sentry["_plot_data"] = plot_data

        # E½ markers
        show_hw = (self.show_half_wave_var.get() if is_active
                   else sentry.get("show_half_wave", True))
        if show_hw:
            for E_pd, Y_pd, lbl_pd, col_pd in plot_data:
                e_half, j_half = _find_half_wave(E_pd, Y_pd)
                if e_half is not None:
                    ax.axvline(e_half, color=col_pd, linestyle="--",
                               linewidth=0.8, alpha=0.55, label="_ehalf")
                    ax.plot([e_half], [j_half], marker="o", color=col_pd,
                            ms=4, zorder=5, linestyle="none", label="_ehalf_dot")
                    ax.annotate(
                        f"E½={e_half:.3f} V",
                        xy=(e_half, j_half),
                        xytext=(4, -14), textcoords="offset points",
                        fontsize=6, color=col_pd,
                        bbox=dict(boxstyle="round,pad=0.2", fc="white",
                                  alpha=0.75, ec=col_pd, lw=0.5),
                    ).set_in_layout(False)

        # Axis labels
        x_label = f"E (V vs {ref})"
        _any_area = any(
            float(sentry.get("catalyst_corrections", {}).get(c, {}).get("area", "") or 0) > 0
            for c in catalyst_order)
        y_label  = "J (mA cm⁻²)" if _any_area else "I (mA)"
        try:
            lbl_sz  = int(sentry.get("font_label_size", "10") if not is_active
                          else self.font_label_size_var.get())
            lbl_wt  = "bold" if (sentry.get("font_label_bold", False) if not is_active
                                 else self.font_label_bold_var.get()) else "normal"
            tick_sz = int(sentry.get("font_tick_size", "8") if not is_active
                          else self.font_tick_size_var.get())
            tick_wt = "bold" if (sentry.get("font_tick_bold", False) if not is_active
                                 else self.font_tick_bold_var.get()) else "normal"
            tit_sz  = int(sentry.get("font_title_size", "10") if not is_active
                          else self.font_title_size_var.get())
            tit_wt  = "bold" if (sentry.get("font_title_bold", False) if not is_active
                                 else self.font_title_bold_var.get()) else "normal"
            t_pad   = int(sentry.get("title_pad", "6") if not is_active
                          else self.title_pad_var.get())
            l_pad   = int(sentry.get("label_pad", "4") if not is_active
                          else self.label_pad_var.get())
        except (ValueError, TypeError):
            lbl_sz = 10; lbl_wt = "normal"; tick_sz = 8; tick_wt = "normal"
            tit_sz = 10; tit_wt = "normal"; t_pad = 6; l_pad = 4

        ax.set_xlabel(sentry.get("custom_xlabel") or x_label,
                      fontsize=lbl_sz, fontweight=lbl_wt, labelpad=l_pad)
        ax.set_ylabel(sentry.get("custom_ylabel") or y_label,
                      fontsize=lbl_sz, fontweight=lbl_wt, labelpad=l_pad)
        ax.tick_params(labelsize=tick_sz)
        for lbl in ax.get_xticklabels() + ax.get_yticklabels():
            lbl.set_fontweight(tick_wt)

        title = (sentry.get("custom_title", "") if not is_active
                 else self.plot_title_var.get())
        ax.set_title(title, fontsize=tit_sz, fontweight=tit_wt, pad=t_pad)

        # Legend
        _leg = None
        if leg_show and plot_data:
            _leg = ax.legend(fontsize=leg_size, frameon=leg_frm, loc=leg_loc)
            _mp  = sentry.get("legend_manual_pos")
            if isinstance(_mp, (tuple, list)):
                _set = getattr(_leg, "_set_loc", None)
                if callable(_set):
                    _set(tuple(_mp))
                else:
                    _leg._loc = tuple(_mp)
            sentry["legend_label_order"] = [t.get_text() for t in _leg.get_texts()]
            _leg.set_draggable(True)
        sentry["legend"] = _leg

        # Layout
        _lv = ax.get_legend()
        if _lv:
            _lv.set_visible(False)
        try:
            sentry["fig"].tight_layout(pad=0.5)
            sentry["fig"].set_layout_engine("none")
        except Exception:
            pass
        if _lv:
            _lv.set_visible(True)

        canvas.draw()

        sentry["auto_xlim"] = ax.get_xlim()
        sentry["auto_ylim"] = ax.get_ylim()

        if _prev_view:
            ax.set_xlim(_prev_view[0])
            ax.set_ylim(_prev_view[1])

        # Manual range
        self._apply_sample_range(sample_name, is_active)

        # Grid + tick intervals
        try:
            x_g  = sentry.get("x_grid", False) if not is_active else self.x_grid_var.get()
            y_g  = sentry.get("y_grid", False) if not is_active else self.y_grid_var.get()
            g_st = sentry.get("grid_style", "dashed") if not is_active else self.grid_style_var.get()
            g_co = sentry.get("grid_color",  "gray")  if not is_active else self.grid_color_var.get()
            g_lw = float(sentry.get("grid_linewidth", "0.8") if not is_active
                         else self.grid_linewidth_var.get())
            xi   = sentry.get("x_grid_int", "0") if not is_active else self.x_grid_int_var.get()
            yi   = sentry.get("y_grid_int", "0") if not is_active else self.y_grid_int_var.get()
            apply_grid(ax, x_g, y_g, xi, yi, g_st, linewidth=g_lw, color=g_co)
        except Exception:
            pass

        # Reference lines
        reflines = (sentry.get("reflines", []) if not is_active
                    else self._get_current_reflines())
        draw_reflines(ax, reflines)

        # Re-apply catalyst highlight (legend handles already drawn at alpha=1)
        self._apply_orr_highlight(sample_name)

        canvas.draw_idle()

    def _apply_sample_range(self, sample_name, is_active=None):
        sentry = self.samples.get(sample_name)
        if sentry is None or "ax" not in sentry:
            return
        if is_active is None:
            is_active = (sample_name == self.active_sample)
        ax = sentry["ax"]

        def _val(key, var_attr):
            if is_active:
                var = getattr(self, var_attr, None)
                return var.get() if var else ""
            return sentry.get(key, "")

        try:
            xmin = float(_val("x_min", "x_min_var")); ax.set_xlim(left=xmin)
        except (ValueError, TypeError): pass
        try:
            xmax = float(_val("x_max", "x_max_var")); ax.set_xlim(right=xmax)
        except (ValueError, TypeError): pass
        try:
            ymin = float(_val("y_min", "y_min_var")); ax.set_ylim(bottom=ymin)
        except (ValueError, TypeError): pass
        try:
            ymax = float(_val("y_max", "y_max_var")); ax.set_ylim(top=ymax)
        except (ValueError, TypeError): pass

        xl = ax.get_xlim()
        yl = ax.get_ylim()
        flip_x = sentry.get("x_flip", False) if not is_active else self.x_flip_var.get()
        flip_y = sentry.get("y_flip", False) if not is_active else self.y_flip_var.get()
        if flip_x and xl[0] < xl[1]:
            ax.set_xlim(xl[1], xl[0])
        elif not flip_x and xl[0] > xl[1]:
            ax.set_xlim(xl[1], xl[0])
        if flip_y and yl[0] < yl[1]:
            ax.set_ylim(yl[1], yl[0])
        elif not flip_y and yl[0] > yl[1]:
            ax.set_ylim(yl[1], yl[0])

    def _reset_sample_view(self, sample_name):
        sentry = self.samples.get(sample_name)
        if sentry and sentry.get("auto_xlim") is not None:
            sentry["ax"].set_xlim(sentry["auto_xlim"])
            sentry["ax"].set_ylim(sentry["auto_ylim"])
            sentry["canvas"].draw_idle()

    # ════════════════════════════════════════════════════════════════
    # Reference lines
    # ════════════════════════════════════════════════════════════════
    def _get_current_reflines(self):
        if not self.active_sample:
            return []
        return self.samples.get(self.active_sample, {}).get("reflines", [])

    def _add_xrefline(self):
        try:
            val = float(self._ref_x_var.get())
        except ValueError:
            return
        self._add_refline("x", val)

    def _add_yrefline(self):
        try:
            val = float(self._ref_y_var.get())
        except ValueError:
            return
        self._add_refline("y", val)

    def _add_refline(self, axis, val):
        if not self.active_sample or self.active_sample not in self.samples:
            return
        self._save_active_sample_state()
        style = self._refline_style_var.get()
        color = self._refline_color_var.get()
        try:
            lw = float(self._refline_lw_var.get())
        except ValueError:
            lw = 1.0
        self.samples[self.active_sample]["reflines"].append((axis, val, style, color, lw))
        self._refresh_reflines_lb()
        self._auto_replot()

    def _remove_refline(self):
        if not self.active_sample:
            return
        sel = self._reflines_lb.curselection()
        if not sel:
            return
        reflines = self.samples[self.active_sample].get("reflines", [])
        idx = sel[0]
        if idx < len(reflines):
            reflines.pop(idx)
        self._refresh_reflines_lb()
        self._auto_replot()

    def _on_refline_select(self):
        if not self.active_sample:
            return
        sel = self._reflines_lb.curselection()
        if not sel:
            return
        reflines = self.samples[self.active_sample].get("reflines", [])
        idx = sel[0]
        if idx < len(reflines):
            rl = reflines[idx]
            style = rl[2] if len(rl) > 2 else "dashed"
            color = rl[3] if len(rl) > 3 else "dimgray"
            lw    = rl[4] if len(rl) > 4 else 1.0
            self._refline_style_var.set(style)
            self._refline_color_var.set(color)
            self._refline_lw_var.set(str(lw))

    def _on_refline_style_change(self):
        if not self.active_sample:
            return
        sel = self._reflines_lb.curselection()
        if not sel:
            return
        reflines = self.samples[self.active_sample].get("reflines", [])
        idx = sel[0]
        if idx < len(reflines):
            rl = list(reflines[idx])
            rl[2] = self._refline_style_var.get()
            rl[3] = self._refline_color_var.get()
            try:
                rl_lw = float(self._refline_lw_var.get())
            except ValueError:
                rl_lw = 1.0
            if len(rl) > 4:
                rl[4] = rl_lw
            else:
                rl.append(rl_lw)
            reflines[idx] = tuple(rl)
        self._auto_replot()

    def _refresh_reflines_lb(self):
        self._reflines_lb.delete(0, tk.END)
        if not self.active_sample:
            return
        for rl in self.samples.get(self.active_sample, {}).get("reflines", []):
            axis = rl[0]; val = rl[1]
            self._reflines_lb.insert(tk.END, f"{axis.upper()}={val:.4g}")

    # ════════════════════════════════════════════════════════════════
    # Figure creation / destruction / layout
    # ════════════════════════════════════════════════════════════════
    def _create_sample_figure(self, sample_name):
        sentry = self.samples.get(sample_name)
        if sentry is None or "fig" in sentry:
            return
        panel_ref = self
        self._placeholder.grid_remove()

        frame  = tk.Frame(self._plots_frame, relief="groove", bd=2)
        header = tk.Frame(frame, bg=_SAMPLE_HDR_BG, cursor="fleur")
        header.pack(fill=tk.X, side=tk.TOP)
        lbl = tk.Label(header, text=f"⠿  {sample_name}",
                       bg=_SAMPLE_HDR_BG, font=("", 9, "bold"), anchor=tk.W)
        lbl.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=3)

        inner = ttk.Frame(frame, padding=(4, 2, 4, 2))
        inner.pack()

        try:
            _fw = float(self.plot_w_var.get())
            _fh = float(self.plot_h_var.get())
        except (ValueError, AttributeError):
            _fw, _fh = 9.5, 5.5

        fig = Figure(figsize=(_fw, _fh), dpi=100)
        ax  = fig.add_subplot(111)
        canvas = FigureCanvasTkAgg(fig, master=inner)
        canvas.get_tk_widget().pack()

        tb_frame = ttk.Frame(inner)
        tb_frame.pack(fill=tk.X)

        class _Toolbar(NavigationToolbar2Tk):
            def home(tb_self, *a):
                panel_ref._reset_sample_view(sample_name)

        _tb = _Toolbar(canvas, tb_frame, pack_toolbar=False)
        _tb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        _tb.update()
        tk.Button(tb_frame, text="Copy",
                  command=lambda f=fig: copy_figure_to_clipboard(f),
                  relief=tk.RAISED, borderwidth=1, padx=6).pack(
            side=tk.LEFT, padx=(4, 2), pady=1)

        def _fwd(e):
            self._right_canvas.yview_scroll(-1 * (e.delta // 120), "units")
        frame.bind("<MouseWheel>",    _fwd)
        header.bind("<MouseWheel>",   _fwd)
        tb_frame.bind("<MouseWheel>", _fwd)

        for _w in (header, lbl):
            _w.bind("<ButtonPress-1>",   lambda e, s=sample_name: self._on_frame_press(e, s))
            _w.bind("<B1-Motion>",       lambda e, s=sample_name: self._on_frame_drag(e, s))
            _w.bind("<ButtonRelease-1>", lambda e, s=sample_name: self._on_frame_release(e, s))
            _w.bind("<Double-Button-1>", lambda e, s=sample_name: self._toggle_zoom(s))

        def _activate(e=None):
            self._activate_sample(sample_name)
        frame.bind("<Button-1>",    _activate, add="+")
        tb_frame.bind("<Button-1>", _activate, add="+")

        canvas.mpl_connect("scroll_event",         lambda ev: self._on_scroll(ev, sample_name))
        canvas.mpl_connect("button_press_event",   lambda ev: self._on_press(ev, sample_name))
        canvas.mpl_connect("button_release_event", lambda ev: self._on_release(ev, sample_name))
        canvas.mpl_connect("motion_notify_event",  lambda ev: self._on_motion(ev, sample_name))
        canvas.mpl_connect("button_press_event",
                           lambda ev, sn=sample_name: self._on_legend_dblclick(ev, sn))

        sentry.update({
            "fig": fig, "ax": ax, "canvas": canvas, "toolbar": _tb,
            "plot_frame": frame, "hdr_frame": header, "hdr_label": lbl,
            "legend": None, "leg_size": 8.0,
            "auto_xlim": None, "auto_ylim": None,
            "panning": False, "pan_ax": None, "pan_start": None, "pan_moved": False,
            "leg_resize": False, "leg_resize_start_y": None, "leg_resize_start_sz": None,
            "ann": None, "ann_dot": None, "ann_last": None, "ann_idx": 0,
            "_plot_data": [],
        })
        ax.set_title("", fontsize=9)
        ax.set_xlabel("E (V vs RHE)")
        ax.set_ylabel("I (mA)")
        canvas.draw()
        self._relayout_figures()

    def _destroy_sample_figure(self, sample_name):
        frame = self.samples.get(sample_name, {}).get("plot_frame")
        if frame is not None:
            frame.destroy()

    def _activate_sample(self, sample_name):
        items = list(self.samples.keys())
        if sample_name in items:
            idx = items.index(sample_name)
            self.sample_lb.selection_clear(0, tk.END)
            self.sample_lb.selection_set(idx)
            self.sample_lb.see(idx)
            self._scroll_left_to_widget(self.sample_lb)
        if sample_name != self.active_sample:
            self._save_active_sample_state()
            self._switch_active_sample(sample_name)

    def _scroll_left_to_widget(self, widget):
        """Scroll the left panel so that widget is visible."""
        try:
            self.update_idletasks()
            y = 0
            w = widget
            while w is not None and w is not self._left_inner:
                y += w.winfo_y()
                w = w.master
            inner_h = self._left_inner.winfo_reqheight()
            if inner_h > 0:
                self._left_lc.yview_moveto(max(0.0, (y - 10) / inner_h))
        except Exception:
            pass

    def _highlight_active_headers(self):
        for sn, sentry in self.samples.items():
            hdr = sentry.get("hdr_frame")
            lbl = sentry.get("hdr_label")
            if hdr is None:
                continue
            color = (_SAMPLE_HDR_ACTIVE if sn == self.active_sample
                     else _SAMPLE_HDR_BG)
            hdr.configure(bg=color)
            if lbl:
                lbl.configure(bg=color)

    # ════════════════════════════════════════════════════════════════
    # Right panel layout
    # ════════════════════════════════════════════════════════════════
    def _on_right_canvas_configure(self, event):
        if self._zoom_sample:
            self._right_canvas.itemconfig(
                self._plots_win, width=event.width, height=event.height)

    def _on_grid_cols_change(self):
        try:
            c = max(1, int(self._grid_cols_var.get()))
            self._grid_cols_var.set(str(c))
        except (ValueError, AttributeError):
            self._grid_cols_var.set("2")
        self._relayout_figures()

    def _relayout_figures(self):
        try:
            MAX_COLS = max(1, int(self._grid_cols_var.get()))
        except (ValueError, AttributeError):
            MAX_COLS = 2
        valid = [(sn, self.samples[sn]) for sn in self.samples
                 if "plot_frame" in self.samples[sn]
                 and not self.samples[sn].get("hidden", False)]

        if self._zoom_sample and any(sn == self._zoom_sample for sn, _ in valid):
            for sn, sentry in valid:
                if sn == self._zoom_sample:
                    sentry["plot_frame"].grid(row=0, column=0,
                                             columnspan=MAX_COLS,
                                             sticky="nsew", padx=4, pady=4)
                else:
                    sentry["plot_frame"].grid_remove()
            self._plots_frame.rowconfigure(0, weight=1)
            return

        for sn in self.samples:
            pf = self.samples[sn].get("plot_frame")
            if pf is not None:
                pf.grid_remove()
        for i, (sname, sentry) in enumerate(valid):
            row = i // MAX_COLS
            col = i % MAX_COLS
            sentry["plot_frame"].grid(row=row, column=col, columnspan=1,
                                     sticky="nsew", padx=4, pady=4)
        n_rows = (len(valid) + MAX_COLS - 1) // MAX_COLS if valid else 0
        for r in range(n_rows):
            self._plots_frame.rowconfigure(r, weight=0)

    def _apply_plot_size(self, event=None):
        try:
            w = float(self.plot_w_var.get())
            h = float(self.plot_h_var.get())
        except ValueError:
            return
        w = max(2.0, min(50.0, w))
        h = max(1.5, min(50.0, h))
        dpi = 100
        self._right_canvas.itemconfig(self._plots_win, width=0, height=0)
        for sentry in self.samples.values():
            fig = sentry.get("fig")
            cv  = sentry.get("canvas")
            if fig and cv:
                fig.set_size_inches(w, h)
                cv.get_tk_widget().config(width=int(w * dpi), height=int(h * dpi))
                _lgs = [a.get_legend() for a in fig.get_axes() if a.get_legend()]
                for _l in _lgs:
                    _l.set_visible(False)
                fig.tight_layout(pad=0.5)
                fig.set_layout_engine("none")
                for _l in _lgs:
                    _l.set_visible(True)
                cv.draw_idle()

        def _upd():
            self._plots_frame.update_idletasks()
            self._right_canvas.configure(scrollregion=self._right_canvas.bbox("all"))
        self._right_canvas.after(100, _upd)

    def _auto_set_initial_size(self):
        w = self._right_canvas.winfo_width()
        h = self._right_canvas.winfo_height()
        if w <= 1 or h <= 1:
            self.after(100, self._auto_set_initial_size)
            return
        try:
            _ncols = max(1, int(self._grid_cols_var.get()))
        except (ValueError, AttributeError):
            _ncols = 2
        plot_w = max(3.0, (w / _ncols - 30) / 100)
        plot_h = max(2.0, round(plot_w * 0.6, 1))
        self.plot_w_var.set(f"{plot_w:.1f}")
        self.plot_h_var.set(f"{plot_h:.1f}")
        self._apply_plot_size()

    # ════════════════════════════════════════════════════════════════
    # Zoom (single-sample full-panel view)
    # ════════════════════════════════════════════════════════════════
    def _toggle_zoom(self, sample_name):
        if self._zoom_sample is None:
            self._zoom_sample_view(sample_name)
        else:
            self._unzoom_sample_view()

    def _zoom_sample_view(self, sample_name):
        self._zoom_sample = sample_name
        self._zoom_bar.grid()
        self.update_idletasks()
        w = self._right_canvas.winfo_width()
        h = self._right_canvas.winfo_height()
        if w > 1 and h > 1:
            self._right_canvas.itemconfig(self._plots_win, width=w, height=h)
            sentry = self.samples.get(sample_name, {})
            fig = sentry.get("fig")
            cv  = sentry.get("canvas")
            if fig and cv:
                fig.set_size_inches(w / 100, h / 100)
                cv.get_tk_widget().config(width=w, height=h)
                cv.draw_idle()
        self._relayout_figures()
        self._right_canvas.yview_moveto(0)

    def _unzoom_sample_view(self):
        self._zoom_sample = None
        self._zoom_bar.grid_remove()
        self._right_canvas.itemconfig(self._plots_win, width=0, height=0)
        self._apply_plot_size()
        self._relayout_figures()
        self._plots_frame.update_idletasks()
        self._right_canvas.configure(scrollregion=self._right_canvas.bbox("all"))

    # ════════════════════════════════════════════════════════════════
    # Drag-to-reorder subplots
    # ════════════════════════════════════════════════════════════════
    def _on_frame_press(self, event, sample_name):
        self._drag = {
            "sample": sample_name,
            "start_x": event.x_root, "start_y": event.y_root,
            "active": False, "target": None, "target_top": True,
        }

    def _on_frame_drag(self, event, sample_name):
        drag = self._drag
        if drag is None or drag["sample"] != sample_name:
            return
        if not drag["active"]:
            if abs(event.x_root - drag["start_x"]) + abs(event.y_root - drag["start_y"]) < 6:
                return
            drag["active"] = True
        target = None; target_top = True
        for sn, sentry in self.samples.items():
            if sn == sample_name or sentry.get("hidden"):
                continue
            pf = sentry.get("plot_frame")
            if pf is None:
                continue
            x0, y0 = pf.winfo_rootx(), pf.winfo_rooty()
            w_,  h_ = pf.winfo_width(), pf.winfo_height()
            if x0 <= event.x_root <= x0 + w_ and y0 <= event.y_root <= y0 + h_:
                target = sn
                target_top = (event.y_root - y0) < h_ / 2
                break
        drag["target"] = target; drag["target_top"] = target_top
        if target is not None:
            pf = self.samples[target]["plot_frame"]
            rx, ry = pf.winfo_x(), pf.winfo_y()
            rw, rh = pf.winfo_width(), pf.winfo_height()
            line_y = ry if target_top else ry + rh - 3
            self._drop_line.place(x=rx, y=line_y, width=rw, height=3)
            self._drop_line.lift()
        else:
            self._drop_line.place_forget()

    def _on_frame_release(self, event, sample_name):
        drag = self._drag; self._drag = None
        self._drop_line.place_forget()
        if drag is None or not drag["active"]:
            return
        target = drag.get("target")
        if target is None or target == sample_name:
            return
        keys = list(self.samples.keys())
        if sample_name not in keys or target not in keys:
            return
        keys.remove(sample_name)
        to_idx = keys.index(target)
        keys.insert(to_idx if drag.get("target_top", True) else to_idx + 1, sample_name)
        self.samples = OrderedDict((k, self.samples[k]) for k in keys)
        self._rebuild_sample_listbox()
        self._relayout_figures()

    # ════════════════════════════════════════════════════════════════
    # Mouse interactions (scroll / pan / annotate)
    # ════════════════════════════════════════════════════════════════
    # Listbox drag-to-resize handlers
    # ════════════════════════════════════════════════════════════════
    def _on_loaded_resize_start(self, event):
        self._loaded_drag = {
            "y0": event.y_root,
            "h0": int(self.loaded_tv.cget("height")),
        }

    def _on_loaded_resize_drag(self, event):
        d = getattr(self, "_loaded_drag", None)
        if d is None:
            return
        dy = event.y_root - d["y0"]
        new_rows = max(2, int(d["h0"] + dy / 20))
        self.loaded_tv.config(height=new_rows)

    def _on_sample_resize_start(self, event):
        h = self.sample_lb._canvas.winfo_height()
        if h <= 1:
            h = self.sample_lb._canvas.cget("height")
        self._sample_drag = {"y0": event.y_root, "h0": int(h)}

    def _on_sample_resize_drag(self, event):
        d = getattr(self, "_sample_drag", None)
        if d is None:
            return
        dy = event.y_root - d["y0"]
        new_h = max(40, d["h0"] + dy)
        self.sample_lb._canvas.config(height=int(new_h))

    def _on_pair_resize_start(self, event):
        h = self._pair_tbl_canvas.winfo_height()
        if h <= 1:
            h = int(self._pair_tbl_canvas.cget("height"))
        self._pair_drag = {"y0": event.y_root, "h0": h}

    def _on_pair_resize_drag(self, event):
        d = getattr(self, "_pair_drag", None)
        if d is None:
            return
        dy = event.y_root - d["y0"]
        new_h = max(40, d["h0"] + dy)
        self._pair_tbl_canvas.config(height=int(new_h))

    # ════════════════════════════════════════════════════════════════
    def _on_scroll(self, event, sample_name):
        sentry = self.samples.get(sample_name)
        if sentry is None or "ax" not in sentry:
            return
        ax = sentry["ax"]
        canvas = sentry["canvas"]
        if event.xdata is None or event.ydata is None:
            return
        factor = (1 / 1.15) if event.button == "up" else 1.15
        cx, cy = event.xdata, event.ydata
        xl = ax.get_xlim(); yl = ax.get_ylim()
        ax.set_xlim(cx + (xl[0] - cx) * factor, cx + (xl[1] - cx) * factor)
        ax.set_ylim(cy + (yl[0] - cy) * factor, cy + (yl[1] - cy) * factor)
        canvas.draw_idle()

    def _on_press(self, event, sample_name):
        sentry = self.samples.get(sample_name)
        if sentry is None:
            return
        sentry["pan_moved"] = False

        # Double-click on axis labels → inline edit
        if event.button == 1 and getattr(event, "dblclick", False):
            ax = sentry.get("ax")
            if ax is not None:
                try:
                    renderer = event.canvas.get_renderer()
                    xl = ax.xaxis.label
                    if xl.get_window_extent(renderer).contains(event.x, event.y):
                        self._edit_axis_label(sample_name, "x")
                        return
                    yl = ax.yaxis.label
                    if yl.get_window_extent(renderer).contains(event.x, event.y):
                        self._edit_axis_label(sample_name, "y")
                        return
                except Exception:
                    pass

        if event.button == 3:
            # Right-drag on legend → resize; right-click elsewhere → dismiss annotation
            leg = sentry.get("legend")
            if leg is not None:
                try:
                    hit, _ = leg.contains(event)
                    if hit:
                        sentry["leg_resize"] = True
                        sentry["leg_resize_start_y"]  = event.y
                        sentry["leg_resize_start_sz"] = sentry.get(
                            "leg_size_live", sentry.get("leg_size", 8.0))
                        return
                except Exception:
                    pass
            sentry["highlighted_cat"] = None
            self._apply_orr_highlight(sample_name)
            self._clear_ann(sample_name, redraw=True)
            return
        if event.button != 1 or event.xdata is None:
            return
        # Don't start panning if the click landed on the legend — let DraggableLegend handle it.
        leg = sentry.get("legend")
        if leg is not None:
            try:
                hit, _ = leg.contains(event)
                if hit:
                    return
            except Exception:
                pass
        sentry["panning"]   = True
        sentry["pan_start"] = (event.xdata, event.ydata,
                               *sentry["ax"].get_xlim(), *sentry["ax"].get_ylim())

    def _edit_axis_label(self, sample_name, which):
        """Double-click X or Y label → askstring dialog; blank reverts to auto."""
        from tkinter.simpledialog import askstring
        sentry = self.samples.get(sample_name)
        if sentry is None:
            return
        ax = sentry.get("ax")
        if ax is None:
            return
        if which == "x":
            current = ax.get_xlabel()
            new = askstring("Edit X Label", "X axis label\n(blank = auto):",
                            initialvalue=current, parent=self)
            if new is not None:
                sentry["custom_xlabel"] = new.strip() or None
                ax.set_xlabel(new.strip() if new.strip() else current)
                sentry["canvas"].draw_idle()
        else:
            current = ax.get_ylabel()
            new = askstring("Edit Y Label", "Y axis label\n(blank = auto):",
                            initialvalue=current, parent=self)
            if new is not None:
                sentry["custom_ylabel"] = new.strip() or None
                ax.set_ylabel(new.strip() if new.strip() else current)
                sentry["canvas"].draw_idle()

    def _on_motion(self, event, sample_name):
        sentry = self.samples.get(sample_name)
        if sentry is None:
            return

        # Legend resize (right-drag on legend)
        if sentry.get("leg_resize"):
            leg = sentry.get("legend")
            if leg is not None and event.y is not None:
                dy = event.y - sentry["leg_resize_start_y"]
                new_sz = sentry["leg_resize_start_sz"] + dy / 5.0
                new_sz = max(4.0, min(30.0, new_sz))
                prev_sz = sentry.get("leg_size_live", sentry.get("leg_size", 8.0))
                sentry["leg_size_live"] = new_sz
                if prev_sz > 0:
                    _scale_legend_spacing(leg, new_sz / prev_sz)
                for txt in leg.get_texts():
                    txt.set_fontsize(new_sz)
                ttl = leg.get_title()
                if ttl:
                    ttl.set_fontsize(new_sz)
                sentry["canvas"].draw()
            return

        if not sentry.get("panning"):
            return
        if event.xdata is None or sentry.get("pan_start") is None:
            return
        sentry["pan_moved"] = True
        x0, y0, x1, x2, y1, y2 = sentry["pan_start"]
        dx = event.xdata - x0; dy = event.ydata - y0
        ax = sentry["ax"]
        ax.set_xlim(x1 - dx, x2 - dx)
        ax.set_ylim(y1 - dy, y2 - dy)
        sentry["canvas"].draw_idle()

    def _on_release(self, event, sample_name):
        sentry = self.samples.get(sample_name)
        if sentry is None:
            return

        # Legend resize release → commit final size
        if sentry.get("leg_resize"):
            sentry["leg_resize"] = False
            live = sentry.get("leg_size_live")
            if live is not None:
                sentry["leg_size"] = live
                sentry["leg_size_live"] = None
            return

        sentry["panning"] = False
        if event.button == 1 and not sentry.get("pan_moved"):
            self._try_annotate(event, sample_name)
        sentry["pan_moved"] = False
        # Deferred legend-position save: fires after matplotlib's DraggableLegend
        # finalize_offset() has completed, so _loc / _loc_real reflects new position.
        if event.button == 1:
            self.after(10, lambda s=sample_name: self._save_leg_pos(s))

    def _save_leg_pos(self, sample_name):
        sentry = self.samples.get(sample_name)
        if sentry is None:
            return
        leg = sentry.get("legend")
        if leg is None:
            return
        for attr in ("_loc_real", "_loc"):
            _loc = getattr(leg, attr, None)
            if isinstance(_loc, (tuple, list)) and len(_loc) == 2:
                sentry["legend_manual_pos"] = tuple(_loc)
                return

    def _try_annotate(self, event, sample_name):
        sentry = self.samples.get(sample_name)
        if sentry is None or event.xdata is None:
            return
        plot_data = sentry.get("_plot_data", [])
        if not plot_data:
            return
        ax = sentry["ax"]; canvas = sentry["canvas"]
        # Find nearest point across all curves
        best_dist = float("inf"); best_x = best_y = None; best_label = ""
        xl = ax.get_xlim(); yl = ax.get_ylim()
        xspan = abs(xl[1] - xl[0]) or 1; yspan = abs(yl[1] - yl[0]) or 1
        for E_arr, Y_arr, label, color in plot_data:
            if len(E_arr) == 0:
                continue
            dx = (E_arr - event.xdata) / xspan
            dy = (Y_arr - event.ydata) / yspan
            dist = dx**2 + dy**2
            idx  = int(np.argmin(dist))
            if dist[idx] < best_dist:
                best_dist = dist[idx]; best_x = E_arr[idx]; best_y = Y_arr[idx]
                best_label = label
        if best_x is None or best_dist > 0.04:
            return
        # Switch catalyst correction display + highlight all lines for that catalyst
        _cat_m = re.match(r'^\[(\w+)\]', best_label)
        _clicked_cat = _cat_m.group(1) if _cat_m else ""
        sentry["highlighted_cat"] = _clicked_cat
        self._apply_orr_highlight(sample_name)
        if _cat_m:
            _avail = list(self._corr_cat_cb["values"])
            if _clicked_cat in _avail and _clicked_cat != getattr(self, "_active_catalyst", None):
                self._corr_catalyst_var.set(_clicked_cat)
                self._load_catalyst_corrections(sample_name, _clicked_cat)
        # Increment annotation index for overlap cycling
        ann_key = (round(best_x, 6), round(best_y, 6))
        if sentry.get("ann_last") == ann_key:
            sentry["ann_idx"] = (sentry.get("ann_idx", 0) + 1) % 4
        else:
            sentry["ann_idx"] = 0; sentry["ann_last"] = ann_key
        self._clear_ann(sample_name, redraw=False)
        offsets = [(8, 8), (-8, 8), (-8, -8), (8, -8)]
        xytext  = offsets[sentry["ann_idx"] % 4]
        ann = ax.annotate(
            f"{best_label}\n({best_x:.4f}, {best_y:.4f})",
            xy=(best_x, best_y), xytext=xytext,
            textcoords="offset points", fontsize=8,
            bbox=dict(boxstyle="round,pad=0.3", fc="lightyellow", alpha=0.85),
            arrowprops=dict(arrowstyle="->", lw=0.8),
        )
        ann.set_in_layout(False)
        dot, = ax.plot([best_x], [best_y], "ko", ms=5,
                       label=_ANN_DOT_LABEL, zorder=10)
        sentry["ann"] = ann; sentry["ann_dot"] = dot
        canvas.draw_idle()

    def _apply_orr_highlight(self, sample_name):
        """Dim all catalyst lines except the highlighted one; bring highlighted to front."""
        sentry = self.samples.get(sample_name)
        if sentry is None or "ax" not in sentry:
            return
        ax = sentry["ax"]
        hl_cat = sentry.get("highlighted_cat")  # None = no highlight; str = catalyst name
        for line in ax.get_lines():
            lbl = line.get_label()
            if lbl.startswith("_"):
                continue  # skip ehalf, annotation dot, glow lines
            if hl_cat is None:
                line.set_alpha(1.0)
                line.set_zorder(2)
            else:
                m = re.match(r'^\[(\w+)\]', lbl)
                line_cat = m.group(1) if m else ""
                if line_cat == hl_cat:
                    line.set_alpha(1.0)
                    line.set_zorder(5)
                else:
                    line.set_alpha(0.15)
                    line.set_zorder(2)

    def _clear_ann(self, sample_name, *, redraw=True):
        sentry = self.samples.get(sample_name)
        if sentry is None:
            return
        for key in ("ann", "ann_dot"):
            obj = sentry.get(key)
            if obj is not None:
                try:
                    obj.remove()
                except Exception:
                    pass
            sentry[key] = None
        if redraw:
            cv = sentry.get("canvas")
            if cv:
                cv.draw_idle()

    def _on_legend_dblclick(self, event, sample_name):
        if event.dblclick and event.button == 1:
            sentry = self.samples.get(sample_name)
            if sentry is None:
                return
            leg = sentry.get("legend")
            if leg is None:
                return
            try:
                hit, _ = leg.contains(event)
            except Exception:
                hit = False
            if not hit:
                return
            leg.set_draggable(False)
            sentry["legend"], perm = open_legend_editor(
                self, leg, sentry["canvas"], sentry.get("leg_size", 8.0))
            if sentry.get("legend") is not None:
                sentry["legend"].set_draggable(True)
            # Store permutation so next replot respects the new order
            orig_labels = sentry.get("legend_label_order", [])
            if perm and orig_labels:
                sentry["legend_label_order"] = [orig_labels[j] for j in perm
                                                if j < len(orig_labels)]
            sentry["canvas"].draw()

    # ════════════════════════════════════════════════════════════════
    # Analysis windows
    # ════════════════════════════════════════════════════════════════
    def _get_active_curves(self):
        """Return list of (E_arr, J_arr, rpm_float, label, color) for the active sample."""
        if not self.active_sample or self.active_sample not in self.samples:
            return []
        sentry = self.samples[self.active_sample]
        cat_corrections = sentry.get("catalyst_corrections", {})
        curves = []
        for i, pair in enumerate(sentry.get("pairs", [])):
            if not pair.get("n2_short") or not pair.get("o2_short"):
                continue
            if not pair.get("enabled", True):
                continue
            cat = pair.get("catalyst_id", "")
            _cc = cat_corrections.get(cat, {})
            try:
                r_n2  = float(_cc.get("r_sol_n2", 0) or 0)
                r_o2  = float(_cc.get("r_sol_o2", 0) or 0)
                e_ref = float(_cc.get("e_ref", 0) or 0)
                area  = float(_cc.get("area", "") or 0)
            except (ValueError, TypeError):
                r_n2 = r_o2 = e_ref = area = 0.0
            result = _process_pair(pair, r_n2, r_o2, e_ref, area)
            if result is None:
                continue
            E_arr, J_arr = result
            try:
                rpm = float(pair.get("rpm_val") or pair.get("rpm_id") or 0)
            except (ValueError, TypeError):
                rpm = 0.0
            cat = pair.get("catalyst_id", "")
            prefix = f"[{cat}] " if cat else ""
            rpm_v = pair.get('rpm_val') or pair.get('rpm_id') or f'#{i+1}'
            label = f"{prefix}{rpm_v} rpm"
            _cat_st = sentry.get("catalyst_styles", {}).get(cat, {})
            _auto_col = _PALETTE[i % len(_PALETTE)]
            color = _cat_st.get("color", "").strip() or _auto_col
            curves.append((E_arr, J_arr, rpm, label, color))
        return curves

    def _open_tafel_window(self):
        curves = self._get_active_curves()
        if not curves:
            messagebox.showwarning("Tafel", "No processed curves for active sample.")
            return
        ref = self.ref_electrode_var.get()
        sname = self.active_sample

        win = tk.Toplevel(self)
        win.title(f"Tafel Analysis — {sname}")
        win.geometry("760x660")

        # ── Theory ──────────────────────────────────────────────────────
        _th = ttk.LabelFrame(win, text="Tafel Theory")
        _th.pack(fill=tk.X, padx=8, pady=(6, 0))
        _th_txt = (
            "  Tafel equation:   E = a + b · log₁₀|J|   (b = Tafel slope, mV/dec)\n"
            "  Kinetic region: low overpotential (far from diffusion plateau)\n"
            "  Diffusion correction:  Jᵏ = J · J_lim / (J_lim − J)   (Koutecky)\n"
            "  Fit:  linear regression of E vs log₁₀|J| → slope = b [V/dec] × 1000 [mV/dec]"
        )
        ttk.Label(_th, text=_th_txt, justify=tk.LEFT,
                  font=("Courier", 8)).pack(anchor=tk.W, padx=6, pady=3)

        # Curve selector
        _csel_fr = ttk.LabelFrame(win, text="Select curves to analyse")
        _csel_fr.pack(fill=tk.X, padx=8, pady=(4, 0))
        _csel_vars = []
        _csel_inner = ttk.Frame(_csel_fr)
        _csel_inner.pack(fill=tk.X, padx=4, pady=2)
        for _cv_idx, (_cv_E, _cv_J, _cv_rpm, _cv_lbl, _cv_col) in enumerate(curves):
            _bv = tk.BooleanVar(value=True)
            _csel_vars.append(_bv)
            ttk.Checkbutton(_csel_inner, text=_cv_lbl, variable=_bv).pack(
                side=tk.LEFT, padx=4)

        # Controls
        ctrl = ttk.Frame(win)
        ctrl.pack(side=tk.TOP, fill=tk.X, padx=8, pady=4)
        ttk.Label(ctrl, text="Kinetic E range (V):").pack(side=tk.LEFT)
        e_lo_var = tk.StringVar(value="0.85")
        e_hi_var = tk.StringVar(value="0.95")
        ttk.Entry(ctrl, textvariable=e_lo_var, width=6).pack(side=tk.LEFT, padx=(2, 0))
        ttk.Label(ctrl, text="to").pack(side=tk.LEFT, padx=3)
        ttk.Entry(ctrl, textvariable=e_hi_var, width=6).pack(side=tk.LEFT, padx=(0, 10))
        use_jk_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(ctrl, text="Diffusion-correct J → Jᵏ",
                        variable=use_jk_var).pack(side=tk.LEFT, padx=(0, 10))
        compute_btn = ttk.Button(ctrl, text="Compute", command=lambda: _compute())
        compute_btn.pack(side=tk.LEFT)

        # Figure
        fig = Figure(figsize=(7.0, 4.2), dpi=100)
        ax  = fig.add_subplot(111)
        cv  = FigureCanvasTkAgg(fig, master=win)
        cv.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=8)
        tb  = NavigationToolbar2Tk(cv, win, pack_toolbar=False)
        tb.pack(side=tk.BOTTOM, fill=tk.X, padx=8)

        # Results
        res = tk.Text(win, height=5, state=tk.DISABLED, font=("Courier", 8),
                      wrap=tk.WORD)
        res.pack(fill=tk.X, padx=8, pady=(0, 4))

        def _compute():
            try:
                e_lo = float(e_lo_var.get())
                e_hi = float(e_hi_var.get())
            except ValueError:
                messagebox.showerror("Tafel", "Invalid E range.", parent=win)
                return
            if e_lo >= e_hi:
                messagebox.showerror("Tafel", "E_lo must be < E_hi.", parent=win)
                return
            ax.clear()
            lines = []
            selected = [(E, J, r, l, c) for (E, J, r, l, c), bv
                        in zip(curves, _csel_vars) if bv.get()]
            for E_arr, J_arr, rpm, label, color in selected:
                j_lim = float(np.min(J_arr))
                mask = (E_arr >= e_lo) & (E_arr <= e_hi)
                if mask.sum() < 3:
                    lines.append(f"{label}: < 3 points in [{e_lo},{e_hi}] V — skipped")
                    continue
                E_k = E_arr[mask]
                J_k = J_arr[mask]
                if use_jk_var.get() and j_lim < 0:
                    with np.errstate(divide="ignore", invalid="ignore"):
                        jk = J_k * j_lim / (j_lim - J_k)
                    good = np.isfinite(jk) & (jk < 0)
                    if good.sum() < 3:
                        lines.append(f"{label}: Jᵏ correction yielded < 3 valid pts — skipped")
                        continue
                    E_k, J_k = E_k[good], jk[good]
                with np.errstate(divide="ignore", invalid="ignore"):
                    log_j = np.log10(np.abs(J_k))
                good = np.isfinite(log_j)
                if good.sum() < 3:
                    lines.append(f"{label}: log|J| not finite — skipped")
                    continue
                E_f, log_j_f = E_k[good], log_j[good]
                # E vs log|J|: slope in V/dec → convert to mV/dec
                coeffs = np.polyfit(log_j_f, E_f, 1)
                b_mV = coeffs[0] * 1000.0
                ax.plot(log_j_f, E_f, color=color, linewidth=1.5, label=label)
                xfit = np.linspace(log_j_f.min(), log_j_f.max(), 60)
                ax.plot(xfit, np.polyval(coeffs, xfit), color=color,
                        linestyle="--", linewidth=0.8, label="_fit")
                lines.append(f"{label:20s}  b = {b_mV:+.1f} mV/dec")
            j_lbl = "Jᵏ" if use_jk_var.get() else "J"
            ax.set_xlabel(f"log₁₀|{j_lbl}|  (mA cm⁻² or mA)")
            ax.set_ylabel(f"E  (V vs {ref})")
            ax.set_title(f"Tafel Analysis — {sname}")
            ax.legend(fontsize=8, frameon=True)
            fig.tight_layout(pad=0.5)
            fig.set_layout_engine("none")
            cv.draw()
            res.configure(state=tk.NORMAL)
            res.delete("1.0", tk.END)
            res.insert(tk.END, "\n".join(lines))
            res.configure(state=tk.DISABLED)

        _compute()

    def _open_kl_window(self):
        curves = self._get_active_curves()
        valid = [(E, J, rpm, lbl, col) for E, J, rpm, lbl, col in curves if rpm > 0]
        if len(valid) < 2:
            messagebox.showwarning(
                "KL Analysis",
                "Need at least 2 RPM pairs with numeric RPM values.")
            return
        ref   = self.ref_electrode_var.get()
        sname = self.active_sample

        win = tk.Toplevel(self)
        win.title(f"Koutecky-Levich Analysis — {sname}")
        win.geometry("800x720")

        # Curve selector
        _ksel_fr = ttk.LabelFrame(win, text="Select curves to analyse")
        _ksel_fr.pack(fill=tk.X, padx=8, pady=(6, 0))
        _ksel_vars = {}
        _ksel_inner = ttk.Frame(_ksel_fr)
        _ksel_inner.pack(fill=tk.X, padx=4, pady=2)
        for (_kE, _kJ, _krpm, _klbl, _kcol) in valid:
            _bv = tk.BooleanVar(value=True)
            _ksel_vars[_krpm] = _bv
            ttk.Checkbutton(_ksel_inner, text=_klbl, variable=_bv).pack(
                side=tk.LEFT, padx=4)

        # ── Theory ──────────────────────────────────────────────────────
        _th = ttk.LabelFrame(win, text="Koutecky-Levich Theory")
        _th.pack(fill=tk.X, padx=8, pady=(6, 0))
        _th_txt = (
            "  K-L equation:   1/J = 1/Jᵏ + 1/J_L = 1/Jᵏ + 1/(B · ω^½)\n"
            "  Levich constant:  B = 0.62 · n · F · D^(2/3) · ν^(−1/6) · C\n"
            "  Plot 1/J vs 1/ω^½ at each chosen E; linear fit → slope = 1/(n·B_unit)\n"
            "  n = 1 / (|slope| · B_unit)  where  B_unit = 0.62·F·D^(2/3)·ν^(−1/6)·C·1000\n"
            "  ω (rad/s) = 2π·RPM/60    F = 96485 C/mol"
        )
        ttk.Label(_th, text=_th_txt, justify=tk.LEFT,
                  font=("Courier", 8)).pack(anchor=tk.W, padx=6, pady=3)

        # Electrochemical parameters
        prm = ttk.LabelFrame(win, text="Electrolyte parameters  (O₂ in 0.1 M KOH, 25 °C)")
        prm.pack(fill=tk.X, padx=8, pady=(6, 2))
        _pr = ttk.Frame(prm)
        _pr.pack(fill=tk.X, padx=6, pady=3)
        d_var  = tk.StringVar(value="1.9e-5")
        nu_var = tk.StringVar(value="0.01")
        c_var  = tk.StringVar(value="1.2e-6")
        for lbl_txt, var, unit in (
            ("Dₒ₂ (cm²/s):", d_var,  None),
            ("ν (cm²/s):",        nu_var, None),
            ("Cₒ₂ (mol/cm³):", c_var, None),
        ):
            ttk.Label(_pr, text=lbl_txt).pack(side=tk.LEFT, padx=(0, 2))
            ttk.Entry(_pr, textvariable=var, width=9).pack(side=tk.LEFT, padx=(0, 10))

        # E-value controls
        ectrl = ttk.Frame(win)
        ectrl.pack(fill=tk.X, padx=8, pady=(2, 0))
        ttk.Label(ectrl, text="E values  (V vs RHE, comma-sep):").pack(side=tk.LEFT)
        e_vals_var = tk.StringVar(value="0.70, 0.75, 0.80, 0.85")
        ttk.Entry(ectrl, textvariable=e_vals_var, width=28).pack(side=tk.LEFT, padx=(4, 10))
        ttk.Button(ectrl, text="Compute KL",
                   command=lambda: _compute()).pack(side=tk.LEFT)

        # Figure
        fig = Figure(figsize=(7.5, 4.3), dpi=100)
        ax  = fig.add_subplot(111)
        cv  = FigureCanvasTkAgg(fig, master=win)
        cv.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=8)
        tb  = NavigationToolbar2Tk(cv, win, pack_toolbar=False)
        tb.pack(side=tk.BOTTOM, fill=tk.X, padx=8)

        # Results
        res = tk.Text(win, height=6, state=tk.DISABLED, font=("Courier", 8),
                      wrap=tk.WORD)
        res.pack(fill=tk.X, padx=8, pady=(0, 4))

        def _compute():
            try:
                D  = float(d_var.get())
                nu = float(nu_var.get())
                C  = float(c_var.get())
            except ValueError:
                messagebox.showerror("KL", "Invalid parameter(s).", parent=win)
                return
            try:
                e_vals = [float(x.strip()) for x in e_vals_var.get().split(",")
                          if x.strip()]
            except ValueError:
                messagebox.showerror("KL", "Invalid E values.", parent=win)
                return
            if not e_vals:
                return
            F = 96485.0
            # Levich slope per electron: B = 0.62 F D^(2/3) nu^(-1/6) C  [A cm^-2 (rad/s)^-1/2]
            # multiply by 1000 to convert to mA
            B_factor = 0.62 * F * (D ** (2.0 / 3.0)) * (nu ** (-1.0 / 6.0)) * C * 1000.0
            ax.clear()
            lines = []
            kl_colors = [_PALETTE[k % len(_PALETTE)] for k in range(len(e_vals))]
            for ei, (e_val, c_kl) in enumerate(zip(e_vals, kl_colors)):
                inv_J = []; inv_sqw = []; rpm_labels = []
                sel_valid = [(E, J, r, l, c) for (E, J, r, l, c) in valid
                             if _ksel_vars.get(r, tk.BooleanVar(value=True)).get()]
                for E_arr, J_arr, rpm, label, _ in sel_valid:
                    if e_val < E_arr[0] or e_val > E_arr[-1]:
                        continue
                    j_at_e = float(np.interp(e_val, E_arr, J_arr))
                    if j_at_e == 0 or not np.isfinite(j_at_e):
                        continue
                    omega = 2.0 * math.pi * rpm / 60.0
                    inv_J.append(1.0 / j_at_e)
                    inv_sqw.append(1.0 / math.sqrt(omega))
                    rpm_labels.append(label)
                if len(inv_J) < 2:
                    lines.append(f"E={e_val:.3f} V: < 2 data points — skipped")
                    continue
                x = np.array(inv_sqw)
                y = np.array(inv_J)
                coeffs = np.polyfit(x, y, 1)
                slope, intercept = coeffs
                # |slope| = 1 / (n * B_factor)  →  n = 1 / (|slope| * B_factor)
                n    = (1.0 / (abs(slope) * B_factor)) if slope != 0 else float("nan")
                j_k  = (1.0 / intercept) if intercept != 0 else float("nan")
                ax.scatter(x, y, color=c_kl, zorder=5, s=40)
                xfit = np.linspace(x.min(), x.max(), 60)
                ax.plot(xfit, np.polyval(coeffs, xfit), color=c_kl,
                        linewidth=1.2, label=f"E={e_val:.3f} V  n={n:.2f}")
                for xi_pt, yi_pt, rl in zip(x, y, rpm_labels):
                    ax.annotate(rl, (xi_pt, yi_pt), fontsize=6, color=c_kl,
                                xytext=(3, 3), textcoords="offset points")
                lines.append(
                    f"E={e_val:.3f} V:  n = {n:.2f}  |  Jᵏ = {j_k:+.3f} mA")
            ax.set_xlabel(r"$\omega^{-1/2}$  (rad s$^{-1}$)$^{-1/2}$")
            ax.set_ylabel("J⁻¹  (mA⁻¹ cm² or mA⁻¹)")
            ax.set_title(f"Koutecky-Levich — {sname}")
            ax.legend(fontsize=8, frameon=True)
            fig.tight_layout(pad=0.5)
            fig.set_layout_engine("none")
            cv.draw()
            res.configure(state=tk.NORMAL)
            res.delete("1.0", tk.END)
            res.insert(tk.END, "\n".join(lines))
            res.configure(state=tk.DISABLED)

        _compute()

    # ════════════════════════════════════════════════════════════════
    # Excel export
    # ════════════════════════════════════════════════════════════════
    def _export_sample_excel(self):
        if not self.active_sample or self.active_sample not in self.samples:
            messagebox.showwarning("ORR Export", "Select a sample first.")
            return
        sentry = self.samples[self.active_sample]
        pairs  = sentry.get("pairs", [])
        if not pairs:
            messagebox.showwarning("ORR Export", "No pairs in active sample.")
            return
        try:
            r_n2  = float(self.r_sol_n2_var.get() or 0)
            r_o2  = float(self.r_sol_o2_var.get() or 0)
            e_ref = float(self.e_ref_var.get() or 0)
            area  = float(self.area_var.get() or 0)
        except ValueError:
            r_n2 = r_o2 = e_ref = area = 0.0
        ref   = self.ref_electrode_var.get()
        e_col = f"E (V vs {ref})"
        y_col = "J (mA cm⁻²)" if area > 0 else "I_net (mA)"

        path = filedialog.asksaveasfilename(
            title="Export ORR Sample to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"{self.active_sample}_ORR.xlsx",
            parent=self)
        if not path:
            return

        try:
            from openpyxl import Workbook
            wb = Workbook()
            wb.remove(wb.active)
            has_data = False
            for i, pair in enumerate(pairs):
                rpm_val    = pair.get("rpm_val") or pair.get("rpm_id") or f"#{i+1}"
                sheet_name = f"{rpm_val} rpm"[:31]
                ws = wb.create_sheet(title=sheet_name)
                ws.append([e_col, y_col,
                           "N2 short", "O2 short",
                           "R_sol_N2 (ohm)", "R_sol_O2 (ohm)",
                           "E_ref (V)", "Area (cm2)"])
                result = _process_pair(pair, r_n2, r_o2, e_ref, area)
                if result is not None:
                    E_pl, Y_pl = result
                    for j, (e_v, y_v) in enumerate(zip(E_pl, Y_pl)):
                        row = [float(e_v), float(y_v)]
                        if j == 0:
                            row += [pair.get("n2_short", ""),
                                    pair.get("o2_short", ""),
                                    r_n2, r_o2, e_ref,
                                    area if area > 0 else ""]
                        ws.append(row)
                    has_data = True
                else:
                    ws.append(["Processing failed for this pair."])
            if not has_data:
                messagebox.showwarning("ORR Export", "No pairs could be processed.")
                return
            wb.save(path)
            self._log(f"Exported '{self.active_sample}' → {os.path.basename(path)}")
            messagebox.showinfo("ORR Export", f"Saved:\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("ORR Export", f"Export failed:\n{exc}", parent=self)

    # ════════════════════════════════════════════════════════════════
    # Logging
    # ════════════════════════════════════════════════════════════════
    def _log(self, msg: str):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    # ════════════════════════════════════════════════════════════════
    # Session save / restore
    # ════════════════════════════════════════════════════════════════
    def get_session_state(self, data_store: dict) -> dict:
        self._save_active_sample_state()

        samples_list = []
        for sname, sentry in self.samples.items():
            rec = {"name": sname}
            pairs_rec = []
            for pair in sentry.get("pairs", []):
                pr = {}
                for k, v in pair.items():
                    if k in _PAIR_RUNTIME:
                        continue
                    if k in ("df_n2", "df_o2"):
                        continue
                    try:
                        import json; json.dumps(v)
                        pr[k] = v
                    except (TypeError, ValueError):
                        pass
                # Store DataFrames by hash
                for gas in ("n2", "o2"):
                    df = pair.get(f"df_{gas}")
                    if df is not None:
                        h = _sm.df_hash(df)
                        data_store[h] = df
                        pr[f"df_{gas}_hash"] = h
                pairs_rec.append(pr)
            rec["pairs"] = pairs_rec
            for k, v in sentry.items():
                if k in _SAMPLE_RUNTIME or k == "pairs":
                    continue
                try:
                    import json; json.dumps(v)
                    rec[k] = v
                except (TypeError, ValueError):
                    pass
            rec["reflines"] = [list(r) for r in sentry.get("reflines", [])]
            samples_list.append(rec)

        return {
            "active_sample": self.active_sample,
            "plot_w_var":    self.plot_w_var.get(),
            "plot_h_var":    self.plot_h_var.get(),
            "grid_cols_var": self._grid_cols_var.get(),
            "samples": samples_list,
        }

    def restore_session_state(self, state: dict, data_store: dict) -> None:
        old = self._suppress_replot
        self._suppress_replot = True

        # Destroy existing sample figures
        for sentry in self.samples.values():
            pf = sentry.get("plot_frame")
            if pf is not None:
                try:
                    pf.destroy()
                except Exception:
                    pass
        self.samples.clear()
        self.active_sample = None
        self.sample_lb.clear()
        self._zoom_sample = None

        # Clear loaded files
        self.loaded_files.clear()
        self._loaded_keys.clear()
        children = self.loaded_tv.get_children()
        if children:
            self.loaded_tv.delete(*children)

        # Panel-level vars
        try:
            self.plot_w_var.set(state.get("plot_w_var", "10.5"))
            self.plot_h_var.set(state.get("plot_h_var", "5.5"))
            self._grid_cols_var.set(state.get("grid_cols_var", "2"))
        except Exception:
            pass

        # Restore samples
        for srec in state.get("samples", []):
            sname = srec.get("name", "")
            if not sname:
                continue
            sentry: dict = {"pairs": [], "reflines": []}
            for k, v in srec.items():
                if k in ("name", "pairs"):
                    continue
                sentry[k] = v
            sentry["reflines"] = [tuple(r) for r in sentry.get("reflines", [])]

            # Restore pairs + re-attach DataFrames from data_store
            for prec in srec.get("pairs", []):
                pair = dict(prec)
                for gas in ("n2", "o2"):
                    h = pair.pop(f"df_{gas}_hash", None)
                    pair[f"df_{gas}"] = data_store.get(h) if h else None
                sentry["pairs"].append(pair)

            self.samples[sname] = sentry
            # Migrate old flat correction values to per-catalyst format
            if "catalyst_corrections" not in sentry:
                sentry["catalyst_corrections"] = {}
            old_r_n2 = sentry.get("r_sol_n2", 0.0)
            old_r_o2 = sentry.get("r_sol_o2", 0.0)
            old_eref = sentry.get("e_ref", 0.0)
            old_area = sentry.get("area", "")
            for p in sentry.get("pairs", []):
                c = p.get("catalyst_id", "")
                if c:
                    sentry["catalyst_corrections"].setdefault(
                        c, {"r_sol_n2": old_r_n2, "r_sol_o2": old_r_o2,
                            "e_ref": old_eref, "area": old_area})
            self.sample_lb.insert(tk.END, sname,
                                  checked=not srec.get("hidden", False))
            self._create_sample_figure(sname)

        self._suppress_replot = old
        active_sample = state.get("active_sample")
        if active_sample and active_sample in self.samples:
            keys = list(self.samples.keys())
            self.sample_lb.selection_set(keys.index(active_sample))
            self._switch_active_sample(active_sample)
        elif self.samples:
            first = next(iter(self.samples))
            self.sample_lb.selection_set(0)
            self._switch_active_sample(first)

        self._apply_plot_size()
        self._relayout_figures()
