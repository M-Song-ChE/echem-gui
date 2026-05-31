"""Hupd ECSA Calculation panel.

Workflow per file:
  1. Load CV file(s); extract the last cycle.
  2. User sets the double-layer (DL) baseline region and Hupd integration range.
  3. Linear baseline is fitted in the DL region and extrapolated.
  4. Q_H [uC] = (1 / v [V/s]) * |integral(I_measured - I_baseline, dE)| over Hupd range.
  5. ECSA [cm2] = Q_H / q_ref;  RF = ECSA / geometric_area.
  6. Results shown in a table; the active file's last cycle is plotted with
     the baseline, the DL region highlighted, and the integration area shaded.
"""

import os
from collections import OrderedDict

import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from .file_manager import _read_mpr
from .plotting import copy_figure_to_clipboard

# ── defaults ────────────────────────────────────────────────────────────────
_DEF = dict(
    scan_rate="50",
    dl_lo="0.40",
    dl_hi="0.50",
    e1="0.05",
    e2="0.40",
    q_ref="210",
    geo_area="0.1963",
)


# ── module-level helpers ─────────────────────────────────────────────────────
def _read_one(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".mpr":
        return _read_mpr(path)
    df = pd.read_csv(path, sep="\t", encoding="latin-1", on_bad_lines="skip")
    df.columns = [c.strip().replace("<", "").replace(">", "") for c in df.columns]
    return df.reset_index(drop=True)


def _fmt_cycle(c) -> str:
    """Display a cycle number as an integer string (EC-Lab stores them as floats)."""
    try:
        return str(int(c))
    except (TypeError, ValueError):
        return str(c)


def _get_cycles(df: pd.DataFrame):
    """Return sorted list of cycle numbers, or [] if no cycle column."""
    if "cycle number" in df.columns and len(df):
        return sorted(df["cycle number"].unique().tolist())
    return []


def _get_cycle(df: pd.DataFrame, cycle_num) -> pd.DataFrame:
    """Return rows for a specific cycle number (or all rows if no cycle column)."""
    if cycle_num is not None and "cycle number" in df.columns:
        return df[df["cycle number"] == cycle_num].copy().reset_index(drop=True)
    return df.copy().reset_index(drop=True)


_trapz = getattr(np, "trapezoid", getattr(np, "trapz", None))


def _split_scans(E: np.ndarray, I: np.ndarray):
    """Split into (E_an, I_an, E_cat, I_cat) each sorted by E ascending.

    Uses whichever extreme (min or max) sits most centrally as the pivot,
    so it handles scans that start from either the cathodic or anodic vertex.
    """
    if len(E) < 4:
        return E, I, E, I
    n = len(E)
    min_idx = int(np.argmin(E))
    max_idx = int(np.argmax(E))
    # Choose the more central extreme as the scan-direction pivot
    if abs(max_idx / n - 0.5) <= abs(min_idx / n - 0.5):
        # Anodic vertex is more central → scan: cathodic↑anodic vertex↓cathodic
        E_an  = E[:max_idx + 1].copy(); I_an  = I[:max_idx + 1].copy()
        E_cat = E[max_idx:].copy();     I_cat = I[max_idx:].copy()
    else:
        # Cathodic vertex is more central → scan: anodic↓cathodic vertex↑anodic
        E_cat = E[:min_idx + 1].copy(); I_cat = I[:min_idx + 1].copy()
        E_an  = E[min_idx:].copy();     I_an  = I[min_idx:].copy()
    o = np.argsort(E_an);  E_an,  I_an  = E_an[o],  I_an[o]
    o = np.argsort(E_cat); E_cat, I_cat = E_cat[o], I_cat[o]
    return E_an, I_an, E_cat, I_cat


def _dl_baseline(E_s, I_s, dl_lo, dl_hi):
    """Two-point baseline through the first and last data points in the DL region.
    Returns coeffs = [slope, intercept] (numpy polyval-compatible), or None."""
    mask = (E_s >= dl_lo) & (E_s <= dl_hi)
    if mask.sum() < 2:
        return None
    E_dl = E_s[mask]; I_dl = I_s[mask]
    dE = E_dl[-1] - E_dl[0]
    slope = (I_dl[-1] - I_dl[0]) / dE if dE != 0 else 0.0
    intercept = I_dl[0] - slope * E_dl[0]
    return np.array([slope, intercept])


def _integrate_one(E_s, I_s, dl_lo, dl_hi, e1, e2, v_mVs):
    """Two-point DL baseline; integrate (I - baseline) in Hupd range.
    Returns (q_uC, coeffs) or (None, None) on failure.
    """
    coeffs = _dl_baseline(E_s, I_s, dl_lo, dl_hi)
    if coeffs is None:
        return None, None

    mask_h = (E_s >= e1) & (E_s <= e2)
    if mask_h.sum() < 2:
        return None, coeffs

    E_h = E_s[mask_h]; I_h = I_s[mask_h]
    I_bl = np.polyval(coeffs, E_h)
    I_net = np.clip(I_h - I_bl, 0, None)  # area ABOVE baseline only

    # Q [C] = integral(I_net_A, dE_V) / v_Vs
    v_si = v_mVs * 1e-3                          # mV/s → V/s
    q_c  = _trapz(I_net * 1e-3, E_h) / v_si      # A·V / (V/s) = C
    return q_c * 1e6, coeffs                      # C → µC


def _compute_result(df_lc, v, dl_lo, dl_hi, e1, e2, q_ref, geo,
                    r_sol=0.0, e_ref=0.0):
    """Full Hupd computation for one last-cycle DataFrame (anodic scan only).
    Returns dict or None."""
    if df_lc is None or len(df_lc) < 10:
        return None
    if "Ewe/V" not in df_lc.columns or "I/mA" not in df_lc.columns:
        return None

    I = df_lc["I/mA"].values.astype(float)
    E = df_lc["Ewe/V"].values.astype(float) - (I * 1e-3) * r_sol + e_ref
    E_an, I_an, _, _ = _split_scans(E, I)

    q_h, coeffs = _integrate_one(E_an, I_an, dl_lo, dl_hi, e1, e2, v)
    if q_h is None:
        return None

    ecsa = q_h / q_ref if q_ref > 0 else float("nan")
    rf   = ecsa / geo  if geo   > 0 else float("nan")
    return dict(q_h=q_h, ecsa=ecsa, rf=rf, coeffs=coeffs)


# ── Panel ────────────────────────────────────────────────────────────────────
class HupdPanel(ttk.Frame):
    """Hupd-based ECSA calculation panel."""

    def __init__(self, master):
        ttk.Frame.__init__(self, master)
        self.files       = OrderedDict()  # short → {path, df, df_lc, result}
        self._keys       = []              # ordered list of short names
        self.active_file  = None
        self._suppress    = False
        self._dragging_var = None        # StringVar currently being dragged
        self._dragging_ann = False       # True when dragging the annotation box
        self._lb_drag_src  = None        # listbox drag: source row index
        self._ann_drag_offset = [0.0, 0.0]
        self._ann_pos  = [0.02, 0.97]   # annotation position in axes fraction
        self._ann_artist = None          # Text artist reference
        self._leg      = None            # Legend artist reference
        self._leg_pos  = None            # saved legend position (axes-frac tuple)
        self._leg_resizing        = False
        self._leg_resize_start_y  = None
        self._leg_resize_start_sz = None
        self._build()

    # ── Construction ─────────────────────────────────────────────────────────
    def _build(self):
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # ── Left (scrollable) ────────────────────────────────────────────
        lo = ttk.Frame(paned, width=340)
        paned.add(lo, weight=0)

        lc = tk.Canvas(lo, highlightthickness=0)
        ls = ttk.Scrollbar(lo, orient=tk.VERTICAL, command=lc.yview)
        lc.configure(yscrollcommand=ls.set)
        ls.pack(side=tk.RIGHT, fill=tk.Y)
        lc.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        left = tk.Frame(lc)
        lw = lc.create_window((0, 0), window=left, anchor=tk.NW)
        left.bind("<Configure>", lambda e: lc.configure(scrollregion=lc.bbox("all")))
        lc.bind("<Configure>",   lambda e: lc.itemconfig(lw, width=e.width))
        lc.bind("<MouseWheel>",  lambda e: lc.yview_scroll(-1*(e.delta//120), "units"))

        # ── File list ────────────────────────────────────────────────────
        ff = ttk.LabelFrame(left, text="Loaded Files")
        ff.pack(fill=tk.X, padx=4, pady=4)

        bf = ttk.Frame(ff)
        bf.pack(fill=tk.X, padx=4, pady=(4, 2))
        ttk.Button(bf, text="Load Files", command=self._load).pack(side=tk.LEFT)
        ttk.Button(bf, text="Remove",     command=self._remove).pack(side=tk.LEFT, padx=4)

        self._lb = tk.Listbox(ff, height=5, selectmode=tk.SINGLE,
                              exportselection=False, font=("", 8))
        lbs = ttk.Scrollbar(ff, orient=tk.VERTICAL, command=self._lb.yview)
        self._lb.configure(yscrollcommand=lbs.set)
        lbs.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 4))
        self._lb.pack(fill=tk.X, padx=4, pady=(0, 2))
        self._lb.bind("<<ListboxSelect>>",  self._on_lb)
        self._lb.bind("<ButtonPress-1>",    self._on_lb_press)
        self._lb.bind("<B1-Motion>",        self._on_lb_drag_motion)
        self._lb.bind("<ButtonRelease-1>",  self._on_lb_drop)

        # Cycle selector
        cr = ttk.Frame(ff)
        cr.pack(fill=tk.X, padx=4, pady=(0, 6))
        ttk.Label(cr, text="Cycle:", anchor=tk.W).pack(side=tk.LEFT)
        self._v_cycle = tk.StringVar()
        self._cycle_cb = ttk.Combobox(cr, textvariable=self._v_cycle,
                                      width=6, state="readonly")
        self._cycle_cb.pack(side=tk.LEFT, padx=(4, 0))
        self._cycle_cb.bind("<<ComboboxSelected>>", self._on_cycle_change)

        # ── Parameters ───────────────────────────────────────────────────
        pf = ttk.LabelFrame(left, text="Parameters")
        pf.pack(fill=tk.X, padx=4, pady=4)

        # ── Correction (per file) ─────────────────────────────────────────
        cf = ttk.LabelFrame(left, text="Correction (per file)")
        cf.pack(fill=tk.X, padx=4, pady=4)

        self._v_rsol = tk.StringVar(value="0.0")
        self._v_eref = tk.StringVar(value="0.0")

        def _corr_row(parent, label, var, unit=""):
            r = ttk.Frame(parent)
            r.pack(fill=tk.X, padx=6, pady=2)
            ttk.Label(r, text=label, width=24, anchor=tk.W).pack(side=tk.LEFT)
            e = ttk.Entry(r, textvariable=var, width=9)
            e.pack(side=tk.LEFT)
            e.bind("<Return>",   lambda ev: (self._save_corr(), self._replot()))
            e.bind("<FocusOut>", lambda ev: (self._save_corr(), self._replot()))
            if unit:
                ttk.Label(r, text=unit, foreground="gray",
                          font=("", 8)).pack(side=tk.LEFT, padx=(3, 0))

        _corr_row(cf, "R_sol (IR correction):", self._v_rsol, unit="Ω")
        _corr_row(cf, "E_ref (RHE offset):",    self._v_eref, unit="V")

        ttk.Label(cf, text="E_corr = E_raw − I·R_sol + E_ref",
                  font=("", 7), foreground="gray").pack(padx=6, pady=(0, 4), anchor=tk.W)

        # ── Parameters ───────────────────────────────────────────────────
        self._v_sr   = tk.StringVar(value=_DEF["scan_rate"])
        self._v_dllo = tk.StringVar(value=_DEF["dl_lo"])
        self._v_dlhi = tk.StringVar(value=_DEF["dl_hi"])
        self._v_e1   = tk.StringVar(value=_DEF["e1"])
        self._v_e2   = tk.StringVar(value=_DEF["e2"])
        self._v_qref = tk.StringVar(value=_DEF["q_ref"])
        self._v_geo  = tk.StringVar(value=_DEF["geo_area"])

        def _entry_row(parent, label, var, width=9, unit=""):
            r = ttk.Frame(parent)
            r.pack(fill=tk.X, padx=6, pady=2)
            ttk.Label(r, text=label, width=24, anchor=tk.W).pack(side=tk.LEFT)
            e = ttk.Entry(r, textvariable=var, width=width)
            e.pack(side=tk.LEFT)
            e.bind("<Return>",   lambda ev: self._replot())
            e.bind("<FocusOut>", lambda ev: self._replot())
            if unit:
                ttk.Label(r, text=unit, foreground="gray",
                          font=("", 8)).pack(side=tk.LEFT, padx=(3, 0))

        _entry_row(pf, "Scan rate:", self._v_sr, unit="mV/s")

        # DL baseline region — two entries on one row
        dr = ttk.Frame(pf)
        dr.pack(fill=tk.X, padx=6, pady=2)
        ttk.Label(dr, text="DL baseline region:", width=24, anchor=tk.W).pack(side=tk.LEFT)
        for v in (self._v_dllo, self._v_dlhi):
            e = ttk.Entry(dr, textvariable=v, width=6)
            e.pack(side=tk.LEFT, padx=(0, 2))
            e.bind("<Return>",   lambda ev: self._replot())
            e.bind("<FocusOut>", lambda ev: self._replot())
        ttk.Label(dr, text="V", foreground="gray", font=("", 8)).pack(side=tk.LEFT)

        # Hupd integration range — two entries on one row
        hr = ttk.Frame(pf)
        hr.pack(fill=tk.X, padx=6, pady=2)
        ttk.Label(hr, text="Hupd range:", width=24, anchor=tk.W).pack(side=tk.LEFT)
        for v in (self._v_e1, self._v_e2):
            e = ttk.Entry(hr, textvariable=v, width=6)
            e.pack(side=tk.LEFT, padx=(0, 2))
            e.bind("<Return>",   lambda ev: self._replot())
            e.bind("<FocusOut>", lambda ev: self._replot())
        ttk.Label(hr, text="V", foreground="gray", font=("", 8)).pack(side=tk.LEFT)

        _entry_row(pf, "qᵣₑf (μC/cm²):", self._v_qref)
        _entry_row(pf, "Geometric area:", self._v_geo, unit="cm²")

        # Legend font size
        leg_row = ttk.Frame(pf)
        leg_row.pack(fill=tk.X, padx=6, pady=2)
        ttk.Label(leg_row, text="Legend size:", width=24, anchor=tk.W).pack(side=tk.LEFT)
        self._v_leg_size = tk.StringVar(value="7.5")
        _leg_sp = ttk.Spinbox(leg_row, textvariable=self._v_leg_size,
                              from_=4.0, to=30.0, increment=0.5, width=6)
        _leg_sp.pack(side=tk.LEFT)
        _leg_sp.bind("<<Increment>>", lambda e: self._replot())
        _leg_sp.bind("<<Decrement>>", lambda e: self._replot())
        _leg_sp.bind("<Return>",      lambda e: self._replot())
        _leg_sp.bind("<FocusOut>",    lambda e: self._replot())
        ttk.Label(leg_row, text="pt  (right-drag legend to resize)",
                  foreground="gray", font=("", 7)).pack(side=tk.LEFT, padx=(4, 0))

        ttk.Button(pf, text="  Compute All  ",
                   command=self._compute_all).pack(padx=6, pady=(6, 8), anchor=tk.W)

        # ── Results table ─────────────────────────────────────────────────
        rf = ttk.LabelFrame(left, text="Results")
        rf.pack(fill=tk.BOTH, padx=4, pady=4, expand=True)

        cols = ("file", "q_h", "ecsa", "rf")
        self._tv = ttk.Treeview(rf, columns=cols, show="headings",
                                height=7, selectmode="browse")
        self._tv.heading("file", text="File")
        self._tv.heading("q_h",  text="Q_H (μC)")
        self._tv.heading("ecsa", text="ECSA (cm²)")
        self._tv.heading("rf",   text="RF")
        self._tv.column("file", width=150, anchor=tk.W, stretch=True)
        self._tv.column("q_h",  width=70,  anchor=tk.CENTER, stretch=False)
        self._tv.column("ecsa", width=80,  anchor=tk.CENTER, stretch=False)
        self._tv.column("rf",   width=45,  anchor=tk.CENTER, stretch=False)
        tvs = ttk.Scrollbar(rf, orient=tk.VERTICAL, command=self._tv.yview)
        self._tv.configure(yscrollcommand=tvs.set)
        tvs.pack(side=tk.RIGHT, fill=tk.Y)
        self._tv.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self._tv.bind("<<TreeviewSelect>>", self._on_tv)

        # ── Right: matplotlib figure ──────────────────────────────────────
        rp = ttk.Frame(paned)
        paned.add(rp, weight=1)

        self._fig = Figure(figsize=(10, 6), dpi=100)
        self._ax  = self._fig.add_subplot(111)
        self._cv  = FigureCanvasTkAgg(self._fig, master=rp)
        self._cv.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        tb_row = ttk.Frame(rp)
        tb_row.pack(fill=tk.X)
        self._tb = NavigationToolbar2Tk(self._cv, tb_row, pack_toolbar=False)
        self._tb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(tb_row, text="Copy",
                   command=lambda: copy_figure_to_clipboard(self._fig)
                   ).pack(side=tk.LEFT, padx=4)

        self._cv.mpl_connect("button_press_event",   self._on_drag_press)
        self._cv.mpl_connect("motion_notify_event",  self._on_drag_motion)
        self._cv.mpl_connect("button_release_event", self._on_drag_release)

    # ── Drag handlers ────────────────────────────────────────────────────────
    def _safe_float(self, var, default=0.0):
        try:
            return float(var.get())
        except ValueError:
            return default

    def _on_drag_press(self, event):
        if getattr(self._tb, "mode", "") != "":
            return  # toolbar pan/zoom active — don't interfere

        # Right-button on legend → resize
        if event.button == 3 and self._leg is not None:
            try:
                contains, _ = self._leg.contains(event)
                if contains:
                    self._leg_resizing        = True
                    self._leg_resize_start_y  = event.y
                    try:
                        self._leg_resize_start_sz = float(self._v_leg_size.get())
                    except ValueError:
                        self._leg_resize_start_sz = 7.5
                    return
            except Exception:
                pass

        if event.inaxes is not self._ax or event.button != 1 or event.xdata is None:
            return

        # Let matplotlib's DraggableLegend handle legend clicks
        if self._leg is not None:
            try:
                contains, _ = self._leg.contains(event)
                if contains:
                    return
            except Exception:
                pass

        # Annotation drag: move the result box in-place (no replot)
        if self._ann_artist is not None:
            try:
                contains, _ = self._ann_artist.contains(event)
                if contains:
                    self._dragging_ann = True
                    ax_x, ax_y = self._ax.transAxes.inverted().transform(
                        (event.x, event.y))
                    ann_x, ann_y = self._ann_artist.get_position()
                    self._ann_drag_offset = [ax_x - ann_x, ax_y - ann_y]
                    return
            except Exception:
                pass

        # Boundary line drag
        x = event.xdata
        xlim = self._ax.get_xlim()
        tol = (xlim[1] - xlim[0]) * 0.025
        candidates = [
            (self._v_dllo, abs(x - self._safe_float(self._v_dllo))),
            (self._v_dlhi, abs(x - self._safe_float(self._v_dlhi))),
            (self._v_e1,   abs(x - self._safe_float(self._v_e1))),
            (self._v_e2,   abs(x - self._safe_float(self._v_e2))),
        ]
        nearest = min(candidates, key=lambda t: t[1])
        if nearest[1] <= tol:
            self._dragging_var = nearest[0]

    def _on_drag_motion(self, event):
        widget = self._cv.get_tk_widget()

        # Legend right-drag resize
        if self._leg_resizing and self._leg is not None:
            dy     = event.y - self._leg_resize_start_y
            new_sz = max(4.0, min(30.0, self._leg_resize_start_sz + dy / 5.0))
            self._v_leg_size.set(f"{new_sz:.1f}")
            for t in self._leg.get_texts():
                t.set_fontsize(new_sz)
            tt = self._leg.get_title()
            if tt:
                tt.set_fontsize(new_sz)
            self._cv.draw_idle()
            return

        # Annotation drag: move artist in-place without a full replot
        if self._dragging_ann:
            if event.x is not None and event.y is not None and self._ann_artist is not None:
                ax_x, ax_y = self._ax.transAxes.inverted().transform(
                    (event.x, event.y))
                new_x = max(0.0, min(1.0, ax_x - self._ann_drag_offset[0]))
                new_y = max(0.0, min(1.0, ax_y - self._ann_drag_offset[1]))
                self._ann_artist.set_position((new_x, new_y))
                self._ann_pos = [new_x, new_y]
                self._cv.draw_idle()
            return

        if event.inaxes is not self._ax or event.xdata is None:
            if self._dragging_var is None:
                widget.config(cursor="")
            return
        x = event.xdata
        xlim = self._ax.get_xlim()
        tol = (xlim[1] - xlim[0]) * 0.025
        if self._dragging_var is not None:
            x_c = round(max(0.0, x), 3)
            self._dragging_var.set(f"{x_c:.3f}")
            self._replot()
        else:
            vals = [self._safe_float(v) for v in
                    (self._v_dllo, self._v_dlhi, self._v_e1, self._v_e2)]
            near = any(abs(x - v) <= tol for v in vals)
            widget.config(cursor="sb_h_double_arrow" if near else "")

    def _on_drag_release(self, event):
        if self._dragging_var is not None:
            self._save_params()  # persist dragged boundary value to entry
        self._dragging_var    = None
        self._dragging_ann    = False
        self._leg_resizing    = False

    # ── File operations ──────────────────────────────────────────────────────
    def _load(self):
        paths = filedialog.askopenfilenames(
            title="Load CV files (Hupd)",
            filetypes=[("Data files", "*.mpr *.txt"), ("All files", "*.*")])
        if not paths:
            return
        for path in paths:
            short = os.path.basename(path)
            base, ext = os.path.splitext(short)
            n = 1
            while short in self.files:
                short = f"{base}_{n}{ext}"; n += 1
            try:
                df     = _read_one(path)
                cycles = _get_cycles(df)
                sel_c  = cycles[-1] if cycles else None
                df_lc  = _get_cycle(df, sel_c)
                self.files[short] = {"path": path, "df": df,
                                     "df_lc": df_lc, "result": None,
                                     "r_sol": 0.0, "e_ref": 0.0,
                                     "cycles": cycles, "sel_cycle": sel_c,
                                     "scan_rate": self._v_sr.get(),
                                     "dl_lo":     self._v_dllo.get(),
                                     "dl_hi":     self._v_dlhi.get(),
                                     "e1":        self._v_e1.get(),
                                     "e2":        self._v_e2.get(),
                                     "q_ref":     self._v_qref.get(),
                                     "geo_area":  self._v_geo.get()}
                self._keys.append(short)
                self._lb.insert(tk.END, short)
            except Exception as ex:
                messagebox.showerror("Load error", f"{short}:\n{ex}")
        if not self.active_file and self._keys:
            self._lb.selection_set(0)
            self._on_lb()

    def _remove(self):
        sel = self._lb.curselection()
        if not sel:
            return
        idx   = sel[0]
        short = self._keys[idx]
        del self.files[short]
        self._keys.pop(idx)
        self._lb.delete(idx)
        for item in self._tv.get_children():
            if self._tv.item(item, "values")[0] == short:
                self._tv.delete(item); break
        if self.active_file == short:
            self.active_file = None
            if self._keys:
                self._lb.selection_set(min(idx, len(self._keys) - 1))
                self._on_lb()
        self._replot()

    def _save_corr(self):
        if not self.active_file or self.active_file not in self.files:
            return
        entry = self.files[self.active_file]
        try:
            entry["r_sol"] = float(self._v_rsol.get())
        except ValueError:
            pass
        try:
            entry["e_ref"] = float(self._v_eref.get())
        except ValueError:
            pass

    def _save_params(self):
        """Persist current UI parameter values to the active file's entry."""
        if not self.active_file or self.active_file not in self.files:
            return
        entry = self.files[self.active_file]
        entry["scan_rate"] = self._v_sr.get()
        entry["dl_lo"]     = self._v_dllo.get()
        entry["dl_hi"]     = self._v_dlhi.get()
        entry["e1"]        = self._v_e1.get()
        entry["e2"]        = self._v_e2.get()
        entry["q_ref"]     = self._v_qref.get()
        entry["geo_area"]  = self._v_geo.get()

    def _on_cycle_change(self, event=None):
        if not self.active_file or self.active_file not in self.files:
            return
        entry = self.files[self.active_file]
        raw = self._v_cycle.get()
        # Match back to original cycle number (stored as float by EC-Lab)
        cycles = entry.get("cycles", [])
        try:
            sel_c = next(c for c in cycles if _fmt_cycle(c) == raw)
        except StopIteration:
            try:
                sel_c = float(raw)
            except ValueError:
                return
        entry["sel_cycle"] = sel_c
        entry["df_lc"]     = _get_cycle(entry["df"], sel_c)
        entry["result"]    = None
        self._replot()

    # ── Listbox drag-to-reorder ──────────────────────────────────────────────
    def _on_lb_press(self, event):
        idx = self._lb.nearest(event.y)
        self._lb_drag_src = idx if 0 <= idx < len(self._keys) else None

    def _on_lb_drag_motion(self, event):
        if self._lb_drag_src is None:
            return
        self._lb.config(cursor="fleur")
        dst = max(0, min(len(self._keys) - 1, self._lb.nearest(event.y)))
        self._lb.selection_clear(0, tk.END)
        self._lb.selection_set(dst)

    def _on_lb_drop(self, event):
        self._lb.config(cursor="")
        src = self._lb_drag_src
        self._lb_drag_src = None
        if src is None:
            return
        dst = max(0, min(len(self._keys) - 1, self._lb.nearest(event.y)))
        if src == dst:
            return  # Plain click — <<ListboxSelect>> already handled it
        keys = list(self._keys)
        key  = keys.pop(src)
        keys.insert(dst, key)
        self._keys = keys
        self.files = OrderedDict((k, self.files[k]) for k in keys)
        old = self._suppress
        self._suppress = True
        self._lb.delete(0, tk.END)
        for k in self._keys:
            self._lb.insert(tk.END, k)
        self._suppress = old
        self.active_file = self._keys[dst]
        self._lb.selection_clear(0, tk.END)
        self._lb.selection_set(dst)
        self._replot()

    def _on_lb(self, event=None):
        sel = self._lb.curselection()
        if not sel:
            return
        short = self._keys[sel[0]]
        if short != self.active_file:
            self._save_corr()
            self._save_params()
            self.active_file = short
            entry = self.files[short]
            old = self._suppress
            self._suppress = True
            self._v_rsol.set(str(entry.get("r_sol", 0.0)))
            self._v_eref.set(str(entry.get("e_ref", 0.0)))
            self._v_sr.set(entry.get("scan_rate", _DEF["scan_rate"]))
            self._v_dllo.set(entry.get("dl_lo",   _DEF["dl_lo"]))
            self._v_dlhi.set(entry.get("dl_hi",   _DEF["dl_hi"]))
            self._v_e1.set(entry.get("e1",         _DEF["e1"]))
            self._v_e2.set(entry.get("e2",         _DEF["e2"]))
            self._v_qref.set(entry.get("q_ref",    _DEF["q_ref"]))
            self._v_geo.set(entry.get("geo_area",  _DEF["geo_area"]))
            # Populate cycle combobox for this file
            cycles = entry.get("cycles", [])
            self._cycle_cb["values"] = [_fmt_cycle(c) for c in cycles]
            sel_c = entry.get("sel_cycle")
            self._v_cycle.set(_fmt_cycle(sel_c) if sel_c is not None else "")
            self._suppress = old
            self._replot()

    def _on_tv(self, event=None):
        sel = self._tv.selection()
        if not sel:
            return
        short = self._tv.item(sel[0], "values")[0]
        if short in self.files and short != self.active_file:
            idx = self._keys.index(short)
            self._lb.selection_clear(0, tk.END)
            self._lb.selection_set(idx)
            self._on_lb()  # handles save+restore+replot

    # ── Parameter parsing ────────────────────────────────────────────────────
    def _params(self):
        try:
            return dict(
                v     = float(self._v_sr.get()),
                dl_lo = float(self._v_dllo.get()),
                dl_hi = float(self._v_dlhi.get()),
                e1    = float(self._v_e1.get()),
                e2    = float(self._v_e2.get()),
                q_ref = float(self._v_qref.get()),
                geo   = float(self._v_geo.get()),
            )
        except ValueError:
            return None

    # ── Computation ──────────────────────────────────────────────────────────
    def _compute_all(self):
        self._save_params()  # persist current UI to active file before batch compute
        any_ok = False
        for short in self._keys:
            entry = self.files[short]
            try:
                p = dict(
                    v     = float(entry.get("scan_rate", _DEF["scan_rate"])),
                    dl_lo = float(entry.get("dl_lo",     _DEF["dl_lo"])),
                    dl_hi = float(entry.get("dl_hi",     _DEF["dl_hi"])),
                    e1    = float(entry.get("e1",         _DEF["e1"])),
                    e2    = float(entry.get("e2",         _DEF["e2"])),
                    q_ref = float(entry.get("q_ref",     _DEF["q_ref"])),
                    geo   = float(entry.get("geo_area",  _DEF["geo_area"])),
                )
            except (ValueError, TypeError):
                continue
            res = _compute_result(
                entry["df_lc"],
                p["v"], p["dl_lo"], p["dl_hi"],
                p["e1"], p["e2"], p["q_ref"], p["geo"],
                r_sol=entry.get("r_sol", 0.0),
                e_ref=entry.get("e_ref", 0.0))
            self.files[short]["result"] = res
            if res:
                any_ok = True
        self._rebuild_tv()
        self._replot()
        if not any_ok and self._keys:
            messagebox.showwarning(
                "Hupd",
                "No valid results.\n"
                "Check that the DL baseline region and Hupd range fall within the CV data.")

    def _rebuild_tv(self):
        for item in self._tv.get_children():
            self._tv.delete(item)
        for short in self._keys:
            r = self.files[short].get("result")
            if r:
                vals = (short,
                        f"{r['q_h']:.1f}",
                        f"{r['ecsa']:.4f}",
                        f"{r['rf']:.2f}")
            else:
                vals = (short, "—", "—", "—")
            self._tv.insert("", tk.END, values=vals)

    # ── Plot ─────────────────────────────────────────────────────────────────
    def _replot(self):
        if self._suppress:
            return
        ax = self._ax

        # Save legend position (if user dragged it) and disconnect handlers
        if self._leg is not None:
            try:
                loc = self._leg._loc
                if hasattr(loc, "__len__") and len(loc) == 2:
                    self._leg_pos = (float(loc[0]), float(loc[1]))
                self._leg.set_draggable(False)
            except Exception:
                pass
            self._leg = None

        self._ann_artist = None
        ax.clear()

        if not self.active_file or self.active_file not in self.files:
            self._cv.draw_idle()
            return

        entry  = self.files[self.active_file]
        df_lc  = entry["df_lc"]
        result = entry.get("result")
        p      = self._params()

        if "Ewe/V" not in df_lc.columns or "I/mA" not in df_lc.columns:
            ax.text(0.5, 0.5, "No Ewe/V or I/mA columns",
                    transform=ax.transAxes, ha="center")
            self._cv.draw_idle()
            return

        r_sol = entry.get("r_sol", 0.0)
        e_ref = entry.get("e_ref", 0.0)
        I = df_lc["I/mA"].values.astype(float)
        E = df_lc["Ewe/V"].values.astype(float) - (I * 1e-3) * r_sol + e_ref
        E_an, I_an, E_cat, I_cat = _split_scans(E, I)

        # ── Full last-cycle background (light gray) ───────────────────────
        ax.plot(E_cat, I_cat, color="#c0c0c0", linewidth=1.0, zorder=1)
        ax.plot(E_an,  I_an,  color="#c0c0c0", linewidth=1.0, zorder=1)

        # ── Anodic scan highlighted (Hupd is anodic peak) ─────────────────
        ax.plot(E_an, I_an, color="steelblue", linewidth=1.8,
                zorder=2, label="Anodic scan")

        # Overlay: baseline and fill always computed on anodic scan
        E_s, I_s = E_an, I_an

        if p:
            # DL region — orange band + draggable edge lines
            ax.axvspan(p["dl_lo"], p["dl_hi"],
                       alpha=0.15, color="orange", zorder=0,
                       label=f"DL region [{p['dl_lo']:.3f}–{p['dl_hi']:.3f} V]")
            ax.axvline(p["dl_lo"], color="darkorange", linewidth=2.0,
                       linestyle="--", zorder=6)
            ax.axvline(p["dl_hi"], color="darkorange", linewidth=2.0,
                       linestyle="--", zorder=6)
            # Hupd boundary draggable lines
            ax.axvline(p["e1"], color="seagreen", linewidth=2.0, linestyle="--",
                       zorder=5, label=f"Hupd [{p['e1']:.3f}–{p['e2']:.3f} V]")
            ax.axvline(p["e2"], color="seagreen", linewidth=2.0, linestyle="--",
                       zorder=5)

            # ── Baseline: two-point line through first/last DL data points ──
            coeffs = _dl_baseline(E_s, I_s, p["dl_lo"], p["dl_hi"])
            if coeffs is not None:
                # Mark the two anchor points
                mask_dl = (E_s >= p["dl_lo"]) & (E_s <= p["dl_hi"])
                E_dl = E_s[mask_dl]; I_dl = I_s[mask_dl]
                ax.plot([E_dl[0], E_dl[-1]], [I_dl[0], I_dl[-1]],
                        "o", color="darkorange", markersize=5, zorder=6)

                # Extrapolated baseline drawn across the full scan range
                E_bl = np.linspace(E_s.min(), E_s.max(), 300)
                ax.plot(E_bl, np.polyval(coeffs, E_bl),
                        color="black", linewidth=1.3, linestyle="--",
                        zorder=4, label="Baseline (two-point DL)")

                # Integration fill — only area ABOVE baseline
                mask_h = (E_s >= p["e1"]) & (E_s <= p["e2"])
                if mask_h.sum() >= 2:
                    E_f    = E_s[mask_h]
                    I_f    = I_s[mask_h]
                    I_bl_f = np.polyval(coeffs, E_f)
                    # fill_between with where=(I_f > I_bl_f) shows only the peak
                    q_label = (f"Q$_{{Hupd}}$ area ({result['q_h']:.1f} μC)"
                               if result else "Q$_{{Hupd}}$ area")
                    ax.fill_between(E_f, I_f, I_bl_f,
                                    where=(I_f > I_bl_f),
                                    alpha=0.45, color="mediumseagreen",
                                    zorder=3, label=q_label)

            # Annotation box — only when Q_H has been computed (draggable)
            if result:
                ann = (f"Q$_H$  = {result['q_h']:.2f} μC\n"
                       f"ECSA = {result['ecsa']:.4f} cm²\n"
                       f"RF    = {result['rf']:.2f}")
                self._ann_artist = ax.text(
                    self._ann_pos[0], self._ann_pos[1], ann,
                    transform=ax.transAxes, fontsize=8.5,
                    verticalalignment="top", family="monospace",
                    bbox=dict(boxstyle="round,pad=0.5", fc="white",
                              alpha=0.90, ec="steelblue", lw=0.9))

        ax.axhline(0, color="black", linewidth=0.5, alpha=0.30, zorder=0)
        xlabel = "E (V vs. RHE)" if e_ref != 0.0 else "E (V vs. Ref)"
        ax.set_xlabel(xlabel)
        ax.set_ylabel("I (mA)")
        entry_title = self.files.get(self.active_file, {})
        cycle_str = _fmt_cycle(entry_title.get("sel_cycle")) if entry_title.get("sel_cycle") is not None else "?"
        ax.set_title(f"Hupd ECSA — {self.active_file}  (cycle {cycle_str})")
        try:
            _leg_sz = float(self._v_leg_size.get())
        except (ValueError, AttributeError):
            _leg_sz = 7.5
        if self._leg_pos is not None:
            self._leg = ax.legend(fontsize=_leg_sz, frameon=True, loc=self._leg_pos)
        else:
            self._leg = ax.legend(fontsize=_leg_sz, frameon=True, loc="upper right")
        self._leg.set_draggable(True)

        self._fig.tight_layout(pad=0.7)
        self._fig.set_layout_engine("none")
        self._cv.draw_idle()

    # ── Session ──────────────────────────────────────────────────────────────
    def get_session_state(self, data_store):
        from .session_manager import df_hash
        self._save_params()  # ensure active file has latest values
        state = {
            "active_file": self.active_file,
            "leg_size":    self._v_leg_size.get(),
            "files":       [],
        }
        for short, entry in self.files.items():
            h = df_hash(entry["df"])
            data_store[h] = entry["df"]
            state["files"].append({
                "short":     short,
                "path":      entry["path"],
                "hash":      h,
                "r_sol":     entry.get("r_sol",     0.0),
                "e_ref":     entry.get("e_ref",     0.0),
                "sel_cycle": entry.get("sel_cycle"),
                "scan_rate": entry.get("scan_rate", _DEF["scan_rate"]),
                "dl_lo":     entry.get("dl_lo",     _DEF["dl_lo"]),
                "dl_hi":     entry.get("dl_hi",     _DEF["dl_hi"]),
                "e1":        entry.get("e1",         _DEF["e1"]),
                "e2":        entry.get("e2",         _DEF["e2"]),
                "q_ref":     entry.get("q_ref",     _DEF["q_ref"]),
                "geo_area":  entry.get("geo_area",  _DEF["geo_area"]),
            })
        return state

    def restore_session_state(self, state, data_store):
        self._suppress = True
        try:
            self._v_leg_size.set(state.get("leg_size", "7.5"))

            self.files.clear()
            self._keys.clear()
            self._lb.delete(0, tk.END)

            for rec in state.get("files", []):
                df = data_store.get(rec["hash"])
                if df is None:
                    continue
                short = rec["short"]
                cycles = _get_cycles(df)
                sel_c  = rec.get("sel_cycle")
                if sel_c is None and cycles:
                    sel_c = cycles[-1]
                df_lc = _get_cycle(df, sel_c)
                self.files[short] = {
                    "path":      rec["path"],
                    "df":        df,
                    "df_lc":     df_lc,
                    "result":    None,
                    "r_sol":     rec.get("r_sol",     0.0),
                    "e_ref":     rec.get("e_ref",     0.0),
                    "cycles":    cycles,
                    "sel_cycle": sel_c,
                    "scan_rate": rec.get("scan_rate", _DEF["scan_rate"]),
                    "dl_lo":     rec.get("dl_lo",     _DEF["dl_lo"]),
                    "dl_hi":     rec.get("dl_hi",     _DEF["dl_hi"]),
                    "e1":        rec.get("e1",         _DEF["e1"]),
                    "e2":        rec.get("e2",         _DEF["e2"]),
                    "q_ref":     rec.get("q_ref",     _DEF["q_ref"]),
                    "geo_area":  rec.get("geo_area",  _DEF["geo_area"]),
                }
                self._keys.append(short)
                self._lb.insert(tk.END, short)

            self.active_file = state.get("active_file")
            if self.active_file not in self.files:
                self.active_file = self._keys[0] if self._keys else None
            if self.active_file:
                self._lb.selection_set(self._keys.index(self.active_file))
                entry = self.files[self.active_file]
                self._v_rsol.set(str(entry.get("r_sol", 0.0)))
                self._v_eref.set(str(entry.get("e_ref", 0.0)))
                self._v_sr.set(entry.get("scan_rate", _DEF["scan_rate"]))
                self._v_dllo.set(entry.get("dl_lo",   _DEF["dl_lo"]))
                self._v_dlhi.set(entry.get("dl_hi",   _DEF["dl_hi"]))
                self._v_e1.set(entry.get("e1",         _DEF["e1"]))
                self._v_e2.set(entry.get("e2",         _DEF["e2"]))
                self._v_qref.set(entry.get("q_ref",    _DEF["q_ref"]))
                self._v_geo.set(entry.get("geo_area",  _DEF["geo_area"]))
                cycles = entry.get("cycles", [])
                self._cycle_cb["values"] = [_fmt_cycle(c) for c in cycles]
                sel_c = entry.get("sel_cycle")
                self._v_cycle.set(_fmt_cycle(sel_c) if sel_c is not None else "")
        finally:
            self._suppress = False
        self._compute_all()
