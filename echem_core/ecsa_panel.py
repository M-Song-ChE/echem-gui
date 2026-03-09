"""ECSA Calculation Panel.

Layout
------
Left (scrollable):
    Files · axis selectors + unit dropdowns · reference electrode ·
    IR correction · cycle checkboxes (9 col) · scan-rate table (8 col) ·
    ECSA parameters · buttons · legend-frame toggle · result label · log.

Right:
    Upper frame – CV figure + its own NavigationToolbar2Tk
    Lower frame – Cdl figure + its own NavigationToolbar2Tk

Each figure is independent: toolbar Home/Zoom/Pan/Save work per-plot.

Interactions (both plots independently)
-----------------------------------------
    Scroll wheel    – zoom centred on cursor
    Left-drag       – pan
    Left-click      – annotate nearest point (cycles overlapping lines)
    Right-click     – dismiss annotation
    Drag legend     – move  (set_draggable)
    Right-drag leg  – resize font
"""

from collections import OrderedDict

import numpy as np
import tkinter as tk
from tkinter import ttk

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from .file_manager import FileManagerMixin, _COLOR_NAMES, _COLOR_HEX, _default_xcol, _default_ycol, _PLOT_STYLES, _PLOT_STYLE_NAMES
from .checklist import CheckableListbox
from .correction import CorrectionMixin
from .plotting import apply_grid, draw_reflines, _cycle_colors, copy_figure_to_clipboard
from .legend_editor import open_legend_editor

_CYCLE_BG        = "#e8f0fe"
_CYCLE_ACTIVE_BG = "#cce0ff"
_CLICK_CYCLE_PX  = 8

# ── Unit tables (shared with EchemPanel) ─────────────────────────────
_UNIT_DIMS = {
    "A": "I",  "mA": "I",  "µA": "I",  "nA": "I",
    "V": "E",  "mV": "E",  "µV": "E",  "nV": "E",
    "s": "t",  "ms": "t",  "µs": "t",  "min": "t", "h": "t",
}
_DIM_OPTS = {
    "I": ["(auto)", "A",  "mA",  "µA",  "nA"],
    "E": ["(auto)", "V",  "mV",  "µV",  "nV"],
    "t": ["(auto)", "s",  "ms",  "µs",  "min", "h"],
}
_ALL_UNITS = ["(auto)", "A", "mA", "µA", "nA",
              "V", "mV", "µV", "nV", "s", "ms", "µs", "min", "h"]
_VOLTAGE_UNITS = frozenset({"V", "mV", "µV", "nV"})


class ECSAPanel(FileManagerMixin, CorrectionMixin, ttk.Frame):
    """Self-contained ECSA extraction panel."""

    def __init__(self, master):
        ttk.Frame.__init__(self, master)
        self.files            = OrderedDict()
        self.active_file      = None
        self._suppress_replot = False
        self._loading_files   = False
        self._cycle_vars      = {}   # {cycle_num: BooleanVar}
        self._sr_vars         = {}   # {cycle_num: StringVar}
        self._sr_traces       = {}   # {cycle_num: (var, trace_id)}
        self._cv_redraw_id    = None
        self._build_panel()

    # ════════════════════════════════════════════════════════════════
    # Panel construction
    # ════════════════════════════════════════════════════════════════
    def _build_panel(self):
        body = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True)

        # ── Scrollable left panel ─────────────────────────────────
        left_outer = ttk.Frame(body, width=310)
        body.add(left_outer, weight=0)

        _lc  = tk.Canvas(left_outer, highlightthickness=0)
        _ls  = ttk.Scrollbar(left_outer, orient=tk.VERTICAL, command=_lc.yview)
        _lc.configure(yscrollcommand=_ls.set)
        _ls.pack(side=tk.RIGHT, fill=tk.Y)
        _lc.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        left = tk.Frame(_lc)
        _lwin = _lc.create_window((0, 0), window=left, anchor=tk.NW)
        left.bind("<Configure>", lambda e: _lc.configure(scrollregion=_lc.bbox("all")))
        _lc.bind("<Configure>", lambda e: _lc.itemconfig(_lwin, width=e.width))
        _lc.bind("<MouseWheel>", lambda e: _lc.yview_scroll(-1 * (e.delta // 120), "units"))

        # ── Files ─────────────────────────────────────────────────
        ttk.Label(left, text="Files:", font=("", 9, "bold")).pack(anchor=tk.W, padx=4, pady=(6, 0))
        ttk.Label(left, text="Load one file per catalyst / scan-rate series",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)
        fb = ttk.Frame(left)
        fb.pack(fill=tk.X, padx=4)
        ttk.Button(fb, text="Load File(s)", command=self._load_files).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(fb, text="Remove",       command=self._remove_file).pack(side=tk.LEFT)
        flf = ttk.Frame(left)
        flf.pack(fill=tk.X, padx=4, pady=2)
        self.file_listbox = CheckableListbox(flf, height=4, show_checkboxes=False,
                                             on_reorder=self._on_file_reorder)
        self.file_listbox.pack(fill=tk.X, expand=True)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        # ── File Color ────────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(left, text="File Color", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        _fc_row = ttk.Frame(left)
        _fc_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_fc_row, text="Color:").pack(side=tk.LEFT)
        self.file_color_var = tk.StringVar(value="Blue")
        _file_color_cb = ttk.Combobox(_fc_row, textvariable=self.file_color_var,
                                      values=_COLOR_NAMES, state="readonly", width=12)
        _file_color_cb.pack(side=tk.LEFT, padx=(4, 0))
        _file_color_cb.bind("<<ComboboxSelected>>", self._on_file_color_change)
        ttk.Label(_fc_row, text="Width:").pack(side=tk.LEFT, padx=(8, 0))
        self.linewidth_var = tk.StringVar(value="1.5")
        _lw_e = ttk.Entry(_fc_row, textvariable=self.linewidth_var, width=4)
        _lw_e.pack(side=tk.LEFT, padx=(2, 0))
        _lw_e.bind("<Return>",   lambda e: self._on_linewidth_change())
        _lw_e.bind("<FocusOut>", lambda e: self._on_linewidth_change())
        ttk.Label(_fc_row, text="Shape:").pack(side=tk.LEFT, padx=(8, 0))
        self.plot_style_var = tk.StringVar(value="Line")
        _style_cb = ttk.Combobox(_fc_row, textvariable=self.plot_style_var,
                                  values=_PLOT_STYLE_NAMES, state="readonly", width=11)
        _style_cb.pack(side=tk.LEFT, padx=(2, 0))
        _style_cb.bind("<<ComboboxSelected>>", lambda e: self._on_plot_style_change())

        # ── Axis selectors + unit dropdowns ──────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)

        def _refresh_unit_opts(col_var, unit_var, unit_cb):
            col      = col_var.get()
            raw_unit = col.rsplit("/", 1)[-1].strip() if "/" in col else ""
            dim      = _UNIT_DIMS.get(raw_unit)
            opts     = _DIM_OPTS.get(dim, _ALL_UNITS)
            unit_cb["values"] = opts
            if unit_var.get() not in opts:
                unit_var.set("(auto)")
            self._auto_replot()

        def _refresh_unit_after(unit_var, unit_cb):
            chosen = unit_var.get()
            if chosen and chosen != "(auto)":
                dim  = _UNIT_DIMS.get(chosen)
                opts = _DIM_OPTS.get(dim, _ALL_UNITS)
                unit_cb["values"] = opts
            self._auto_replot()

        # X-axis
        ttk.Label(left, text="X-axis (potential):").pack(anchor=tk.W, padx=4)
        x_row = ttk.Frame(left)
        x_row.pack(fill=tk.X, padx=4, pady=2)
        self.x_var   = tk.StringVar()
        self.x_combo = ttk.Combobox(x_row, textvariable=self.x_var, state="readonly", width=16)
        self.x_combo.pack(side=tk.LEFT)
        self.x_unit_var = tk.StringVar(value="V")
        x_unit_cb = ttk.Combobox(x_row, textvariable=self.x_unit_var,
                                  values=_DIM_OPTS["E"], state="readonly", width=6)
        x_unit_cb.pack(side=tk.LEFT, padx=(4, 0))
        self.x_combo.bind("<<ComboboxSelected>>",
                          lambda e: _refresh_unit_opts(self.x_var, self.x_unit_var, x_unit_cb))
        x_unit_cb.bind("<<ComboboxSelected>>",
                       lambda e: _refresh_unit_after(self.x_unit_var, x_unit_cb))

        # Y-axis
        ttk.Label(left, text="Y-axis (current):").pack(anchor=tk.W, padx=4)
        y_row = ttk.Frame(left)
        y_row.pack(fill=tk.X, padx=4, pady=2)
        self.y_var   = tk.StringVar()
        self.y_combo = ttk.Combobox(y_row, textvariable=self.y_var, state="readonly", width=16)
        self.y_combo.pack(side=tk.LEFT)
        self.y_unit_var = tk.StringVar(value="mA")
        y_unit_cb = ttk.Combobox(y_row, textvariable=self.y_unit_var,
                                  values=_DIM_OPTS["I"], state="readonly", width=6)
        y_unit_cb.pack(side=tk.LEFT, padx=(4, 0))
        self.y_combo.bind("<<ComboboxSelected>>",
                          lambda e: _refresh_unit_opts(self.y_var, self.y_unit_var, y_unit_cb))
        y_unit_cb.bind("<<ComboboxSelected>>",
                       lambda e: _refresh_unit_after(self.y_unit_var, y_unit_cb))

        def _swap_xy():
            xc, yc = self.x_var.get(),     self.y_var.get()
            xu, yu = self.x_unit_var.get(), self.y_unit_var.get()
            xn, yn = self.x_min_var.get(),  self.y_min_var.get()
            xx, yx = self.x_max_var.get(),  self.y_max_var.get()
            xf, yf = self.x_flip_var.get(), self.y_flip_var.get()
            self.x_var.set(yc);      self.y_var.set(xc)
            self.x_unit_var.set(yu); self.y_unit_var.set(xu)
            self.x_min_var.set(yn);  self.y_min_var.set(xn)
            self.x_max_var.set(yx);  self.y_max_var.set(xx)
            self.x_flip_var.set(yf); self.y_flip_var.set(xf)
            self._suppress_replot = True
            _refresh_unit_opts(self.x_var, self.x_unit_var, x_unit_cb)
            self._suppress_replot = False
            _refresh_unit_opts(self.y_var, self.y_unit_var, y_unit_cb)

        ttk.Button(left, text="⇄  Swap X↔Y", command=_swap_xy).pack(
            anchor=tk.W, padx=4, pady=(0, 4))

        # CV plot range (min / max per axis)
        ttk.Label(left, text="CV Plot Range:", font=("", 8)).pack(anchor=tk.W, padx=4, pady=(6, 0))
        xr_f = ttk.Frame(left)
        xr_f.pack(fill=tk.X, padx=4, pady=(1, 0))
        ttk.Label(xr_f, text="X min:").pack(side=tk.LEFT)
        self.x_min_var = tk.StringVar()
        _xmin = ttk.Entry(xr_f, textvariable=self.x_min_var, width=7)
        _xmin.pack(side=tk.LEFT, padx=(2, 4))
        ttk.Label(xr_f, text="X max:").pack(side=tk.LEFT)
        self.x_max_var = tk.StringVar()
        _xmax = ttk.Entry(xr_f, textvariable=self.x_max_var, width=7)
        _xmax.pack(side=tk.LEFT, padx=(2, 4))
        ttk.Label(xr_f, text="Int:").pack(side=tk.LEFT)
        self.cv_x_grid_int_var = tk.StringVar(value="0")
        _cvxgi = ttk.Entry(xr_f, textvariable=self.cv_x_grid_int_var, width=5)
        _cvxgi.pack(side=tk.LEFT, padx=(2, 0))
        _cvxgi.bind("<Return>",   lambda e: self._auto_replot())
        _cvxgi.bind("<FocusOut>", lambda e: self._auto_replot())
        yr_f = ttk.Frame(left)
        yr_f.pack(fill=tk.X, padx=4, pady=(1, 0))
        ttk.Label(yr_f, text="Y min:").pack(side=tk.LEFT)
        self.y_min_var = tk.StringVar()
        _ymin = ttk.Entry(yr_f, textvariable=self.y_min_var, width=7)
        _ymin.pack(side=tk.LEFT, padx=(2, 4))
        ttk.Label(yr_f, text="Y max:").pack(side=tk.LEFT)
        self.y_max_var = tk.StringVar()
        _ymax = ttk.Entry(yr_f, textvariable=self.y_max_var, width=7)
        _ymax.pack(side=tk.LEFT, padx=(2, 4))
        ttk.Label(yr_f, text="Int:").pack(side=tk.LEFT)
        self.cv_y_grid_int_var = tk.StringVar(value="0")
        _cvygi = ttk.Entry(yr_f, textvariable=self.cv_y_grid_int_var, width=5)
        _cvygi.pack(side=tk.LEFT, padx=(2, 0))
        _cvygi.bind("<Return>",   lambda e: self._auto_replot())
        _cvygi.bind("<FocusOut>", lambda e: self._auto_replot())
        ttk.Label(left, text="(blank = auto)", foreground="gray",
                  font=("", 8)).pack(anchor=tk.W, padx=4)
        for _re in (_xmin, _xmax, _ymin, _ymax):
            _re.bind("<Return>",   lambda e: self._plot_cv())
            _re.bind("<FocusOut>", lambda e: self._plot_cv())

        flip_row = ttk.Frame(left)
        flip_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        self.x_flip_var = tk.BooleanVar(value=False)
        self.y_flip_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(flip_row, text="Flip X", variable=self.x_flip_var,
                        command=self._plot_cv).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Checkbutton(flip_row, text="Flip Y", variable=self.y_flip_var,
                        command=self._plot_cv).pack(side=tk.LEFT)

        # Reference electrode
        ttk.Label(left, text="Reference Electrode:").pack(anchor=tk.W, padx=4, pady=(4, 0))
        self.ref_electrode_var = tk.StringVar(value="Ag/AgCl")
        _ref_cb = ttk.Combobox(
            left, textvariable=self.ref_electrode_var,
            values=["Ag/AgCl", "SCE", "SHE", "NHE", "RHE",
                    "Hg/HgO", "Hg/HgSO4 (MSE)", "Fc/Fc+", "Ag/Ag+", "Li/Li+"],
            state="readonly", width=24,
        )
        _ref_cb.pack(padx=4, pady=2)
        _ref_cb.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())

        # Hidden vars required by FileManagerMixin (IR/RHE correction not used in ECSA)
        self.r_sol_var = tk.StringVar(value="0")
        self.e_ref_var = tk.StringVar(value="0")

        # ── Cycle checkboxes (9 columns) ──────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Cycles:", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        cb_row = ttk.Frame(left)
        cb_row.pack(fill=tk.X, padx=4)
        ttk.Button(cb_row, text="Select All",   command=self._select_all).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(cb_row, text="Deselect All", command=self._deselect_all).pack(side=tk.LEFT)

        cyc_outer = ttk.Frame(left)
        cyc_outer.pack(fill=tk.X, padx=4, pady=2)
        cyc_canvas = tk.Canvas(cyc_outer, background=_CYCLE_BG, highlightthickness=0, height=90)
        cyc_vs = ttk.Scrollbar(cyc_outer, orient=tk.VERTICAL,   command=cyc_canvas.yview)
        cyc_hs = ttk.Scrollbar(cyc_outer, orient=tk.HORIZONTAL, command=cyc_canvas.xview)
        cyc_canvas.configure(yscrollcommand=cyc_vs.set, xscrollcommand=cyc_hs.set)
        cyc_vs.pack(side=tk.RIGHT,  fill=tk.Y)
        cyc_hs.pack(side=tk.BOTTOM, fill=tk.X)
        cyc_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._cycle_inner  = tk.Frame(cyc_canvas, background=_CYCLE_BG)
        self._cycle_canvas = cyc_canvas
        cyc_canvas.create_window((0, 0), window=self._cycle_inner, anchor=tk.NW)
        self._cycle_inner.bind("<Configure>",
                               lambda e: cyc_canvas.configure(scrollregion=cyc_canvas.bbox("all")))

        def _cyc_wheel(e):
            cyc_canvas.yview_scroll(-1 * (e.delta // 120), "units")
            return "break"
        cyc_canvas.bind("<MouseWheel>", _cyc_wheel)
        self._cycle_inner.bind("<MouseWheel>", _cyc_wheel)

        # ── Cycle Colors ─────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(left, text="Cycle Colors (CV)", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        _cc_row1 = ttk.Frame(left)
        _cc_row1.pack(fill=tk.X, padx=4, pady=2)
        self.cycle_gradient_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(_cc_row1, text="Gradient", variable=self.cycle_gradient_var,
                        command=self._on_gradient_change).pack(side=tk.LEFT)
        self.cycle_reverse_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(_cc_row1, text="Reverse", variable=self.cycle_reverse_var,
                        command=self._on_gradient_change).pack(side=tk.LEFT, padx=(8, 0))
        _cc_row2 = ttk.Frame(left)
        _cc_row2.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(_cc_row2, text="Step:").pack(side=tk.LEFT)
        self.lightness_step_var = tk.StringVar(value="0.08")
        _step_spin = ttk.Spinbox(_cc_row2, textvariable=self.lightness_step_var,
                                  from_=0.01, to=0.30, increment=0.01, width=6)
        _step_spin.pack(side=tk.LEFT, padx=(4, 0))
        _step_spin.bind("<<Increment>>", lambda e: self._on_gradient_change())
        _step_spin.bind("<<Decrement>>", lambda e: self._on_gradient_change())
        _step_spin.bind("<Return>",      lambda e: self._on_gradient_change())
        _step_spin.bind("<FocusOut>",    lambda e: self._on_gradient_change())

        # ── Scan-rate per cycle (8-column grid) ───────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Scan Rate per Cycle", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        ttk.Label(left, text="Enter scan rate (mV/s) — legend updates automatically",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)

        sr_outer = ttk.Frame(left)
        sr_outer.pack(fill=tk.X, padx=4, pady=2)
        sr_canvas = tk.Canvas(sr_outer, highlightthickness=0, height=90)
        sr_sc     = ttk.Scrollbar(sr_outer, orient=tk.VERTICAL, command=sr_canvas.yview)
        sr_canvas.configure(yscrollcommand=sr_sc.set)
        sr_sc.pack(side=tk.RIGHT, fill=tk.Y)
        sr_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._sr_inner  = tk.Frame(sr_canvas)
        self._sr_canvas = sr_canvas
        _sr_win = sr_canvas.create_window((0, 0), window=self._sr_inner, anchor=tk.NW)
        self._sr_inner.bind("<Configure>",
                            lambda e: sr_canvas.configure(scrollregion=sr_canvas.bbox("all")))
        sr_canvas.bind("<Configure>", lambda e: sr_canvas.itemconfig(_sr_win, width=e.width))

        # ── ECSA parameters ───────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="ECSA Parameters", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        estd_f = ttk.Frame(left)
        estd_f.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Label(estd_f, text="E_std (V):").pack(side=tk.LEFT)
        self.e_std_var  = tk.StringVar(value="")
        e_std_entry     = ttk.Entry(estd_f, textvariable=self.e_std_var, width=10)
        e_std_entry.pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(estd_f, text="  Rec:").pack(side=tk.LEFT, padx=(8, 0))
        self.e_std_rec_var = tk.StringVar(value="—")
        ttk.Label(estd_f, textvariable=self.e_std_rec_var,
                  foreground="#1a7a30", font=("", 9, "bold")).pack(side=tk.LEFT, padx=(2, 0))
        ttk.Label(estd_f, text="V", foreground="#1a7a30",
                  font=("", 8)).pack(side=tk.LEFT, padx=(2, 0))
        ttk.Label(left, text="  ← ja/jc read here; Rec = mid of plotted data range",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4, pady=(0, 2))

        def _on_estd_change(e=None):
            # Persist immediately so file-switch always sees the latest value
            if self.active_file and self.active_file in self.files:
                self.files[self.active_file]["e_std"] = self.e_std_var.get()
            self._plot_cv()

        e_std_entry.bind("<Return>",   _on_estd_change)
        e_std_entry.bind("<FocusOut>", _on_estd_change)

        cs_f = ttk.Frame(left)
        cs_f.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(cs_f, text="Cs (mF/cm²):").pack(side=tk.LEFT)
        self.cs_var = tk.StringVar(value="0.040")
        cs_entry = ttk.Entry(cs_f, textvariable=self.cs_var, width=10)
        cs_entry.pack(side=tk.LEFT, padx=4)
        ttk.Label(cs_f, text="typical 0.040", foreground="gray", font=("", 8)).pack(side=tk.LEFT)

        def _on_cs_change(e=None):
            # Persist immediately so file-switch always sees the latest value
            if self.active_file and self.active_file in self.files:
                self.files[self.active_file]["cs"] = self.cs_var.get()

        cs_entry.bind("<Return>",   _on_cs_change)
        cs_entry.bind("<FocusOut>", _on_cs_change)

        # ── Action buttons ────────────────────────────────────────
        act_f = ttk.Frame(left)
        act_f.pack(fill=tk.X, padx=4, pady=6)
        ttk.Button(act_f, text="Plot CV",
                   command=self._plot_cv).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(act_f, text="Extract Cdl & ECSA",
                   command=self._extract_cdl_ecsa).pack(side=tk.LEFT)

        # Legend frame toggles
        leg_opt_row = ttk.Frame(left)
        leg_opt_row.pack(fill=tk.X, padx=4, pady=(0, 4))
        self.legend_frame_cv_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            leg_opt_row, text="Show CV Legend Frame",
            variable=self.legend_frame_cv_var,
            command=self._toggle_cv_legend_frame,
        ).pack(side=tk.LEFT)
        self.legend_frame_cdl_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            leg_opt_row, text="Show Cdl Legend Frame",
            variable=self.legend_frame_cdl_var,
            command=self._toggle_cdl_legend_frame,
        ).pack(side=tk.LEFT, padx=(8, 0))

        # CV Grid
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(left, text="CV Grid", font=("", 8, "bold")).pack(anchor=tk.W, padx=4)
        cv_grid_row = ttk.Frame(left)
        cv_grid_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        self.cv_x_grid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(cv_grid_row, text="X", variable=self.cv_x_grid_var,
                        command=self._auto_replot).pack(side=tk.LEFT)
        self.cv_y_grid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(cv_grid_row, text="Y", variable=self.cv_y_grid_var,
                        command=self._auto_replot).pack(side=tk.LEFT, padx=(8, 0))
        cv_grid_style_row = ttk.Frame(left)
        cv_grid_style_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(cv_grid_style_row, text="Style:").pack(side=tk.LEFT)
        self.cv_grid_style_var = tk.StringVar(value="dashed")
        _cvgscb = ttk.Combobox(cv_grid_style_row, textvariable=self.cv_grid_style_var,
                               values=["dashed", "dotted", "solid", "dash-dot"],
                               state="readonly", width=9)
        _cvgscb.pack(side=tk.LEFT, padx=(2, 6))
        _cvgscb.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())
        ttk.Label(cv_grid_style_row, text="Color:").pack(side=tk.LEFT)
        self.cv_grid_color_var = tk.StringVar(value="gray")
        _cvgcol_cb = ttk.Combobox(cv_grid_style_row, textvariable=self.cv_grid_color_var,
                                   values=["gray", "black", "red", "blue", "green",
                                           "orange", "purple", "crimson", "royalblue",
                                           "darkorange", "teal"],
                                   state="readonly", width=9)
        _cvgcol_cb.pack(side=tk.LEFT, padx=(2, 6))
        _cvgcol_cb.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())
        ttk.Label(cv_grid_style_row, text="Width:").pack(side=tk.LEFT)
        self.cv_grid_linewidth_var = tk.StringVar(value="0.8")
        _cvglw = ttk.Entry(cv_grid_style_row, textvariable=self.cv_grid_linewidth_var, width=4)
        _cvglw.pack(side=tk.LEFT, padx=(2, 0))
        _cvglw.bind("<Return>",   lambda e: self._auto_replot())
        _cvglw.bind("<FocusOut>", lambda e: self._auto_replot())

        # CV Reference Lines
        ttk.Label(left, text="CV Ref Lines", font=("", 8, "bold")).pack(anchor=tk.W, padx=4)
        cv_ref_add_row = ttk.Frame(left)
        cv_ref_add_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Label(cv_ref_add_row, text="X:").pack(side=tk.LEFT)
        self._cv_ref_x_var = tk.StringVar()
        _cvref_x_e = ttk.Entry(cv_ref_add_row, textvariable=self._cv_ref_x_var, width=7)
        _cvref_x_e.pack(side=tk.LEFT, padx=(2, 0))
        ttk.Button(cv_ref_add_row, text="+X", width=3,
                   command=self._add_cv_xrefline).pack(side=tk.LEFT, padx=(2, 8))
        ttk.Label(cv_ref_add_row, text="Y:").pack(side=tk.LEFT)
        self._cv_ref_y_var = tk.StringVar()
        _cvref_y_e = ttk.Entry(cv_ref_add_row, textvariable=self._cv_ref_y_var, width=7)
        _cvref_y_e.pack(side=tk.LEFT, padx=(2, 0))
        ttk.Button(cv_ref_add_row, text="+Y", width=3,
                   command=self._add_cv_yrefline).pack(side=tk.LEFT, padx=2)
        _cvref_x_e.bind("<Return>", lambda e: self._add_cv_xrefline())
        _cvref_y_e.bind("<Return>", lambda e: self._add_cv_yrefline())
        cv_ref_list_row = ttk.Frame(left)
        cv_ref_list_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        self._cv_reflines_lb = tk.Listbox(cv_ref_list_row, height=3,
                                           selectmode=tk.SINGLE, exportselection=False)
        self._cv_reflines_lb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._cv_reflines_lb.bind("<<ListboxSelect>>", lambda e: self._on_cv_refline_select())
        ttk.Button(cv_ref_list_row, text="Remove",
                   command=self._remove_cv_refline).pack(side=tk.RIGHT, padx=(4, 0))
        cv_ref_opt_row = ttk.Frame(left)
        cv_ref_opt_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(cv_ref_opt_row, text="Style:").pack(side=tk.LEFT)
        self._cv_refline_style_var = tk.StringVar(value="dashed")
        _cvrl_style_cb = ttk.Combobox(cv_ref_opt_row, textvariable=self._cv_refline_style_var,
                                       values=["dashed", "dotted", "solid", "dash-dot"],
                                       state="readonly", width=9)
        _cvrl_style_cb.pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(cv_ref_opt_row, text="Color:").pack(side=tk.LEFT)
        self._cv_refline_color_var = tk.StringVar(value="dimgray")
        _cvrl_color_cb = ttk.Combobox(cv_ref_opt_row, textvariable=self._cv_refline_color_var,
                                       values=["dimgray", "black", "red", "blue", "green",
                                               "orange", "purple", "crimson", "royalblue",
                                               "darkorange", "teal", "saddlebrown"],
                                       state="readonly", width=9)
        _cvrl_color_cb.pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(cv_ref_opt_row, text="Width:").pack(side=tk.LEFT)
        self._cv_refline_linewidth_var = tk.StringVar(value="1.0")
        _cvrl_lw = ttk.Entry(cv_ref_opt_row, textvariable=self._cv_refline_linewidth_var, width=4)
        _cvrl_lw.pack(side=tk.LEFT, padx=(2, 0))
        _cvrl_style_cb.bind("<<ComboboxSelected>>", lambda e: self._on_cv_refline_style_color_change())
        _cvrl_color_cb.bind("<<ComboboxSelected>>", lambda e: self._on_cv_refline_style_color_change())
        _cvrl_lw.bind("<Return>",   lambda e: self._on_cv_refline_style_color_change())
        _cvrl_lw.bind("<FocusOut>", lambda e: self._on_cv_refline_style_color_change())

        # Cdl Grid
        ttk.Label(left, text="Cdl Grid", font=("", 8, "bold")).pack(anchor=tk.W, padx=4)
        cdl_grid_row = ttk.Frame(left)
        cdl_grid_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        self.cdl_x_grid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(cdl_grid_row, text="X", variable=self.cdl_x_grid_var,
                        command=self._replot_active_cdl).pack(side=tk.LEFT)
        ttk.Label(cdl_grid_row, text="Interval:").pack(side=tk.LEFT, padx=(6, 0))
        self.cdl_x_grid_int_var = tk.StringVar(value="0")
        _cdlxgi = ttk.Entry(cdl_grid_row, textvariable=self.cdl_x_grid_int_var, width=5)
        _cdlxgi.pack(side=tk.LEFT, padx=(2, 0))
        self.cdl_y_grid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(cdl_grid_row, text="Y", variable=self.cdl_y_grid_var,
                        command=self._replot_active_cdl).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Label(cdl_grid_row, text="Interval:").pack(side=tk.LEFT, padx=(6, 0))
        self.cdl_y_grid_int_var = tk.StringVar(value="0")
        _cdlygi = ttk.Entry(cdl_grid_row, textvariable=self.cdl_y_grid_int_var, width=5)
        _cdlygi.pack(side=tk.LEFT, padx=(2, 0))
        for _gi in (_cdlxgi, _cdlygi):
            _gi.bind("<Return>",   lambda e: self._replot_active_cdl())
            _gi.bind("<FocusOut>", lambda e: self._replot_active_cdl())
        cdl_grid_style_row = ttk.Frame(left)
        cdl_grid_style_row.pack(fill=tk.X, padx=4, pady=(0, 4))
        ttk.Label(cdl_grid_style_row, text="Style:").pack(side=tk.LEFT)
        self.cdl_grid_style_var = tk.StringVar(value="dashed")
        _cdlgscb = ttk.Combobox(cdl_grid_style_row, textvariable=self.cdl_grid_style_var,
                                values=["dashed", "dotted", "solid", "dash-dot"],
                                state="readonly", width=9)
        _cdlgscb.pack(side=tk.LEFT, padx=(2, 6))
        _cdlgscb.bind("<<ComboboxSelected>>", lambda e: self._replot_active_cdl())
        ttk.Label(cdl_grid_style_row, text="Color:").pack(side=tk.LEFT)
        self.cdl_grid_color_var = tk.StringVar(value="gray")
        _cdlgcol_cb = ttk.Combobox(cdl_grid_style_row, textvariable=self.cdl_grid_color_var,
                                    values=["gray", "black", "red", "blue", "green",
                                            "orange", "purple", "crimson", "royalblue",
                                            "darkorange", "teal"],
                                    state="readonly", width=9)
        _cdlgcol_cb.pack(side=tk.LEFT, padx=(2, 6))
        _cdlgcol_cb.bind("<<ComboboxSelected>>", lambda e: self._replot_active_cdl())
        ttk.Label(cdl_grid_style_row, text="Width:").pack(side=tk.LEFT)
        self.cdl_grid_linewidth_var = tk.StringVar(value="0.8")
        _cdlglw = ttk.Entry(cdl_grid_style_row, textvariable=self.cdl_grid_linewidth_var, width=4)
        _cdlglw.pack(side=tk.LEFT, padx=(2, 0))
        _cdlglw.bind("<Return>",   lambda e: self._replot_active_cdl())
        _cdlglw.bind("<FocusOut>", lambda e: self._replot_active_cdl())

        # Cdl Reference Lines
        ttk.Label(left, text="Cdl Ref Lines", font=("", 8, "bold")).pack(anchor=tk.W, padx=4)
        cdl_ref_add_row = ttk.Frame(left)
        cdl_ref_add_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Label(cdl_ref_add_row, text="X:").pack(side=tk.LEFT)
        self._cdl_ref_x_var = tk.StringVar()
        _cdlref_x_e = ttk.Entry(cdl_ref_add_row, textvariable=self._cdl_ref_x_var, width=7)
        _cdlref_x_e.pack(side=tk.LEFT, padx=(2, 0))
        ttk.Button(cdl_ref_add_row, text="+X", width=3,
                   command=self._add_cdl_xrefline).pack(side=tk.LEFT, padx=(2, 8))
        ttk.Label(cdl_ref_add_row, text="Y:").pack(side=tk.LEFT)
        self._cdl_ref_y_var = tk.StringVar()
        _cdlref_y_e = ttk.Entry(cdl_ref_add_row, textvariable=self._cdl_ref_y_var, width=7)
        _cdlref_y_e.pack(side=tk.LEFT, padx=(2, 0))
        ttk.Button(cdl_ref_add_row, text="+Y", width=3,
                   command=self._add_cdl_yrefline).pack(side=tk.LEFT, padx=2)
        _cdlref_x_e.bind("<Return>", lambda e: self._add_cdl_xrefline())
        _cdlref_y_e.bind("<Return>", lambda e: self._add_cdl_yrefline())
        cdl_ref_list_row = ttk.Frame(left)
        cdl_ref_list_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        self._cdl_reflines_lb = tk.Listbox(cdl_ref_list_row, height=3,
                                            selectmode=tk.SINGLE, exportselection=False)
        self._cdl_reflines_lb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._cdl_reflines_lb.bind("<<ListboxSelect>>", lambda e: self._on_cdl_refline_select())
        ttk.Button(cdl_ref_list_row, text="Remove",
                   command=self._remove_cdl_refline).pack(side=tk.RIGHT, padx=(4, 0))
        cdl_ref_opt_row = ttk.Frame(left)
        cdl_ref_opt_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(cdl_ref_opt_row, text="Style:").pack(side=tk.LEFT)
        self._cdl_refline_style_var = tk.StringVar(value="dashed")
        _cdlrl_style_cb = ttk.Combobox(cdl_ref_opt_row, textvariable=self._cdl_refline_style_var,
                                        values=["dashed", "dotted", "solid", "dash-dot"],
                                        state="readonly", width=9)
        _cdlrl_style_cb.pack(side=tk.LEFT, padx=(2, 8))
        ttk.Label(cdl_ref_opt_row, text="Color:").pack(side=tk.LEFT)
        self._cdl_refline_color_var = tk.StringVar(value="dimgray")
        _cdlrl_color_cb = ttk.Combobox(cdl_ref_opt_row, textvariable=self._cdl_refline_color_var,
                                        values=["dimgray", "black", "red", "blue", "green",
                                                "orange", "purple", "crimson", "royalblue",
                                                "darkorange", "teal", "saddlebrown"],
                                        state="readonly", width=10)
        _cdlrl_color_cb.pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(cdl_ref_opt_row, text="Width:").pack(side=tk.LEFT)
        self._cdl_refline_linewidth_var = tk.StringVar(value="1.0")
        _cdlrl_lw = ttk.Entry(cdl_ref_opt_row, textvariable=self._cdl_refline_linewidth_var, width=4)
        _cdlrl_lw.pack(side=tk.LEFT, padx=(2, 0))
        _cdlrl_style_cb.bind("<<ComboboxSelected>>", lambda e: self._on_cdl_refline_style_color_change())
        _cdlrl_color_cb.bind("<<ComboboxSelected>>", lambda e: self._on_cdl_refline_style_color_change())
        _cdlrl_lw.bind("<Return>",   lambda e: self._on_cdl_refline_style_color_change())
        _cdlrl_lw.bind("<FocusOut>", lambda e: self._on_cdl_refline_style_color_change())

        # Font
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(left, text="Font", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        self.font_title_size_var = tk.StringVar(value="10")
        self.font_title_bold_var = tk.BooleanVar(value=False)
        self.font_label_size_var = tk.StringVar(value="10")
        self.font_label_bold_var = tk.BooleanVar(value=False)
        self.font_tick_size_var  = tk.StringVar(value="8")
        self.font_tick_bold_var  = tk.BooleanVar(value=False)
        _font_title_row = ttk.Frame(left)
        _font_title_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Label(_font_title_row, text="Title:      Size").pack(side=tk.LEFT)
        _fts_e = ttk.Entry(_font_title_row, textvariable=self.font_title_size_var, width=4)
        _fts_e.pack(side=tk.LEFT, padx=(2, 4))
        _fts_e.bind("<Return>",   lambda e: self._auto_replot())
        _fts_e.bind("<FocusOut>", lambda e: self._auto_replot())
        ttk.Checkbutton(_font_title_row, text="Bold", variable=self.font_title_bold_var,
                        command=self._auto_replot).pack(side=tk.LEFT)
        _font_label_row = ttk.Frame(left)
        _font_label_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Label(_font_label_row, text="Axis Lbl: Size").pack(side=tk.LEFT)
        _fls_e = ttk.Entry(_font_label_row, textvariable=self.font_label_size_var, width=4)
        _fls_e.pack(side=tk.LEFT, padx=(2, 4))
        _fls_e.bind("<Return>",   lambda e: self._auto_replot())
        _fls_e.bind("<FocusOut>", lambda e: self._auto_replot())
        ttk.Checkbutton(_font_label_row, text="Bold", variable=self.font_label_bold_var,
                        command=self._auto_replot).pack(side=tk.LEFT)
        _font_tick_row = ttk.Frame(left)
        _font_tick_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Label(_font_tick_row, text="Tick Nos: Size").pack(side=tk.LEFT)
        _fks_e = ttk.Entry(_font_tick_row, textvariable=self.font_tick_size_var, width=4)
        _fks_e.pack(side=tk.LEFT, padx=(2, 4))
        _fks_e.bind("<Return>",   lambda e: self._auto_replot())
        _fks_e.bind("<FocusOut>", lambda e: self._auto_replot())
        ttk.Checkbutton(_font_tick_row, text="Bold", variable=self.font_tick_bold_var,
                        command=self._auto_replot).pack(side=tk.LEFT)
        _spacing_row = ttk.Frame(left)
        _spacing_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Label(_spacing_row, text="Spacing (pt): Title").pack(side=tk.LEFT)
        self.title_pad_var = tk.StringVar(value="6")
        _tpad_e = ttk.Entry(_spacing_row, textvariable=self.title_pad_var, width=4)
        _tpad_e.pack(side=tk.LEFT, padx=(2, 6))
        _tpad_e.bind("<Return>",   lambda e: self._auto_replot())
        _tpad_e.bind("<FocusOut>", lambda e: self._auto_replot())
        ttk.Label(_spacing_row, text="Label").pack(side=tk.LEFT)
        self.label_pad_var = tk.StringVar(value="4")
        _lpad_e = ttk.Entry(_spacing_row, textvariable=self.label_pad_var, width=4)
        _lpad_e.pack(side=tk.LEFT, padx=(2, 0))
        _lpad_e.bind("<Return>",   lambda e: self._auto_replot())
        _lpad_e.bind("<FocusOut>", lambda e: self._auto_replot())

        self.result_label = ttk.Label(left, text="", wraplength=290, justify=tk.LEFT)
        self.result_label.pack(anchor=tk.W, padx=4, pady=2)

        # ── Log ───────────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Log", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        log_f = ttk.Frame(left)
        log_f.pack(fill=tk.X, padx=4, pady=2)
        self.log_text = tk.Text(log_f, height=6, state=tk.DISABLED,
                                wrap=tk.WORD, relief=tk.SUNKEN, borderwidth=1)
        log_sc = ttk.Scrollbar(log_f, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sc.set)
        log_sc.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # ── Right panel: two independent figures ──────────────────
        right = ttk.Frame(body)
        body.add(right, weight=1)

        # Upper: CV plot
        upper = ttk.Frame(right)
        upper.pack(fill=tk.BOTH, expand=True)
        self.fig_cv    = Figure(figsize=(6, 3.5), dpi=100, constrained_layout=True)
        self.ax_cv     = self.fig_cv.add_subplot(1, 1, 1)
        self.canvas_cv = FigureCanvasTkAgg(self.fig_cv, master=upper)
        self.canvas_cv.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        tb_cv = ttk.Frame(upper)
        tb_cv.pack(fill=tk.X)
        _panel = self
        class _CVToolbar(NavigationToolbar2Tk):
            def home(tb_self, *args):
                _panel._reset_cv_view()
        _cv_tb = _CVToolbar(self.canvas_cv, tb_cv, pack_toolbar=False)
        _cv_tb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        _cv_tb.update()
        tk.Button(
            tb_cv, text="Copy",
            command=lambda: copy_figure_to_clipboard(self.fig_cv),
            relief=tk.RAISED, borderwidth=1, padx=6,
        ).pack(side=tk.LEFT, padx=(4, 2), pady=1)

        # Separator between plots
        ttk.Separator(right, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=2)

        # Lower: Cdl extraction plot
        lower = ttk.Frame(right)
        lower.pack(fill=tk.BOTH, expand=True)
        self.fig_cdl    = Figure(figsize=(6, 3.5), dpi=100, constrained_layout=True)
        self.ax_cdl     = self.fig_cdl.add_subplot(1, 1, 1)
        self.canvas_cdl = FigureCanvasTkAgg(self.fig_cdl, master=lower)
        self.canvas_cdl.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        tb_cdl = ttk.Frame(lower)
        tb_cdl.pack(fill=tk.X)
        class _CdlToolbar(NavigationToolbar2Tk):
            def home(tb_self, *args):
                _panel._reset_cdl_view()
        _cdl_tb = _CdlToolbar(self.canvas_cdl, tb_cdl, pack_toolbar=False)
        _cdl_tb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        _cdl_tb.update()
        tk.Button(
            tb_cdl, text="Copy",
            command=lambda: copy_figure_to_clipboard(self.fig_cdl),
            relief=tk.RAISED, borderwidth=1, padx=6,
        ).pack(side=tk.LEFT, padx=(4, 2), pady=1)

        self._reset_cv_axes_labels()
        self._reset_cdl_axes_labels()
        self.canvas_cv.draw()
        self.canvas_cdl.draw()

        self._init_plot_interactions()

    # ── Axis label helpers ───────────────────────────────────────────
    def _reset_cv_axes_labels(self):
        self.ax_cv.set_title("CV Curves  (non-Faradaic region)")
        self.ax_cv.set_xlabel("Potential (V)")
        self.ax_cv.set_ylabel("Current (mA)")

    def _reset_cdl_axes_labels(self):
        self.ax_cdl.set_title("Cdl Extraction")
        self.ax_cdl.set_xlabel("Scan Rate  (mV/s)")
        self.ax_cdl.set_ylabel("Δj/2  (mA)")

    # ── Unit conversion helper ───────────────────────────────────────
    def _get_unit_scale(self, col, target_unit):
        """Return (scale_factor, display_label). Mirrors PlottingMixin logic."""
        if not target_unit or target_unit == "(auto)":
            if "/" in col:
                _cb, _cu = col.rsplit("/", 1)
                return 1.0, f"{_cb.strip()} ({_cu.strip()})"
            return 1.0, col
        _FACTORS = {
            "A": 1.0,  "mA": 1e-3, "µA": 1e-6, "nA": 1e-9,
            "V": 1.0,  "mV": 1e-3, "µV": 1e-6, "nV": 1e-9,
            "s": 1.0,  "ms": 1e-3, "µs": 1e-6,
            "min": 60.0, "h": 3600.0,
        }
        _DIMS = {
            "A": "I", "mA": "I", "µA": "I", "nA": "I",
            "V": "E", "mV": "E", "µV": "E", "nV": "E",
            "s": "t", "ms": "t", "µs": "t", "min": "t", "h": "t",
        }
        if "/" in col:
            col_base, src_unit = col.rsplit("/", 1)
            col_base  = col_base.strip()
            src_unit  = src_unit.strip()
        else:
            col_base  = col
            src_unit  = None
        display_label = f"{col_base} ({target_unit})"
        src_f = _FACTORS.get(src_unit)
        tgt_f = _FACTORS.get(target_unit)
        if (src_f is not None and tgt_f is not None
                and _DIMS.get(src_unit) == _DIMS.get(target_unit)):
            return src_f / tgt_f, display_label
        return 1.0, display_label

    # ════════════════════════════════════════════════════════════════
    # Interactive plot interactions
    # ════════════════════════════════════════════════════════════════
    def _init_plot_interactions(self):
        self._legend_cv         = None
        self._legend_cdl        = None
        self._leg_size_cv       = 7.0
        self._leg_size_cdl      = 7.0
        self._legend_labels_cv  = []    # persisted custom CV legend labels
        self._legend_manual_pos_cv = None  # saved dragged CV legend position

        # Stored auto-scaled limits for Home button restore
        self._auto_xlim_cv  = None
        self._auto_ylim_cv  = None
        self._auto_xlim_cdl = None
        self._auto_ylim_cdl = None

        self._panning   = False
        self._pan_ax    = None
        self._pan_start = None
        self._pan_moved = False

        self._leg_resizing    = False
        self._resize_leg      = None
        self._resize_is_cv    = True
        self._resize_start_y  = None
        self._resize_start_sz = None

        self._ann            = None
        self._ann_dot        = None
        self._ann_ax         = None
        self._last_click_pos = None
        self._cand_idx       = 0

        for cv in (self.canvas_cv, self.canvas_cdl):
            cv.mpl_connect("scroll_event",         self._ei_scroll)
            cv.mpl_connect("button_press_event",   self._ei_press)
            cv.mpl_connect("button_release_event", self._ei_release)
            cv.mpl_connect("motion_notify_event",  self._ei_motion)

    def _get_canvas(self, ax):
        return self.canvas_cv if ax is self.ax_cv else self.canvas_cdl

    # ── Legend hit-test ──────────────────────────────────────────────
    def _ei_leg_hit(self, event):
        try:
            r = event.canvas.get_renderer()
            for leg, is_cv in ((self._legend_cv, True), (self._legend_cdl, False)):
                if leg is None:
                    continue
                if leg.get_window_extent(r).contains(event.x, event.y):
                    return leg, is_cv
        except Exception:
            pass
        return None, None

    # ── Scroll = zoom ────────────────────────────────────────────────
    def _ei_scroll(self, event):
        ax = event.inaxes
        if ax not in (self.ax_cv, self.ax_cdl):
            return
        scale = 0.8 if event.step > 0 else 1.25
        xl, yl = ax.get_xlim(), ax.get_ylim()
        xd, yd = event.xdata, event.ydata
        xf = (xd - xl[0]) / (xl[1] - xl[0]) if xl[1] != xl[0] else 0.5
        yf = (yd - yl[0]) / (yl[1] - yl[0]) if yl[1] != yl[0] else 0.5
        nxr = (xl[1] - xl[0]) * scale
        nyr = (yl[1] - yl[0]) * scale
        ax.set_xlim(xd - nxr * xf,      xd + nxr * (1 - xf))
        ax.set_ylim(yd - nyr * yf,      yd + nyr * (1 - yf))
        self._get_canvas(ax).draw_idle()

    # ── Press ────────────────────────────────────────────────────────
    def _ei_press(self, event):
        # Handle dblclick on title strip (may be outside axes)
        if event.button == 1 and getattr(event, 'dblclick', False):
            ax = event.inaxes if event.inaxes in (self.ax_cv, self.ax_cdl) else None
            # Determine which axes the click is near based on y-position
            if ax is None:
                for candidate in (self.ax_cv, self.ax_cdl):
                    try:
                        canvas = (self.canvas_cv if candidate is self.ax_cv
                                  else self.canvas_cdl)
                        r       = canvas.get_renderer()
                        ax_bbox = candidate.get_window_extent(r)
                        if ax_bbox.x0 <= event.x <= ax_bbox.x1:
                            ax = candidate
                            break
                    except Exception:
                        pass
            if ax is not None:
                canvas = self.canvas_cv if ax is self.ax_cv else self.canvas_cdl
                try:
                    r        = canvas.get_renderer()
                    ax_bbox  = ax.get_window_extent(r)
                    fig_bbox = ax.get_figure().get_window_extent(r)
                    t_bbox   = ax.title.get_window_extent(r)
                    on_title = (
                        (t_bbox.width > 2 and t_bbox.contains(event.x, event.y))
                        or (ax_bbox.x0 <= event.x <= ax_bbox.x1
                            and ax_bbox.y1 <= event.y <= fig_bbox.y1)
                    )
                    if on_title:
                        self._edit_plot_title(ax, canvas)
                        return
                except Exception:
                    pass

        leg, is_cv = self._ei_leg_hit(event)
        if event.button == 1:
            self._pan_moved = False
            if leg is not None and getattr(event, 'dblclick', False):
                self._edit_legend_labels(leg, is_cv)
                return
            if leg is None and event.inaxes in (self.ax_cv, self.ax_cdl):
                if getattr(event, 'dblclick', False):
                    return   # dblclick not on title — ignore, don't start pan
                self._panning   = True
                self._pan_ax    = event.inaxes
                self._pan_start = (event.xdata, event.ydata)
        elif event.button == 3:
            if leg is not None:
                self._leg_resizing    = True
                self._resize_leg      = leg
                self._resize_is_cv    = is_cv
                self._resize_start_y  = event.y
                self._resize_start_sz = (self._leg_size_cv if is_cv else self._leg_size_cdl)

    # ── Release ──────────────────────────────────────────────────────
    def _ei_release(self, event):
        self._panning = False
        self._pan_ax  = None

        was_resizing = self._leg_resizing
        self._leg_resizing = False
        if was_resizing:
            return

        leg, _ = self._ei_leg_hit(event)
        if (event.button == 1
                and not self._pan_moved
                and event.inaxes in (self.ax_cv, self.ax_cdl)
                and leg is None):
            self._ei_annotate(event)
        elif event.button == 3 and leg is None:
            self._ei_clear_ann()

    # ── Motion ───────────────────────────────────────────────────────
    def _ei_motion(self, event):
        if self._panning and self._pan_ax is not None:
            if event.inaxes != self._pan_ax or event.xdata is None:
                return
            self._pan_moved = True
            dx = self._pan_start[0] - event.xdata
            dy = self._pan_start[1] - event.ydata
            xl = self._pan_ax.get_xlim()
            yl = self._pan_ax.get_ylim()
            self._pan_ax.set_xlim(xl[0] + dx, xl[1] + dx)
            self._pan_ax.set_ylim(yl[0] + dy, yl[1] + dy)
            self._get_canvas(self._pan_ax).draw_idle()
            return

        if self._leg_resizing and self._resize_leg is not None:
            dy     = event.y - self._resize_start_y
            new_sz = max(4.0, min(30.0, self._resize_start_sz + dy / 5.0))
            if self._resize_is_cv:
                self._leg_size_cv  = new_sz
            else:
                self._leg_size_cdl = new_sz
            for t in self._resize_leg.get_texts():
                t.set_fontsize(new_sz)
            tt = self._resize_leg.get_title()
            if tt:
                tt.set_fontsize(new_sz)
            canvas = self.canvas_cv if self._resize_is_cv else self.canvas_cdl
            canvas.draw()

    # ── Click annotate ───────────────────────────────────────────────
    def _ei_annotate(self, event):
        ax    = event.inaxes
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
                and abs(event.x - self._last_click_pos[0]) <= _CLICK_CYCLE_PX
                and abs(event.y - self._last_click_pos[1]) <= _CLICK_CYCLE_PX):
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
        self._ei_clear_ann(redraw=False)
        self._ann_ax = ax
        self._ann    = ax.annotate(
            text, xy=(x, y), xytext=(xoff, yoff), textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.4", fc="lightyellow", ec="gray", alpha=0.92),
            arrowprops=dict(arrowstyle="->", color="gray", lw=1.2),
            fontsize=8, zorder=10,
        )
        self._ann_dot, = ax.plot(x, y, "o", color=ln.get_color(),
                                  markersize=7, zorder=11, label="_ann_dot")
        self._get_canvas(ax).draw_idle()

    def _ei_clear_ann(self, redraw=True):
        canvas = self._get_canvas(self._ann_ax) if self._ann_ax is not None else None
        for artist in (self._ann, self._ann_dot):
            if artist is not None:
                try:
                    artist.remove()
                except Exception:
                    pass
        self._ann = self._ann_dot = None
        self._last_click_pos = None
        self._cand_idx       = 0
        self._ann_ax         = None
        if redraw and canvas is not None:
            canvas.draw_idle()

    # ════════════════════════════════════════════════════════════════
    # FileManagerMixin overrides
    # ════════════════════════════════════════════════════════════════
    def _on_file_reorder(self, new_order):
        """Rebuild self.files in the dragged order (no replot needed — ECSA shows one file)."""
        new_files = OrderedDict()
        for name in new_order:
            if name in self.files:
                new_files[name] = self.files[name]
        for name, entry in self.files.items():
            if name not in new_files:
                new_files[name] = entry
        self.files = new_files

    def _clear_plot(self):
        """Clear both plots when all files are removed."""
        self._ei_clear_ann(redraw=False)
        self._legend_cv = None
        self._legend_cdl = None
        self.ax_cv.clear()
        self.ax_cdl.clear()
        self._reset_cv_axes_labels()
        self._reset_cdl_axes_labels()
        self.canvas_cv.draw()
        self.canvas_cdl.draw()
        self.result_label.config(text="")

    def _on_file_color_change(self, event=None):
        if not self.active_file:
            return
        self.files[self.active_file]["color"] = _COLOR_HEX.get(
            self.file_color_var.get(), "#1f77b4")
        self._auto_replot()

    def _on_linewidth_change(self):
        if self.active_file and self.active_file in self.files:
            self.files[self.active_file]["linewidth"] = self.linewidth_var.get()
        self._auto_replot()

    def _on_plot_style_change(self):
        if self.active_file and self.active_file in self.files:
            self.files[self.active_file]["plot_style"] = self.plot_style_var.get()
        self._auto_replot()

    def _on_gradient_change(self):
        """Persist gradient settings to the active file's entry, then replot."""
        if self.active_file and self.active_file in self.files:
            self.files[self.active_file]["cycle_gradient"] = self.cycle_gradient_var.get()
            self.files[self.active_file]["cycle_reverse"]  = self.cycle_reverse_var.get()
            self.files[self.active_file]["lightness_step"] = self.lightness_step_var.get()
        self._auto_replot()

    def _save_active_state(self):
        """Extend base save to include all ECSA-panel per-file state."""
        super()._save_active_state()   # saves selected_cycles, r_sol, e_ref
        if self.active_file and self.active_file in self.files:
            entry = self.files[self.active_file]
            entry["sr_data"]        = {c: var.get() for c, var in self._sr_vars.items()}
            entry["e_std"]          = self.e_std_var.get()
            entry["cs"]             = self.cs_var.get()
            entry["x_col"]          = self.x_var.get()
            entry["y_col"]          = self.y_var.get()
            entry["x_unit"]         = self.x_unit_var.get()
            entry["y_unit"]         = self.y_unit_var.get()
            entry["x_min"]          = self.x_min_var.get()
            entry["x_max"]          = self.x_max_var.get()
            entry["y_min"]          = self.y_min_var.get()
            entry["y_max"]          = self.y_max_var.get()
            entry["ref_electrode"]  = self.ref_electrode_var.get()
            entry["legend_frame_cv"]  = self.legend_frame_cv_var.get()
            entry["legend_frame_cdl"] = self.legend_frame_cdl_var.get()
            entry["result_text"]    = self.result_label.cget("text")
            entry["cv_x_grid"]      = self.cv_x_grid_var.get()
            entry["cv_y_grid"]      = self.cv_y_grid_var.get()
            entry["cv_x_grid_int"]  = self.cv_x_grid_int_var.get()
            entry["cv_y_grid_int"]  = self.cv_y_grid_int_var.get()
            entry["cv_grid_style"]  = self.cv_grid_style_var.get()
            entry["cdl_x_grid"]     = self.cdl_x_grid_var.get()
            entry["cdl_y_grid"]     = self.cdl_y_grid_var.get()
            entry["cdl_x_grid_int"] = self.cdl_x_grid_int_var.get()
            entry["cdl_y_grid_int"] = self.cdl_y_grid_int_var.get()
            entry["cdl_grid_style"]    = self.cdl_grid_style_var.get()
            entry["cv_grid_color"]      = self.cv_grid_color_var.get()
            entry["cv_grid_linewidth"]  = self.cv_grid_linewidth_var.get()
            entry["cdl_grid_color"]     = self.cdl_grid_color_var.get()
            entry["cdl_grid_linewidth"] = self.cdl_grid_linewidth_var.get()
            entry["cycle_gradient"] = self.cycle_gradient_var.get()
            entry["cycle_reverse"]  = self.cycle_reverse_var.get()
            entry["lightness_step"] = self.lightness_step_var.get()
            entry["linewidth"]      = self.linewidth_var.get()
            entry["plot_style"]     = self.plot_style_var.get()
            # Preserve current zoom/pan for both plots
            entry["view_xlim_cv"]  = self.ax_cv.get_xlim()
            entry["view_ylim_cv"]  = self.ax_cv.get_ylim()
            entry["view_xlim_cdl"] = self.ax_cdl.get_xlim()
            entry["view_ylim_cdl"] = self.ax_cdl.get_ylim()

    def _switch_active_file(self, short):
        self.active_file = short
        entry = self.files[short]

        # Initialise per-file fields that may be absent on first visit
        entry.setdefault("sr_data",         {})
        entry.setdefault("e_std",           "")
        entry.setdefault("cs",              "0.040")
        entry.setdefault("x_col",           None)
        entry.setdefault("y_col",           None)
        entry.setdefault("x_unit",          "V")
        entry.setdefault("y_unit",          "mA")
        entry.setdefault("x_min",           "")
        entry.setdefault("x_max",           "")
        entry.setdefault("y_min",           "")
        entry.setdefault("y_max",           "")
        entry.setdefault("ref_electrode",   "Ag/AgCl")
        entry.setdefault("legend_frame_cv",  True)
        entry.setdefault("legend_frame_cdl", True)
        entry.setdefault("cdl_data",        None)
        entry.setdefault("result_text",     "")
        entry.setdefault("cv_x_grid",       False)
        entry.setdefault("cv_y_grid",       False)
        entry.setdefault("cv_x_grid_int",   "0")
        entry.setdefault("cv_y_grid_int",   "0")
        entry.setdefault("cv_grid_style",   "dashed")
        entry.setdefault("cdl_x_grid",      False)
        entry.setdefault("cdl_y_grid",      False)
        entry.setdefault("cdl_x_grid_int",  "0")
        entry.setdefault("cdl_y_grid_int",  "0")
        entry.setdefault("cdl_grid_style",  "dashed")
        entry.setdefault("cv_grid_color",      "gray")
        entry.setdefault("cv_grid_linewidth",  "0.8")
        entry.setdefault("cdl_grid_color",     "gray")
        entry.setdefault("cdl_grid_linewidth", "0.8")
        entry.setdefault("cv_reflines",      [])
        entry.setdefault("cdl_reflines",     [])
        entry.setdefault("cycle_gradient", True)
        entry.setdefault("cycle_reverse",  True)
        entry.setdefault("lightness_step", "0.08")
        entry.setdefault("linewidth",      "1.5")

        df   = entry["df"]
        cols = list(df.columns)

        self.x_combo["values"] = cols
        self.y_combo["values"] = cols

        # Restore saved column selection, or auto-detect on first visit
        x_col = entry["x_col"]
        if x_col and x_col in cols:
            self.x_var.set(x_col)
        else:
            self.x_var.set(_default_xcol(cols))

        y_col = entry["y_col"]
        if y_col and y_col in cols:
            self.y_var.set(y_col)
        else:
            self.y_var.set(_default_ycol(cols, self.x_var.get()))

        self.x_unit_var.set(entry["x_unit"])
        self.y_unit_var.set(entry["y_unit"])
        self.x_min_var.set(entry["x_min"])
        self.x_max_var.set(entry["x_max"])
        self.y_min_var.set(entry["y_min"])
        self.y_max_var.set(entry["y_max"])
        self.ref_electrode_var.set(entry["ref_electrode"])
        self.legend_frame_cv_var.set(entry["legend_frame_cv"])
        self.legend_frame_cdl_var.set(entry["legend_frame_cdl"])
        self.r_sol_var.set(str(entry["r_sol"]))
        self.e_ref_var.set(str(entry["e_ref"]))
        self.e_std_var.set(entry["e_std"])
        self.cs_var.set(entry["cs"])
        self.cv_x_grid_var.set(entry["cv_x_grid"])
        self.cv_y_grid_var.set(entry["cv_y_grid"])
        self.cv_x_grid_int_var.set(entry["cv_x_grid_int"])
        self.cv_y_grid_int_var.set(entry["cv_y_grid_int"])
        self.cv_grid_style_var.set(entry["cv_grid_style"])
        self.cdl_x_grid_var.set(entry["cdl_x_grid"])
        self.cdl_y_grid_var.set(entry["cdl_y_grid"])
        self.cdl_x_grid_int_var.set(entry["cdl_x_grid_int"])
        self.cdl_y_grid_int_var.set(entry["cdl_y_grid_int"])
        self.cdl_grid_style_var.set(entry["cdl_grid_style"])
        self.cv_grid_color_var.set(entry["cv_grid_color"])
        self.cv_grid_linewidth_var.set(entry["cv_grid_linewidth"])
        self.cdl_grid_color_var.set(entry["cdl_grid_color"])
        self.cdl_grid_linewidth_var.set(entry["cdl_grid_linewidth"])

        # Clear per-panel scan-rate vars so the new file starts completely fresh
        self._sr_vars.clear()

        # Populate cycle checkboxes + scan-rate table, with sr restore all inside suppress
        old = self._suppress_replot
        self._suppress_replot = True
        if "cycle number" in df.columns:
            cycles = sorted(int(c) for c in df["cycle number"].unique())
            self._populate_cycle_checkboxes(cycles, entry["selected_cycles"])
        else:
            self._populate_cycle_checkboxes([], [])

        # Restore saved scan-rate values into the freshly created StringVars
        for c, val in entry["sr_data"].items():
            if c in self._sr_vars:
                self._sr_vars[c].set(val)
        self._suppress_replot = old

        # Restore Cdl plot or show empty placeholder
        self._legend_cdl = None
        self.ax_cdl.clear()
        if entry["cdl_data"] is not None:
            self._replot_cdl(entry["cdl_data"])
        else:
            self._reset_cdl_axes_labels()
            self.canvas_cdl.draw()

        # Restore Cdl zoom/pan if the user had previously modified the view
        if "view_xlim_cdl" in entry:
            self.ax_cdl.set_xlim(entry["view_xlim_cdl"])
            self.ax_cdl.set_ylim(entry["view_ylim_cdl"])
            self.canvas_cdl.draw_idle()

        # Restore result label
        self.result_label.config(text=entry["result_text"])

        # Call _plot_cv() directly (not _auto_replot) so the CV plot always
        # redraws when switching files, even when _suppress_replot is True
        # (e.g. during _load_files).  The Cdl plot is handled the same way above.
        self._plot_cv()

        # Restore CV zoom/pan if the user had previously modified the view
        if "view_xlim_cv" in entry:
            self.ax_cv.set_xlim(entry["view_xlim_cv"])
            self.ax_cv.set_ylim(entry["view_ylim_cv"])
            self.canvas_cv.draw_idle()

        self._refresh_cv_reflines_lb()
        self._refresh_cdl_reflines_lb()

        # Restore color combobox and gradient controls to match this file's stored settings
        color = entry.get("color", "#1f77b4")
        name = next((n for n, h in _COLOR_HEX.items() if h == color), "Blue")
        self.file_color_var.set(name)
        self.cycle_gradient_var.set(entry.get("cycle_gradient", True))
        self.cycle_reverse_var.set(entry.get("cycle_reverse", True))
        self.lightness_step_var.set(entry.get("lightness_step", "0.08"))
        self.linewidth_var.set(entry.get("linewidth", "1.5"))
        self.plot_style_var.set(entry.get("plot_style", "Line"))

    def _auto_replot(self):
        if self._suppress_replot:
            return
        if not self.files or not self.x_var.get() or not self.y_var.get():
            return
        self._plot_cv()

    def _plot(self):
        self._plot_cv()

    # ════════════════════════════════════════════════════════════════
    # Font helpers
    # ════════════════════════════════════════════════════════════════
    def _read_font(self):
        def _f(v, d):
            try:
                return float(v.get())
            except Exception:
                return d
        return (
            _f(self.font_title_size_var, 10.0),
            'bold' if self.font_title_bold_var.get() else 'normal',
            _f(self.font_label_size_var, 10.0),
            'bold' if self.font_label_bold_var.get() else 'normal',
            _f(self.font_tick_size_var,  8.0),
            self.font_tick_bold_var.get(),
        )

    def _apply_font_to_ax(self, ax, canvas):
        ts, tb, ls, lb, ks, kb = self._read_font()
        try: title_pad = float(self.title_pad_var.get())
        except Exception: title_pad = 6.0
        try: label_pad = float(self.label_pad_var.get())
        except Exception: label_pad = 4.0
        ax.set_title(ax.get_title(),   fontsize=ts, fontweight=tb, pad=title_pad)
        ax.set_xlabel(ax.get_xlabel(), fontsize=ls, fontweight=lb, labelpad=label_pad)
        ax.set_ylabel(ax.get_ylabel(), fontsize=ls, fontweight=lb, labelpad=label_pad)
        ax.tick_params(axis='both', labelsize=ks)
        ax.figure.tight_layout()
        canvas.draw()
        if kb:
            for lbl in ax.get_xticklabels() + ax.get_yticklabels():
                lbl.set_fontweight('bold')
            ax.figure.tight_layout()
            canvas.draw()

    # ════════════════════════════════════════════════════════════════
    # CV plot (upper figure)
    # ════════════════════════════════════════════════════════════════
    def _plot_cv(self):
        if not self.active_file:
            return
        xcol = self.x_var.get()
        ycol = self.y_var.get()
        if not xcol or not ycol:
            return

        df       = self.files[self.active_file]["df"]
        selected = self._selected_cycles()

        # Update recommended E_std (midpoint of actual plotted data range)
        self._update_e_std_rec()

        # Apply unit scaling for display
        x_scale, x_label = self._get_unit_scale(xcol, self.x_unit_var.get())
        y_scale, y_label = self._get_unit_scale(ycol, self.y_unit_var.get())

        self._ei_clear_ann(redraw=False)
        # Save dragged CV legend position before clearing
        if self._legend_cv is not None:
            _loc = getattr(self._legend_cv, '_loc', None)
            if isinstance(_loc, (tuple, list)):
                self._legend_manual_pos_cv = tuple(_loc)
        self._legend_cv = None
        self.ax_cv.clear()

        entry_ref = self.files[self.active_file]
        base_color = entry_ref.get("color", "#1f77b4")
        _grad  = entry_ref.get("cycle_gradient", True)
        _rev   = entry_ref.get("cycle_reverse",  True)
        try:    _step = float(entry_ref.get("lightness_step", "0.08"))
        except: _step = 0.08

        try:
            _cv_lw = float(self.files[self.active_file].get("linewidth", "1.5"))
        except (ValueError, TypeError):
            _cv_lw = 1.5
        _ls, _mk, _ms = _PLOT_STYLES.get(
            self.files[self.active_file].get("plot_style", "Line"), ("-", "", 0))

        if "cycle number" in df.columns:
            if not selected:
                # No cycles selected → show placeholder only
                self._reset_cv_axes_labels()
                self.canvas_cv.draw()
                return
            cycle_cols = (_cycle_colors(base_color, len(selected), _step, _rev)
                          if _grad else [base_color] * len(selected))
            for i, c in enumerate(selected):
                sub = df[df["cycle number"] == c]
                sr  = self._sr_vars.get(c, tk.StringVar()).get().strip()
                lbl = f"C{c}" + (f"  ({sr} mV/s)" if sr else "")
                self.ax_cv.plot(sub[xcol] * x_scale, sub[ycol] * y_scale,
                                color=cycle_cols[i], label=lbl, linewidth=_cv_lw,
                                linestyle=_ls, marker=_mk or None,
                                markersize=_ms if _mk else 0)
        else:
            self.ax_cv.plot(df[xcol] * x_scale, df[ycol] * y_scale,
                            color=base_color, linewidth=_cv_lw,
                            linestyle=_ls, marker=_mk or None,
                            markersize=_ms if _mk else 0)

        ref = self.ref_electrode_var.get().strip()
        _x_src = xcol.rsplit("/", 1)[-1].strip() if "/" in xcol else ""
        _x_unit_str = self.x_unit_var.get()
        _x_is_V = (_x_unit_str in _VOLTAGE_UNITS if _x_unit_str != "(auto)"
                   else _x_src in _VOLTAGE_UNITS)
        self.ax_cv.set_xlabel(f"{x_label}  (vs {ref})" if (ref and _x_is_V) else x_label)

        _y_src = ycol.rsplit("/", 1)[-1].strip() if "/" in ycol else ""
        _y_unit_str = self.y_unit_var.get()
        _y_is_V = (_y_unit_str in _VOLTAGE_UNITS if _y_unit_str != "(auto)"
                   else _y_src in _VOLTAGE_UNITS)
        self.ax_cv.set_ylabel(f"{y_label}  (vs {ref})" if (ref and _y_is_V) else y_label)
        self.ax_cv.set_title(f"{self.active_file}  —  CV Curves  (non-Faradaic region)")

        # Red dashed line at E_std (in display units)
        try:
            e_std_raw = float(self.e_std_var.get())
            e_std_disp = e_std_raw * x_scale
            self.ax_cv.axvline(e_std_disp, color="red", linestyle="--",
                               linewidth=1.2, label=f"E_std = {e_std_raw:.3f} V")
        except ValueError:
            pass

        if self.ax_cv.get_lines():
            self._legend_cv = self.ax_cv.legend(fontsize=self._leg_size_cv)
            self._legend_cv.set_draggable(True)
            self._legend_cv.get_frame().set_visible(self.legend_frame_cv_var.get())
            # Restore custom CV legend labels if count matches
            _cv_custom = getattr(self, '_legend_labels_cv', [])
            if _cv_custom:
                for text_obj, lbl in zip(self._legend_cv.get_texts(), _cv_custom):
                    if lbl:
                        text_obj.set_text(lbl)
            # Restore dragged position
            if self._legend_manual_pos_cv is not None:
                self._legend_cv._loc = self._legend_manual_pos_cv

        if self.active_file and self.active_file in self.files:
            draw_reflines(self.ax_cv,
                          self.files[self.active_file].get("cv_reflines", []))

        apply_grid(self.ax_cv,
                   self.cv_x_grid_var.get(), self.cv_y_grid_var.get(),
                   self.cv_x_grid_int_var.get(), self.cv_y_grid_int_var.get(),
                   self.cv_grid_style_var.get(),
                   linewidth=self.cv_grid_linewidth_var.get(),
                   color=self.cv_grid_color_var.get())
        self._apply_font_to_ax(self.ax_cv, self.canvas_cv)
        self._auto_xlim_cv = self.ax_cv.get_xlim()
        self._auto_ylim_cv = self.ax_cv.get_ylim()
        self._apply_cv_range()

    def _toggle_cv_legend_frame(self):
        if self._legend_cv is not None:
            self._legend_cv.get_frame().set_visible(self.legend_frame_cv_var.get())
            self.canvas_cv.draw()

    def _toggle_cdl_legend_frame(self):
        if self._legend_cdl is not None:
            self._legend_cdl.get_frame().set_visible(self.legend_frame_cdl_var.get())
            self.canvas_cdl.draw()

    def _reset_cv_view(self):
        """Restore CV plot to the auto-scaled limits from the last draw."""
        if self._auto_xlim_cv is not None:
            self.ax_cv.set_xlim(self._auto_xlim_cv)
            self.ax_cv.set_ylim(self._auto_ylim_cv)
            self.canvas_cv.draw_idle()

    def _reset_cdl_view(self):
        """Restore Cdl plot to the auto-scaled limits from the last draw."""
        if self._auto_xlim_cdl is not None:
            self.ax_cdl.set_xlim(self._auto_xlim_cdl)
            self.ax_cdl.set_ylim(self._auto_ylim_cdl)
            self.canvas_cdl.draw_idle()

    def _apply_cv_range(self):
        """Apply manual axis limits from range entries; no-op if all blank."""
        changed = False
        try:
            self.ax_cv.set_xlim(left=float(self.x_min_var.get()))
            changed = True
        except ValueError:
            pass
        try:
            self.ax_cv.set_xlim(right=float(self.x_max_var.get()))
            changed = True
        except ValueError:
            pass
        try:
            self.ax_cv.set_ylim(bottom=float(self.y_min_var.get()))
            changed = True
        except ValueError:
            pass
        try:
            self.ax_cv.set_ylim(top=float(self.y_max_var.get()))
            changed = True
        except ValueError:
            pass

        # Flip axes if requested
        xl = self.ax_cv.get_xlim()
        if self.x_flip_var.get() != (xl[0] > xl[1]):
            self.ax_cv.set_xlim(xl[1], xl[0])
            changed = True
        yl = self.ax_cv.get_ylim()
        if self.y_flip_var.get() != (yl[0] > yl[1]):
            self.ax_cv.set_ylim(yl[1], yl[0])
            changed = True

        if changed:
            self.canvas_cv.draw_idle()

    def _replot_active_cdl(self):
        """Re-render the Cdl plot for the active file (grid change, no re-extraction)."""
        if not self.active_file or self.active_file not in self.files:
            return
        cdl_data = self.files[self.active_file].get("cdl_data")
        if cdl_data is None:
            return
        self._legend_cdl = None
        self.ax_cdl.clear()
        self._replot_cdl(cdl_data)

    def _replot_cdl(self, cdl_data):
        """Replot the Cdl extraction figure from stored per-file data."""
        sr_arr = cdl_data["sr_arr"]
        dj_arr = cdl_data["dj_arr"]
        coeffs = cdl_data["coeffs"]
        cdl_mF = cdl_data["cdl_mF"]
        r_sq   = cdl_data["r_sq"]
        y_unit = cdl_data["y_unit"]
        ecsa   = cdl_data.get("ecsa")
        slope, intercept = float(coeffs[0]), float(coeffs[1])
        sr_fit = np.linspace(0, sr_arr.max() * 1.1, 300)
        dj_fit = np.polyval(coeffs, sr_fit)

        _fit_label = (f"y = {slope:.4g}x + {intercept:.4g}\n"
                      f"Cdl = {cdl_mF:.4f} mF    R² = {r_sq:.4f}")
        if ecsa is not None:
            _fit_label += f"\nECSA = {ecsa:.2f} cm²"

        try:
            _cdl_lw = float(self.files[self.active_file].get("linewidth", "1.5"))
        except (ValueError, TypeError):
            _cdl_lw = 1.5

        self.ax_cdl.scatter(sr_arr, dj_arr, color="steelblue", zorder=5, label="Data")
        self.ax_cdl.plot(
            sr_fit, dj_fit, color="tomato", linewidth=_cdl_lw,
            label=_fit_label,
        )
        self.ax_cdl.axhline(0, color="gray", linewidth=0.5, linestyle="--")
        self.ax_cdl.set_xlabel("Scan Rate  (mV/s)")
        self.ax_cdl.set_ylabel(f"Δj/2  ({y_unit})")
        _cdl_title = (f"{self.active_file}  —  Cdl Extraction"
                      if self.active_file else "Cdl Extraction")
        self.ax_cdl.set_title(_cdl_title)
        self._legend_cdl = self.ax_cdl.legend(fontsize=self._leg_size_cdl)
        self._legend_cdl.set_draggable(True)
        self._legend_cdl.get_frame().set_visible(self.legend_frame_cdl_var.get())

        if self.active_file and self.active_file in self.files:
            draw_reflines(self.ax_cdl,
                          self.files[self.active_file].get("cdl_reflines", []))

        apply_grid(self.ax_cdl,
                   self.cdl_x_grid_var.get(), self.cdl_y_grid_var.get(),
                   self.cdl_x_grid_int_var.get(), self.cdl_y_grid_int_var.get(),
                   self.cdl_grid_style_var.get(),
                   linewidth=self.cdl_grid_linewidth_var.get(),
                   color=self.cdl_grid_color_var.get())
        self._apply_font_to_ax(self.ax_cdl, self.canvas_cdl)
        self._auto_xlim_cdl = self.ax_cdl.get_xlim()
        self._auto_ylim_cdl = self.ax_cdl.get_ylim()

    # ════════════════════════════════════════════════════════════════
    # Debounced CV redraw (triggered by scan rate edits)
    # ════════════════════════════════════════════════════════════════
    def _schedule_cv_redraw(self):
        if self._suppress_replot:
            return
        if self._cv_redraw_id is not None:
            try:
                self.after_cancel(self._cv_redraw_id)
            except Exception:
                pass
        self._cv_redraw_id = self.after(300, self._plot_cv)

    # ════════════════════════════════════════════════════════════════
    # Scan-rate table  (8-column grid, with live-update traces)
    # ════════════════════════════════════════════════════════════════
    def _rebuild_sr_table(self):
        # Remove stale traces before destroying widgets
        for c, (var, tid) in list(self._sr_traces.items()):
            try:
                var.trace_remove("write", tid)
            except Exception:
                pass
        self._sr_traces.clear()

        for w in self._sr_inner.winfo_children():
            w.destroy()

        selected = sorted(self._selected_cycles())
        ncols    = 8

        for i, c in enumerate(selected):
            if c not in self._sr_vars:
                self._sr_vars[c] = tk.StringVar()
            row_g = (i // ncols) * 2
            col_g = i % ncols
            ttk.Label(self._sr_inner, text=f"C{c}:").grid(
                row=row_g, column=col_g, padx=3, pady=(3, 0), sticky=tk.W)
            ttk.Entry(self._sr_inner, textvariable=self._sr_vars[c], width=5).grid(
                row=row_g + 1, column=col_g, padx=3, pady=(0, 3))
            # Live trace: schedule a debounced replot on every keystroke
            tid = self._sr_vars[c].trace_add("write",
                                              lambda *_: self._schedule_cv_redraw())
            self._sr_traces[c] = (self._sr_vars[c], tid)

        self._sr_inner.update_idletasks()
        self._sr_canvas.configure(scrollregion=self._sr_canvas.bbox("all"))

    # ════════════════════════════════════════════════════════════════
    # Cdl extraction & ECSA calculation
    # ════════════════════════════════════════════════════════════════
    def _extract_cdl_ecsa(self):
        if not self.active_file:
            self._log("No file loaded.")
            return

        xcol = self.x_var.get()
        ycol = self.y_var.get()
        if not xcol or not ycol:
            self._log("Select X and Y axes first.")
            return

        df = self.files[self.active_file]["df"]
        if "cycle number" not in df.columns:
            self._log("ERROR: 'cycle number' column not found.")
            return

        selected = sorted(self._selected_cycles())
        if len(selected) < 2:
            self._log("Select at least 2 cycles (one per scan rate).")
            return

        scan_rates = {}
        for c in selected:
            sr_str = self._sr_vars.get(c, tk.StringVar()).get().strip()
            try:
                scan_rates[c] = float(sr_str)
            except ValueError:
                self._log(f"C{c}: missing or invalid scan rate '{sr_str}'.")
                return

        try:
            e_std = float(self.e_std_var.get())
        except ValueError:
            self._log("Invalid E_std — enter a numeric value (V).")
            return

        try:
            cs = float(self.cs_var.get())
            if cs <= 0:
                raise ValueError
        except ValueError:
            self._log("Invalid Cs — must be > 0 (mF/cm²).")
            return

        self._log(f"\n── {self.active_file} ──")
        self._log(f"E_std = {e_std} V    Cs = {cs} mF/cm²")
        self._log(f"(Extraction uses raw column units: {xcol}, {ycol})")

        sr_list = []
        dj_list = []

        for c in selected:
            sub  = df[df["cycle number"] == c]
            ewe  = sub[xcol].values.astype(float)
            I    = sub[ycol].values.astype(float)

            if len(ewe) < 4:
                self._log(f"  C{c}: too few points — skipped.")
                continue

            emin, emax = ewe.min(), ewe.max()
            if not (emin <= e_std <= emax):
                self._log(f"  C{c}: E_std={e_std:.3f}V outside "
                          f"[{emin:.3f}, {emax:.3f}]V — skipped.")
                continue

            imax  = int(np.argmax(ewe))
            an_e, an_i = ewe[:imax + 1], I[:imax + 1]
            ca_e, ca_i = ewe[imax:],      I[imax:]

            try:
                idx = np.argsort(an_e)
                ja  = float(np.interp(e_std, an_e[idx], an_i[idx]))
            except Exception:
                self._log(f"  C{c}: anodic interpolation failed — skipped.")
                continue

            try:
                idx = np.argsort(ca_e)
                jc  = float(np.interp(e_std, ca_e[idx], ca_i[idx]))
            except Exception:
                self._log(f"  C{c}: cathodic interpolation failed — skipped.")
                continue

            dj = (ja - jc) / 2.0
            self._log(f"  C{c}  ν={scan_rates[c]} mV/s  "
                      f"ja={ja:.5f}  jc={jc:.5f}  Δj/2={dj:.5f}")
            sr_list.append(scan_rates[c])
            dj_list.append(dj)

        if len(sr_list) < 2:
            self._log("  Not enough valid cycles for linear fit.")
            return

        sr_arr           = np.array(sr_list)
        dj_arr           = np.array(dj_list)
        coeffs           = np.polyfit(sr_arr, dj_arr, 1)
        slope, intercept = float(coeffs[0]), float(coeffs[1])
        r_sq             = float(np.corrcoef(sr_arr, dj_arr)[0, 1] ** 2)

        # slope [mA/(mV/s)] = F  →  ×1000 → mF
        cdl_mF = slope * 1000.0
        ecsa   = cdl_mF / cs

        self._log(f"  Fit: Δj/2 = {slope:.5g}·ν + {intercept:.5g}")
        self._log(f"  Cdl = {cdl_mF:.4f} mF    R² = {r_sq:.4f}")
        self._log(f"  ECSA = {ecsa:.2f} cm²")

        self.result_label.config(
            text=(f"Cdl  = {cdl_mF:.4f} mF\n"
                  f"ECSA = {ecsa:.2f} cm²\n"
                  f"(Cs = {cs} mF/cm²,  R² = {r_sq:.4f})")
        )

        # ── Update lower plot ─────────────────────────────────────
        self._legend_cdl = None
        self.ax_cdl.clear()

        sr_fit = np.linspace(0, sr_arr.max() * 1.1, 300)
        dj_fit = np.polyval(coeffs, sr_fit)
        y_unit = ycol.split("/")[-1] if "/" in ycol else ycol

        try:
            _cdl_lw2 = float(self.files[self.active_file].get("linewidth", "1.5"))
        except (ValueError, TypeError):
            _cdl_lw2 = 1.5

        self.ax_cdl.scatter(sr_arr, dj_arr, color="steelblue", zorder=5, label="Data")
        self.ax_cdl.plot(
            sr_fit, dj_fit, color="tomato", linewidth=_cdl_lw2,
            label=(f"y = {slope:.4g}x + {intercept:.4g}\n"
                   f"Cdl = {cdl_mF:.4f} mF    R² = {r_sq:.4f}\n"
                   f"ECSA = {ecsa:.2f} cm²"),
        )
        self.ax_cdl.axhline(0, color="gray", linewidth=0.5, linestyle="--")
        self.ax_cdl.set_xlabel("Scan Rate  (mV/s)")
        self.ax_cdl.set_ylabel(f"Δj/2  ({y_unit})")
        self.ax_cdl.set_title(f"{self.active_file}  —  Cdl Extraction")

        self._legend_cdl = self.ax_cdl.legend(fontsize=self._leg_size_cdl)
        self._legend_cdl.set_draggable(True)
        self._legend_cdl.get_frame().set_visible(self.legend_frame_cdl_var.get())

        if self.active_file and self.active_file in self.files:
            draw_reflines(self.ax_cdl,
                          self.files[self.active_file].get("cdl_reflines", []))

        apply_grid(self.ax_cdl,
                   self.cdl_x_grid_var.get(), self.cdl_y_grid_var.get(),
                   self.cdl_x_grid_int_var.get(), self.cdl_y_grid_int_var.get(),
                   self.cdl_grid_style_var.get(),
                   linewidth=self.cdl_grid_linewidth_var.get(),
                   color=self.cdl_grid_color_var.get())
        self._apply_font_to_ax(self.ax_cdl, self.canvas_cdl)
        self._auto_xlim_cdl = self.ax_cdl.get_xlim()
        self._auto_ylim_cdl = self.ax_cdl.get_ylim()

        # Persist Cdl data and result text to the file entry for file-switch restore
        if self.active_file in self.files:
            self.files[self.active_file]["cdl_data"] = {
                "sr_arr": sr_arr,
                "dj_arr": dj_arr,
                "coeffs": coeffs,
                "cdl_mF": cdl_mF,
                "r_sq":   r_sq,
                "y_unit": y_unit,
                "ecsa":   ecsa,
                "cs":     cs,
            }
            self.files[self.active_file]["result_text"] = (
                f"Cdl  = {cdl_mF:.4f} mF\n"
                f"ECSA = {ecsa:.2f} cm²\n"
                f"(Cs = {cs} mF/cm²,  R² = {r_sq:.4f})"
            )

    # ════════════════════════════════════════════════════════════════
    # Reference line helpers (CV)
    # ════════════════════════════════════════════════════════════════
    def _add_cv_xrefline(self):
        if not self.active_file:
            return
        try:
            v = float(self._cv_ref_x_var.get())
        except ValueError:
            return
        style = self._cv_refline_style_var.get()
        color = self._cv_refline_color_var.get()
        lw    = self._cv_refline_linewidth_var.get()
        self.files[self.active_file].setdefault("cv_reflines", []).append(('x', v, style, color, lw))
        self._cv_reflines_lb.insert(tk.END, f"X = {v:.4g}")
        self._auto_replot()

    def _add_cv_yrefline(self):
        if not self.active_file:
            return
        try:
            v = float(self._cv_ref_y_var.get())
        except ValueError:
            return
        style = self._cv_refline_style_var.get()
        color = self._cv_refline_color_var.get()
        lw    = self._cv_refline_linewidth_var.get()
        self.files[self.active_file].setdefault("cv_reflines", []).append(('y', v, style, color, lw))
        self._cv_reflines_lb.insert(tk.END, f"Y = {v:.4g}")
        self._auto_replot()

    def _remove_cv_refline(self):
        sel = self._cv_reflines_lb.curselection()
        if not sel or not self.active_file:
            return
        idx = sel[0]
        self.files[self.active_file]["cv_reflines"].pop(idx)
        self._cv_reflines_lb.delete(idx)
        self._auto_replot()

    def _refresh_cv_reflines_lb(self):
        self._cv_reflines_lb.delete(0, tk.END)
        if not self.active_file:
            return
        for axis, val, *_ in self.files.get(self.active_file, {}).get("cv_reflines", []):
            self._cv_reflines_lb.insert(tk.END, f"{'X' if axis == 'x' else 'Y'} = {val:.4g}")

    def _on_cv_refline_select(self):
        """Populate CV style/color comboboxes from the selected line's settings."""
        sel = self._cv_reflines_lb.curselection()
        if not sel or not self.active_file:
            return
        reflines = self.files.get(self.active_file, {}).get("cv_reflines", [])
        if sel[0] >= len(reflines):
            return
        entry = reflines[sel[0]]
        self._cv_refline_style_var.set(entry[2])
        self._cv_refline_color_var.set(entry[3])
        self._cv_refline_linewidth_var.set(str(entry[4]) if len(entry) > 4 else "1.0")

    def _on_cv_refline_style_color_change(self):
        """Apply new style/color to the currently selected CV reference line."""
        sel = self._cv_reflines_lb.curselection()
        if not sel or not self.active_file:
            return
        reflines = self.files.get(self.active_file, {}).get("cv_reflines", [])
        idx = sel[0]
        if idx >= len(reflines):
            return
        axis, val = reflines[idx][:2]
        reflines[idx] = (axis, val,
                         self._cv_refline_style_var.get(),
                         self._cv_refline_color_var.get(),
                         self._cv_refline_linewidth_var.get())
        self._auto_replot()

    # ════════════════════════════════════════════════════════════════
    # Reference line helpers (Cdl)
    # ════════════════════════════════════════════════════════════════
    def _add_cdl_xrefline(self):
        if not self.active_file:
            return
        try:
            v = float(self._cdl_ref_x_var.get())
        except ValueError:
            return
        style = self._cdl_refline_style_var.get()
        color = self._cdl_refline_color_var.get()
        lw    = self._cdl_refline_linewidth_var.get()
        self.files[self.active_file].setdefault("cdl_reflines", []).append(('x', v, style, color, lw))
        self._cdl_reflines_lb.insert(tk.END, f"X = {v:.4g}")
        self._replot_active_cdl()

    def _add_cdl_yrefline(self):
        if not self.active_file:
            return
        try:
            v = float(self._cdl_ref_y_var.get())
        except ValueError:
            return
        style = self._cdl_refline_style_var.get()
        color = self._cdl_refline_color_var.get()
        lw    = self._cdl_refline_linewidth_var.get()
        self.files[self.active_file].setdefault("cdl_reflines", []).append(('y', v, style, color, lw))
        self._cdl_reflines_lb.insert(tk.END, f"Y = {v:.4g}")
        self._replot_active_cdl()

    def _remove_cdl_refline(self):
        sel = self._cdl_reflines_lb.curselection()
        if not sel or not self.active_file:
            return
        idx = sel[0]
        self.files[self.active_file]["cdl_reflines"].pop(idx)
        self._cdl_reflines_lb.delete(idx)
        self._replot_active_cdl()

    def _refresh_cdl_reflines_lb(self):
        self._cdl_reflines_lb.delete(0, tk.END)
        if not self.active_file:
            return
        for axis, val, *_ in self.files.get(self.active_file, {}).get("cdl_reflines", []):
            self._cdl_reflines_lb.insert(tk.END, f"{'X' if axis == 'x' else 'Y'} = {val:.4g}")

    def _on_cdl_refline_select(self):
        """Populate Cdl style/color comboboxes from the selected line's settings."""
        sel = self._cdl_reflines_lb.curselection()
        if not sel or not self.active_file:
            return
        reflines = self.files.get(self.active_file, {}).get("cdl_reflines", [])
        if sel[0] >= len(reflines):
            return
        entry = reflines[sel[0]]
        self._cdl_refline_style_var.set(entry[2])
        self._cdl_refline_color_var.set(entry[3])
        self._cdl_refline_linewidth_var.set(str(entry[4]) if len(entry) > 4 else "1.0")

    def _on_cdl_refline_style_color_change(self):
        """Apply new style/color to the currently selected Cdl reference line."""
        sel = self._cdl_reflines_lb.curselection()
        if not sel or not self.active_file:
            return
        reflines = self.files.get(self.active_file, {}).get("cdl_reflines", [])
        idx = sel[0]
        if idx >= len(reflines):
            return
        axis, val = reflines[idx][:2]
        reflines[idx] = (axis, val,
                         self._cdl_refline_style_var.get(),
                         self._cdl_refline_color_var.get(),
                         self._cdl_refline_linewidth_var.get())
        self._replot_active_cdl()

    # ════════════════════════════════════════════════════════════════
    # Log helper
    # ════════════════════════════════════════════════════════════════
    def _log(self, message: str):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    # ════════════════════════════════════════════════════════════════
    # Cycle helpers
    # ════════════════════════════════════════════════════════════════
    def _populate_cycle_checkboxes(self, cycles, selected):
        for w in self._cycle_inner.winfo_children():
            w.destroy()
        self._cycle_vars.clear()

        selected_set = set(selected)
        ncols        = 9

        def _wheel(e):
            self._cycle_canvas.yview_scroll(-1 * (e.delta // 120), "units")
            return "break"

        for i, c in enumerate(cycles):
            r, col = divmod(i, ncols)
            var = tk.BooleanVar(value=(c in selected_set))
            var.trace_add("write", self._on_cycle_toggle)
            cb = tk.Checkbutton(
                self._cycle_inner, text=f"C{c}", variable=var,
                background=_CYCLE_BG, activebackground=_CYCLE_ACTIVE_BG,
                selectcolor=_CYCLE_BG, anchor=tk.W,
            )
            cb.grid(row=r, column=col, sticky=tk.W, padx=2, pady=1)
            cb.bind("<MouseWheel>", _wheel)
            self._cycle_vars[c] = var

        self._cycle_inner.update_idletasks()
        self._cycle_canvas.configure(scrollregion=self._cycle_canvas.bbox("all"))
        self._cycle_canvas.yview_moveto(0)
        self._rebuild_sr_table()

    def _on_cycle_toggle(self, *_args):
        if not self._suppress_replot:
            self._rebuild_sr_table()
            self._auto_replot()

    def _select_all(self):
        self._suppress_replot = True
        for v in self._cycle_vars.values():
            v.set(True)
        self._suppress_replot = False
        self._rebuild_sr_table()
        self._auto_replot()

    def _deselect_all(self):
        self._suppress_replot = True
        for v in self._cycle_vars.values():
            v.set(False)
        self._suppress_replot = False
        self._rebuild_sr_table()
        # Clear CV plot to placeholder
        self._legend_cv = None
        self.ax_cv.clear()
        self._reset_cv_axes_labels()
        self.canvas_cv.draw()

    def _selected_cycles(self):
        return [c for c, v in self._cycle_vars.items() if v.get()]

    def _edit_legend_labels(self, leg=None, is_cv=True):
        """Open the legend editor for the CV or Cdl legend.

        Called from double-click on legend (leg and is_cv come from _ei_leg_hit)
        or programmatically (leg=None → defaults to CV legend).
        """
        if leg is None:
            leg   = self._legend_cv if self._legend_cv is not None else self._legend_cdl
            is_cv = (leg is self._legend_cv)
        if leg is None:
            from tkinter import messagebox
            messagebox.showinfo("Info", "Plot CV first to create a legend.")
            return
        canvas    = self.canvas_cv if is_cv else self.canvas_cdl
        font_size = self._leg_size_cv if is_cv else self._leg_size_cdl
        leg.set_draggable(False)
        new_leg = open_legend_editor(self, leg, canvas, font_size)
        if is_cv:
            self._legend_cv = new_leg
        else:
            self._legend_cdl = new_leg
        if new_leg is not None:
            new_leg.set_draggable(True)
            # Persist CV labels so they survive replots (Cdl is auto-generated)
            if is_cv:
                self._legend_labels_cv = [
                    t.get_text() for t in new_leg.get_texts()
                ]

    def _update_e_std_rec(self):
        """Compute and display the recommended E_std (midpoint of plotted data range).

        The recommendation is (E_min + E_max) / 2 across the selected cycles' X-column
        data.  It is shown in raw column units (same units E_std is entered in).
        """
        if not self.active_file or self.active_file not in self.files:
            self.e_std_rec_var.set("—")
            return
        xcol = self.x_var.get()
        if not xcol:
            self.e_std_rec_var.set("—")
            return
        df       = self.files[self.active_file]["df"]
        selected = self._selected_cycles()
        try:
            if "cycle number" in df.columns and selected:
                mask   = df["cycle number"].isin(selected)
                x_data = df.loc[mask, xcol].values.astype(float)
            else:
                x_data = df[xcol].values.astype(float)
            if len(x_data) == 0:
                self.e_std_rec_var.set("—")
                return
            e_mid = (float(x_data.min()) + float(x_data.max())) / 2.0
            self.e_std_rec_var.set(f"{e_mid:.3f}")
        except Exception:
            self.e_std_rec_var.set("—")

    def _edit_plot_title(self, ax, canvas):
        """Prompt the user to edit a plot title (double-click on title area)."""
        from tkinter.simpledialog import askstring
        current = ax.title.get_text()
        new_title = askstring("Edit Title", "Plot title:", initialvalue=current, parent=self)
        if new_title is not None:
            ax.set_title(new_title, fontsize=9)
            canvas.draw_idle()
