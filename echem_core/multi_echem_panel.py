"""Multi E.Chem Panel — one independent subplot per loaded txt file.

Layout
------
Left (scrollable):
    Files listbox · axis selectors + unit dropdowns · plot range ·
    reference electrode · IR / RHE correction · cycle checkboxes (9 col) ·
    legend options · Plot button · log.

Right (scrollable):
    One labelled Figure per loaded file, stacked vertically.
    Each figure has its own NavigationToolbar (Home / Zoom / Pan / Save).
    All figures stay visible simultaneously; selecting a file in the listbox
    updates the left-panel controls for that file only.
"""

from collections import OrderedDict
import math

import numpy as np
import tkinter as tk
from tkinter import ttk

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from .file_manager import FileManagerMixin, _COLOR_NAMES, _COLOR_HEX, _default_xcol, _default_ycol, _PLOT_STYLES, _PLOT_STYLE_NAMES
from .correction import CorrectionMixin
from .plotting import apply_grid, draw_reflines, _cycle_colors, copy_figure_to_clipboard, _scale_legend_spacing
from .legend_editor import open_legend_editor
from .checklist import CheckableListbox

_CYCLE_BG        = "#e8f0fe"


# Maps J density unit-range labels to the underlying current base unit
_J_TO_BASE = {"A/cm²": "A", "mA/cm²": "mA", "µA/cm²": "µA", "nA/cm²": "nA"}

_CYCLE_ACTIVE_BG = "#cce0ff"
_CLICK_CYCLE_PX  = 8

_UNIT_DIMS = {
    "A": "I",  "mA": "I",  "µA": "I",  "nA": "I",
    "V": "E",  "mV": "E",  "µV": "E",  "nV": "E",
    "s": "t",  "ms": "t",  "µs": "t",  "min": "t", "h": "t",
    "Ohm": "Z", "Ω": "Z", "mΩ": "Z", "kΩ": "Z", "MΩ": "Z",
    "Hz": "f",  "kHz": "f", "MHz": "f",
    "deg": "φ", "rad": "φ",
}
_DIM_OPTS = {
    "I": ["(auto)", "A",  "mA",  "µA",  "nA"],
    "E": ["(auto)", "V",  "mV",  "µV",  "nV"],
    "t": ["(auto)", "s",  "ms",  "µs",  "min", "h"],
    "J": ["(auto)", "A/cm²", "mA/cm²", "µA/cm²", "nA/cm²"],
    "Z": ["(auto)", "mΩ", "Ω",  "kΩ",  "MΩ"],
    "f": ["(auto)", "Hz", "kHz", "MHz"],
    "φ": ["(auto)", "deg", "rad"],
}
_ALL_UNITS = ["(auto)", "A", "mA", "µA", "nA",
              "V", "mV", "µV", "nV", "s", "ms", "µs", "min", "h"]

_VOLTAGE_UNITS = frozenset({"V", "mV", "µV", "nV"})
_CURRENT_UNITS = frozenset({"A", "mA", "µA", "nA"})


class MultiEchemPanel(FileManagerMixin, CorrectionMixin, ttk.Frame):
    """Multi-file panel: each txt file gets its own independent labelled figure."""

    def __init__(self, master):
        ttk.Frame.__init__(self, master)
        self.files            = OrderedDict()
        self.active_file      = None
        self._suppress_replot = False
        self._loading_files   = False
        self._cycle_vars      = {}
        self._zoom_file       = None
        self._drag            = None   # drag-to-reorder state
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

        _lc = tk.Canvas(left_outer, highlightthickness=0)
        _ls = ttk.Scrollbar(left_outer, orient=tk.VERTICAL, command=_lc.yview)
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
        ttk.Label(left, text="Each file gets its own independent plot",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)
        fb = ttk.Frame(left)
        fb.pack(fill=tk.X, padx=4)
        ttk.Button(fb, text="Load File(s)", command=self._load_files).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(fb, text="Remove",       command=self._remove_file).pack(side=tk.LEFT)

        flf = ttk.Frame(left)
        flf.pack(fill=tk.X, padx=4, pady=2)
        self.file_listbox = CheckableListbox(flf, height=4,
                                             on_check=self._on_file_visibility_change,
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
        ttk.Label(left, text="Axes  (active file)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        def _refresh_unit_opts(col_var, unit_var, unit_cb):
            col = col_var.get()
            if col == "J":
                dim = "J"
            else:
                raw_unit = col.rsplit("/", 1)[-1].strip() if "/" in col else ""
                dim = _UNIT_DIMS.get(raw_unit)
            opts = list(_DIM_OPTS.get(dim, ["(auto)"]))
            unit_cb["values"] = opts
            if unit_var.get() not in opts:
                unit_var.set("(auto)")
            self._auto_replot()

        def _refresh_unit_after(unit_var, unit_cb):
            chosen = unit_var.get()
            if chosen and chosen != "(auto)":
                if chosen.endswith("/cm²"):
                    opts = list(_DIM_OPTS["J"])
                else:
                    dim = _UNIT_DIMS.get(chosen)
                    opts = list(_DIM_OPTS.get(dim, _ALL_UNITS))
                unit_cb["values"] = opts
            self._auto_replot()

        def _refresh_j_in_active_combo():
            """Add/remove 'J' from column combos based on active file's area."""
            try:
                _has_area = float(self.area_var.get()) > 0
            except (ValueError, TypeError):
                _has_area = False
            for combo, var in ((self.x_combo, self.x_var),
                               (self.y_combo, self.y_var)):
                vals = list(combo["values"])
                if _has_area and "J" not in vals:
                    vals.append("J")
                    combo["values"] = vals
                elif not _has_area and "J" in vals:
                    vals.remove("J")
                    combo["values"] = vals
                    if var.get() == "J" and vals:
                        var.set(vals[0])

        # X-axis
        ttk.Label(left, text="X-axis:").pack(anchor=tk.W, padx=4)
        x_row = ttk.Frame(left)
        x_row.pack(fill=tk.X, padx=4, pady=2)
        self.x_var   = tk.StringVar()
        self.x_combo = ttk.Combobox(x_row, textvariable=self.x_var,
                                    state="readonly", width=16)
        self.x_combo.pack(side=tk.LEFT)
        self.x_unit_var = tk.StringVar(value="V")
        x_unit_cb = ttk.Combobox(x_row, textvariable=self.x_unit_var,
                                  values=_DIM_OPTS["E"], state="readonly", width=12)
        x_unit_cb.pack(side=tk.LEFT, padx=(4, 0))
        self.x_combo.bind("<<ComboboxSelected>>",
                          lambda e: _refresh_unit_opts(self.x_var, self.x_unit_var, x_unit_cb))
        x_unit_cb.bind("<<ComboboxSelected>>",
                       lambda e: _refresh_unit_after(self.x_unit_var, x_unit_cb))

        # Y-axis
        ttk.Label(left, text="Y-axis:").pack(anchor=tk.W, padx=4)
        y_row = ttk.Frame(left)
        y_row.pack(fill=tk.X, padx=4, pady=2)
        self.y_var   = tk.StringVar()
        self.y_combo = ttk.Combobox(y_row, textvariable=self.y_var,
                                    state="readonly", width=16)
        self.y_combo.pack(side=tk.LEFT)
        self.y_unit_var = tk.StringVar(value="mA")
        y_unit_cb = ttk.Combobox(y_row, textvariable=self.y_unit_var,
                                  values=_DIM_OPTS["I"], state="readonly", width=12)
        y_unit_cb.pack(side=tk.LEFT, padx=(4, 0))
        self.y_combo.bind("<<ComboboxSelected>>",
                          lambda e: _refresh_unit_opts(self.y_var, self.y_unit_var, y_unit_cb))
        y_unit_cb.bind("<<ComboboxSelected>>",
                       lambda e: _refresh_unit_after(self.y_unit_var, y_unit_cb))

        def _do_refresh_unit_combos():
            """Silently update unit combobox options to match current column selections."""
            for col_var, unit_var, cb in (
                (self.x_var, self.x_unit_var, x_unit_cb),
                (self.y_var, self.y_unit_var, y_unit_cb),
            ):
                col = col_var.get()
                if col == "J":
                    dim = "J"
                else:
                    raw_unit = col.rsplit("/", 1)[-1].strip() if "/" in col else ""
                    dim = _UNIT_DIMS.get(raw_unit)
                opts = list(_DIM_OPTS.get(dim, ["(auto)"]))
                cb["values"] = opts
                if unit_var.get() not in opts:
                    unit_var.set("(auto)")
        self._do_refresh_unit_combos = _do_refresh_unit_combos

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

        # ── Plot range ────────────────────────────────────────────
        ttk.Label(left, text="Plot Range:", font=("", 8)).pack(
            anchor=tk.W, padx=4, pady=(6, 0))
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
        self.x_grid_int_var = tk.StringVar(value="0")
        _xgi = ttk.Entry(xr_f, textvariable=self.x_grid_int_var, width=5)
        _xgi.pack(side=tk.LEFT, padx=(2, 0))
        _xgi.bind("<Return>",   lambda e: self._auto_replot())
        _xgi.bind("<FocusOut>", lambda e: self._auto_replot())

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
        self.y_grid_int_var = tk.StringVar(value="0")
        _ygi = ttk.Entry(yr_f, textvariable=self.y_grid_int_var, width=5)
        _ygi.pack(side=tk.LEFT, padx=(2, 0))
        _ygi.bind("<Return>",   lambda e: self._auto_replot())
        _ygi.bind("<FocusOut>", lambda e: self._auto_replot())

        ttk.Label(left, text="(blank = auto)", foreground="gray",
                  font=("", 8)).pack(anchor=tk.W, padx=4)
        for _re in (_xmin, _xmax, _ymin, _ymax):
            _re.bind("<Return>",   lambda e: self._auto_replot())
            _re.bind("<FocusOut>", lambda e: self._auto_replot())

        flip_row = ttk.Frame(left)
        flip_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        self.x_flip_var = tk.BooleanVar(value=False)
        self.y_flip_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(flip_row, text="Flip X", variable=self.x_flip_var,
                        command=self._auto_replot).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Checkbutton(flip_row, text="Flip Y", variable=self.y_flip_var,
                        command=self._auto_replot).pack(side=tk.LEFT)

        # ── Current density (per-file electrode area — unlocks J units) ──
        area_row = ttk.Frame(left)
        area_row.pack(fill=tk.X, padx=4, pady=(4, 0))
        ttk.Label(area_row, text="Area (cm²):").pack(side=tk.LEFT)
        self.area_var = tk.StringVar()
        _area_e = ttk.Entry(area_row, textvariable=self.area_var, width=8)
        _area_e.pack(side=tk.LEFT, padx=(4, 0))

        def _on_area_change(e=None):
            # Update "J" availability in column combos, then refresh unit range combos
            _refresh_j_in_active_combo()
            self._suppress_replot = True
            _refresh_unit_opts(self.x_var, self.x_unit_var, x_unit_cb)
            self._suppress_replot = False
            _refresh_unit_opts(self.y_var, self.y_unit_var, y_unit_cb)

        _area_e.bind("<Return>",   _on_area_change)
        _area_e.bind("<FocusOut>", _on_area_change)
        ttk.Label(area_row, text="unlocks J units",
                  foreground="gray", font=("", 8)).pack(side=tk.LEFT, padx=6)

        # ── Reference electrode ───────────────────────────────────
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

        # ── IR / RHE correction ───────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="IR / RHE Correction",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        rf = ttk.Frame(left)
        rf.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(rf, text="R_sol (Ohm):").pack(side=tk.LEFT)
        self.r_sol_var = tk.StringVar(value="0")
        _rsol_e = ttk.Entry(rf, textvariable=self.r_sol_var, width=10)
        _rsol_e.pack(side=tk.LEFT, padx=4)
        _rsol_e.bind("<Return>",   lambda e: self._apply_correction())
        _rsol_e.bind("<FocusOut>", lambda e: self._apply_correction())

        ef = ttk.Frame(left)
        ef.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(ef, text="E_ref (V vs RHE):").pack(side=tk.LEFT)
        self.e_ref_var = tk.StringVar(value="0")
        _eref_e = ttk.Entry(ef, textvariable=self.e_ref_var, width=10)
        _eref_e.pack(side=tk.LEFT, padx=4)
        _eref_e.bind("<Return>",   lambda e: self._apply_correction())
        _eref_e.bind("<FocusOut>", lambda e: self._apply_correction())

        ttk.Label(left, text="(auto-applied on Enter / focus change)",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)
        ttk.Button(left, text="Reset Correction",
                   command=self._reset_correction).pack(anchor=tk.W, padx=4, pady=(2, 0))

        # ── Cycle checkboxes (9 columns) ──────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Cycles:", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        cb_row = ttk.Frame(left)
        cb_row.pack(fill=tk.X, padx=4)
        ttk.Button(cb_row, text="Select All",
                   command=self._select_all).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(cb_row, text="Deselect All",
                   command=self._deselect_all).pack(side=tk.LEFT)

        cyc_outer = ttk.Frame(left)
        cyc_outer.pack(fill=tk.X, padx=4, pady=2)
        cyc_canvas = tk.Canvas(cyc_outer, background=_CYCLE_BG,
                               highlightthickness=0, height=90)
        cyc_vs = ttk.Scrollbar(cyc_outer, orient=tk.VERTICAL,   command=cyc_canvas.yview)
        cyc_hs = ttk.Scrollbar(cyc_outer, orient=tk.HORIZONTAL, command=cyc_canvas.xview)
        cyc_canvas.configure(yscrollcommand=cyc_vs.set, xscrollcommand=cyc_hs.set)
        cyc_vs.pack(side=tk.RIGHT,  fill=tk.Y)
        cyc_hs.pack(side=tk.BOTTOM, fill=tk.X)
        cyc_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._cycle_inner  = tk.Frame(cyc_canvas, background=_CYCLE_BG)
        self._cycle_canvas = cyc_canvas
        cyc_canvas.create_window((0, 0), window=self._cycle_inner, anchor=tk.NW)
        self._cycle_inner.bind(
            "<Configure>",
            lambda e: cyc_canvas.configure(scrollregion=cyc_canvas.bbox("all")))

        def _cyc_wheel(e):
            cyc_canvas.yview_scroll(-1 * (e.delta // 120), "units")
            return "break"
        cyc_canvas.bind("<MouseWheel>", _cyc_wheel)
        self._cycle_inner.bind("<MouseWheel>", _cyc_wheel)

        # ── Cycle Colors ─────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(left, text="Cycle Colors", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        _cc_row1 = ttk.Frame(left)
        _cc_row1.pack(fill=tk.X, padx=4, pady=2)
        self.cycle_gradient_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(_cc_row1, text="Gradient", variable=self.cycle_gradient_var,
                        command=self._on_gradient_change).pack(side=tk.LEFT)
        self.cycle_reverse_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(_cc_row1, text="Reverse", variable=self.cycle_reverse_var,
                        command=self._on_gradient_change).pack(side=tk.LEFT, padx=(8, 0))
        _cc_row2 = ttk.Frame(left)
        _cc_row2.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(_cc_row2, text="Step:").pack(side=tk.LEFT)
        self.lightness_step_var = tk.StringVar(value="0.15")
        _step_spin = ttk.Spinbox(_cc_row2, textvariable=self.lightness_step_var,
                                  from_=0.01, to=0.30, increment=0.01, width=6)
        _step_spin.pack(side=tk.LEFT, padx=(4, 0))
        _step_spin.bind("<<Increment>>", lambda e: self._on_gradient_change())
        _step_spin.bind("<<Decrement>>", lambda e: self._on_gradient_change())
        _step_spin.bind("<Return>",      lambda e: self._on_gradient_change())
        _step_spin.bind("<FocusOut>",    lambda e: self._on_gradient_change())

        # ── Title ─────────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        _title_row = ttk.Frame(left)
        _title_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(_title_row, text="Title:").pack(side=tk.LEFT)
        self.plot_title_var = tk.StringVar(value="")
        _title_entry = ttk.Entry(_title_row, textvariable=self.plot_title_var)
        _title_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
        _title_entry.bind("<Return>",   lambda e: self._auto_replot())
        _title_entry.bind("<FocusOut>", lambda e: self._auto_replot())

        # ── Legend options ────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Legend  (active file)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        leg_row1 = ttk.Frame(left)
        leg_row1.pack(fill=tk.X, padx=4, pady=2)
        self.legend_show_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(leg_row1, text="Show Legend",
                        variable=self.legend_show_var,
                        command=self._auto_replot).pack(side=tk.LEFT)
        self.legend_frame_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(leg_row1, text="Show Frame",
                        variable=self.legend_frame_var,
                        command=self._toggle_legend_frame).pack(side=tk.LEFT, padx=(8, 0))

        leg_row2 = ttk.Frame(left)
        leg_row2.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(leg_row2, text="Size:").pack(side=tk.LEFT)
        self.legend_size_var = tk.StringVar(value="8")
        _leg_sz_e = ttk.Entry(leg_row2, textvariable=self.legend_size_var, width=4)
        _leg_sz_e.pack(side=tk.LEFT, padx=(2, 8))
        _leg_sz_e.bind("<Return>",   lambda e: self._auto_replot())
        _leg_sz_e.bind("<FocusOut>", lambda e: self._auto_replot())
        ttk.Label(leg_row2, text="Loc:").pack(side=tk.LEFT)
        self.legend_loc_var = tk.StringVar(value="best")
        _leg_loc_cb = ttk.Combobox(
            leg_row2, textvariable=self.legend_loc_var,
            values=["best", "upper right", "upper left", "lower left", "lower right",
                    "right", "center left", "center right", "lower center",
                    "upper center", "center"],
            state="readonly", width=11,
        )
        _leg_loc_cb.pack(side=tk.LEFT, padx=2)
        def _on_leg_loc_select_multi(e=None):
            if self.active_file and self.active_file in self.files:
                self.files[self.active_file].pop("legend_manual_pos", None)
            self._auto_replot()
        _leg_loc_cb.bind("<<ComboboxSelected>>", _on_leg_loc_select_multi)
        ttk.Label(left, text="(left-drag to move, right-drag to resize; dbl-click to edit)",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)

        # ── Grid ──────────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Grid  (active file)", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        grid_xy_row = ttk.Frame(left)
        grid_xy_row.pack(fill=tk.X, padx=4, pady=2)
        self.x_grid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(grid_xy_row, text="X", variable=self.x_grid_var,
                        command=self._auto_replot).pack(side=tk.LEFT)
        self.y_grid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(grid_xy_row, text="Y", variable=self.y_grid_var,
                        command=self._auto_replot).pack(side=tk.LEFT, padx=(10, 0))
        grid_style_row = ttk.Frame(left)
        grid_style_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(grid_style_row, text="Style:").pack(side=tk.LEFT)
        self.grid_style_var = tk.StringVar(value="dashed")
        _gscb = ttk.Combobox(grid_style_row, textvariable=self.grid_style_var,
                             values=["dashed", "dotted", "solid", "dash-dot"],
                             state="readonly", width=9)
        _gscb.pack(side=tk.LEFT, padx=(2, 6))
        _gscb.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())
        ttk.Label(grid_style_row, text="Color:").pack(side=tk.LEFT)
        self.grid_color_var = tk.StringVar(value="gray")
        _gcol_cb = ttk.Combobox(grid_style_row, textvariable=self.grid_color_var,
                                values=["gray", "black", "red", "blue", "green",
                                        "orange", "purple", "crimson", "royalblue",
                                        "darkorange", "teal"],
                                state="readonly", width=9)
        _gcol_cb.pack(side=tk.LEFT, padx=(2, 6))
        _gcol_cb.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())
        ttk.Label(grid_style_row, text="Width:").pack(side=tk.LEFT)
        self.grid_linewidth_var = tk.StringVar(value="0.8")
        _glw = ttk.Entry(grid_style_row, textvariable=self.grid_linewidth_var, width=4)
        _glw.pack(side=tk.LEFT, padx=(2, 0))
        _glw.bind("<Return>",   lambda e: self._auto_replot())
        _glw.bind("<FocusOut>", lambda e: self._auto_replot())

        # ── Font ──────────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
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

        # ── Reference Lines ───────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Reference Lines  (active file)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        ref_add_row = ttk.Frame(left)
        ref_add_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(ref_add_row, text="X:").pack(side=tk.LEFT)
        self._ref_x_var = tk.StringVar()
        _ref_x_e = ttk.Entry(ref_add_row, textvariable=self._ref_x_var, width=7)
        _ref_x_e.pack(side=tk.LEFT, padx=(2, 0))
        ttk.Button(ref_add_row, text="+X", width=3,
                   command=self._add_xrefline).pack(side=tk.LEFT, padx=(2, 8))
        ttk.Label(ref_add_row, text="Y:").pack(side=tk.LEFT)
        self._ref_y_var = tk.StringVar()
        _ref_y_e = ttk.Entry(ref_add_row, textvariable=self._ref_y_var, width=7)
        _ref_y_e.pack(side=tk.LEFT, padx=(2, 0))
        ttk.Button(ref_add_row, text="+Y", width=3,
                   command=self._add_yrefline).pack(side=tk.LEFT, padx=2)
        _ref_x_e.bind("<Return>", lambda e: self._add_xrefline())
        _ref_y_e.bind("<Return>", lambda e: self._add_yrefline())
        ref_list_row = ttk.Frame(left)
        ref_list_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        self._reflines_lb = tk.Listbox(ref_list_row, height=3,
                                        selectmode=tk.SINGLE, exportselection=False)
        self._reflines_lb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._reflines_lb.bind("<<ListboxSelect>>", lambda e: self._on_refline_select())
        ttk.Button(ref_list_row, text="Remove",
                   command=self._remove_refline).pack(side=tk.RIGHT, padx=(4, 0))
        ref_opt_row = ttk.Frame(left)
        ref_opt_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(ref_opt_row, text="Style:").pack(side=tk.LEFT)
        self._refline_style_var = tk.StringVar(value="dashed")
        _rl_style_cb = ttk.Combobox(ref_opt_row, textvariable=self._refline_style_var,
                                     values=["dashed", "dotted", "solid", "dash-dot"],
                                     state="readonly", width=9)
        _rl_style_cb.pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(ref_opt_row, text="Color:").pack(side=tk.LEFT)
        self._refline_color_var = tk.StringVar(value="dimgray")
        _rl_color_cb = ttk.Combobox(ref_opt_row, textvariable=self._refline_color_var,
                                     values=["dimgray", "black", "red", "blue", "green",
                                             "orange", "purple", "crimson", "royalblue",
                                             "darkorange", "teal", "saddlebrown"],
                                     state="readonly", width=9)
        _rl_color_cb.pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(ref_opt_row, text="Width:").pack(side=tk.LEFT)
        self._refline_linewidth_var = tk.StringVar(value="1.0")
        _rl_lw = ttk.Entry(ref_opt_row, textvariable=self._refline_linewidth_var, width=4)
        _rl_lw.pack(side=tk.LEFT, padx=(2, 0))
        _rl_style_cb.bind("<<ComboboxSelected>>", lambda e: self._on_refline_style_color_change())
        _rl_color_cb.bind("<<ComboboxSelected>>", lambda e: self._on_refline_style_color_change())
        _rl_lw.bind("<Return>",   lambda e: self._on_refline_style_color_change())
        _rl_lw.bind("<FocusOut>", lambda e: self._on_refline_style_color_change())

        # ── Plot button ───────────────────────────────────────────
        ttk.Button(left, text="Plot Active File",
                   command=self._auto_replot).pack(padx=4, pady=6, anchor=tk.W)

        # ── Log ───────────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Log", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        log_f = ttk.Frame(left)
        log_f.pack(fill=tk.X, padx=4, pady=2)
        self.log_text = tk.Text(log_f, height=5, state=tk.DISABLED,
                                wrap=tk.WORD, relief=tk.SUNKEN, borderwidth=1)
        log_sc = ttk.Scrollbar(log_f, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sc.set)
        log_sc.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # ── Right panel: scrollable, one figure per file ──────────
        right_outer = ttk.Frame(body)
        body.add(right_outer, weight=1)

        # row 0: size bar (permanent), row 1: zoom bar, row 2: canvas area
        right_outer.rowconfigure(0, weight=0)
        right_outer.rowconfigure(1, weight=0)
        right_outer.rowconfigure(2, weight=1)
        right_outer.columnconfigure(0, weight=1)

        # ── Plot size controls (always visible) ──────────────────────
        self.plot_w_var = tk.StringVar(value="10.5")
        self.plot_h_var = tk.StringVar(value="5.5")
        _size_bar = ttk.Frame(right_outer)
        _size_bar.grid(row=0, column=0, sticky="ew", padx=4, pady=2)
        ttk.Label(_size_bar, text="Plot size (in):").pack(side=tk.LEFT, padx=(4, 2))
        ttk.Label(_size_bar, text="W").pack(side=tk.LEFT)
        _pw_e = ttk.Entry(_size_bar, textvariable=self.plot_w_var, width=5)
        _pw_e.pack(side=tk.LEFT, padx=(1, 6))
        ttk.Label(_size_bar, text="H").pack(side=tk.LEFT)
        _ph_e = ttk.Entry(_size_bar, textvariable=self.plot_h_var, width=5)
        _ph_e.pack(side=tk.LEFT, padx=(1, 0))
        for _e in (_pw_e, _ph_e):
            _e.bind("<Return>",   lambda ev: self._apply_plot_size())
            _e.bind("<FocusOut>", lambda ev: self._apply_plot_size())

        # Zoom bar (hidden initially; shown via grid() in _zoom_file_view)
        self._zoom_bar = ttk.Frame(right_outer)
        ttk.Button(self._zoom_bar, text="← Back to Grid",
                   command=self._unzoom_file_view).pack(side=tk.LEFT, padx=6, pady=3)
        self._zoom_bar.grid(row=1, column=0, sticky="ew")
        self._zoom_bar.grid_remove()   # hidden until first zoom

        # Canvas container sits in row 2 and fills all remaining space
        _right_inner = ttk.Frame(right_outer)
        _right_inner.grid(row=2, column=0, sticky="nsew")
        _right_inner.rowconfigure(0, weight=1)
        _right_inner.columnconfigure(0, weight=1)

        self._right_canvas = tk.Canvas(_right_inner, highlightthickness=0)
        right_vscroll = ttk.Scrollbar(_right_inner, orient=tk.VERTICAL,
                                      command=self._right_canvas.yview)
        right_hscroll = ttk.Scrollbar(_right_inner, orient=tk.HORIZONTAL,
                                      command=self._right_canvas.xview)
        self._right_canvas.configure(yscrollcommand=right_vscroll.set,
                                     xscrollcommand=right_hscroll.set)
        right_vscroll.grid(row=0, column=1, sticky="ns")
        right_hscroll.grid(row=1, column=0, sticky="ew")
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

        # No column weights — each plot keeps its fixed size

        # Drop indicator: a thin colored bar shown during drag-to-reorder
        self._drop_line = tk.Frame(self._plots_frame, bg="#1a73e8", height=3)

        # Placeholder shown when no files are loaded (uses grid like everything else)
        self._placeholder = ttk.Label(
            self._plots_frame,
            text="Load files to display individual plots here.",
            foreground="gray", font=("", 10),
        )
        self._placeholder.grid(row=0, column=0, columnspan=2, pady=60)

    # ════════════════════════════════════════════════════════════════
    # Unit conversion helper
    # ════════════════════════════════════════════════════════════════
    def _get_unit_scale(self, col, target_unit):
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
            "Ohm": 1.0, "Ω": 1.0, "mΩ": 1e-3, "kΩ": 1e3, "MΩ": 1e6,
            "Hz": 1.0, "kHz": 1e3, "MHz": 1e6,
            "rad": 1.0, "deg": math.pi / 180.0,
        }
        _DIMS = {
            "A": "I", "mA": "I", "µA": "I", "nA": "I",
            "V": "E", "mV": "E", "µV": "E", "nV": "E",
            "s": "t", "ms": "t", "µs": "t", "min": "t", "h": "t",
            "Ohm": "Z", "Ω": "Z", "mΩ": "Z", "kΩ": "Z", "MΩ": "Z",
            "Hz": "f", "kHz": "f", "MHz": "f",
            "rad": "φ", "deg": "φ",
        }
        if "/" in col:
            col_base, src_unit = col.rsplit("/", 1)
            col_base = col_base.strip()
            src_unit = src_unit.strip()
        else:
            col_base = col
            src_unit = None
        display_label = f"{col_base} ({target_unit})"
        src_f = _FACTORS.get(src_unit)
        tgt_f = _FACTORS.get(target_unit)
        if (src_f is not None and tgt_f is not None
                and _DIMS.get(src_unit) == _DIMS.get(target_unit)):
            return src_f / tgt_f, display_label
        return 1.0, display_label

    # ════════════════════════════════════════════════════════════════
    # Per-file figure creation / destruction
    # ════════════════════════════════════════════════════════════════
    def _create_file_figure(self, short):
        """Create a labelled frame + matplotlib figure for *short* in the right panel."""
        entry = self.files.get(short)
        if entry is None or "fig" in entry:
            return   # already created or file removed

        self._placeholder.grid_remove()

        panel_ref = self

        # Outer bordered frame (replaces LabelFrame so we can add a visible drag strip)
        _HDR_BG = "#c0cfe4"
        frame = tk.Frame(self._plots_frame, relief="groove", bd=2)
        # Placement is handled by _relayout_figures(); do not pack/grid here

        # ── Header strip: drag handle + filename ────────────────────
        header = tk.Frame(frame, bg=_HDR_BG, cursor="fleur")
        header.pack(fill=tk.X, side=tk.TOP)
        handle_lbl = tk.Label(header, text=f"⠿  {short}",
                              bg=_HDR_BG, cursor="fleur",
                              font=("", 9, "bold"), anchor=tk.W)
        handle_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=3)

        # ── Content area (canvas + toolbar) ─────────────────────────
        inner = ttk.Frame(frame, padding=(4, 2, 4, 2))
        inner.pack()

        try:
            _fw = float(self.plot_w_var.get())
            _fh = float(self.plot_h_var.get())
        except (ValueError, AttributeError):
            _fw, _fh = 9.5, 5.5
        fig = Figure(figsize=(_fw, _fh), dpi=100, constrained_layout=True)
        ax  = fig.add_subplot(111)
        canvas = FigureCanvasTkAgg(fig, master=inner)
        canvas.get_tk_widget().pack()

        tb_frame = ttk.Frame(inner)
        tb_frame.pack(fill=tk.X)

        class _Toolbar(NavigationToolbar2Tk):
            def home(tb_self, *args):
                panel_ref._reset_file_view(short)

        _tb = _Toolbar(canvas, tb_frame, pack_toolbar=False)
        _tb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        _tb.update()
        tk.Button(
            tb_frame, text="Copy",
            command=lambda f=fig: copy_figure_to_clipboard(f),
            relief=tk.RAISED, borderwidth=1, padx=6,
        ).pack(side=tk.LEFT, padx=(4, 2), pady=1)

        # Forward mouse-wheel on the frame and toolbar (not the canvas) to
        # the right-side scroll panel.  The canvas widget must NOT forward wheel
        # events so that matplotlib's scroll-to-zoom handler works undisturbed.
        def _fwd_scroll(e):
            self._right_canvas.yview_scroll(-1 * (e.delta // 120), "units")

        frame.bind("<MouseWheel>", _fwd_scroll)
        header.bind("<MouseWheel>", _fwd_scroll)
        tb_frame.bind("<MouseWheel>", _fwd_scroll)

        # Click on the frame / toolbar → select this file in the listbox
        def _activate(e=None):
            self._activate_file(short)
        frame.bind("<Button-1>",    _activate, add="+")
        tb_frame.bind("<Button-1>", _activate, add="+")

        # Drag on the header strip → reorder subplots
        for _w in (header, handle_lbl):
            _w.bind("<ButtonPress-1>",   lambda e, s=short: self._on_frame_press(e, s))
            _w.bind("<B1-Motion>",       lambda e, s=short: self._on_frame_drag(e, s))
            _w.bind("<ButtonRelease-1>", lambda e, s=short: self._on_frame_release(e, s))

        # Per-figure matplotlib interactions
        canvas.mpl_connect("scroll_event",         lambda ev: self._on_scroll(ev, short))
        canvas.mpl_connect("button_press_event",   lambda ev: self._on_press(ev, short))
        canvas.mpl_connect("button_release_event", lambda ev: self._on_release(ev, short))
        canvas.mpl_connect("motion_notify_event",  lambda ev: self._on_motion(ev, short))

        # Store figure objects and interaction state in the file entry
        entry.update({
            "fig":        fig,
            "ax":         ax,
            "canvas":     canvas,
            "plot_frame": frame,
            "legend":     None,
            "leg_size":   8.0,
            "auto_xlim":  None,
            "auto_ylim":  None,
            # Interaction state
            "panning":    False,
            "pan_ax":     None,
            "pan_start":  None,
            "pan_moved":  False,
            "leg_resize": False,
            "leg_resize_start_y":  None,
            "leg_resize_start_sz": None,
            "ann":        None,
            "ann_dot":    None,
            "ann_last":   None,
            "ann_idx":    0,
        })

        ax.set_title(short, fontsize=9)
        ax.set_xlabel("Potential (V)")
        ax.set_ylabel("Current (mA)")
        canvas.draw()

        # Place the new frame in the grid alongside existing ones
        self._relayout_figures()

    def _destroy_file_figure(self, short):
        """Tear down the figure panel for *short* before its entry is deleted."""
        frame = self.files.get(short, {}).get("plot_frame")
        if frame is not None:
            frame.destroy()

    # ════════════════════════════════════════════════════════════════
    # FileManagerMixin overrides
    # ════════════════════════════════════════════════════════════════
    def _load_files(self):
        existing = set(self.files.keys())
        super()._load_files()
        # Create figure panels for any files added by the parent loader
        for short in self.files:
            if short not in existing:
                self._create_file_figure(short)

    def _remove_file(self):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        short = self.file_listbox.get(sel[0])
        self._destroy_file_figure(short)
        super()._remove_file()
        if not self.files:
            self.area_var.set("")
            self._placeholder.grid(row=0, column=0, columnspan=2, pady=60)
        else:
            self._relayout_figures()

    def _on_file_visibility_change(self, short, visible):
        """Toggle a file's hidden flag and update the subplot grid."""
        if short not in self.files:
            return
        entry = self.files[short]
        # Snapshot the current zoom before hiding so it survives the replot on unhide
        if not visible and "ax" in entry:
            entry["view_xlim"] = entry["ax"].get_xlim()
            entry["view_ylim"] = entry["ax"].get_ylim()
        entry["hidden"] = not visible
        # If hiding the currently zoomed file, exit zoom mode
        if not visible and self._zoom_file == short:
            self._zoom_file = None
            self._zoom_bar.grid_remove()
            self._right_canvas.itemconfig(self._plots_win, height=0)
        self._relayout_figures()
        if visible:
            self._plot_file(short)
            # Restore zoom/pan after the replot resets the view to auto-scale
            if "view_xlim" in entry and "ax" in entry:
                entry["ax"].set_xlim(entry["view_xlim"])
                entry["ax"].set_ylim(entry["view_ylim"])
                entry["canvas"].draw_idle()

    def _on_file_reorder(self, new_order):
        """Called when the file list is drag-reordered; rebuilds self.files and re-layouts."""
        from collections import OrderedDict
        new_files = OrderedDict()
        for name in new_order:
            if name in self.files:
                new_files[name] = self.files[name]
        for name, entry in self.files.items():
            if name not in new_files:
                new_files[name] = entry
        self.files = new_files
        self._relayout_figures()

    def _on_right_canvas_configure(self, event):
        if self._zoom_file:
            self._right_canvas.itemconfig(self._plots_win,
                                          width=event.width, height=event.height)

    def _relayout_figures(self):
        """Place every file's plot frame in a max-2-column grid, in load order.
        When self._zoom_file is set, expand that subplot to fill the full canvas area.
        """
        MAX_COLS = 2
        valid = [(s, self.files[s]) for s in self.files
                 if "plot_frame" in self.files[s]
                 and not self.files[s].get("hidden", False)]

        # ── Zoom mode: one subplot fills the whole canvas area ───────
        if self._zoom_file and any(s == self._zoom_file for s, _ in valid):
            for s, entry in valid:
                if s == self._zoom_file:
                    entry["plot_frame"].grid(row=0, column=0, columnspan=2,
                                             sticky="nsew", padx=4, pady=4)
                else:
                    entry["plot_frame"].grid_remove()
            self._plots_frame.rowconfigure(0, weight=1)
            return

        # ── Normal grid layout ───────────────────────────────────────
        # Remove all frames (including hidden ones) before re-placing visible ones
        for s in self.files:
            pf = self.files[s].get("plot_frame")
            if pf is not None:
                pf.grid_remove()
        for i, (short, entry) in enumerate(valid):
            row = i // MAX_COLS
            col = i % MAX_COLS
            # columnspan=1 must be explicit to reset any prior columnspan=2 from zoom mode
            entry["plot_frame"].grid(row=row, column=col, columnspan=1,
                                     sticky="nsew", padx=4, pady=4)
        # Rows have no extra weight — height is driven by figure content
        n_rows = (len(valid) + MAX_COLS - 1) // MAX_COLS if valid else 0
        for r in range(n_rows):
            self._plots_frame.rowconfigure(r, weight=0)

    def _save_active_state(self):
        """Save all per-file UI state when switching away from a file."""
        if not self.active_file or self.active_file not in self.files:
            return
        entry = self.files[self.active_file]
        # Preserve current zoom/pan view so it can be restored on file switch-back
        if "ax" in entry:
            entry["view_xlim"] = entry["ax"].get_xlim()
            entry["view_ylim"] = entry["ax"].get_ylim()
        entry["selected_cycles"] = self._selected_cycles()
        try:
            entry["r_sol"] = float(self.r_sol_var.get())
        except ValueError:
            pass
        try:
            entry["e_ref"] = float(self.e_ref_var.get())
        except ValueError:
            pass
        entry["x_col"]        = self.x_var.get()
        entry["y_col"]        = self.y_var.get()
        entry["x_unit"]       = self.x_unit_var.get()
        entry["y_unit"]       = self.y_unit_var.get()
        entry["x_min"]        = self.x_min_var.get()
        entry["x_max"]        = self.x_max_var.get()
        entry["y_min"]        = self.y_min_var.get()
        entry["y_max"]        = self.y_max_var.get()
        entry["area"]         = self.area_var.get()
        entry["ref_electrode"]  = self.ref_electrode_var.get()
        entry["legend_show"]    = self.legend_show_var.get()
        entry["legend_frame"]   = self.legend_frame_var.get()
        try:
            entry["leg_size"] = float(self.legend_size_var.get())
        except ValueError:
            pass
        entry["legend_loc"] = self.legend_loc_var.get()
        entry["x_grid"]          = self.x_grid_var.get()
        entry["y_grid"]          = self.y_grid_var.get()
        entry["x_grid_int"]      = self.x_grid_int_var.get()
        entry["y_grid_int"]      = self.y_grid_int_var.get()
        entry["grid_style"]      = self.grid_style_var.get()
        entry["grid_color"]      = self.grid_color_var.get()
        entry["grid_linewidth"]  = self.grid_linewidth_var.get()
        entry["cycle_gradient"] = self.cycle_gradient_var.get()
        entry["cycle_reverse"]  = self.cycle_reverse_var.get()
        entry["lightness_step"] = self.lightness_step_var.get()
        entry["linewidth"]      = self.linewidth_var.get()
        entry["plot_style"]     = self.plot_style_var.get()
        entry["x_flip"]         = self.x_flip_var.get()
        entry["y_flip"]         = self.y_flip_var.get()
        entry["custom_title"]   = self.plot_title_var.get()

    def _switch_active_file(self, short):
        self.active_file = short
        entry = self.files[short]

        # Ensure figure exists (lazy creation for files added before panel was ready)
        self._create_file_figure(short)

        # Initialise per-file fields absent on first visit
        entry.setdefault("x_col",         None)
        entry.setdefault("y_col",         None)
        entry.setdefault("x_unit",        "V")
        entry.setdefault("y_unit",        "mA")
        entry.setdefault("x_min",         "")
        entry.setdefault("x_max",         "")
        entry.setdefault("y_min",         "")
        entry.setdefault("y_max",         "")
        entry.setdefault("area",          "")
        entry.setdefault("ref_electrode", "Ag/AgCl")
        entry.setdefault("legend_show",   True)
        entry.setdefault("legend_frame",  True)
        entry.setdefault("leg_size",      8.0)
        entry.setdefault("legend_loc",    "best")
        entry.setdefault("x_grid",        False)
        entry.setdefault("y_grid",        False)
        entry.setdefault("x_grid_int",    "0")
        entry.setdefault("y_grid_int",    "0")
        entry.setdefault("grid_style",    "dashed")
        entry.setdefault("grid_color",    "gray")
        entry.setdefault("grid_linewidth","0.8")
        entry.setdefault("reflines",      [])
        entry.setdefault("cycle_gradient", True)
        entry.setdefault("cycle_reverse",  False)
        entry.setdefault("lightness_step", "0.15")
        entry.setdefault("linewidth",      "1.5")
        entry.setdefault("x_flip",         False)
        entry.setdefault("y_flip",         False)
        entry.setdefault("custom_title",   "")

        df   = entry["df"]

        # Restore area FIRST so J-availability check below is correct
        self.area_var.set(entry["area"])
        self.x_unit_var.set(entry["x_unit"])
        self.y_unit_var.set(entry["y_unit"])
        self.x_min_var.set(entry["x_min"])
        self.x_max_var.set(entry["x_max"])
        self.y_min_var.set(entry["y_min"])
        self.y_max_var.set(entry["y_max"])

        # Build column list; include "J" when this file has area
        cols = list(df.columns)
        try:
            _has_area = float(entry["area"]) > 0
        except (ValueError, TypeError):
            _has_area = False
        if _has_area and "J" not in cols:
            cols.append("J")

        self.x_combo["values"] = cols
        self.y_combo["values"] = cols

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

        # Update unit combobox options to match the incoming file's columns
        fn = getattr(self, '_do_refresh_unit_combos', None)
        if fn:
            fn()

        self.ref_electrode_var.set(entry["ref_electrode"])
        self.legend_show_var.set(entry["legend_show"])
        self.legend_frame_var.set(entry["legend_frame"])
        self.legend_size_var.set(str(int(entry["leg_size"])
                                     if float(entry["leg_size"]) == int(entry["leg_size"])
                                     else entry["leg_size"]))
        self.legend_loc_var.set(entry["legend_loc"])
        self.r_sol_var.set(str(entry["r_sol"]))
        self.e_ref_var.set(str(entry["e_ref"]))
        self.x_grid_var.set(entry["x_grid"])
        self.y_grid_var.set(entry["y_grid"])
        self.x_grid_int_var.set(entry["x_grid_int"])
        self.y_grid_int_var.set(entry["y_grid_int"])
        self.grid_style_var.set(entry["grid_style"])
        self.grid_color_var.set(entry["grid_color"])
        self.grid_linewidth_var.set(entry["grid_linewidth"])

        old = self._suppress_replot
        self._suppress_replot = True
        if "cycle number" in df.columns:
            cycles = sorted(int(c) for c in df["cycle number"].unique())
            self._populate_cycle_checkboxes(cycles, entry["selected_cycles"])
        else:
            self._populate_cycle_checkboxes([], [])
        self._suppress_replot = old

        # Restore color and gradient UI vars BEFORE _auto_replot so _plot_file
        # (which reads from UI vars for the active file) uses the correct settings.
        color = entry.get("color", "#1f77b4")
        name = next((n for n, h in _COLOR_HEX.items() if h == color), "Blue")
        self.file_color_var.set(name)
        self.cycle_gradient_var.set(entry.get("cycle_gradient", True))
        self.cycle_reverse_var.set(entry.get("cycle_reverse", False))
        self.lightness_step_var.set(entry.get("lightness_step", "0.15"))
        self.linewidth_var.set(entry.get("linewidth", "1.5"))
        self.plot_style_var.set(entry.get("plot_style", "Line"))
        self.x_flip_var.set(entry.get("x_flip", False))
        self.y_flip_var.set(entry.get("y_flip", False))
        self.plot_title_var.set(entry.get("custom_title", ""))

        self._auto_replot()

        # Restore per-file zoom/pan view if the user had previously zoomed/panned
        _entry = self.files.get(short)
        if _entry and "view_xlim" in _entry and "ax" in _entry:
            _entry["ax"].set_xlim(_entry["view_xlim"])
            _entry["ax"].set_ylim(_entry["view_ylim"])
            _entry["canvas"].draw_idle()

        self._refresh_reflines_lb()

    # ════════════════════════════════════════════════════════════════
    # Replot
    # ════════════════════════════════════════════════════════════════
    def _auto_replot(self):
        if self._suppress_replot:
            return
        if not self.active_file or not self.x_var.get() or not self.y_var.get():
            return
        self._plot_file(self.active_file)

    def _plot(self):
        self._auto_replot()

    def _plot_file(self, short):
        """Replot the figure for *short* using that file's stored settings."""
        entry = self.files.get(short)
        if entry is None or "ax" not in entry:
            return

        # Read settings: use live UI vars for the active file, stored values otherwise
        is_active = (short == self.active_file)

        def _pick(attr, ui_var, default=""):
            return ui_var.get() if is_active else entry.get(attr, default)

        xcol     = _pick("x_col",        self.x_var,            "")
        ycol     = _pick("y_col",        self.y_var,            "")
        x_unit   = _pick("x_unit",       self.x_unit_var,       "(auto)")
        y_unit   = _pick("y_unit",       self.y_unit_var,       "(auto)")
        x_min_s  = _pick("x_min",        self.x_min_var,        "")
        x_max_s  = _pick("x_max",        self.x_max_var,        "")
        y_min_s  = _pick("y_min",        self.y_min_var,        "")
        y_max_s  = _pick("y_max",        self.y_max_var,        "")
        ref      = _pick("ref_electrode", self.ref_electrode_var, "")
        leg_show = entry.get("legend_show",  True)  if not is_active else self.legend_show_var.get()
        leg_frm  = entry.get("legend_frame", True)  if not is_active else self.legend_frame_var.get()
        leg_loc  = entry.get("legend_loc",  "best") if not is_active else self.legend_loc_var.get()
        try:
            leg_size = float(self.legend_size_var.get()) if is_active else float(entry.get("leg_size", 8.0))
        except ValueError:
            leg_size = 8.0

        selected = (self._selected_cycles() if is_active
                    else entry.get("selected_cycles", []))

        if not xcol or not ycol:
            return

        df = entry["df"]

        # Per-file area for J density conversion
        area_s = (self.area_var.get() if is_active else entry.get("area", ""))
        try:
            _farea = float(area_s) if area_s else 0.0
        except (ValueError, TypeError):
            _farea = 0.0

        # Resolve "J" virtual column → find the actual current column in df
        _x_is_J = (xcol == "J")
        _y_is_J = (ycol == "J")
        _real_xcol = xcol
        _real_ycol = ycol
        if _x_is_J or _y_is_J:
            for c in df.columns:
                if "/" in c and c.rsplit("/", 1)[-1].strip() in _CURRENT_UNITS:
                    if _x_is_J:
                        _real_xcol = c
                    if _y_is_J:
                        _real_ycol = c
                    break

        # X-axis scale
        if _x_is_J:
            _xbase = _J_TO_BASE.get(x_unit)
            if _xbase:
                x_scale, _ = self._get_unit_scale(_real_xcol, _xbase)
                x_label = f"J ({x_unit})"
            else:
                x_scale = 1.0
                _src = _real_xcol.rsplit("/", 1)[-1].strip() if "/" in _real_xcol else "?"
                x_label = f"J ({_src}/cm²)"
            if _farea > 0:
                x_scale = x_scale / _farea
        else:
            x_scale, x_label = self._get_unit_scale(_real_xcol, x_unit)

        # Y-axis scale
        if _y_is_J:
            _ybase = _J_TO_BASE.get(y_unit)
            if _ybase:
                y_scale, _ = self._get_unit_scale(_real_ycol, _ybase)
                y_label = f"J ({y_unit})"
            else:
                y_scale = 1.0
                _src = _real_ycol.rsplit("/", 1)[-1].strip() if "/" in _real_ycol else "?"
                y_label = f"J ({_src}/cm²)"
            if _farea > 0:
                y_scale = y_scale / _farea
        else:
            y_scale, y_label = self._get_unit_scale(_real_ycol, y_unit)

        ax     = entry["ax"]
        canvas = entry["canvas"]

        self._clear_ann(short, redraw=False)
        # Save dragged legend position before discarding old legend
        _old_leg = entry.get("legend")
        if _old_leg is not None:
            _loc = getattr(_old_leg, '_loc', None)
            if isinstance(_loc, (tuple, list)):
                entry["legend_manual_pos"] = tuple(_loc)
        entry["legend"] = None
        ax.clear()

        base_color = entry.get("color", "#1f77b4")
        _grad  = self.cycle_gradient_var.get() if is_active else entry.get("cycle_gradient", True)
        _rev   = self.cycle_reverse_var.get()  if is_active else entry.get("cycle_reverse", False)
        try:    _step = float(self.lightness_step_var.get() if is_active else entry.get("lightness_step", "0.15"))
        except: _step = 0.08

        _lw_s = self.linewidth_var.get() if is_active else entry.get("linewidth", "1.5")
        try:
            _lw = float(_lw_s)
        except (ValueError, TypeError):
            _lw = 1.5
        _sname = (self.plot_style_var.get() if is_active else entry.get("plot_style", "Line"))
        _ls, _mk, _ms = _PLOT_STYLES.get(_sname, ("-", "", 0))

        has_data = False
        if "cycle number" in df.columns:
            if selected:
                cycle_cols = (_cycle_colors(base_color, len(selected), _step, _rev)
                              if _grad else [base_color] * len(selected))
                for i, c in enumerate(selected):
                    sub = df[df["cycle number"] == c]
                    ax.plot(sub[_real_xcol] * x_scale,
                            sub[_real_ycol] * y_scale,
                            color=cycle_cols[i], label=f"C{c}", linewidth=_lw,
                            linestyle=_ls, marker=_mk or None,
                            markersize=_ms if _mk else 0)
                has_data = True
        else:
            ax.plot(df[_real_xcol] * x_scale, df[_real_ycol] * y_scale,
                    color=base_color, label=short, linewidth=_lw,
                    linestyle=_ls, marker=_mk or None,
                    markersize=_ms if _mk else 0)
            has_data = True

        # Append "(vs Ref)" only to voltage-type axes; J is never voltage
        if _x_is_J:
            _x_is_V = False
        else:
            _x_src = _real_xcol.rsplit("/", 1)[-1].strip() if "/" in _real_xcol else ""
            _x_is_V = (x_unit in _VOLTAGE_UNITS if x_unit != "(auto)"
                       else _x_src in _VOLTAGE_UNITS)
        ax.set_xlabel(f"{x_label}  (vs {ref})" if (ref and _x_is_V) else x_label)

        if _y_is_J:
            _y_is_V = False
        else:
            _y_src = _real_ycol.rsplit("/", 1)[-1].strip() if "/" in _real_ycol else ""
            _y_is_V = (y_unit in _VOLTAGE_UNITS if y_unit != "(auto)"
                       else _y_src in _VOLTAGE_UNITS)
        ax.set_ylabel(f"{y_label}  (vs {ref})" if (ref and _y_is_V) else y_label)
        _title = (self.plot_title_var.get() if short == self.active_file
                  else entry.get("custom_title", ""))
        ax.set_title(_title, fontsize=9)

        canvas.draw()
        entry["auto_xlim"] = ax.get_xlim()
        entry["auto_ylim"] = ax.get_ylim()

        # Apply manual axis range
        self._apply_range(short, x_min_s, x_max_s, y_min_s, y_max_s)

        # Reference lines — each entry carries its own style and color
        draw_reflines(ax, entry.get("reflines", []))

        # Grid — read from live UI if active file, else from stored entry
        _xg  = self.x_grid_var.get()        if is_active else entry.get("x_grid",        False)
        _yg  = self.y_grid_var.get()        if is_active else entry.get("y_grid",        False)
        _xgi = self.x_grid_int_var.get()   if is_active else entry.get("x_grid_int",    "0")
        _ygi = self.y_grid_int_var.get()   if is_active else entry.get("y_grid_int",    "0")
        _gs  = self.grid_style_var.get()   if is_active else entry.get("grid_style",    "dashed")
        _gc  = self.grid_color_var.get()   if is_active else entry.get("grid_color",    "gray")
        _glw = self.grid_linewidth_var.get() if is_active else entry.get("grid_linewidth", "0.8")
        apply_grid(ax, _xg, _yg, _xgi, _ygi, _gs, linewidth=_glw, color=_gc)

        # Legend
        if leg_show and has_data and ax.get_lines():
            entry["legend"] = ax.legend(fontsize=leg_size, loc=leg_loc)
            entry["legend"].set_draggable(True)
            entry["legend"].get_frame().set_visible(leg_frm)
            entry["leg_size"] = leg_size
            # Restore custom labels if count matches
            custom = entry.get("legend_labels", [])
            if custom:
                for text_obj, lbl in zip(entry["legend"].get_texts(), custom):
                    if lbl:
                        text_obj.set_text(lbl)
            # Restore dragged position
            if entry.get("legend_manual_pos") is not None:
                entry["legend"]._loc = entry["legend_manual_pos"]
            canvas.draw()

        self._apply_font_to_ax(ax, canvas)

    def _apply_range(self, short, x_min_s, x_max_s, y_min_s, y_max_s):
        """Apply manual axis limits to the figure for *short*."""
        entry = self.files.get(short)
        if entry is None or "ax" not in entry:
            return
        ax = entry["ax"]
        changed = False
        for setter, kwarg, val_s in (
            (ax.set_xlim, "left",   x_min_s),
            (ax.set_xlim, "right",  x_max_s),
            (ax.set_ylim, "bottom", y_min_s),
            (ax.set_ylim, "top",    y_max_s),
        ):
            try:
                setter(**{kwarg: float(val_s)})
                changed = True
            except (ValueError, TypeError):
                pass
        # Flip axes if requested (only for the active file whose vars reflect current state)
        if short == self.active_file:
            xl = ax.get_xlim()
            if self.x_flip_var.get() != (xl[0] > xl[1]):
                ax.set_xlim(xl[1], xl[0])
                changed = True
            yl = ax.get_ylim()
            if self.y_flip_var.get() != (yl[0] > yl[1]):
                ax.set_ylim(yl[1], yl[0])
                changed = True
        else:
            # Non-active files: apply saved flip state from entry dict
            xl = ax.get_xlim()
            if entry.get("x_flip", False) != (xl[0] > xl[1]):
                ax.set_xlim(xl[1], xl[0])
                changed = True
            yl = ax.get_ylim()
            if entry.get("y_flip", False) != (yl[0] > yl[1]):
                ax.set_ylim(yl[1], yl[0])
                changed = True

        if changed:
            entry["canvas"].draw_idle()

    # ════════════════════════════════════════════════════════════════
    # Reset view (Home button)
    # ════════════════════════════════════════════════════════════════
    def _reset_file_view(self, short):
        entry = self.files.get(short)
        if entry is None or entry.get("auto_xlim") is None:
            return
        entry["ax"].set_xlim(entry["auto_xlim"])
        entry["ax"].set_ylim(entry["auto_ylim"])
        entry["canvas"].draw_idle()

    # ════════════════════════════════════════════════════════════════
    # Subplot zoom (double-click to expand / Back to Grid)
    # ════════════════════════════════════════════════════════════════
    def _apply_plot_size(self, event=None):
        """Resize all file figures to the current plot_w_var × plot_h_var (inches)."""
        try:
            w = float(self.plot_w_var.get())
            h = float(self.plot_h_var.get())
        except ValueError:
            return
        w = max(2.0, min(50.0, w))
        h = max(1.5, min(50.0, h))
        dpi = 100
        for entry in self.files.values():
            fig = entry.get("fig")
            cv  = entry.get("canvas")
            if fig and cv:
                fig.set_size_inches(w, h)
                cv.get_tk_widget().config(width=int(w * dpi), height=int(h * dpi))
                cv.draw_idle()
        self._right_canvas.after(
            50, lambda: self._right_canvas.configure(
                scrollregion=self._right_canvas.bbox("all")))

    def _zoom_file_view(self, short):
        """Expand the subplot for *short* to fill the full right panel."""
        self._zoom_file = short
        self._zoom_bar.grid()
        w = self._right_canvas.winfo_width()
        h = self._right_canvas.winfo_height()
        if w > 1 and h > 1:
            self._right_canvas.itemconfig(self._plots_win, width=w, height=h)
            entry = self.files.get(short, {})
            fig   = entry.get("fig")
            cv    = entry.get("canvas")
            if fig and cv:
                dpi = fig.get_dpi()
                fig.set_size_inches(w / dpi, h / dpi)
                cv.get_tk_widget().config(width=w, height=h)
                cv.draw_idle()
        self._relayout_figures()

    def _unzoom_file_view(self):
        """Restore the 2-column grid layout."""
        self._zoom_file = None
        self._zoom_bar.grid_remove()
        self._right_canvas.itemconfig(self._plots_win, height=0)
        self._apply_plot_size()
        self._relayout_figures()
        self._right_canvas.configure(scrollregion=self._right_canvas.bbox("all"))

    # ════════════════════════════════════════════════════════════════
    # File color helper
    # ════════════════════════════════════════════════════════════════
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
        _leg = ax.get_legend()
        if _leg is not None:
            _leg.set_visible(False)
        ax.figure.tight_layout()
        if _leg is not None:
            _leg.set_visible(True)
        canvas.draw()
        if kb:
            for lbl in ax.get_xticklabels() + ax.get_yticklabels():
                lbl.set_fontweight('bold')
            if _leg is not None:
                _leg.set_visible(False)
            ax.figure.tight_layout()
            if _leg is not None:
                _leg.set_visible(True)
            canvas.draw()

    # ════════════════════════════════════════════════════════════════
    # Legend frame toggle
    # ════════════════════════════════════════════════════════════════
    def _toggle_legend_frame(self):
        if not self.active_file:
            return
        leg = self.files[self.active_file].get("legend")
        if leg is not None:
            leg.get_frame().set_visible(self.legend_frame_var.get())
            self.files[self.active_file]["canvas"].draw()

    # ════════════════════════════════════════════════════════════════
    # Per-figure matplotlib interactions
    # ════════════════════════════════════════════════════════════════
    def _on_scroll(self, event, short):
        entry = self.files.get(short)
        if entry is None or event.inaxes is not entry.get("ax"):
            return
        ax     = entry["ax"]
        canvas = entry["canvas"]
        scale  = 0.8 if event.step > 0 else 1.25
        xl, yl = ax.get_xlim(), ax.get_ylim()
        xd, yd = event.xdata, event.ydata
        xf = (xd - xl[0]) / (xl[1] - xl[0]) if xl[1] != xl[0] else 0.5
        yf = (yd - yl[0]) / (yl[1] - yl[0]) if yl[1] != yl[0] else 0.5
        nxr = (xl[1] - xl[0]) * scale
        nyr = (yl[1] - yl[0]) * scale
        ax.set_xlim(xd - nxr * xf,  xd + nxr * (1 - xf))
        ax.set_ylim(yd - nyr * yf,  yd + nyr * (1 - yf))
        canvas.draw_idle()

    def _on_press(self, event, short):
        entry = self.files.get(short)
        if entry is None:
            return
        # Any click on this figure → select it as the active file
        self._activate_file(short)
        entry["pan_moved"] = False
        ax = entry.get("ax")

        # Handle double-click anywhere in the figure (title may be outside axes proper)
        if event.button == 1 and getattr(event, 'dblclick', False) and ax is not None:
            canvas = entry["canvas"]
            # 1. Legend editor?
            if event.inaxes is ax:
                leg = entry.get("legend")
                if leg is not None:
                    try:
                        r = canvas.get_renderer()
                        if leg.get_window_extent(r).contains(event.x, event.y):
                            self._edit_legend_labels()
                            return
                    except Exception:
                        pass
            # 2. Title area (works inside or outside axes)?
            try:
                r = canvas.get_renderer()
                ax_bbox  = ax.get_window_extent(r)
                fig_bbox = ax.get_figure().get_window_extent(r)
                t_bbox   = ax.title.get_window_extent(r)
                on_title = (
                    (t_bbox.width > 2 and t_bbox.contains(event.x, event.y))
                    or (ax_bbox.x0 <= event.x <= ax_bbox.x1
                        and ax_bbox.y1 <= event.y <= fig_bbox.y1)
                )
                if on_title:
                    self._edit_subplot_title(short, ax, canvas)
                    return
            except Exception:
                pass
            # 3. Zoom toggle (only when click is inside axes)
            if event.inaxes is ax:
                if self._zoom_file is None:
                    self._zoom_file_view(short)
                else:
                    self._unzoom_file_view()
            return

        if event.button == 1 and event.inaxes is ax:
            leg = entry.get("legend")
            if leg is not None:
                try:
                    r = event.canvas.get_renderer()
                    if leg.get_window_extent(r).contains(event.x, event.y):
                        return
                except Exception:
                    pass
            entry["panning"]   = True
            entry["pan_ax"]    = ax
            entry["pan_start"] = (event.xdata, event.ydata)
        elif event.button == 3:
            leg = entry.get("legend")
            if leg is not None:
                try:
                    r = event.canvas.get_renderer()
                    if leg.get_window_extent(r).contains(event.x, event.y):
                        entry["leg_resize"]          = True
                        entry["leg_resize_start_y"]  = event.y
                        entry["leg_resize_start_sz"] = entry.get("leg_size", 8.0)
                except Exception:
                    pass

    def _on_release(self, event, short):
        entry = self.files.get(short)
        if entry is None:
            return

        was_resizing = entry.get("leg_resize", False)
        entry["leg_resize"] = False
        entry["panning"]    = False
        entry["pan_ax"]     = None

        if was_resizing:
            if short == self.active_file:
                new_sz = entry.get("leg_size", 8.0)
                self.legend_size_var.set(
                    str(int(new_sz)) if new_sz == int(new_sz) else str(new_sz))
            return

        ax = entry.get("ax")
        if (event.button == 1
                and not entry.get("pan_moved", False)
                and event.inaxes is ax):
            leg = entry.get("legend")
            if leg is not None:
                try:
                    r = event.canvas.get_renderer()
                    if leg.get_window_extent(r).contains(event.x, event.y):
                        return
                except Exception:
                    pass
            self._annotate(event, short)
        elif event.button == 3 and event.inaxes is ax:
            self._clear_ann(short)

    def _on_motion(self, event, short):
        entry = self.files.get(short)
        if entry is None:
            return

        if entry.get("panning") and entry.get("pan_ax") is not None:
            if event.inaxes is not entry["pan_ax"] or event.xdata is None:
                return
            entry["pan_moved"] = True
            ax = entry["ax"]
            dx = entry["pan_start"][0] - event.xdata
            dy = entry["pan_start"][1] - event.ydata
            ax.set_xlim(ax.get_xlim()[0] + dx, ax.get_xlim()[1] + dx)
            ax.set_ylim(ax.get_ylim()[0] + dy, ax.get_ylim()[1] + dy)
            entry["canvas"].draw_idle()
            return

        if entry.get("leg_resize") and entry.get("legend") is not None:
            dy     = event.y - entry["leg_resize_start_y"]
            new_sz = max(4.0, min(30.0, entry["leg_resize_start_sz"] + dy / 5.0))
            prev_sz = entry.get("leg_size", 8.0)
            entry["leg_size"] = new_sz
            leg = entry["legend"]
            if prev_sz > 0:
                _scale_legend_spacing(leg, new_sz / prev_sz)
            for t in leg.get_texts():
                t.set_fontsize(new_sz)
            tt = leg.get_title()
            if tt:
                tt.set_fontsize(new_sz)
            entry["canvas"].draw()

    # ════════════════════════════════════════════════════════════════
    # Click annotate (per figure)
    # ════════════════════════════════════════════════════════════════
    def _annotate(self, event, short):
        entry = self.files.get(short)
        if entry is None:
            return
        ax    = entry["ax"]
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
        last = entry.get("ann_last")
        if (last is not None
                and abs(event.x - last[0]) <= _CLICK_CYCLE_PX
                and abs(event.y - last[1]) <= _CLICK_CYCLE_PX):
            entry["ann_idx"] = (entry["ann_idx"] + 1) % len(candidates)
        else:
            entry["ann_idx"] = 0
        entry["ann_last"] = (event.x, event.y)
        n   = len(candidates)
        idx = entry["ann_idx"]
        _, ln, x, y = candidates[idx]
        label = ln.get_label() or "?"
        xl, yl = ax.get_xlim(), ax.get_ylim()
        xf = (x - xl[0]) / (xl[1] - xl[0]) if xl[1] != xl[0] else 0.5
        yf = (y - yl[0]) / (yl[1] - yl[0]) if yl[1] != yl[0] else 0.5
        xoff = -95 if xf > 0.65 else 15
        yoff = -60 if yf > 0.65 else 15
        hint = f"  [{idx + 1}/{n}]" if n > 1 else ""
        text = f"{short}\nx = {x:.4g}\ny = {y:.4g}"
        if label != short:
            text += f"  ({label})"
        text += hint
        if n > 1 and idx == 0:
            text += "\n↻ click again to cycle"
        self._clear_ann(short, redraw=False)
        entry["ann"] = ax.annotate(
            text, xy=(x, y), xytext=(xoff, yoff), textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.4", fc="lightyellow", ec="gray", alpha=0.92),
            arrowprops=dict(arrowstyle="->", color="gray", lw=1.2),
            fontsize=8, zorder=10,
        )
        entry["ann_dot"], = ax.plot(x, y, "o", color=ln.get_color(),
                                    markersize=7, zorder=11, label="_ann_dot")
        entry["canvas"].draw_idle()

    def _clear_ann(self, short, redraw=True):
        entry = self.files.get(short)
        if entry is None:
            return
        for key in ("ann", "ann_dot"):
            artist = entry.get(key)
            if artist is not None:
                try:
                    artist.remove()
                except Exception:
                    pass
                entry[key] = None
        entry["ann_last"] = None
        entry["ann_idx"]  = 0
        if redraw and "canvas" in entry:
            entry["canvas"].draw_idle()

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

    def _on_cycle_toggle(self, *_args):
        if not self._suppress_replot:
            self._auto_replot()

    def _select_all(self):
        self._suppress_replot = True
        for v in self._cycle_vars.values():
            v.set(True)
        self._suppress_replot = False
        self._auto_replot()

    def _deselect_all(self):
        self._suppress_replot = True
        for v in self._cycle_vars.values():
            v.set(False)
        self._suppress_replot = False
        if self.active_file and "ax" in self.files.get(self.active_file, {}):
            entry = self.files[self.active_file]
            self._clear_ann(self.active_file, redraw=False)
            entry["legend"] = None
            entry["ax"].clear()
            entry["ax"].set_title(self.active_file, fontsize=9)
            entry["canvas"].draw()

    def _selected_cycles(self):
        return [c for c, v in self._cycle_vars.items() if v.get()]

    # ════════════════════════════════════════════════════════════════
    # Drag-to-reorder: drag a subplot frame to change grid order
    # ════════════════════════════════════════════════════════════════
    def _on_frame_press(self, event, short):
        self._drag = {
            "short":      short,
            "start_x":    event.x_root,
            "start_y":    event.y_root,
            "active":     False,
            "target":     None,
            "target_top": True,
        }

    def _on_frame_drag(self, event, short):
        drag = self._drag
        if drag is None or drag["short"] != short:
            return
        if not drag["active"]:
            if abs(event.x_root - drag["start_x"]) + abs(event.y_root - drag["start_y"]) < 6:
                return
            drag["active"] = True

        # Detect which visible frame the cursor is over
        target = None
        target_top = True
        for s, entry in self.files.items():
            if s == short or entry.get("hidden"):
                continue
            pf = entry.get("plot_frame")
            if pf is None:
                continue
            x0 = pf.winfo_rootx()
            y0 = pf.winfo_rooty()
            w  = pf.winfo_width()
            h  = pf.winfo_height()
            if x0 <= event.x_root <= x0 + w and y0 <= event.y_root <= y0 + h:
                target = s
                target_top = (event.y_root - y0) < h / 2
                break
        drag["target"]     = target
        drag["target_top"] = target_top

        # Move / show the drop-indicator line
        if target is not None:
            pf = self.files[target]["plot_frame"]
            rx = pf.winfo_x()
            ry = pf.winfo_y()
            rw = pf.winfo_width()
            rh = pf.winfo_height()
            line_y = ry if target_top else ry + rh - 3
            self._drop_line.place(x=rx, y=line_y, width=rw, height=3)
            self._drop_line.lift()
        else:
            self._drop_line.place_forget()

    def _on_frame_release(self, event, short):
        drag = self._drag
        self._drag = None
        self._drop_line.place_forget()
        if drag is None or not drag["active"]:
            return
        target = drag.get("target")
        if target is None or target == short:
            return
        self._reorder_files(short, target, before=drag.get("target_top", True))

    def _reorder_files(self, from_short, to_short, *, before=True):
        """Move from_short to just before (or after) to_short in self.files."""
        keys = list(self.files.keys())
        if from_short not in keys or to_short not in keys:
            return
        keys.remove(from_short)
        to_idx = keys.index(to_short)
        keys.insert(to_idx if before else to_idx + 1, from_short)
        self.files = OrderedDict((k, self.files[k]) for k in keys)
        self._rebuild_listbox()
        # Restore active-file highlight without triggering _on_file_select
        if self.active_file in self.files:
            idx = list(self.files.keys()).index(self.active_file)
            self.file_listbox.selection_clear(0, tk.END)
            self._loading_files = True
            try:
                self.file_listbox.selection_set(idx)
            finally:
                self._loading_files = False
        self._relayout_figures()

    def _rebuild_listbox(self):
        """Rebuild CheckableListbox rows to match current self.files order."""
        self.file_listbox.clear()
        for short, entry in self.files.items():
            self.file_listbox.insert(tk.END, short,
                                     checked=not entry.get("hidden", False))

    # ════════════════════════════════════════════════════════════════
    # Click-to-select: clicking any plot selects that file
    # ════════════════════════════════════════════════════════════════
    def _activate_file(self, short):
        """Select *short* in the listbox and make it the active file."""
        keys = list(self.files.keys())
        if short not in keys:
            return
        idx = keys.index(short)
        self.file_listbox.selection_clear(0, tk.END)
        self.file_listbox.selection_set(idx)
        self.file_listbox.see(idx)
        if short != self.active_file:
            self._save_active_state()
            self._switch_active_file(short)

    def _edit_subplot_title(self, short, ax, canvas):
        """Prompt the user to edit the title of the subplot for *short*."""
        from tkinter.simpledialog import askstring
        entry = self.files.get(short, {})
        current = entry.get("custom_title", ax.title.get_text() or short)
        new_title = askstring("Edit Title", "Plot title:", initialvalue=current, parent=self)
        if new_title is not None:
            entry["custom_title"] = new_title
            ax.set_title(new_title, fontsize=9)
            canvas.draw_idle()

    def _edit_legend_labels(self):
        if not self.active_file:
            return
        entry = self.files.get(self.active_file)
        if entry is None:
            return
        leg = entry.get("legend")
        if leg is None:
            from tkinter import messagebox
            messagebox.showinfo("Info", "Plot data first to create a legend.")
            return
        leg.set_draggable(False)
        entry["legend"] = open_legend_editor(
            self, leg, entry["canvas"], entry.get("leg_size", 8.0))
        if entry.get("legend") is not None:
            entry["legend"].set_draggable(True)
            # Persist labels so they survive the next replot
            entry["legend_labels"] = [
                t.get_text() for t in entry["legend"].get_texts()
            ]

    # ════════════════════════════════════════════════════════════════
    # Reference line helpers
    # ════════════════════════════════════════════════════════════════
    def _add_xrefline(self):
        if not self.active_file:
            return
        try:
            v = float(self._ref_x_var.get())
        except ValueError:
            return
        style = self._refline_style_var.get()
        color = self._refline_color_var.get()
        lw    = self._refline_linewidth_var.get()
        self.files[self.active_file].setdefault("reflines", []).append(('x', v, style, color, lw))
        self._reflines_lb.insert(tk.END, f"X = {v:.4g}")
        self._auto_replot()

    def _add_yrefline(self):
        if not self.active_file:
            return
        try:
            v = float(self._ref_y_var.get())
        except ValueError:
            return
        style = self._refline_style_var.get()
        color = self._refline_color_var.get()
        lw    = self._refline_linewidth_var.get()
        self.files[self.active_file].setdefault("reflines", []).append(('y', v, style, color, lw))
        self._reflines_lb.insert(tk.END, f"Y = {v:.4g}")
        self._auto_replot()

    def _remove_refline(self):
        sel = self._reflines_lb.curselection()
        if not sel or not self.active_file:
            return
        idx = sel[0]
        self.files[self.active_file]["reflines"].pop(idx)
        self._reflines_lb.delete(idx)
        self._auto_replot()

    def _refresh_reflines_lb(self):
        self._reflines_lb.delete(0, tk.END)
        if not self.active_file:
            return
        for e in self.files.get(self.active_file, {}).get("reflines", []):
            axis, val = e[0], e[1]
            self._reflines_lb.insert(tk.END, f"{'X' if axis == 'x' else 'Y'} = {val:.4g}")

    def _on_refline_select(self):
        """Populate style/color/width widgets from the selected line's own settings."""
        sel = self._reflines_lb.curselection()
        if not sel or not self.active_file:
            return
        reflines = self.files.get(self.active_file, {}).get("reflines", [])
        if sel[0] >= len(reflines):
            return
        e = reflines[sel[0]]
        self._refline_style_var.set(e[2])
        self._refline_color_var.set(e[3])
        self._refline_linewidth_var.set(e[4] if len(e) > 4 else "1.0")

    def _on_refline_style_color_change(self):
        """Apply new style/color/width to the currently selected reference line."""
        sel = self._reflines_lb.curselection()
        if not sel or not self.active_file:
            return  # no line selected — widgets just set defaults for next add
        reflines = self.files.get(self.active_file, {}).get("reflines", [])
        idx = sel[0]
        if idx >= len(reflines):
            return
        axis, val = reflines[idx][:2]
        reflines[idx] = (axis, val,
                         self._refline_style_var.get(),
                         self._refline_color_var.get(),
                         self._refline_linewidth_var.get())
        self._auto_replot()

    # ════════════════════════════════════════════════════════════════
    # Log helper
    # ════════════════════════════════════════════════════════════════
    def _log(self, message: str):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
