"""Multi E.Chem 2 — Grouped overlay panel.

Each GROUP gets its own subplot (2-column grid, like Multi E.Chem).
Each group overlays multiple files on one plot (like General E.Chem).

Per-file settings : color · linewidth · style · area · IR/RHE correction · cycles
Per-group settings: axes · range · legend · grid · font · reference lines
"""

from collections import OrderedDict
import math

import numpy as np
import tkinter as tk
from tkinter import ttk, simpledialog

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from .file_manager import (FileManagerMixin, _COLOR_NAMES, _COLOR_HEX,
                            _default_xcol, _default_ycol,
                            _PLOT_STYLES, _PLOT_STYLE_NAMES)
from .correction import CorrectionMixin
from .plotting import apply_grid, draw_reflines, _cycle_colors, copy_figure_to_clipboard, _scale_legend_spacing, _reorder_legend_handles, _build_legend_order
from .legend_editor import open_legend_editor
from .checklist import CheckableListbox
from . import session_manager as _sm

_CYCLE_BG        = "#e8f0fe"
_CYCLE_ACTIVE_BG = "#cce0ff"
_GROUP_HDR_BG     = "#c8e6c9"   # light green — distinct from Multi E.Chem blue
_GROUP_HDR_ACTIVE = "#ffd54f"   # gold — matches Multi E.Chem 1 active header
_CLICK_CYCLE_PX  = 8

_J_TO_BASE = {"A/cm²": "A", "mA/cm²": "mA", "µA/cm²": "µA", "nA/cm²": "nA"}

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
_ALL_UNITS     = ["(auto)", "A", "mA", "µA", "nA",
                  "V", "mV", "µV", "nV", "s", "ms", "µs", "min", "h"]
_VOLTAGE_UNITS = frozenset({"V", "mV", "µV", "nV"})
_CURRENT_UNITS = frozenset({"A", "mA", "µA", "nA"})


class MultiEchem2Panel(FileManagerMixin, CorrectionMixin, ttk.Frame):
    """Grouped multi-file overlay panel.

    • Each GROUP → one subplot in a 2-column grid.
    • Multiple files can be assigned to a group and are overlaid on its plot.
    """

    def __init__(self, master):
        ttk.Frame.__init__(self, master)
        self.files            = OrderedDict()   # filename → file entry
        self.groups           = OrderedDict()   # group_name → group entry
        self.active_file      = None
        self.active_group     = None
        self._suppress_replot = False
        self._loading_files   = False
        self._cycle_vars      = {}
        self._zoom_group      = None
        self._drag            = None   # drag-to-reorder group subplot frames
        self._plot_highlight      = False
        self._active_cycle        = None   # specific cycle number highlighted (None = whole file)
        self._copied_group_params = None   # clipboard for Copy/Paste group settings
        self._build_panel()
        self.after(500, self._auto_set_initial_size)

    # ════════════════════════════════════════════════════════════════
    # Panel construction
    # ════════════════════════════════════════════════════════════════
    def _build_panel(self):
        body = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True)

        # ── Scrollable left panel ─────────────────────────────────
        left_outer = ttk.Frame(body, width=340)
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

        # ══ FILES ══════════════════════════════════════════════════
        ttk.Label(left, text="Files:", font=("", 9, "bold")).pack(anchor=tk.W, padx=4, pady=(6, 0))
        ttk.Label(left, text="Load files, then assign to groups below",
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

        # ══ GROUPS ══════════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        _grp_hdr = ttk.Frame(left)
        _grp_hdr.pack(fill=tk.X, padx=4)
        ttk.Label(_grp_hdr, text="Groups:", font=("", 9, "bold")).pack(side=tk.LEFT)
        gb = ttk.Frame(left)
        gb.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Button(gb, text="New Group", command=self._new_group).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(gb, text="Rename",    command=self._rename_group).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(gb, text="Delete",    command=self._delete_group).pack(side=tk.LEFT)

        glf = ttk.Frame(left)
        glf.pack(fill=tk.X, padx=4, pady=2)
        self.group_listbox = CheckableListbox(glf, height=3,
                                              on_check=self._on_group_visibility_change,
                                              on_reorder=self._on_group_reorder)
        self.group_listbox.pack(fill=tk.X, expand=True)
        self.group_listbox.bind("<<ListboxSelect>>", self._on_group_select)

        ttk.Button(left, text="↓ Add Selected Files to Group",
                   command=self._add_files_to_group).pack(fill=tk.X, padx=4, pady=(2, 0))

        _fgrp_hdr = ttk.Frame(left)
        _fgrp_hdr.pack(fill=tk.X, padx=4, pady=(4, 0))
        ttk.Label(_fgrp_hdr, text="Files in selected group:",
                  font=("", 8), foreground="gray").pack(side=tk.LEFT)
        gif = ttk.Frame(left)
        gif.pack(fill=tk.X, padx=4, pady=2)
        self.group_files_lb = CheckableListbox(gif, height=3,
                                               on_check=self._on_group_file_visibility_change,
                                               on_reorder=self._on_group_file_reorder)
        self.group_files_lb.pack(fill=tk.X, expand=True)
        self.group_files_lb.bind("<<ListboxSelect>>", self._on_group_file_select)

        ttk.Button(left, text="↑ Remove Selected from Group",
                   command=self._remove_files_from_group).pack(fill=tk.X, padx=4, pady=(0, 2))

        # Copy / Paste group display settings
        _cp_row = ttk.Frame(left)
        _cp_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Button(_cp_row, text="Copy Settings",
                   command=self._copy_group_settings).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(_cp_row, text="Paste Settings",
                   command=self._paste_group_settings).pack(side=tk.LEFT)
        ttk.Label(_cp_row, text="(font/grid/legend)",
                  foreground="gray", font=("", 7)).pack(side=tk.LEFT, padx=(6, 0))

        # ══ FILE STYLE ═════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(left, text="File Style  (active file)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        _fc_row = ttk.Frame(left)
        _fc_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_fc_row, text="Color:").pack(side=tk.LEFT)
        self.file_color_var = tk.StringVar(value="Blue")
        _fccb = ttk.Combobox(_fc_row, textvariable=self.file_color_var,
                              values=_COLOR_NAMES, state="readonly", width=10)
        _fccb.pack(side=tk.LEFT, padx=(4, 0))
        _fccb.bind("<<ComboboxSelected>>", self._on_file_color_change)
        ttk.Label(_fc_row, text="Width:").pack(side=tk.LEFT, padx=(6, 0))
        self.linewidth_var = tk.StringVar(value="1.5")
        _lw_e = ttk.Entry(_fc_row, textvariable=self.linewidth_var, width=4)
        _lw_e.pack(side=tk.LEFT, padx=(2, 0))
        _lw_e.bind("<Return>",   lambda e: self._on_linewidth_change())
        _lw_e.bind("<FocusOut>", lambda e: self._on_linewidth_change())

        _sty_row = ttk.Frame(left)
        _sty_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(_sty_row, text="Shape:").pack(side=tk.LEFT)
        self.plot_style_var = tk.StringVar(value="Line")
        _stcb = ttk.Combobox(_sty_row, textvariable=self.plot_style_var,
                              values=_PLOT_STYLE_NAMES, state="readonly", width=11)
        _stcb.pack(side=tk.LEFT, padx=(2, 0))
        _stcb.bind("<<ComboboxSelected>>", lambda e: self._on_plot_style_change())

        # ══ AREA + IR/RHE (per file) ════════════════════════════════
        area_row = ttk.Frame(left)
        area_row.pack(fill=tk.X, padx=4, pady=(4, 0))
        ttk.Label(area_row, text="Area (cm²):").pack(side=tk.LEFT)
        self.area_var = tk.StringVar()
        _ae = ttk.Entry(area_row, textvariable=self.area_var, width=8)
        _ae.pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(area_row, text="(J units)", foreground="gray",
                  font=("", 8)).pack(side=tk.LEFT, padx=4)
        _ae.bind("<Return>",   lambda e: self._on_area_change())
        _ae.bind("<FocusOut>", lambda e: self._on_area_change())

        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(left, text="IR / RHE Correction  (per file)",
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

        # ══ CYCLES (per file) ═══════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Cycles  (active file)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        cb_row = ttk.Frame(left)
        cb_row.pack(fill=tk.X, padx=4)
        ttk.Button(cb_row, text="Select All",
                   command=self._select_all).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(cb_row, text="Deselect All",
                   command=self._deselect_all).pack(side=tk.LEFT)

        cyc_outer = ttk.Frame(left)
        cyc_outer.pack(fill=tk.X, padx=4, pady=2)
        cyc_canvas = tk.Canvas(cyc_outer, background=_CYCLE_BG,
                               highlightthickness=0, height=80)
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

        # Cycle colors (per file)
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(left, text="Cycle Colors  (active file)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        _cc1 = ttk.Frame(left)
        _cc1.pack(fill=tk.X, padx=4, pady=2)
        self.cycle_gradient_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(_cc1, text="Gradient", variable=self.cycle_gradient_var,
                        command=self._on_gradient_change).pack(side=tk.LEFT)
        self.cycle_reverse_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(_cc1, text="Reverse", variable=self.cycle_reverse_var,
                        command=self._on_gradient_change).pack(side=tk.LEFT, padx=(8, 0))
        _cc2 = ttk.Frame(left)
        _cc2.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(_cc2, text="Step:").pack(side=tk.LEFT)
        self.lightness_step_var = tk.StringVar(value="0.15")
        _step_spin = ttk.Spinbox(_cc2, textvariable=self.lightness_step_var,
                                  from_=0.01, to=0.30, increment=0.01, width=6)
        _step_spin.pack(side=tk.LEFT, padx=(4, 0))
        _step_spin.bind("<<Increment>>", lambda e: self._on_gradient_change())
        _step_spin.bind("<<Decrement>>", lambda e: self._on_gradient_change())
        _step_spin.bind("<Return>",      lambda e: self._on_gradient_change())
        _step_spin.bind("<FocusOut>",    lambda e: self._on_gradient_change())

        # ══ AXES (per group) ═══════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Axes  (active group)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        def _refresh_unit_opts(col_var, unit_var, unit_cb):
            col = col_var.get()
            dim = "J" if col == "J" else _UNIT_DIMS.get(
                col.rsplit("/", 1)[-1].strip() if "/" in col else "")
            opts = list(_DIM_OPTS.get(dim, ["(auto)"]))
            unit_cb["values"] = opts
            cur = unit_var.get()
            if cur not in opts:
                unit_var.set("mA/cm²" if col == "J" else "(auto)")
            elif col == "J" and cur == "(auto)":
                unit_var.set("mA/cm²")
            self._auto_replot()

        def _refresh_unit_after(unit_var, unit_cb):
            chosen = unit_var.get()
            if chosen and chosen != "(auto)":
                opts = list(_DIM_OPTS["J"] if chosen.endswith("/cm²")
                            else _DIM_OPTS.get(_UNIT_DIMS.get(chosen), _ALL_UNITS))
                unit_cb["values"] = opts
            self._auto_replot()

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
            for col_var, unit_var, cb in (
                (self.x_var, self.x_unit_var, x_unit_cb),
                (self.y_var, self.y_unit_var, y_unit_cb),
            ):
                col = col_var.get()
                dim = "J" if col == "J" else _UNIT_DIMS.get(
                    col.rsplit("/", 1)[-1].strip() if "/" in col else "")
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

        # ══ PLOT RANGE ═════════════════════════════════════════════
        ttk.Label(left, text="Plot Range:", font=("", 8)).pack(anchor=tk.W, padx=4, pady=(4, 0))
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

        ttk.Label(left, text="(blank = auto)", foreground="gray",
                  font=("", 8)).pack(anchor=tk.W, padx=4)
        for _re in (_xmin, _xmax, _ymin, _ymax, _xgi, _ygi):
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

        ttk.Label(left, text="Reference Electrode:", font=("", 8)).pack(
            anchor=tk.W, padx=4, pady=(4, 0))
        self.ref_electrode_var = tk.StringVar(value="Ag/AgCl")
        _ref_cb = ttk.Combobox(
            left, textvariable=self.ref_electrode_var,
            values=["Ag/AgCl", "SCE", "SHE", "NHE", "RHE",
                    "Hg/HgO", "Hg/HgSO4 (MSE)", "Fc/Fc+", "Ag/Ag+", "Li/Li+"],
            state="readonly", width=24,
        )
        _ref_cb.pack(padx=4, pady=2)
        _ref_cb.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())

        # ══ TITLE ═════════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        _title_row = ttk.Frame(left)
        _title_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        ttk.Label(_title_row, text="Title:").pack(side=tk.LEFT)
        self.plot_title_var = tk.StringVar(value="")
        _title_entry = ttk.Entry(_title_row, textvariable=self.plot_title_var)
        _title_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
        _title_entry.bind("<Return>",   lambda e: self._auto_replot())
        _title_entry.bind("<FocusOut>", lambda e: self._auto_replot())

        # ══ LEGEND ════════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Legend  (active group)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        leg_row1 = ttk.Frame(left)
        leg_row1.pack(fill=tk.X, padx=4, pady=2)
        self.legend_show_var  = tk.BooleanVar(value=True)
        self.legend_frame_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(leg_row1, text="Show Legend",
                        variable=self.legend_show_var,
                        command=self._auto_replot).pack(side=tk.LEFT)
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
        def _on_leg_loc_select(e=None):
            if self.active_group and self.active_group in self.groups:
                self.groups[self.active_group].pop("legend_manual_pos", None)
                _leg = self.groups[self.active_group].get("legend")
                if _leg is not None:
                    _leg._loc = 0  # clear dragged tuple so it doesn't get re-saved
            self._auto_replot()
        _leg_loc_cb.bind("<<ComboboxSelected>>", _on_leg_loc_select)
        ttk.Label(left, text="(left-drag to move, dbl-click to edit labels)",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)

        # ══ GRID ═══════════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Grid  (active group)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
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

        # ══ FONT ════════════════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Font  (active group)",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        self.font_title_size_var = tk.StringVar(value="10")
        self.font_title_bold_var = tk.BooleanVar(value=False)
        self.font_label_size_var = tk.StringVar(value="10")
        self.font_label_bold_var = tk.BooleanVar(value=False)
        self.font_tick_size_var  = tk.StringVar(value="8")
        self.font_tick_bold_var  = tk.BooleanVar(value=False)
        for lbl, sz_var, bold_var in (
            ("Title:      Size", self.font_title_size_var, self.font_title_bold_var),
            ("Axis Lbl: Size",   self.font_label_size_var, self.font_label_bold_var),
            ("Tick Nos: Size",   self.font_tick_size_var,  self.font_tick_bold_var),
        ):
            _fr = ttk.Frame(left)
            _fr.pack(fill=tk.X, padx=4, pady=(2, 0))
            ttk.Label(_fr, text=lbl).pack(side=tk.LEFT)
            _e = ttk.Entry(_fr, textvariable=sz_var, width=4)
            _e.pack(side=tk.LEFT, padx=(2, 4))
            _e.bind("<Return>",   lambda e: self._auto_replot())
            _e.bind("<FocusOut>", lambda e: self._auto_replot())
            ttk.Checkbutton(_fr, text="Bold", variable=bold_var,
                            command=self._auto_replot).pack(side=tk.LEFT)
        _spc_row = ttk.Frame(left)
        _spc_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Label(_spc_row, text="Spacing: Title").pack(side=tk.LEFT)
        self.title_pad_var = tk.StringVar(value="6")
        _tpe = ttk.Entry(_spc_row, textvariable=self.title_pad_var, width=4)
        _tpe.pack(side=tk.LEFT, padx=(2, 6))
        _tpe.bind("<Return>",   lambda e: self._auto_replot())
        _tpe.bind("<FocusOut>", lambda e: self._auto_replot())
        ttk.Label(_spc_row, text="Label").pack(side=tk.LEFT)
        self.label_pad_var = tk.StringVar(value="4")
        _lpe = ttk.Entry(_spc_row, textvariable=self.label_pad_var, width=4)
        _lpe.pack(side=tk.LEFT, padx=(2, 0))
        _lpe.bind("<Return>",   lambda e: self._auto_replot())
        _lpe.bind("<FocusOut>", lambda e: self._auto_replot())

        # ══ REFERENCE LINES ════════════════════════════════════════
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Reference Lines  (active group)",
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

        # ══ PLOT BUTTON + EXPORT + LOG ══════════════════════════════
        _btn2_row = ttk.Frame(left)
        _btn2_row.pack(fill=tk.X, padx=4, pady=6)
        ttk.Button(_btn2_row, text="Plot Active Group",
                   command=self._auto_replot).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(_btn2_row, text="Export Group Cycles",
                   command=self._export_group_cycles_excel).pack(side=tk.LEFT)
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Log", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        log_f = ttk.Frame(left)
        log_f.pack(fill=tk.X, padx=4, pady=2)
        self.log_text = tk.Text(log_f, height=4, state=tk.DISABLED,
                                wrap=tk.WORD, relief=tk.SUNKEN, borderwidth=1)
        log_sc = ttk.Scrollbar(log_f, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_sc.set)
        log_sc.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # ── Right panel ────────────────────────────────────────────
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
                   command=self._unzoom_group_view).pack(side=tk.LEFT, padx=6, pady=3)
        self._zoom_bar.grid(row=1, column=0, sticky="ew")
        self._zoom_bar.grid_remove()

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

        # Drop indicator: thin colored bar shown during drag-to-reorder
        self._drop_line = tk.Frame(self._plots_frame, bg="#1a73e8", height=3)

        self._placeholder = ttk.Label(
            self._plots_frame,
            text="Create groups and assign files to display overlay plots here.",
            foreground="gray", font=("", 10),
        )
        self._placeholder.grid(row=0, column=0, columnspan=2, pady=60)

    # ════════════════════════════════════════════════════════════════
    # Group management
    # ════════════════════════════════════════════════════════════════
    def _rebuild_group_listbox(self):
        """Rebuild the groups CheckableListbox from self.groups (preserves visibility states)."""
        self.group_listbox.clear()
        for gn, gentry in self.groups.items():
            vis = not gentry.get("hidden", False)
            self.group_listbox.insert(tk.END, gn, checked=vis)
        if self.active_group and self.active_group in self.groups:
            idx = list(self.groups.keys()).index(self.active_group)
            self._loading_files = True
            try:
                self.group_listbox.selection_clear(0, tk.END)
                self.group_listbox.selection_set(idx)
            finally:
                self._loading_files = False

    def _new_group(self):
        name = simpledialog.askstring("New Group", "Group name:", parent=self)
        if not name:
            return
        name = name.strip()
        if not name or name in self.groups:
            return
        self.groups[name] = {"files": [], "reflines": []}
        self._rebuild_group_listbox()
        self._create_group_figure(name)
        self._save_active_group_state()
        self._switch_active_group(name)

    def _rename_group(self):
        sel = self.group_listbox.curselection()
        if not sel:
            return
        old = list(self.groups.keys())[sel[0]]
        new = simpledialog.askstring("Rename Group", "New name:",
                                     initialvalue=old, parent=self)
        if not new:
            return
        new = new.strip()
        if not new or new == old or new in self.groups:
            return
        # Rebuild OrderedDict preserving order
        new_groups = OrderedDict()
        for k, v in self.groups.items():
            new_groups[new if k == old else k] = v
        self.groups = new_groups
        if self.active_group == old:
            self.active_group = new
        self._rebuild_group_listbox()
        gentry = self.groups.get(new, {})
        if "ax" in gentry:
            gentry["ax"].set_title(gentry.get("custom_title", new), fontsize=9)
            gentry["canvas"].draw_idle()

    def _delete_group(self):
        sel = self.group_listbox.curselection()
        if not sel:
            return
        name = list(self.groups.keys())[sel[0]]
        self._destroy_group_figure(name)
        del self.groups[name]
        if self.active_group == name:
            self.active_group = None
            self.group_files_lb.clear()
        if not self.groups:
            self._rebuild_group_listbox()
            self._placeholder.grid(row=0, column=0, columnspan=2, pady=60)
        else:
            self._rebuild_group_listbox()
            self._relayout_figures()
            new_idx = min(sel[0], self.group_listbox.size() - 1)
            if new_idx >= 0:
                self._switch_active_group(list(self.groups.keys())[new_idx])

    def _add_files_to_group(self):
        if not self.active_group:
            return
        sel = self.file_listbox.curselection()
        if not sel:
            return
        gentry = self.groups[self.active_group]
        for i in sel:
            fname = self.file_listbox.get(i)
            if fname and fname not in gentry["files"]:
                gentry["files"].append(fname)
                # Seed independent per-group state from current file defaults
                fentry = self.files.get(fname, {})
                fp = gentry.setdefault("file_params", {}).setdefault(fname, {})
                fp.setdefault("selected_cycles", list(fentry.get("selected_cycles", [])))
                fp.setdefault("r_sol", float(fentry.get("r_sol", 0.0)))
                fp.setdefault("e_ref", float(fentry.get("e_ref", 0.0)))
        self._update_group_files_lb()
        self._update_group_column_combos()
        self._plot_group(self.active_group)

    def _remove_files_from_group(self):
        if not self.active_group:
            return
        sel = self.group_files_lb.curselection()
        if not sel:
            return
        gentry = self.groups[self.active_group]
        files = gentry["files"]
        file_hidden = gentry.get("file_hidden", {})
        for i in sorted(sel, reverse=True):
            if i < len(files):
                fname = files[i]
                files.pop(i)
                file_hidden.pop(fname, None)
        self._update_group_files_lb()
        self._update_group_column_combos()
        self._plot_group(self.active_group)

    def _update_group_files_lb(self):
        self.group_files_lb.clear()
        if not self.active_group or self.active_group not in self.groups:
            return
        gentry = self.groups[self.active_group]
        file_hidden = gentry.get("file_hidden", {})
        for fname in gentry["files"]:
            vis = not file_hidden.get(fname, False)
            self.group_files_lb.insert(tk.END, fname, checked=vis)

    def _update_group_column_combos(self):
        """Populate x/y combos with union of columns from group's files."""
        if not self.active_group or self.active_group not in self.groups:
            return
        gentry = self.groups[self.active_group]
        all_cols, seen = [], set()
        for fname in gentry["files"]:
            fentry = self.files.get(fname)
            if fentry is None:
                continue
            df = fentry["df"]
            try:
                _has_area = float(fentry.get("area", "")) > 0
            except (ValueError, TypeError):
                _has_area = False
            for c in df.columns:
                if c not in seen:
                    all_cols.append(c)
                    seen.add(c)
            if _has_area and "J" not in seen:
                all_cols.append("J")
                seen.add("J")
        self.x_combo["values"] = all_cols
        self.y_combo["values"] = all_cols
        x_col = gentry.get("x_col")
        y_col = gentry.get("y_col")
        if x_col and x_col in all_cols:
            self.x_var.set(x_col)
        elif all_cols:
            self.x_var.set(_default_xcol(all_cols))
        if y_col and y_col in all_cols:
            self.y_var.set(y_col)
        elif all_cols:
            self.y_var.set(_default_ycol(all_cols, self.x_var.get()))
        fn = getattr(self, '_do_refresh_unit_combos', None)
        if fn:
            fn()

    def _on_group_select(self, event):
        if self._loading_files:
            return
        sel = self.group_listbox.curselection()
        if not sel:
            return
        keys = list(self.groups.keys())
        if sel[0] >= len(keys):
            return
        name = keys[sel[0]]
        if name != self.active_group:
            self._save_active_state()       # flush current file's corrections into group params
            self._save_active_group_state()
            self._switch_active_group(name)

    def _load_files(self):
        """Override: reset highlight state before and after loading files."""
        self._plot_highlight = False
        self._active_cycle = None
        super()._load_files()
        self._plot_highlight = False
        self._active_cycle = None
        for gname in list(self.groups.keys()):
            self._apply_highlight_to_group(gname)

    def _on_file_select(self, event):
        if self._loading_files:
            return
        self._plot_highlight = True
        self._active_cycle = None   # listbox → highlight whole file
        super()._on_file_select(event)

    def _on_group_file_select(self, event):
        if self._loading_files:
            return
        sel = self.group_files_lb.curselection()
        if not sel:
            return
        gf_list = self.groups.get(self.active_group, {}).get("files", [])
        if sel[0] >= len(gf_list):
            return
        fname = gf_list[sel[0]]
        if fname in self.files and fname != self.active_file:
            self._plot_highlight = True
            self._active_cycle = None   # group listbox → highlight whole file
            self._save_active_state()
            self._switch_active_file(fname)
            file_keys = list(self.files.keys())
            if fname in file_keys:
                idx = file_keys.index(fname)
                self.file_listbox.selection_clear(0, tk.END)
                self.file_listbox.selection_set(idx)

    # ════════════════════════════════════════════════════════════════
    # FileManagerMixin overrides
    # ════════════════════════════════════════════════════════════════
    def _on_file_visibility_change(self, short, visible):
        if short not in self.files:
            return
        self.files[short]["hidden"] = not visible
        self._replot_groups_for_file(short)

    def _on_file_reorder(self, new_order):
        new_files = OrderedDict()
        for name in new_order:
            if name in self.files:
                new_files[name] = self.files[name]
        for name, entry in self.files.items():
            if name not in new_files:
                new_files[name] = entry
        self.files = new_files
        # Replot all groups
        for gname in self.groups:
            self._plot_group(gname)

    def _remove_file(self):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        idx  = sel[0]
        short = self.file_listbox.get(idx)
        self.file_listbox.delete(idx)
        # Remove from all groups
        affected = []
        for gname, gentry in self.groups.items():
            if short in gentry["files"]:
                gentry["files"].remove(short)
                affected.append(gname)
        del self.files[short]
        if self.active_file == short:
            self.active_file = None
            self._suppress_replot = True
            self._populate_cycle_checkboxes([], [])
            self._suppress_replot = False
            self.r_sol_var.set("0")
            self.e_ref_var.set("0")
            self.area_var.set("")
            if self.files:
                self.file_listbox.selection_set(0)
                self._switch_active_file(list(self.files.keys())[0])
        for gname in affected:
            self._update_group_files_lb()
            self._plot_group(gname)

    def _clear_plot(self):
        pass  # no single plot to clear; groups handle their own figures

    def _replot_groups_for_file(self, fname):
        """Replot every group that contains fname."""
        for gname, gentry in self.groups.items():
            if fname in gentry.get("files", []):
                self._plot_group(gname)

    # ════════════════════════════════════════════════════════════════
    # Figure creation / destruction
    # ════════════════════════════════════════════════════════════════
    def _create_group_figure(self, group_name):
        gentry = self.groups.get(group_name)
        if gentry is None or "fig" in gentry:
            return
        self._placeholder.grid_remove()
        panel_ref = self

        frame = tk.Frame(self._plots_frame, relief="groove", bd=2)
        header = tk.Frame(frame, bg=_GROUP_HDR_BG, cursor="fleur")
        header.pack(fill=tk.X, side=tk.TOP)
        lbl = tk.Label(header, text=f"⠿  {group_name}",
                       bg=_GROUP_HDR_BG, font=("", 9, "bold"), anchor=tk.W)
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
            def home(tb_self, *args):
                panel_ref._reset_group_view(group_name)

        _tb = _Toolbar(canvas, tb_frame, pack_toolbar=False)
        _tb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        _tb.update()
        tk.Button(tb_frame, text="Copy",
                  command=lambda f=fig: copy_figure_to_clipboard(f),
                  relief=tk.RAISED, borderwidth=1, padx=6).pack(
            side=tk.LEFT, padx=(4, 2), pady=1)

        def _fwd_scroll(e):
            self._right_canvas.yview_scroll(-1 * (e.delta // 120), "units")
        frame.bind("<MouseWheel>",    _fwd_scroll)
        header.bind("<MouseWheel>",   _fwd_scroll)
        tb_frame.bind("<MouseWheel>", _fwd_scroll)

        # Drag on the header strip → reorder group subplots; double-click → zoom
        for _w in (header, lbl):
            _w.bind("<ButtonPress-1>",   lambda e, g=group_name: self._on_frame_press(e, g))
            _w.bind("<B1-Motion>",       lambda e, g=group_name: self._on_frame_drag(e, g))
            _w.bind("<ButtonRelease-1>", lambda e, g=group_name: self._on_frame_release(e, g))
            _w.bind("<Double-Button-1>", lambda e, g=group_name: self._toggle_zoom(g))

        def _activate(e=None):
            self._activate_group(group_name)
        frame.bind("<Button-1>",    _activate, add="+")
        tb_frame.bind("<Button-1>", _activate, add="+")

        canvas.mpl_connect("scroll_event",         lambda ev: self._on_scroll(ev, group_name))
        canvas.mpl_connect("button_press_event",   lambda ev: self._on_press(ev, group_name))
        canvas.mpl_connect("button_release_event", lambda ev: self._on_release(ev, group_name))
        canvas.mpl_connect("motion_notify_event",  lambda ev: self._on_motion(ev, group_name))

        gentry.update({
            "fig":        fig,
            "ax":         ax,
            "canvas":     canvas,
            "plot_frame": frame,
            "hdr_frame":  header,
            "hdr_label":  lbl,
            "legend":     None,
            "leg_size":   8.0,
            "auto_xlim":  None,
            "auto_ylim":  None,
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
        ax.set_title(group_name, fontsize=9)
        ax.set_xlabel("Potential (V)")
        ax.set_ylabel("Current (mA)")
        canvas.draw()
        self._relayout_figures()

    def _destroy_group_figure(self, group_name):
        frame = self.groups.get(group_name, {}).get("plot_frame")
        if frame is not None:
            frame.destroy()

    # ════════════════════════════════════════════════════════════════
    # Right-panel layout
    # ════════════════════════════════════════════════════════════════
    def _on_right_canvas_configure(self, event):
        if self._zoom_group:
            self._right_canvas.itemconfig(self._plots_win,
                                          width=event.width, height=event.height)

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
        valid = [(gn, self.groups[gn]) for gn in self.groups
                 if "plot_frame" in self.groups[gn]
                 and not self.groups[gn].get("hidden", False)]

        if self._zoom_group and any(gn == self._zoom_group for gn, _ in valid):
            for gn, gentry in valid:
                if gn == self._zoom_group:
                    gentry["plot_frame"].grid(row=0, column=0, columnspan=MAX_COLS,
                                             sticky="nsew", padx=4, pady=4)
                else:
                    gentry["plot_frame"].grid_remove()
            self._plots_frame.rowconfigure(0, weight=1)
            return

        for gn in self.groups:
            pf = self.groups[gn].get("plot_frame")
            if pf is not None:
                pf.grid_remove()
        for i, (gname, gentry) in enumerate(valid):
            row = i // MAX_COLS
            col = i % MAX_COLS
            gentry["plot_frame"].grid(row=row, column=col, columnspan=1,
                                     sticky="nsew", padx=4, pady=4)
        n_rows = (len(valid) + MAX_COLS - 1) // MAX_COLS if valid else 0
        for r in range(n_rows):
            self._plots_frame.rowconfigure(r, weight=0)

    # ════════════════════════════════════════════════════════════════
    # Drag-to-reorder: drag a group subplot frame to change grid order
    # ════════════════════════════════════════════════════════════════
    def _on_frame_press(self, event, group_name):
        self._drag = {
            "group":      group_name,
            "start_x":    event.x_root,
            "start_y":    event.y_root,
            "active":     False,
            "target":     None,
            "target_top": True,
        }

    def _on_frame_drag(self, event, group_name):
        drag = self._drag
        if drag is None or drag["group"] != group_name:
            return
        if not drag["active"]:
            if abs(event.x_root - drag["start_x"]) + abs(event.y_root - drag["start_y"]) < 6:
                return
            drag["active"] = True

        # Detect which visible group frame the cursor is over
        target = None
        target_top = True
        for gn, gentry in self.groups.items():
            if gn == group_name or gentry.get("hidden"):
                continue
            pf = gentry.get("plot_frame")
            if pf is None:
                continue
            x0 = pf.winfo_rootx()
            y0 = pf.winfo_rooty()
            w  = pf.winfo_width()
            h  = pf.winfo_height()
            if x0 <= event.x_root <= x0 + w and y0 <= event.y_root <= y0 + h:
                target = gn
                target_top = (event.y_root - y0) < h / 2
                break
        drag["target"]     = target
        drag["target_top"] = target_top

        # Move / show the drop-indicator line
        if target is not None:
            pf = self.groups[target]["plot_frame"]
            rx = pf.winfo_x()
            ry = pf.winfo_y()
            rw = pf.winfo_width()
            rh = pf.winfo_height()
            line_y = ry if target_top else ry + rh - 3
            self._drop_line.place(x=rx, y=line_y, width=rw, height=3)
            self._drop_line.lift()
        else:
            self._drop_line.place_forget()

    def _on_frame_release(self, event, group_name):
        drag = self._drag
        self._drag = None
        self._drop_line.place_forget()
        if drag is None or not drag["active"]:
            return
        target = drag.get("target")
        if target is None or target == group_name:
            return
        self._reorder_groups(group_name, target, before=drag.get("target_top", True))

    def _reorder_groups(self, from_group, to_group, *, before=True):
        """Move from_group to just before (or after) to_group in self.groups."""
        keys = list(self.groups.keys())
        if from_group not in keys or to_group not in keys:
            return
        keys.remove(from_group)
        to_idx = keys.index(to_group)
        keys.insert(to_idx if before else to_idx + 1, from_group)
        self.groups = OrderedDict((k, self.groups[k]) for k in keys)
        self._rebuild_group_listbox()
        self._relayout_figures()

    def _on_group_visibility_change(self, group_name, visible):
        """Called when a group's checkbox is toggled — hide/show its subplot."""
        gentry = self.groups.get(group_name)
        if gentry is None:
            return
        gentry["hidden"] = not visible
        self._relayout_figures()
        self._right_canvas.after(50, lambda: self._right_canvas.configure(
            scrollregion=self._right_canvas.bbox("all")))

    def _on_group_reorder(self, new_names):
        """Called when CheckableListbox drag-to-reorder finishes for groups."""
        self.groups = OrderedDict((n, self.groups[n]) for n in new_names if n in self.groups)
        self._relayout_figures()

    def _on_group_file_visibility_change(self, fname, visible):
        """Called when a file's checkbox is toggled in the group files list."""
        if not self.active_group or self.active_group not in self.groups:
            return
        gentry = self.groups[self.active_group]
        gentry.setdefault("file_hidden", {})[fname] = not visible
        self._plot_group(self.active_group)

    def _on_group_file_reorder(self, new_names):
        """Called when CheckableListbox drag-to-reorder finishes for files in a group."""
        if not self.active_group or self.active_group not in self.groups:
            return
        gentry = self.groups[self.active_group]
        gentry["files"] = [n for n in new_names if n in gentry.get("files", [])]
        # Reset legend order so it rebuilds from new file rank
        gentry["legend_order"] = []
        self._plot_group(self.active_group)

    def _reset_group_view(self, group_name):
        gentry = self.groups.get(group_name)
        if gentry and gentry.get("auto_xlim") is not None:
            gentry["ax"].set_xlim(gentry["auto_xlim"])
            gentry["ax"].set_ylim(gentry["auto_ylim"])
            gentry["canvas"].draw_idle()

    def _toggle_zoom(self, group_name):
        """Toggle full-screen view for group (called from header double-click)."""
        if self._zoom_group is None:
            self._zoom_group_view(group_name)
        else:
            self._unzoom_group_view()

    def _auto_set_initial_size(self):
        """Set default figure size to fill the right panel on first show."""
        w = self._right_canvas.winfo_width()
        h = self._right_canvas.winfo_height()
        if w <= 1 or h <= 1:
            self.after(100, self._auto_set_initial_size)
            return
        dpi = 100
        try:
            _ncols = max(1, int(self._grid_cols_var.get()))
        except (ValueError, AttributeError):
            _ncols = 2
        plot_w = max(3.0, (w / _ncols - 30) / dpi)
        plot_h = max(2.0, round(plot_w * 0.6, 1))
        self.plot_w_var.set(f"{plot_w:.1f}")
        self.plot_h_var.set(f"{plot_h:.1f}")
        self._apply_plot_size()

    # ════════════════════════════════════════════════════════════════
    # Correction overrides (group-scoped per-file R_sol / E_ref)
    # ════════════════════════════════════════════════════════════════
    def _apply_correction(self):
        """Save R_sol/E_ref to active group's file_params and replot."""
        if not self.active_file or self.active_file not in self.files:
            return
        if not self.active_group or self.active_group not in self.groups:
            return
        try:
            r_sol = float(self.r_sol_var.get())
        except ValueError:
            r_sol = 0.0
        try:
            e_ref = float(self.e_ref_var.get())
        except ValueError:
            e_ref = 0.0
        gentry = self.groups[self.active_group]
        fp = gentry.setdefault("file_params", {}).setdefault(self.active_file, {})
        fp["r_sol"] = r_sol
        fp["e_ref"] = e_ref
        self._auto_replot()

    def _reset_correction(self):
        """Clear group-scoped corrections for the active (group, file) pair."""
        if not self.active_file or not self.active_group:
            return
        gentry = self.groups.get(self.active_group)
        if gentry:
            gentry.setdefault("file_params", {}).pop(self.active_file, None)
        self.r_sol_var.set("0")
        self.e_ref_var.set("0")
        self._auto_replot()


    def _apply_plot_size(self, event=None):
        """Resize all group figures to the current plot_w_var × plot_h_var (inches)."""
        try:
            w = float(self.plot_w_var.get())
            h = float(self.plot_h_var.get())
        except ValueError:
            return
        w = max(2.0, min(50.0, w))
        h = max(1.5, min(50.0, h))
        dpi = 100
        # Reset any width constraint from zoom mode before resizing
        self._right_canvas.itemconfig(self._plots_win, width=0, height=0)
        for gentry in self.groups.values():
            fig = gentry.get("fig")
            cv  = gentry.get("canvas")
            if fig and cv:
                fig.set_size_inches(w, h)
                cv.get_tk_widget().config(width=int(w * dpi), height=int(h * dpi))
                _legs = [a.get_legend() for a in fig.get_axes() if a.get_legend() is not None]
                for _l in _legs: _l.set_visible(False)
                fig.tight_layout(pad=0.5)
                fig.set_layout_engine('none')
                for _l in _legs: _l.set_visible(True)
                cv.draw_idle()

        def _update_scrollregion():
            self._plots_frame.update_idletasks()
            self._right_canvas.configure(
                scrollregion=self._right_canvas.bbox("all"))
        self._right_canvas.after(100, _update_scrollregion)

    def _zoom_group_view(self, group_name):
        self._zoom_group = group_name
        self._zoom_bar.grid()
        self.update_idletasks()
        w = self._right_canvas.winfo_width()
        h = self._right_canvas.winfo_height()
        if w > 1 and h > 1:
            self._right_canvas.itemconfig(self._plots_win, width=w, height=h)
            gentry = self.groups.get(group_name, {})
            fig    = gentry.get("fig")
            cv     = gentry.get("canvas")
            if fig and cv:
                dpi = fig.get_dpi()
                # Subtract header strip (~28px), toolbar (~32px), padding (~12px)
                fig_w = max(100, w - 8)
                fig_h = max(100, h - 72)
                fig.set_size_inches(fig_w / dpi, fig_h / dpi)
                cv.get_tk_widget().config(width=fig_w, height=fig_h)
                cv.draw_idle()
        self._relayout_figures()
        self._right_canvas.xview_moveto(0)
        self._right_canvas.yview_moveto(0)

    def _unzoom_group_view(self):
        self._zoom_group = None
        self._zoom_bar.grid_remove()
        self._right_canvas.itemconfig(self._plots_win, width=0, height=0)
        self._apply_plot_size()
        self._relayout_figures()
        self._plots_frame.update_idletasks()
        self._right_canvas.configure(scrollregion=self._right_canvas.bbox("all"))

    def _activate_group(self, group_name):
        items = list(self.groups.keys())
        if group_name in items:
            idx = items.index(group_name)
            self.group_listbox.selection_clear(0, tk.END)
            self.group_listbox.selection_set(idx)
        if group_name != self.active_group:
            self._save_active_group_state()
            self._switch_active_group(group_name)

    # ════════════════════════════════════════════════════════════════
    # Copy / Paste group display settings
    # ════════════════════════════════════════════════════════════════
    _COPYABLE_KEYS = (
        "font_title_size", "font_title_bold",
        "font_label_size", "font_label_bold",
        "font_tick_size",  "font_tick_bold",
        "title_pad", "label_pad",
        "ref_electrode",
        "x_grid", "y_grid", "grid_style", "grid_color", "grid_linewidth",
        "legend_frame", "leg_size", "legend_loc", "legend_show",
    )

    def _copy_group_settings(self):
        if not self.active_group or self.active_group not in self.groups:
            return
        self._save_active_group_state()
        g = self.groups[self.active_group]
        self._copied_group_params = {k: g.get(k) for k in self._COPYABLE_KEYS}
        self._copied_group_params["reflines"] = list(g.get("reflines", []))

    def _paste_group_settings(self):
        if not self._copied_group_params:
            return
        if not self.active_group or self.active_group not in self.groups:
            return
        g = self.groups[self.active_group]
        p = self._copied_group_params
        for k in self._COPYABLE_KEYS:
            if k in p and p[k] is not None:
                g[k] = p[k]
        g["reflines"] = list(p.get("reflines", []))
        # Refresh UI vars for the active group
        self.font_title_size_var.set(g.get("font_title_size", "10"))
        self.font_title_bold_var.set(g.get("font_title_bold", False))
        self.font_label_size_var.set(g.get("font_label_size", "10"))
        self.font_label_bold_var.set(g.get("font_label_bold", False))
        self.font_tick_size_var.set(g.get("font_tick_size",  "8"))
        self.font_tick_bold_var.set(g.get("font_tick_bold",  False))
        self.title_pad_var.set(g.get("title_pad", "6"))
        self.label_pad_var.set(g.get("label_pad", "4"))
        self.ref_electrode_var.set(g.get("ref_electrode", "Ag/AgCl"))
        self.x_grid_var.set(g.get("x_grid", False))
        self.y_grid_var.set(g.get("y_grid", False))
        self.grid_style_var.set(g.get("grid_style", "dashed"))
        self.grid_color_var.set(g.get("grid_color", "gray"))
        self.grid_linewidth_var.set(g.get("grid_linewidth", "0.8"))
        self.legend_frame_var.set(g.get("legend_frame", True))
        _ls = g.get("leg_size", 8.0)
        self.legend_size_var.set(str(int(_ls) if float(_ls) == int(_ls) else _ls))
        self.legend_loc_var.set(g.get("legend_loc", "best"))
        self.legend_show_var.set(g.get("legend_show", True))
        self._refresh_reflines_lb()
        self._auto_replot()

    # ════════════════════════════════════════════════════════════════
    # Export group cycles to Excel (one file per source file)
    # ════════════════════════════════════════════════════════════════
    def _export_group_cycles_excel(self):
        import os
        import pandas as pd
        from tkinter import filedialog, messagebox
        if not self.active_group or self.active_group not in self.groups:
            messagebox.showinfo("Info", "Select a group first.")
            return
        self._save_active_group_state()
        gentry = self.groups[self.active_group]
        files = gentry.get("files", [])
        if not files:
            messagebox.showinfo("Info", "No files in the active group.")
            return
        dir_path = filedialog.askdirectory(title="Select folder to save Excel files")
        if not dir_path:
            return
        exported, errors = [], []
        for fname in files:
            fentry = self.files.get(fname)
            if fentry is None:
                continue
            fp     = gentry.get("file_params", {}).get(fname, {})
            cycles = fp.get("selected_cycles", fentry.get("selected_cycles", []))
            if not cycles:
                continue
            _rsol = fp.get("r_sol", 0.0)
            _eref = fp.get("e_ref", 0.0)
            df = fentry["df_raw"].copy()
            if _rsol != 0.0 and "Ewe/V" in df.columns and "I/mA" in df.columns:
                df["Ewe/V"] = df["Ewe/V"] - (df["I/mA"] / 1000.0) * _rsol
            if _eref != 0.0 and "Ewe/V" in df.columns:
                df["Ewe/V"] = df["Ewe/V"] + _eref
            if "cycle number" not in df.columns:
                continue
            base = os.path.splitext(fname)[0]
            out_path = os.path.join(dir_path, f"{base}.xlsx")
            try:
                with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                    for c in cycles:
                        sub = df[df["cycle number"] == c].reset_index(drop=True)
                        sub.to_excel(writer, sheet_name=f"C{c}", index=False)
                exported.append(fname)
            except Exception as exc:
                errors.append(f"{fname}: {exc}")
        if exported:
            msg = f"Exported {len(exported)} file(s) to:\n{dir_path}"
            if errors:
                msg += f"\n\nErrors:\n" + "\n".join(errors)
            messagebox.showinfo("Export complete", msg)
        elif errors:
            messagebox.showerror("Export failed", "\n".join(errors))
        else:
            messagebox.showinfo("Info", "No cycles selected for any file in the group.")

    # ════════════════════════════════════════════════════════════════
    # State save / restore
    # ════════════════════════════════════════════════════════════════
    def _save_active_state(self):
        """Save per-file UI state (overrides FileManagerMixin — no axis saving)."""
        if not self.active_file or self.active_file not in self.files:
            return
        entry = self.files[self.active_file]
        entry["selected_cycles"] = self._selected_cycles()
        # Corrections are group-scoped: save to the active group's file_params
        if self.active_group and self.active_group in self.groups:
            gentry = self.groups[self.active_group]
            fp = gentry.setdefault("file_params", {}).setdefault(self.active_file, {})
            try:    fp["r_sol"] = float(self.r_sol_var.get())
            except ValueError: fp.setdefault("r_sol", 0.0)
            try:    fp["e_ref"] = float(self.e_ref_var.get())
            except ValueError: fp.setdefault("e_ref", 0.0)
            fp["selected_cycles"] = self._selected_cycles()
        entry["area"]           = self.area_var.get()
        entry["color"]          = _COLOR_HEX.get(self.file_color_var.get(), "#1f77b4")
        entry["linewidth"]      = self.linewidth_var.get()
        entry["plot_style"]     = self.plot_style_var.get()
        entry["cycle_gradient"] = self.cycle_gradient_var.get()
        entry["cycle_reverse"]  = self.cycle_reverse_var.get()
        entry["lightness_step"] = self.lightness_step_var.get()

    def _save_active_group_state(self):
        if not self.active_group or self.active_group not in self.groups:
            return
        g = self.groups[self.active_group]
        g["x_col"]          = self.x_var.get()
        g["y_col"]          = self.y_var.get()
        g["x_unit"]         = self.x_unit_var.get()
        g["y_unit"]         = self.y_unit_var.get()
        g["x_min"]          = self.x_min_var.get()
        g["x_max"]          = self.x_max_var.get()
        g["y_min"]          = self.y_min_var.get()
        g["y_max"]          = self.y_max_var.get()
        g["x_grid_int"]     = self.x_grid_int_var.get()
        g["y_grid_int"]     = self.y_grid_int_var.get()
        g["x_flip"]         = self.x_flip_var.get()
        g["y_flip"]         = self.y_flip_var.get()
        g["ref_electrode"]  = self.ref_electrode_var.get()
        g["legend_show"]    = self.legend_show_var.get()
        g["legend_frame"]   = self.legend_frame_var.get()
        try:    g["leg_size"] = float(self.legend_size_var.get())
        except: pass
        g["legend_loc"]     = self.legend_loc_var.get()
        g["x_grid"]         = self.x_grid_var.get()
        g["y_grid"]         = self.y_grid_var.get()
        g["grid_style"]     = self.grid_style_var.get()
        g["grid_color"]     = self.grid_color_var.get()
        g["grid_linewidth"] = self.grid_linewidth_var.get()
        g["font_title_size"]= self.font_title_size_var.get()
        g["font_title_bold"]= self.font_title_bold_var.get()
        g["font_label_size"]= self.font_label_size_var.get()
        g["font_label_bold"]= self.font_label_bold_var.get()
        g["font_tick_size"] = self.font_tick_size_var.get()
        g["font_tick_bold"] = self.font_tick_bold_var.get()
        g["title_pad"]      = self.title_pad_var.get()
        g["label_pad"]      = self.label_pad_var.get()
        g["custom_title"]   = self.plot_title_var.get()

    # ── Session save / restore ────────────────────────────────────────
    def get_session_state(self, data_store: dict) -> dict:
        self._save_active_state()
        self._save_active_group_state()
        files_list  = [_sm.serialise_file_entry(n, e, data_store)
                       for n, e in self.files.items()]
        groups_list = [_sm.serialise_group_entry(n, g)
                       for n, g in self.groups.items()]
        return {
            "active_file":   self.active_file,
            "active_group":  self.active_group,
            "plot_w_var":    self.plot_w_var.get(),
            "plot_h_var":    self.plot_h_var.get(),
            "grid_cols_var": self._grid_cols_var.get(),
            "files":  files_list,
            "groups": groups_list,
        }

    def restore_session_state(self, state: dict, data_store: dict) -> None:
        old = self._suppress_replot
        self._suppress_replot = True

        # Clear existing groups (destroy their frames)
        for gentry in self.groups.values():
            pf = gentry.get("plot_frame")
            if pf is not None:
                try:
                    pf.destroy()
                except Exception:
                    pass
        self.groups.clear()
        self.active_group = None
        self.group_listbox.clear()
        self._zoom_group = None

        # Clear existing files
        self.files.clear()
        self.active_file = None
        self.file_listbox.clear()
        self._populate_cycle_checkboxes([], [])

        # Restore panel-level vars
        try:
            self.plot_w_var.set(state.get("plot_w_var", "10.5"))
            self.plot_h_var.set(state.get("plot_h_var", "5.5"))
            self._grid_cols_var.set(state.get("grid_cols_var", "2"))
        except Exception:
            pass

        # Restore files (entries only; no per-file figures in ME2)
        for rec in state.get("files", []):
            name = rec.get("name", "")
            df_raw = data_store.get(rec.get("data_hash", ""))
            if df_raw is None or not name:
                continue
            entry = {
                "path":           rec.get("path", ""),
                "df_raw":         df_raw.copy(),
                "df":             df_raw.copy(),
                "selected_cycles": rec.get("selected_cycles", []),
                "r_sol":          rec.get("r_sol", 0.0),
                "e_ref":          rec.get("e_ref", 0.0),
                "area":           rec.get("area", ""),
                "color":          rec.get("color", "#1f77b4"),
                "marker":         rec.get("marker", "o"),
                "cycle_gradient": rec.get("cycle_gradient", True),
                "cycle_reverse":  rec.get("cycle_reverse", False),
                "lightness_step": rec.get("lightness_step", "0.15"),
                "hidden":         rec.get("hidden", False),
                "linewidth":      rec.get("linewidth", "1.5"),
                "plot_style":     rec.get("plot_style", "Line"),
            }
            self.files[name] = entry
            self.file_listbox.insert(tk.END, name,
                                     checked=not rec.get("hidden", False))

        # Restore groups (create group dicts + figures)
        for grec in state.get("groups", []):
            gname = grec.get("name", "")
            if not gname:
                continue
            gentry: dict = {"files": [], "reflines": []}
            # Restore all non-runtime saved keys
            for k, v in grec.items():
                if k == "name":
                    continue
                gentry[k] = v
            # Ensure reflines are tuples
            gentry["reflines"] = [tuple(r) for r in gentry.get("reflines", [])]
            self.groups[gname] = gentry
            self.group_listbox.insert(tk.END, gname,
                                      checked=not grec.get("hidden", False))
            self._create_group_figure(gname)

        # Switch to saved active group and file
        self._suppress_replot = old
        active_group = state.get("active_group")
        active_file  = state.get("active_file")
        if active_group and active_group in self.groups:
            keys = list(self.groups.keys())
            self.group_listbox.selection_set(keys.index(active_group))
            self._switch_active_group(active_group)
        elif self.groups:
            first_g = next(iter(self.groups))
            self.group_listbox.selection_set(0)
            self._switch_active_group(first_g)

        if active_file and active_file in self.files:
            keys_f = list(self.files.keys())
            self._loading_files = True
            try:
                self.file_listbox.selection_set(keys_f.index(active_file))
            finally:
                self._loading_files = False
            self._switch_active_file(active_file)
        elif self.files:
            self._loading_files = True
            try:
                self.file_listbox.selection_set(0)
            finally:
                self._loading_files = False
            self._switch_active_file(next(iter(self.files)))

        self._apply_plot_size()
        self._relayout_figures()

    def _switch_active_file(self, short):
        """Override: update only per-file UI, not group axes."""
        self.active_file = short
        entry = self.files[short]
        # Load corrections from the active group's file_params (group-scoped)
        fp = {}
        if self.active_group and self.active_group in self.groups:
            fp = self.groups[self.active_group].get("file_params", {}).get(short, {})
        self.r_sol_var.set(str(fp.get("r_sol", 0.0)))
        self.e_ref_var.set(str(fp.get("e_ref", 0.0)))
        self.area_var.set(entry.get("area", ""))
        color = entry.get("color", "#1f77b4")
        cname = next((n for n, h in _COLOR_HEX.items() if h == color), "Blue")
        self.file_color_var.set(cname)
        self.linewidth_var.set(entry.get("linewidth", "1.5"))
        self.plot_style_var.set(entry.get("plot_style", "Line"))
        self.cycle_gradient_var.set(entry.get("cycle_gradient", True))
        self.cycle_reverse_var.set(entry.get("cycle_reverse", False))
        self.lightness_step_var.set(entry.get("lightness_step", "0.15"))

        old = self._suppress_replot
        self._suppress_replot = True
        df = entry["df"]
        if "cycle number" in df.columns:
            cycles = sorted(int(c) for c in df["cycle number"].unique())
            sel_cycles = fp.get("selected_cycles", entry.get("selected_cycles", []))
            self._populate_cycle_checkboxes(cycles, sel_cycles)
        else:
            self._populate_cycle_checkboxes([], [])
        self._suppress_replot = old

        self._auto_replot()

    def _switch_active_group(self, group_name):
        self.active_group = group_name
        g = self.groups[group_name]
        g.setdefault("x_col",           None)
        g.setdefault("y_col",           None)
        g.setdefault("x_unit",          "V")
        g.setdefault("y_unit",          "mA")
        g.setdefault("x_min",           "")
        g.setdefault("x_max",           "")
        g.setdefault("y_min",           "")
        g.setdefault("y_max",           "")
        g.setdefault("x_grid_int",      "0")
        g.setdefault("y_grid_int",      "0")
        g.setdefault("x_flip",          False)
        g.setdefault("y_flip",          False)
        g.setdefault("ref_electrode",   "Ag/AgCl")
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
        g.setdefault("file_params",     {})   # per-(group,file) corrections

        self.x_unit_var.set(g["x_unit"])
        self.y_unit_var.set(g["y_unit"])
        self.x_min_var.set(g["x_min"])
        self.x_max_var.set(g["x_max"])
        self.y_min_var.set(g["y_min"])
        self.y_max_var.set(g["y_max"])
        self.x_grid_int_var.set(g["x_grid_int"])
        self.y_grid_int_var.set(g["y_grid_int"])
        self.x_flip_var.set(g["x_flip"])
        self.y_flip_var.set(g["y_flip"])
        self.ref_electrode_var.set(g["ref_electrode"])
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

        self._update_group_column_combos()
        self._update_group_files_lb()
        self._refresh_reflines_lb()

        # Switch active_file to a file in this group so that per-file fields
        # (area, r_sol, e_ref) reflect the selected group.
        group_files = g.get("files", [])
        if group_files:
            target_file = (self.active_file
                           if self.active_file in group_files
                           else group_files[0])
            # Always refresh per-file UI from the new group's file_params,
            # even if target_file == active_file (cycles/corrections are group-scoped).
            # Do NOT call _save_active_state() here — active_group is already the new
            # group at this point, so saving would contaminate the new group's file_params
            # with the old group's UI state.  Callers are responsible for saving first.
            old_suppress = self._suppress_replot
            self._suppress_replot = True
            try:
                self._switch_active_file(target_file)
            finally:
                self._suppress_replot = old_suppress
            # Highlight target_file in the group files listbox (suppress _on_group_file_select)
            gf_list = g.get("files", [])
            if target_file in gf_list:
                lb_idx = gf_list.index(target_file)
                self._loading_files = True
                try:
                    self.group_files_lb.selection_clear(0, tk.END)
                    self.group_files_lb.selection_set(lb_idx)
                finally:
                    self._loading_files = False
                self.group_files_lb.see(lb_idx)
            # Sync the main file listbox highlight (suppress _on_file_select)
            file_keys = list(self.files.keys())
            if target_file in file_keys:
                fidx = file_keys.index(target_file)
                self._loading_files = True
                try:
                    self.file_listbox.selection_clear(0, tk.END)
                    self.file_listbox.selection_set(fidx)
                finally:
                    self._loading_files = False

        self._auto_replot()

    # ════════════════════════════════════════════════════════════════
    # Plotting
    # ════════════════════════════════════════════════════════════════
    def _auto_replot(self):
        if self._suppress_replot or not self.active_group:
            return
        self._plot_group(self.active_group)

    def _plot(self):
        self._auto_replot()

    def _plot_group(self, group_name):
        gentry = self.groups.get(group_name)
        if gentry is None or "ax" not in gentry:
            return

        is_active = (group_name == self.active_group)

        def _gv(key, default=""):
            var = getattr(self, key + "_var", None)
            return var.get() if (is_active and var is not None) else gentry.get(key, default)

        xcol    = self.x_var.get()     if is_active else gentry.get("x_col", "")
        ycol    = self.y_var.get()     if is_active else gentry.get("y_col", "")
        x_unit  = self.x_unit_var.get() if is_active else gentry.get("x_unit", "(auto)")
        y_unit  = self.y_unit_var.get() if is_active else gentry.get("y_unit", "(auto)")
        x_min_s = self.x_min_var.get() if is_active else gentry.get("x_min", "")
        x_max_s = self.x_max_var.get() if is_active else gentry.get("x_max", "")
        y_min_s = self.y_min_var.get() if is_active else gentry.get("y_min", "")
        y_max_s = self.y_max_var.get() if is_active else gentry.get("y_max", "")
        ref      = self.ref_electrode_var.get() if is_active else gentry.get("ref_electrode", "")
        leg_show = self.legend_show_var.get()  if is_active else gentry.get("legend_show", True)
        leg_frm  = self.legend_frame_var.get() if is_active else gentry.get("legend_frame", True)
        leg_loc  = self.legend_loc_var.get()   if is_active else gentry.get("legend_loc", "best")
        try:
            leg_size = float(self.legend_size_var.get() if is_active
                             else gentry.get("leg_size", 8.0))
        except (ValueError, TypeError):
            leg_size = 8.0

        if not xcol or not ycol:
            return

        ax     = gentry["ax"]
        canvas = gentry["canvas"]

        # Capture current view before clearing so user's zoom/pan survives the replot
        _prev_view = (ax.get_xlim(), ax.get_ylim()) if gentry.get("auto_xlim") is not None else None

        # Save legend position before clearing
        _old_leg = gentry.get("legend")
        if _old_leg is not None:
            _loc = getattr(_old_leg, '_loc', None)
            if isinstance(_loc, (tuple, list)):
                gentry["legend_manual_pos"] = tuple(_loc)
        gentry["legend"] = None
        self._clear_ann(group_name, redraw=False)
        ax.clear()

        has_data   = False
        multi_file = len([f for f in gentry["files"] if f in self.files]) > 1
        gentry["line_to_file"]  = {}   # reset line→file map for click detection
        gentry["line_to_cycle"] = {}   # reset line→cycle map for cycle-specific highlight

        # Determine highlight state
        _fhidden = gentry.get("file_hidden", {})
        _visible_fnames = [f for f in gentry["files"]
                           if self.files.get(f)
                           and not self.files[f].get("hidden", False)
                           and not _fhidden.get(f, False)]
        # Rank 1 (index 0, top of list) is drawn last → appears in front
        _visible_fnames = list(reversed(_visible_fnames))
        _af_in_group    = self.active_file in _visible_fnames
        _highlight      = self._plot_highlight and _af_in_group

        # Draw all files in INSERTION ORDER at full alpha (stable legend)
        for fname in _visible_fnames:
            fentry = self.files[fname]
            # Apply group-scoped corrections on top of raw data (independent per group)
            fp    = gentry.get("file_params", {}).get(fname, {})
            _rsol = fp.get("r_sol", 0.0)
            _eref = fp.get("e_ref", 0.0)
            if _rsol != 0.0 or _eref != 0.0:
                df = fentry["df_raw"].copy()
                if _rsol != 0.0 and "Ewe/V" in df.columns and "I/mA" in df.columns:
                    df["Ewe/V"] = df["Ewe/V"] - (df["I/mA"] / 1000.0) * _rsol
                if _eref != 0.0 and "Ewe/V" in df.columns:
                    df["Ewe/V"] = df["Ewe/V"] + _eref
            else:
                df = fentry["df_raw"]

            is_af = (fname == self.active_file)
            base_color = (_COLOR_HEX.get(self.file_color_var.get(), "#1f77b4")
                          if is_af else fentry.get("color", "#1f77b4"))
            try:    _lw = float(self.linewidth_var.get() if is_af else fentry.get("linewidth", "1.5"))
            except: _lw = 1.5
            _sname = (self.plot_style_var.get() if is_af else fentry.get("plot_style", "Line"))
            _ls, _mk, _ms = _PLOT_STYLES.get(_sname, ("-", "", 0))
            _grad  = self.cycle_gradient_var.get() if is_af else fentry.get("cycle_gradient", True)
            _rev   = self.cycle_reverse_var.get()  if is_af else fentry.get("cycle_reverse", False)
            try:    _step = float(self.lightness_step_var.get() if is_af else fentry.get("lightness_step", "0.15"))
            except: _step = 0.15
            # Use live UI only when rendering the active file in its active group;
            # all other combinations (same file, different group) use group-scoped state.
            _is_af_ui = is_af and (group_name == self.active_group)
            selected = (self._selected_cycles() if _is_af_ui
                        else fp.get("selected_cycles", []))
            area_s = self.area_var.get() if is_af else fentry.get("area", "")
            try:    _farea = float(area_s) if area_s else 0.0
            except: _farea = 0.0

            # Resolve "J" virtual column
            _x_is_J = (xcol == "J")
            _y_is_J = (ycol == "J")
            _real_xcol = xcol
            _real_ycol = ycol
            if _x_is_J or _y_is_J:
                for c in df.columns:
                    if "/" in c and c.rsplit("/", 1)[-1].strip() in _CURRENT_UNITS:
                        if _x_is_J: _real_xcol = c
                        if _y_is_J: _real_ycol = c
                        break

            if _real_xcol not in df.columns or _real_ycol not in df.columns:
                continue

            # Scale
            if _x_is_J:
                _xbase = _J_TO_BASE.get(x_unit)
                if _xbase:
                    x_scale, _ = self._get_unit_scale(_real_xcol, _xbase)
                else:
                    x_scale = 1.0
                if _farea > 0:
                    x_scale /= _farea
            else:
                x_scale, _ = self._get_unit_scale(_real_xcol, x_unit)

            if _y_is_J:
                _ybase = _J_TO_BASE.get(y_unit)
                if _ybase:
                    y_scale, _ = self._get_unit_scale(_real_ycol, _ybase)
                else:
                    y_scale = 1.0
                if _farea > 0:
                    y_scale /= _farea
            else:
                y_scale, _ = self._get_unit_scale(_real_ycol, y_unit)

            # Plot (all at alpha=1.0; glow/zorder applied in post-draw pass)
            if "cycle number" in df.columns and selected:
                cyc_cols = (_cycle_colors(base_color, len(selected), _step, _rev)
                            if _grad else [base_color] * len(selected))
                for i, c in enumerate(selected):
                    sub = df[df["cycle number"] == c]
                    if sub.empty:
                        continue
                    lbl = f"{fname} C{c}" if multi_file else f"C{c}"
                    ln, = ax.plot(sub[_real_xcol] * x_scale, sub[_real_ycol] * y_scale,
                                  color=cyc_cols[i], label=lbl, linewidth=_lw,
                                  linestyle=_ls, marker=_mk or None,
                                  markersize=_ms if _mk else 0)
                    gentry["line_to_file"][ln]  = fname
                    gentry["line_to_cycle"][ln] = c
                has_data = True
            elif "cycle number" not in df.columns:
                # No cycle column — plot all data as one line
                ln, = ax.plot(df[_real_xcol] * x_scale, df[_real_ycol] * y_scale,
                              color=base_color, label=fname, linewidth=_lw,
                              linestyle=_ls, marker=_mk or None,
                              markersize=_ms if _mk else 0)
                gentry["line_to_file"][ln] = fname
                has_data = True
            # else: cycle column exists but no cycles selected → plot nothing

        # Axis labels from group column + unit settings
        def _group_label(col, unit, is_J):
            if is_J:
                return f"J ({unit})" if unit and unit != "(auto)" else "J"
            _, lbl = self._get_unit_scale(col, unit)
            return lbl

        x_lbl = _group_label(xcol, x_unit, xcol == "J")
        y_lbl = _group_label(ycol, y_unit, ycol == "J")

        def _is_V(col, unit):
            src = col.rsplit("/", 1)[-1].strip() if "/" in col else ""
            return (unit in _VOLTAGE_UNITS if unit != "(auto)" else src in _VOLTAGE_UNITS)

        x_is_V = _is_V(xcol, x_unit) if xcol != "J" else False
        y_is_V = _is_V(ycol, y_unit) if ycol != "J" else False
        ax.set_xlabel(f"{x_lbl}  (vs {ref})" if (ref and x_is_V) else x_lbl)
        ax.set_ylabel(f"{y_lbl}  (vs {ref})" if (ref and y_is_V) else y_lbl)
        _title = (self.plot_title_var.get() if group_name == self.active_group
                  else gentry.get("custom_title", ""))
        ax.set_title(_title, fontsize=9)

        canvas.draw()
        gentry["auto_xlim"] = ax.get_xlim()
        gentry["auto_ylim"] = ax.get_ylim()

        # Restore the user's previous zoom/pan before applying manual range
        if _prev_view is not None:
            ax.set_xlim(_prev_view[0])
            ax.set_ylim(_prev_view[1])

        # Apply manual axis range (overrides restored view if values are set)
        self._apply_group_range(group_name, x_min_s, x_max_s, y_min_s, y_max_s, is_active)
        draw_reflines(ax, gentry.get("reflines", []))

        _xg  = self.x_grid_var.get()       if is_active else gentry.get("x_grid", False)
        _yg  = self.y_grid_var.get()       if is_active else gentry.get("y_grid", False)
        _xgi = self.x_grid_int_var.get()   if is_active else gentry.get("x_grid_int", "0")
        _ygi = self.y_grid_int_var.get()   if is_active else gentry.get("y_grid_int", "0")
        _gs  = self.grid_style_var.get()   if is_active else gentry.get("grid_style", "dashed")
        _gc  = self.grid_color_var.get()   if is_active else gentry.get("grid_color", "gray")
        _glw = self.grid_linewidth_var.get() if is_active else gentry.get("grid_linewidth", "0.8")
        apply_grid(ax, _xg, _yg, _xgi, _ygi, _gs, linewidth=_glw, color=_gc)

        if leg_show and has_data and ax.get_lines():
            # Build legend: rank-1 file first, cycles ascending; then apply custom order.
            _lh, _ll = ax.get_legend_handles_labels()
            _l2f = gentry.get("line_to_file", {})
            _l2c = gentry.get("line_to_cycle", {})
            def _h2k(h):
                f = _l2f.get(h); c = _l2c.get(h)
                return f"{f}:C{c}" if (f and c is not None) else (f or None)
            _h2k_map = {h: _h2k(h) for h in _lh}
            _fhidden = gentry.get("file_hidden", {})
            _rank_fnames = [f for f in gentry["files"]
                            if self.files.get(f)
                            and not self.files[f].get("hidden", False)
                            and not _fhidden.get(f, False)]
            _lh, _ll = _build_legend_order(
                _lh, _ll, _h2k_map, _rank_fnames)
            _lh, _ll = _reorder_legend_handles(
                _lh, _ll, gentry.get("legend_order", []), _h2k_map)
            gentry["legend"] = ax.legend(_lh, _ll,
                fontsize=leg_size, loc=leg_loc)
            gentry["legend"].set_draggable(True)
            gentry["legend"].get_frame().set_visible(leg_frm)
            gentry["leg_size"] = leg_size
            # Store key order (position → stable key) for this legend
            _key_order = [_h2k_map.get(h) for h in _lh]
            gentry["legend_key_order"] = _key_order
            # Restore custom labels — position-based (legend_key_order[i] → label_map)
            label_map = gentry.get("legend_labels", {})
            if label_map and isinstance(label_map, dict):
                for i, text_obj in enumerate(gentry["legend"].get_texts()):
                    if i < len(_key_order):
                        lbl = label_map.get(_key_order[i], "")
                        if lbl:
                            text_obj.set_text(lbl)
            if gentry.get("legend_manual_pos") is not None:
                gentry["legend"]._loc = gentry["legend_manual_pos"]
            canvas.draw()

        # Dim non-active lines AFTER legend is built so legend handles keep alpha=1.0
        self._apply_highlight_to_group(group_name)

        self._apply_font_to_group(group_name)

    def _apply_group_range(self, group_name, x_min_s, x_max_s, y_min_s, y_max_s, is_active):
        gentry = self.groups.get(group_name)
        if not gentry or "ax" not in gentry:
            return
        ax = gentry["ax"]
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
        if is_active:
            xl = ax.get_xlim()
            if self.x_flip_var.get() != (xl[0] > xl[1]):
                ax.set_xlim(xl[1], xl[0]); changed = True
            yl = ax.get_ylim()
            if self.y_flip_var.get() != (yl[0] > yl[1]):
                ax.set_ylim(yl[1], yl[0]); changed = True
        else:
            xl = ax.get_xlim()
            if gentry.get("x_flip", False) != (xl[0] > xl[1]):
                ax.set_xlim(xl[1], xl[0]); changed = True
            yl = ax.get_ylim()
            if gentry.get("y_flip", False) != (yl[0] > yl[1]):
                ax.set_ylim(yl[1], yl[0]); changed = True
        if changed:
            gentry["canvas"].draw_idle()

    # ════════════════════════════════════════════════════════════════
    # Unit conversion
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
    # Font / legend helpers
    # ════════════════════════════════════════════════════════════════
    def _apply_font_to_group(self, group_name):
        gentry = self.groups.get(group_name)
        if not gentry or "ax" not in gentry:
            return
        ax     = gentry["ax"]
        canvas = gentry["canvas"]
        is_active = (group_name == self.active_group)
        try:    ts = float(self.font_title_size_var.get() if is_active else gentry.get("font_title_size", "10"))
        except: ts = 10.0
        tb = ('bold' if (self.font_title_bold_var.get() if is_active else gentry.get("font_title_bold", False))
              else 'normal')
        try:    ls = float(self.font_label_size_var.get() if is_active else gentry.get("font_label_size", "10"))
        except: ls = 10.0
        lb = ('bold' if (self.font_label_bold_var.get() if is_active else gentry.get("font_label_bold", False))
              else 'normal')
        try:    ks = float(self.font_tick_size_var.get() if is_active else gentry.get("font_tick_size", "8"))
        except: ks = 8.0
        kb = (self.font_tick_bold_var.get() if is_active else gentry.get("font_tick_bold", False))
        try:    tpad = float(self.title_pad_var.get() if is_active else gentry.get("title_pad", "6"))
        except: tpad = 6.0
        try:    lpad = float(self.label_pad_var.get() if is_active else gentry.get("label_pad", "4"))
        except: lpad = 4.0

        _leg = ax.get_legend()
        if _leg is not None: _leg.set_visible(False)
        ax.set_title(ax.get_title(),   fontsize=ts, fontweight=tb, pad=tpad)
        ax.set_xlabel(ax.get_xlabel(), fontsize=ls, fontweight=lb, labelpad=lpad)
        ax.set_ylabel(ax.get_ylabel(), fontsize=ls, fontweight=lb, labelpad=lpad)
        ax.tick_params(axis='both', labelsize=ks)
        ax.figure.tight_layout()
        ax.figure.set_layout_engine('none')
        if _leg is not None: _leg.set_visible(True)
        canvas.draw()
        if kb:
            for lbl in ax.get_xticklabels() + ax.get_yticklabels():
                lbl.set_fontweight('bold')
            if _leg is not None: _leg.set_visible(False)
            ax.figure.tight_layout()
            ax.figure.set_layout_engine('none')
            if _leg is not None: _leg.set_visible(True)
            canvas.draw()

    def _toggle_legend_frame(self):
        if not self.active_group:
            return
        leg = self.groups[self.active_group].get("legend")
        if leg is not None:
            leg.get_frame().set_visible(self.legend_frame_var.get())
            self.groups[self.active_group]["canvas"].draw()

    def _edit_group_title(self, group_name):
        gentry = self.groups.get(group_name, {})
        ax     = gentry.get("ax")
        canvas = gentry.get("canvas")
        if ax is None:
            return
        current  = gentry.get("custom_title", ax.title.get_text() or group_name)
        new_title = simpledialog.askstring("Edit Title", "Plot title:",
                                           initialvalue=current, parent=self)
        if new_title is not None:
            gentry["custom_title"] = new_title
            ax.set_title(new_title, fontsize=9)
            canvas.draw_idle()

    def _edit_legend_labels(self):
        if not self.active_group:
            return
        gentry = self.groups.get(self.active_group)
        if gentry is None:
            return
        leg = gentry.get("legend")
        if leg is None:
            from tkinter import messagebox
            messagebox.showinfo("Info", "Plot data first to create a legend.")
            return
        leg.set_draggable(False)
        gentry["legend"], perm = open_legend_editor(
            self, leg, gentry["canvas"], gentry.get("leg_size", 8.0))
        if gentry.get("legend") is not None:
            gentry["legend"].set_draggable(True)
            orig_key_order = gentry.get("legend_key_order", [])
            # perm[new_pos] = orig_pos — produced by legend_editor
            legend_order = [orig_key_order[j] for j in perm
                            if j < len(orig_key_order) and orig_key_order[j] is not None]
            gentry["legend_order"] = legend_order
            # Persist custom label text: new position i → key = legend_order[i]
            label_map = gentry.get("legend_labels") if isinstance(gentry.get("legend_labels"), dict) else {}
            for i, text_obj in enumerate(gentry["legend"].get_texts()):
                if i < len(legend_order):
                    label_map[legend_order[i]] = text_obj.get_text()
            gentry["legend_labels"] = label_map

    # ════════════════════════════════════════════════════════════════
    # Per-file change handlers
    # ════════════════════════════════════════════════════════════════
    def _on_file_color_change(self, event=None):
        if self.active_file and self.active_file in self.files:
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

    def _on_area_change(self):
        if self.active_file and self.active_file in self.files:
            self.files[self.active_file]["area"] = self.area_var.get()
        self._update_group_column_combos()
        self._auto_replot()

    def _on_gradient_change(self):
        if self.active_file and self.active_file in self.files:
            self.files[self.active_file]["cycle_gradient"] = self.cycle_gradient_var.get()
            self.files[self.active_file]["cycle_reverse"]  = self.cycle_reverse_var.get()
            self.files[self.active_file]["lightness_step"] = self.lightness_step_var.get()
        self._auto_replot()

    # ════════════════════════════════════════════════════════════════
    # ── Highlight active file within a group (direct alpha update) ───
    def _apply_highlight_to_group(self, group_name):
        """Apply or remove the highlight effect on all lines in a group's axes.

        Removes stale glow lines, redraws glow for the active file, and
        updates alpha / z-order.  Called after every file selection and
        after right-click to reset.
        """
        gentry = self.groups.get(group_name)
        if not gentry or "ax" not in gentry:
            return
        ax           = gentry["ax"]
        canvas       = gentry["canvas"]
        active       = self.active_file
        _fh      = gentry.get("file_hidden", {})
        visible  = [f for f in gentry.get("files", [])
                    if self.files.get(f)
                    and not self.files[f].get("hidden", False)
                    and not _fh.get(f, False)]
        on = self._plot_highlight and bool(active) and active in visible

        # Remove stale glow lines
        for ln in list(ax.get_lines()):
            if ln.get_label() == '_glow':
                ln.remove()

        line_to_file  = gentry.get("line_to_file",  {})
        line_to_cycle = gentry.get("line_to_cycle", {})
        for ln in ax.get_lines():
            lbl = ln.get_label() or ""
            if lbl.startswith("_"):
                continue
            fname = line_to_file.get(ln)
            is_af_file = (fname == active) if fname else (lbl == active or lbl.startswith(active + " "))
            if on and self._active_cycle is not None:
                # Cycle-specific: only highlight the exact file + cycle combination
                cycle_n = line_to_cycle.get(ln)
                is_af = (is_af_file and cycle_n == self._active_cycle)
            else:
                is_af = is_af_file
            ln.set_alpha(1.0 if (not on or is_af) else 0.55)
            ln.set_zorder(3.0 if (on and is_af) else 2.0)
            if on and is_af:
                ax.plot(ln.get_xdata(), ln.get_ydata(),
                        color=ln.get_color(),
                        linewidth=ln.get_linewidth() * 2.5,
                        linestyle=ln.get_linestyle(),
                        alpha=0.18, label='_glow', zorder=1.9)

        canvas.draw_idle()
        self._highlight_active_headers()

    def _highlight_active_headers(self):
        """Gold header on all group plots that contain the active file; green otherwise."""
        for gname, gentry in self.groups.items():
            is_active = (self._plot_highlight and bool(self.active_file)
                         and self.active_file in gentry.get("files", []))
            color = _GROUP_HDR_ACTIVE if is_active else _GROUP_HDR_BG
            hdr = gentry.get("hdr_frame")
            lbl_w = gentry.get("hdr_label")
            if hdr and hdr.winfo_exists():
                hdr.configure(bg=color)
            if lbl_w and lbl_w.winfo_exists():
                lbl_w.configure(bg=color)

    # Matplotlib interactions (per group figure)
    # ════════════════════════════════════════════════════════════════
    def _on_scroll(self, event, group_name):
        gentry = self.groups.get(group_name)
        if gentry is None or event.inaxes is not gentry.get("ax"):
            return
        ax     = gentry["ax"]
        canvas = gentry["canvas"]
        scale  = 0.8 if event.step > 0 else 1.25
        xl, yl = ax.get_xlim(), ax.get_ylim()
        xd, yd = event.xdata, event.ydata
        xf = (xd - xl[0]) / (xl[1] - xl[0]) if xl[1] != xl[0] else 0.5
        yf = (yd - yl[0]) / (yl[1] - yl[0]) if yl[1] != yl[0] else 0.5
        nxr = (xl[1] - xl[0]) * scale
        nyr = (yl[1] - yl[0]) * scale
        ax.set_xlim(xd - nxr * xf, xd + nxr * (1 - xf))
        ax.set_ylim(yd - nyr * yf, yd + nyr * (1 - yf))
        canvas.draw_idle()

    def _on_press(self, event, group_name):
        gentry = self.groups.get(group_name)
        if gentry is None:
            return
        self._activate_group(group_name)
        gentry["pan_moved"] = False
        ax = gentry.get("ax")

        # Single left-click on a data line → select that file
        if (event.button == 1 and ax is not None and event.inaxes is ax
                and not getattr(event, 'dblclick', False)):
            for line, fname in gentry.get("line_to_file", {}).items():
                try:
                    hit, _ = line.contains(event)
                except Exception:
                    hit = False
                if hit and fname in self.files:
                    # Capture specific cycle (None for non-cycle files)
                    clicked_cycle = gentry.get("line_to_cycle", {}).get(line)
                    if fname != self.active_file:
                        self._save_active_state()
                        old_supp = self._suppress_replot
                        self._suppress_replot = True
                        try:
                            self._switch_active_file(fname)
                        finally:
                            self._suppress_replot = old_supp
                        # Update "files in group" listbox highlight
                        gf_list = self.groups.get(self.active_group, {}).get("files", [])
                        if fname in gf_list:
                            idx = gf_list.index(fname)
                            self.group_files_lb.selection_clear(0, tk.END)
                            self.group_files_lb.selection_set(idx)
                            self.group_files_lb.see(idx)
                        # Sync main file listbox (suppress _on_file_select)
                        file_keys = list(self.files.keys())
                        if fname in file_keys:
                            self._loading_files = True
                            try:
                                self.file_listbox.selection_clear(0, tk.END)
                                self.file_listbox.selection_set(file_keys.index(fname))
                            finally:
                                self._loading_files = False
                    self._active_cycle   = clicked_cycle
                    self._plot_highlight = True
                    self._auto_replot()
                    break

        if event.button == 1 and getattr(event, 'dblclick', False) and ax is not None:
            canvas = gentry["canvas"]
            if event.inaxes is ax:
                leg = gentry.get("legend")
                if leg is not None:
                    try:
                        r = canvas.get_renderer()
                        if leg.get_window_extent(r).contains(event.x, event.y):
                            self._edit_legend_labels()
                            return
                    except Exception:
                        pass
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
                    self._edit_group_title(group_name)
                    return
            except Exception:
                pass
            if event.inaxes is ax:
                if self._zoom_group is None:
                    self._zoom_group_view(group_name)
                else:
                    self._unzoom_group_view()
            return

        if event.button == 1 and event.inaxes is ax:
            leg = gentry.get("legend")
            if leg is not None:
                try:
                    r = event.canvas.get_renderer()
                    if leg.get_window_extent(r).contains(event.x, event.y):
                        return
                except Exception:
                    pass
            gentry["panning"]   = True
            gentry["pan_ax"]    = ax
            gentry["pan_start"] = (event.xdata, event.ydata)
        elif event.button == 3:
            leg = gentry.get("legend")
            if leg is not None:
                try:
                    r = event.canvas.get_renderer()
                    if leg.get_window_extent(r).contains(event.x, event.y):
                        gentry["leg_resize"]          = True
                        gentry["leg_resize_start_y"]  = event.y
                        gentry["leg_resize_start_sz"] = gentry.get("leg_size", 8.0)
                except Exception:
                    pass

    def _on_release(self, event, group_name):
        gentry = self.groups.get(group_name)
        if gentry is None:
            return
        was_resizing = gentry.get("leg_resize", False)
        gentry["leg_resize"] = False
        gentry["panning"]    = False
        gentry["pan_ax"]     = None
        if was_resizing:
            if group_name == self.active_group:
                new_sz = gentry.get("leg_size", 8.0)
                self.legend_size_var.set(
                    str(int(new_sz)) if new_sz == int(new_sz) else str(new_sz))
            return
        ax = gentry.get("ax")
        if (event.button == 1
                and not gentry.get("pan_moved", False)
                and event.inaxes is ax):
            leg = gentry.get("legend")
            if leg is not None:
                try:
                    r = event.canvas.get_renderer()
                    if leg.get_window_extent(r).contains(event.x, event.y):
                        return
                except Exception:
                    pass
            self._annotate(event, group_name)
        elif event.button == 3 and event.inaxes is ax:
            self._plot_highlight = False
            self._active_cycle = None
            self._clear_ann(group_name)
            self._apply_highlight_to_group(group_name)

    def _on_motion(self, event, group_name):
        gentry = self.groups.get(group_name)
        if gentry is None:
            return
        if gentry.get("panning") and gentry.get("pan_ax") is not None:
            if event.inaxes is not gentry["pan_ax"] or event.xdata is None:
                return
            gentry["pan_moved"] = True
            ax = gentry["ax"]
            dx = gentry["pan_start"][0] - event.xdata
            dy = gentry["pan_start"][1] - event.ydata
            ax.set_xlim(ax.get_xlim()[0] + dx, ax.get_xlim()[1] + dx)
            ax.set_ylim(ax.get_ylim()[0] + dy, ax.get_ylim()[1] + dy)
            gentry["canvas"].draw_idle()
            return
        if gentry.get("leg_resize") and gentry.get("legend") is not None:
            dy     = event.y - gentry["leg_resize_start_y"]
            new_sz = max(4.0, min(30.0, gentry["leg_resize_start_sz"] + dy / 5.0))
            prev_sz = gentry.get("leg_size", 8.0)
            gentry["leg_size"] = new_sz
            leg = gentry["legend"]
            if prev_sz > 0:
                _scale_legend_spacing(leg, new_sz / prev_sz)
            for t in leg.get_texts():
                t.set_fontsize(new_sz)
            gentry["canvas"].draw()

    def _annotate(self, event, group_name):
        gentry = self.groups.get(group_name)
        if gentry is None:
            return
        ax    = gentry["ax"]
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
        last = gentry.get("ann_last")
        if (last is not None
                and abs(event.x - last[0]) <= _CLICK_CYCLE_PX
                and abs(event.y - last[1]) <= _CLICK_CYCLE_PX):
            gentry["ann_idx"] = (gentry["ann_idx"] + 1) % len(candidates)
        else:
            gentry["ann_idx"] = 0
        gentry["ann_last"] = (event.x, event.y)
        n   = len(candidates)
        idx = gentry["ann_idx"]
        _, ln, x, y = candidates[idx]
        label = ln.get_label() or "?"
        xl, yl = ax.get_xlim(), ax.get_ylim()
        xf = (x - xl[0]) / (xl[1] - xl[0]) if xl[1] != xl[0] else 0.5
        yf = (y - yl[0]) / (yl[1] - yl[0]) if yl[1] != yl[0] else 0.5
        xoff = -95 if xf > 0.65 else 15
        yoff = -60 if yf > 0.65 else 15
        hint = f"  [{idx + 1}/{n}]" if n > 1 else ""
        text = f"x = {x:.4g}\ny = {y:.4g}  ({label}){hint}"
        if n > 1 and idx == 0:
            text += "\n↻ click again to cycle"
        self._clear_ann(group_name, redraw=False)
        gentry["ann"] = ax.annotate(
            text, xy=(x, y), xytext=(xoff, yoff), textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.4", fc="lightyellow", ec="gray", alpha=0.92),
            arrowprops=dict(arrowstyle="->", color="gray", lw=1.2),
            fontsize=8, zorder=10,
        )
        gentry["ann"].set_in_layout(False)  # exclude from tight_layout bounding box
        gentry["ann_dot"], = ax.plot(x, y, "o", color=ln.get_color(),
                                     markersize=7, zorder=11, label="_ann_dot")
        gentry["canvas"].draw_idle()

    def _clear_ann(self, group_name, redraw=True):
        gentry = self.groups.get(group_name)
        if gentry is None:
            return
        for key in ("ann", "ann_dot"):
            artist = gentry.get(key)
            if artist is not None:
                try:    artist.remove()
                except: pass
                gentry[key] = None
        gentry["ann_last"] = None
        gentry["ann_idx"]  = 0
        if redraw and "canvas" in gentry:
            gentry["canvas"].draw_idle()

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
        self._auto_replot()

    def _selected_cycles(self):
        return [c for c, v in self._cycle_vars.items() if v.get()]

    # ════════════════════════════════════════════════════════════════
    # Reference lines (per group)
    # ════════════════════════════════════════════════════════════════
    def _add_xrefline(self):
        if not self.active_group:
            return
        try:    v = float(self._ref_x_var.get())
        except: return
        style = self._refline_style_var.get()
        color = self._refline_color_var.get()
        lw    = self._refline_linewidth_var.get()
        self.groups[self.active_group].setdefault("reflines", []).append(
            ('x', v, style, color, lw))
        self._reflines_lb.insert(tk.END, f"X = {v:.4g}")
        self._auto_replot()

    def _add_yrefline(self):
        if not self.active_group:
            return
        try:    v = float(self._ref_y_var.get())
        except: return
        style = self._refline_style_var.get()
        color = self._refline_color_var.get()
        lw    = self._refline_linewidth_var.get()
        self.groups[self.active_group].setdefault("reflines", []).append(
            ('y', v, style, color, lw))
        self._reflines_lb.insert(tk.END, f"Y = {v:.4g}")
        self._auto_replot()

    def _remove_refline(self):
        sel = self._reflines_lb.curselection()
        if not sel or not self.active_group:
            return
        idx = sel[0]
        self.groups[self.active_group]["reflines"].pop(idx)
        self._reflines_lb.delete(idx)
        self._auto_replot()

    def _refresh_reflines_lb(self):
        self._reflines_lb.delete(0, tk.END)
        if not self.active_group:
            return
        for e in self.groups.get(self.active_group, {}).get("reflines", []):
            axis, val = e[0], e[1]
            self._reflines_lb.insert(tk.END, f"{'X' if axis == 'x' else 'Y'} = {val:.4g}")

    def _on_refline_select(self):
        sel = self._reflines_lb.curselection()
        if not sel or not self.active_group:
            return
        reflines = self.groups.get(self.active_group, {}).get("reflines", [])
        if sel[0] >= len(reflines):
            return
        e = reflines[sel[0]]
        self._refline_style_var.set(e[2])
        self._refline_color_var.set(e[3])
        self._refline_linewidth_var.set(e[4] if len(e) > 4 else "1.0")

    def _on_refline_style_color_change(self):
        sel = self._reflines_lb.curselection()
        if not sel or not self.active_group:
            return
        reflines = self.groups.get(self.active_group, {}).get("reflines", [])
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
    # Log
    # ════════════════════════════════════════════════════════════════
    def _log(self, message: str):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
