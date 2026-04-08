"""EIS Panel — Nyquist plot for electrochemical impedance spectroscopy.

Multiple files can be loaded and overlaid on one shared plot.
All settings (columns, units, range, display, legend, grid, reflines) are
stored per file; the active file's settings drive the shared overlay.
"""

import math
from collections import OrderedDict

import tkinter as tk
from tkinter import ttk

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from .file_manager import FileManagerMixin, _COLOR_NAMES, _COLOR_HEX, _PLOT_STYLES, _PLOT_STYLE_NAMES
from .plotting import apply_grid, draw_reflines, copy_figure_to_clipboard
from .legend_editor import open_legend_editor
from .checklist import CheckableListbox


# ── Unit option lists by physical dimension ──────────────────────────────────
_OHM_UNITS   = ["(auto)", "Ω",  "kΩ",  "MΩ",  "mΩ"]
_PHASE_UNITS = ["(auto)", "°",  "rad"]
_FREQ_UNITS  = ["(auto)", "Hz", "kHz", "MHz"]

# Column-type → unit option list
_UNITS_BY_TYPE = {
    "impedance": _OHM_UNITS,
    "phase":     _PHASE_UNITS,
    "frequency": _FREQ_UNITS,
}

# Impedance: source-unit → SI-ohm factor (for columns like "Re(Z)/Ohm")
_IMP_RAW_MAP = {
    "Ohm": 1.0,   "ohm": 1.0,   "Ω": 1.0,
    "kOhm": 1e3,  "kΩ":  1e3,
    "MOhm": 1e6,  "MΩ":  1e6,
    "mOhm": 1e-3, "mΩ":  1e-3,
}
# Impedance: display-unit → SI-ohm factor
_IMP_DISP_MAP = {"Ω": 1.0, "kΩ": 1e3, "MΩ": 1e6, "mΩ": 1e-3}

# Frequency: unit → SI-Hz factor
_FREQ_MAP = {"Hz": 1.0, "kHz": 1e3, "MHz": 1e6, "GHz": 1e9}

# Phase: source-unit → radians factor (for conversion to display unit)
_PHASE_SRC_RAD = {"rad": 1.0, "Rad": 1.0}
_PHASE_SRC_DEG = {"°": 1.0, "deg": 1.0, "Deg": 1.0, "degree": 1.0}


def _col_type(col):
    """Return 'phase', 'frequency', or 'impedance' based on the column name."""
    cl = col.lower()
    if any(k in cl for k in ("phase", "phi", "angle")):
        return "phase"
    if any(k in cl for k in ("freq", "/hz", "/khz", "/mhz", "hertz")):
        return "frequency"
    return "impedance"


class EISPanel(FileManagerMixin, ttk.Frame):
    """EIS tab: shared Nyquist overlay with per-file column/unit/range/display settings."""

    def __init__(self, master):
        ttk.Frame.__init__(self, master)

        # Core state
        self.files            = OrderedDict()
        self.active_file      = None
        self._suppress_replot = False
        self._loading_files   = False

        # Plot / legend state
        self._legend_obj           = None
        self._current_legend_size  = 8.0
        self._legend_stable_map  = {}    # {stable_key: custom_label}
        self._legend_stable_keys = []    # stable keys from the last _plot()
        self._legend_auto_labels = []    # auto-labels from the last _plot() (for edit diffing)
        self._auto_xlim            = None
        self._auto_ylim            = None

        # Click annotation state
        self._ann                  = None
        self._ann_dot              = None
        self._last_click_pos       = None
        self._click_candidate_idx  = 0

        # Debounce timer for range entries
        self._range_replot_id = None

        # Required by FileManagerMixin (not shown in UI for EIS)
        self.r_sol_var = tk.StringVar(value="0.0")
        self.e_ref_var = tk.StringVar(value="0.0")

        self._build_panel()
        self.after(500, self._auto_set_initial_size)

    # ════════════════════════════════════════════════════════════════
    # Panel construction
    # ════════════════════════════════════════════════════════════════
    def _build_panel(self):
        body = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True)

        # ── Scrollable left panel ─────────────────────────────────
        left_outer = ttk.Frame(body, width=260)
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
        ttk.Label(left, text="Files:", font=("", 9, "bold")).pack(
            anchor=tk.W, padx=4, pady=(6, 0))
        fb = ttk.Frame(left)
        fb.pack(fill=tk.X, padx=4)
        ttk.Button(fb, text="Load File(s)", command=self._load_files).pack(
            side=tk.LEFT, padx=(0, 4))
        ttk.Button(fb, text="Remove", command=self._remove_file).pack(side=tk.LEFT)

        flf = ttk.Frame(left)
        flf.pack(fill=tk.X, padx=4, pady=2)
        self.file_listbox = CheckableListbox(flf, height=5,
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

        # ── Columns ───────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Columns", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        x_row = ttk.Frame(left)
        x_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(x_row, text="X:").pack(side=tk.LEFT)
        self.x_var   = tk.StringVar()
        self.x_combo = ttk.Combobox(x_row, textvariable=self.x_var,
                                    state="readonly", width=14)
        self.x_combo.pack(side=tk.LEFT, padx=(2, 4))
        self.x_unit_var = tk.StringVar(value="(auto)")
        self.x_unit_cb = ttk.Combobox(x_row, textvariable=self.x_unit_var,
                                      values=_OHM_UNITS, state="readonly", width=6)
        self.x_unit_cb.pack(side=tk.LEFT)
        self.x_combo.bind("<<ComboboxSelected>>", self._on_x_col_change)
        self.x_unit_cb.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())

        y_row = ttk.Frame(left)
        y_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(y_row, text="Y:").pack(side=tk.LEFT)
        self.y_var   = tk.StringVar()
        self.y_combo = ttk.Combobox(y_row, textvariable=self.y_var,
                                    state="readonly", width=14)
        self.y_combo.pack(side=tk.LEFT, padx=(2, 4))
        self.y_unit_var = tk.StringVar(value="(auto)")
        self.y_unit_cb = ttk.Combobox(y_row, textvariable=self.y_unit_var,
                                      values=_OHM_UNITS, state="readonly", width=6)
        self.y_unit_cb.pack(side=tk.LEFT)
        self.y_combo.bind("<<ComboboxSelected>>", self._on_y_col_change)
        self.y_unit_cb.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())

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
            self._plot()

        ttk.Button(left, text="⇄  Swap X↔Y", command=_swap_xy).pack(
            anchor=tk.W, padx=4, pady=(0, 4))

        # ── Display ───────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Display", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        self.show_markers_var  = tk.BooleanVar(value=True)
        self.connect_lines_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(left, text="Show markers",
                        variable=self.show_markers_var,
                        command=self._auto_replot).pack(anchor=tk.W, padx=4)
        ttk.Checkbutton(left, text="Connect with lines",
                        variable=self.connect_lines_var,
                        command=self._auto_replot).pack(anchor=tk.W, padx=4)

        lw_row = ttk.Frame(left)
        lw_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Label(lw_row, text="Line Width:").pack(side=tk.LEFT)
        self.linewidth_var = tk.StringVar(value="1.5")
        _lw_e = ttk.Entry(lw_row, textvariable=self.linewidth_var, width=4)
        _lw_e.pack(side=tk.LEFT, padx=(2, 0))
        _lw_e.bind("<Return>",   lambda e: self._on_linewidth_change())
        _lw_e.bind("<FocusOut>", lambda e: self._on_linewidth_change())
        ttk.Label(lw_row, text="Shape:").pack(side=tk.LEFT, padx=(8, 0))
        self.plot_style_var = tk.StringVar(value="Line+Circle")
        _style_cb = ttk.Combobox(lw_row, textvariable=self.plot_style_var,
                                  values=_PLOT_STYLE_NAMES, state="readonly", width=11)
        _style_cb.pack(side=tk.LEFT, padx=(2, 0))
        _style_cb.bind("<<ComboboxSelected>>", lambda e: self._on_plot_style_change())

        # ── Plot Range ────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Plot Range", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        xr_f = ttk.Frame(left)
        xr_f.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(xr_f, text="X:").pack(side=tk.LEFT)
        self.x_min_var = tk.StringVar()
        _xmin = ttk.Entry(xr_f, textvariable=self.x_min_var, width=6)
        _xmin.pack(side=tk.LEFT, padx=(2, 2))
        ttk.Label(xr_f, text="–").pack(side=tk.LEFT)
        self.x_max_var = tk.StringVar()
        _xmax = ttk.Entry(xr_f, textvariable=self.x_max_var, width=6)
        _xmax.pack(side=tk.LEFT, padx=(2, 4))
        ttk.Label(xr_f, text="Int:").pack(side=tk.LEFT)
        self.x_grid_int_var = tk.StringVar(value="0")
        _xgi = ttk.Entry(xr_f, textvariable=self.x_grid_int_var, width=5)
        _xgi.pack(side=tk.LEFT, padx=(2, 0))
        _xgi.bind("<Return>",   lambda e: self._auto_replot())
        _xgi.bind("<FocusOut>", lambda e: self._auto_replot())

        yr_f = ttk.Frame(left)
        yr_f.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(yr_f, text="Y:").pack(side=tk.LEFT)
        self.y_min_var = tk.StringVar()
        _ymin = ttk.Entry(yr_f, textvariable=self.y_min_var, width=6)
        _ymin.pack(side=tk.LEFT, padx=(2, 2))
        ttk.Label(yr_f, text="–").pack(side=tk.LEFT)
        self.y_max_var = tk.StringVar()
        _ymax = ttk.Entry(yr_f, textvariable=self.y_max_var, width=6)
        _ymax.pack(side=tk.LEFT, padx=(2, 4))
        ttk.Label(yr_f, text="Int:").pack(side=tk.LEFT)
        self.y_grid_int_var = tk.StringVar(value="0")
        _ygi = ttk.Entry(yr_f, textvariable=self.y_grid_int_var, width=5)
        _ygi.pack(side=tk.LEFT, padx=(2, 0))
        _ygi.bind("<Return>",   lambda e: self._auto_replot())
        _ygi.bind("<FocusOut>", lambda e: self._auto_replot())

        ttk.Label(left, text="(blank = auto)", foreground="gray",
                  font=("", 8)).pack(anchor=tk.W, padx=4)

        for _e in (_xmin, _xmax, _ymin, _ymax):
            _e.bind("<Return>",   lambda e: self._schedule_range_replot())
            _e.bind("<FocusOut>", lambda e: self._schedule_range_replot())

        flip_row = ttk.Frame(left)
        flip_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        self.x_flip_var = tk.BooleanVar(value=False)
        self.y_flip_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(flip_row, text="Flip X", variable=self.x_flip_var,
                        command=self._plot).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Checkbutton(flip_row, text="Flip Y", variable=self.y_flip_var,
                        command=self._plot).pack(side=tk.LEFT)

        # ── Legend ────────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Legend", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        leg_row1 = ttk.Frame(left)
        leg_row1.pack(fill=tk.X, padx=4, pady=2)
        self.legend_show_var  = tk.BooleanVar(value=True)
        self.legend_frame_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(leg_row1, text="Show",
                        variable=self.legend_show_var,
                        command=self._toggle_legend).pack(side=tk.LEFT)
        ttk.Checkbutton(leg_row1, text="Frame",
                        variable=self.legend_frame_var,
                        command=self._toggle_legend_frame).pack(side=tk.LEFT, padx=(8, 0))

        leg_row2 = ttk.Frame(left)
        leg_row2.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(leg_row2, text="Size:").pack(side=tk.LEFT)
        self.legend_size_var = tk.StringVar(value="8")
        _leg_sz = ttk.Entry(leg_row2, textvariable=self.legend_size_var, width=4)
        _leg_sz.pack(side=tk.LEFT, padx=(2, 8))
        _leg_sz.bind("<Return>",   lambda e: self._on_legend_size_change())
        _leg_sz.bind("<FocusOut>", lambda e: self._on_legend_size_change())
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
        def _on_leg_loc_select_eis(e=None):
            self._legend_manual_pos = None  # user explicitly chose a location
            # Also neutralise the live legend's _loc so _plot() doesn't
            # re-capture the old tuple before ax.clear() runs.
            if self._legend_obj is not None:
                self._legend_obj._loc = 0
            self._auto_replot()
        _leg_loc_cb.bind("<<ComboboxSelected>>", _on_leg_loc_select_eis)

        ttk.Button(left, text="Edit Labels",
                   command=self._edit_legend_labels).pack(anchor=tk.W, padx=4, pady=2)
        ttk.Label(left,
                  text="(left-drag to move, right-drag to resize; dbl-click to edit)",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)

        # ── Grid ──────────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Grid", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        grid_row = ttk.Frame(left)
        grid_row.pack(fill=tk.X, padx=4, pady=2)
        self.x_grid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(grid_row, text="X", variable=self.x_grid_var,
                        command=self._auto_replot).pack(side=tk.LEFT)
        self.y_grid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(grid_row, text="Y", variable=self.y_grid_var,
                        command=self._auto_replot).pack(side=tk.LEFT, padx=(8, 0))

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
        _rl_style_cb.pack(side=tk.LEFT, padx=(2, 8))
        ttk.Label(ref_opt_row, text="Color:").pack(side=tk.LEFT)
        self._refline_color_var = tk.StringVar(value="dimgray")
        _rl_color_cb = ttk.Combobox(ref_opt_row, textvariable=self._refline_color_var,
                                    values=["dimgray", "black", "red", "blue", "green",
                                            "orange", "purple", "crimson", "royalblue",
                                            "darkorange", "teal", "saddlebrown"],
                                    state="readonly", width=10)
        _rl_color_cb.pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(ref_opt_row, text="Width:").pack(side=tk.LEFT)
        self._refline_linewidth_var = tk.StringVar(value="1.0")
        _rl_lw = ttk.Entry(ref_opt_row, textvariable=self._refline_linewidth_var, width=4)
        _rl_lw.pack(side=tk.LEFT, padx=(2, 0))
        _rl_style_cb.bind("<<ComboboxSelected>>",
                          lambda e: self._on_refline_style_color_change())
        _rl_color_cb.bind("<<ComboboxSelected>>",
                          lambda e: self._on_refline_style_color_change())
        _rl_lw.bind("<Return>",   lambda e: self._on_refline_style_color_change())
        _rl_lw.bind("<FocusOut>", lambda e: self._on_refline_style_color_change())

        # ── Plot height ───────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=4)
        _ph_row = ttk.Frame(left)
        _ph_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(_ph_row, text="Plot height (px):").pack(side=tk.LEFT)
        self._plot_height_var = tk.StringVar(value="")
        _ph_entry = ttk.Entry(_ph_row, textvariable=self._plot_height_var, width=6)
        _ph_entry.pack(side=tk.LEFT, padx=(4, 0))
        _ph_entry.bind("<Return>",   lambda e: self._apply_plot_height())
        _ph_entry.bind("<FocusOut>", lambda e: self._apply_plot_height())
        ttk.Label(_ph_row, text="(blank = auto)").pack(side=tk.LEFT, padx=(4, 0))

        # ── Plot button ───────────────────────────────────────────
        ttk.Button(left, text="Plot",
                   command=self._plot).pack(padx=4, pady=6, anchor=tk.W)

        # ── Right panel: single shared Nyquist figure ─────────────
        right = ttk.Frame(body)
        body.add(right, weight=1)
        right.rowconfigure(0, weight=0)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        # ── Plot size controls (always visible) ──────────────────
        self.plot_w_var = tk.StringVar(value="21.0")
        self.plot_h_var = tk.StringVar(value="12.5")
        _size_bar = ttk.Frame(right)
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

        # ── Scrollable plot area ──────────────────────────────────
        _right_inner = ttk.Frame(right)
        _right_inner.grid(row=1, column=0, sticky="nsew")
        _right_inner.rowconfigure(0, weight=1)
        _right_inner.columnconfigure(0, weight=1)
        _plot_sc = tk.Canvas(_right_inner, highlightthickness=0)
        _right_vs = ttk.Scrollbar(_right_inner, orient=tk.VERTICAL,   command=_plot_sc.yview)
        _right_hs = ttk.Scrollbar(_right_inner, orient=tk.HORIZONTAL, command=_plot_sc.xview)
        _plot_sc.configure(yscrollcommand=_right_vs.set, xscrollcommand=_right_hs.set)
        _right_vs.grid(row=0, column=1, sticky="ns")
        _right_hs.grid(row=1, column=0, sticky="ew")
        _plot_sc.grid(row=0, column=0, sticky="nsew")
        _plot_sc.bind("<MouseWheel>",
                      lambda e: _plot_sc.yview_scroll(-1*(e.delta//120), "units"))
        _plot_sc.bind("<Shift-MouseWheel>",
                      lambda e: _plot_sc.xview_scroll(-1*(e.delta//120), "units"))
        _plots_frame = ttk.Frame(_plot_sc)
        _plots_win = _plot_sc.create_window((0, 0), window=_plots_frame, anchor=tk.NW)
        _plots_frame.bind("<Configure>",
                          lambda e: _plot_sc.configure(scrollregion=_plot_sc.bbox("all")))
        self._plot_sc = _plot_sc

        _fw = float(self.plot_w_var.get())
        _fh = float(self.plot_h_var.get())
        self.fig = Figure(figsize=(_fw, _fh), dpi=100)
        self.ax  = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=_plots_frame)
        self.canvas.get_tk_widget().pack()
        self.canvas.get_tk_widget().config(width=int(_fw * 100), height=int(_fh * 100))

        tb_frame = ttk.Frame(_plots_frame)
        tb_frame.pack(fill=tk.X)

        panel_ref = self

        class _EISToolbar(NavigationToolbar2Tk):
            def home(tb_self, *args):
                panel_ref._reset_view()

        tb = _EISToolbar(self.canvas, tb_frame, pack_toolbar=False)
        tb.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tb.update()
        tk.Button(
            tb_frame, text="Copy",
            command=lambda: copy_figure_to_clipboard(self.fig),
            relief=tk.RAISED, borderwidth=1, padx=6,
        ).pack(side=tk.LEFT, padx=(4, 2), pady=1)

        # Mouse interactions
        self.canvas.mpl_connect("scroll_event",          self._on_scroll)
        self.canvas.mpl_connect("button_press_event",    self._on_press)
        self.canvas.mpl_connect("button_release_event",  self._on_release)
        self.canvas.mpl_connect("motion_notify_event",   self._on_motion)

        # Pan / legend-resize state
        self._panning          = False
        self._pan_start        = None
        self._pan_moved        = False
        self._legend_resizing  = False
        self._resize_start_y   = None
        self._resize_start_sz  = None

    # ════════════════════════════════════════════════════════════════
    # FileManagerMixin override — auto-plot immediately on load
    # ════════════════════════════════════════════════════════════════
    def _load_files(self):
        """Load files and plot immediately (EIS data has no cycles to select first)."""
        super()._load_files()
        if self.active_file:
            self._plot()

    # ════════════════════════════════════════════════════════════════
    # FileManagerMixin stubs  (EIS has no cycles or IR/RHE correction)
    # ════════════════════════════════════════════════════════════════
    def _populate_cycle_checkboxes(self, cycles, selected):
        pass

    def _selected_cycles(self):
        return []

    # ════════════════════════════════════════════════════════════════
    # Column-type → unit combo helpers
    # ════════════════════════════════════════════════════════════════
    def _update_unit_combo(self, col, unit_cb, unit_var):
        """Set unit_cb options to match the physical dimension of col.

        If the currently selected unit is no longer valid for the new column
        type it is reset to '(auto)'.
        """
        opts = _UNITS_BY_TYPE.get(_col_type(col) if col else "impedance",
                                  _OHM_UNITS)
        unit_cb["values"] = opts
        if unit_var.get() not in opts:
            unit_var.set("(auto)")

    def _on_x_col_change(self, event=None):
        self._update_unit_combo(self.x_var.get(), self.x_unit_cb, self.x_unit_var)
        self._auto_replot()

    def _on_y_col_change(self, event=None):
        self._update_unit_combo(self.y_var.get(), self.y_unit_cb, self.y_unit_var)
        self._auto_replot()

    # ════════════════════════════════════════════════════════════════
    # Unit conversion
    # ════════════════════════════════════════════════════════════════
    def _get_unit_scale(self, col, target_unit):
        """Return (scale_factor, display_label) for any EIS column.

        Handles impedance (Ω/kΩ/MΩ/mΩ), phase (°/rad), and
        frequency (Hz/kHz/MHz) columns.  "(auto)" → no conversion.
        """
        if not target_unit or target_unit == "(auto)":
            if "/" in col:
                base, src = col.rsplit("/", 1)
                return 1.0, f"{base.strip()} ({src.strip()})"
            return 1.0, col

        if "/" in col:
            base, src_unit = col.rsplit("/", 1)
            base     = base.strip()
            src_unit = src_unit.strip()
        else:
            base     = col
            src_unit = None

        display_label = f"{base} ({target_unit})"

        # ── Impedance ─────────────────────────────────────────────
        src_f = _IMP_RAW_MAP.get(src_unit)
        tgt_f = _IMP_DISP_MAP.get(target_unit)
        if src_f is not None and tgt_f is not None:
            return src_f / tgt_f, display_label

        # ── Phase ─────────────────────────────────────────────────
        if target_unit in ("°", "rad"):
            if src_unit in _PHASE_SRC_DEG:
                scale = 1.0 if target_unit == "°" else math.pi / 180.0
                return scale, display_label
            if src_unit in _PHASE_SRC_RAD:
                scale = 180.0 / math.pi if target_unit == "°" else 1.0
                return scale, display_label

        # ── Frequency ─────────────────────────────────────────────
        if target_unit in _FREQ_MAP:
            src_hz = _FREQ_MAP.get(src_unit)
            tgt_hz = _FREQ_MAP[target_unit]
            if src_hz is not None:
                return src_hz / tgt_hz, display_label

        return 1.0, display_label

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

    # ════════════════════════════════════════════════════════════════
    # State save / restore
    # ════════════════════════════════════════════════════════════════
    def _save_active_state(self):
        """Save per-file UI state for the currently active file."""
        super()._save_active_state()   # saves r_sol, e_ref, selected_cycles (harmless)
        if not self.active_file or self.active_file not in self.files:
            return
        entry = self.files[self.active_file]
        entry["x_col"]         = self.x_var.get()
        entry["y_col"]         = self.y_var.get()
        entry["x_unit"]        = self.x_unit_var.get()
        entry["y_unit"]        = self.y_unit_var.get()
        entry["x_min"]         = self.x_min_var.get()
        entry["x_max"]         = self.x_max_var.get()
        entry["y_min"]         = self.y_min_var.get()
        entry["y_max"]         = self.y_max_var.get()
        entry["show_markers"]  = self.show_markers_var.get()
        entry["connect_lines"] = self.connect_lines_var.get()
        entry["show_legend"]   = self.legend_show_var.get()
        entry["legend_frame"]  = self.legend_frame_var.get()
        try:
            entry["legend_size"] = float(self.legend_size_var.get())
        except ValueError:
            pass
        entry["legend_loc"]    = self.legend_loc_var.get()
        entry["x_grid"]        = self.x_grid_var.get()
        entry["y_grid"]        = self.y_grid_var.get()
        entry["x_grid_int"]    = self.x_grid_int_var.get()
        entry["y_grid_int"]    = self.y_grid_int_var.get()
        entry["grid_style"]    = self.grid_style_var.get()
        entry["grid_color"]     = self.grid_color_var.get()
        entry["grid_linewidth"] = self.grid_linewidth_var.get()
        entry["plot_style"]     = self.plot_style_var.get()
        entry["linewidth"]      = self.linewidth_var.get()
        # Persist current view so it can be restored after a file switch
        entry["view_xlim"] = self.ax.get_xlim()
        entry["view_ylim"] = self.ax.get_ylim()

    def _switch_active_file(self, short):
        """Switch the UI to *short*, restoring all per-file EIS settings."""
        self.active_file = short
        entry = self.files[short]
        df    = entry["df"]

        # Initialise missing per-file fields on the first visit
        entry.setdefault("x_col",         None)
        entry.setdefault("y_col",         None)
        entry.setdefault("x_unit",        "(auto)")
        entry.setdefault("y_unit",        "(auto)")
        entry.setdefault("x_min",         "0")
        entry.setdefault("x_max",         "100")
        entry.setdefault("y_min",         "0")
        entry.setdefault("y_max",         "100")
        entry.setdefault("show_markers",  True)
        entry.setdefault("connect_lines", True)
        entry.setdefault("show_legend",   True)
        entry.setdefault("legend_frame",  True)
        entry.setdefault("legend_size",   8.0)
        entry.setdefault("legend_loc",    "best")
        entry.setdefault("x_grid",        False)
        entry.setdefault("y_grid",        False)
        entry.setdefault("x_grid_int",    "0")
        entry.setdefault("y_grid_int",    "0")
        entry.setdefault("grid_style",    "dashed")
        entry.setdefault("grid_color",     "gray")
        entry.setdefault("grid_linewidth", "0.8")
        entry.setdefault("reflines",      [])
        entry.setdefault("plot_style",    "Line+Circle")

        cols = list(df.columns)
        self.x_combo["values"] = cols
        self.y_combo["values"] = cols

        # Resolve X column: prefer saved, then auto-detect Re(Z), fallback to col[0]
        x_col = entry["x_col"]
        if not x_col or x_col not in cols:
            x_col = next(
                (c for c in cols if any(k in c.lower()
                 for k in ("re(z)", "zre", "z'", "real"))),
                cols[0],
            )
            entry["x_col"] = x_col

        # Resolve Y column: prefer saved, then auto-detect -Im(Z), fallback to col[1]
        y_col = entry["y_col"]
        if not y_col or y_col not in cols:
            y_col = next(
                (c for c in cols if any(k in c.lower()
                 for k in ("-im(z)", "zim", "-im", "z''", "imag"))),
                cols[1] if len(cols) > 1 else cols[0],
            )
            entry["y_col"] = y_col

        # Refresh unit combo options for the resolved columns
        self._update_unit_combo(x_col, self.x_unit_cb, self.x_unit_var)
        self._update_unit_combo(y_col, self.y_unit_cb, self.y_unit_var)

        # Restore all UI vars while suppressing intermediate replots
        old = self._suppress_replot
        self._suppress_replot = True

        self.x_var.set(x_col)
        self.y_var.set(y_col)
        # Only restore saved unit if it is valid for the current column type
        _xopts = list(self.x_unit_cb["values"])
        _yopts = list(self.y_unit_cb["values"])
        self.x_unit_var.set(entry["x_unit"] if entry["x_unit"] in _xopts else "(auto)")
        self.y_unit_var.set(entry["y_unit"] if entry["y_unit"] in _yopts else "(auto)")
        self.x_min_var.set(entry["x_min"])
        self.x_max_var.set(entry["x_max"])
        self.y_min_var.set(entry["y_min"])
        self.y_max_var.set(entry["y_max"])
        self.show_markers_var.set(entry["show_markers"])
        self.connect_lines_var.set(entry["connect_lines"])
        self.legend_show_var.set(entry["show_legend"])
        self.legend_frame_var.set(entry["legend_frame"])
        sz = entry["legend_size"]
        try:
            sz_f = float(sz)
            self.legend_size_var.set(str(int(sz_f)) if sz_f == int(sz_f) else str(sz_f))
        except (ValueError, TypeError):
            self.legend_size_var.set("8")
        self.legend_loc_var.set(entry["legend_loc"])
        self.x_grid_var.set(entry["x_grid"])
        self.y_grid_var.set(entry["y_grid"])
        self.x_grid_int_var.set(entry["x_grid_int"])
        self.y_grid_int_var.set(entry["y_grid_int"])
        self.grid_style_var.set(entry["grid_style"])
        self.grid_color_var.set(entry["grid_color"])
        self.grid_linewidth_var.set(entry["grid_linewidth"])

        self._suppress_replot = old

        self._auto_replot()

        # Restore per-file zoom/pan view saved before the last file switch
        if "view_xlim" in entry:
            self.ax.set_xlim(entry["view_xlim"])
            self.ax.set_ylim(entry["view_ylim"])
            self.canvas.draw_idle()

        self._refresh_reflines_lb()

        # Restore color and shape comboboxes to match this file's stored settings
        color = entry.get("color", "#1f77b4")
        name = next((n for n, h in _COLOR_HEX.items() if h == color), "Blue")
        self.file_color_var.set(name)
        self.plot_style_var.set(entry.get("plot_style", "Line+Circle"))
        self.linewidth_var.set(entry.get("linewidth", "1.5"))

    # ════════════════════════════════════════════════════════════════
    # Plot
    # ════════════════════════════════════════════════════════════════
    def _auto_replot(self):
        if self._suppress_replot:
            return
        self._plot()

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

    # ── Plot size helper ─────────────────────────────────────────────
    def _auto_set_initial_size(self):
        """Resize the Nyquist figure to fill the right panel on first show."""
        w = self._plot_sc.winfo_width()
        h = self._plot_sc.winfo_height()
        if w <= 1 or h <= 1:
            self.after(100, self._auto_set_initial_size)
            return
        dpi = 100
        plot_w = max(4.0, (w - 20) / dpi)
        plot_h = max(3.0, (h - 50) / dpi)
        self.plot_w_var.set(f"{plot_w:.1f}")
        self.plot_h_var.set(f"{plot_h:.1f}")
        self._apply_plot_size()

    def _apply_plot_size(self, event=None):
        """Resize the figure to the current plot_w_var × plot_h_var (inches)."""
        try:
            w = float(self.plot_w_var.get())
            h = float(self.plot_h_var.get())
        except ValueError:
            return
        w = max(2.0, min(50.0, w))
        h = max(1.5, min(50.0, h))
        dpi = 100
        self.fig.set_size_inches(w, h)
        self.canvas.get_tk_widget().config(width=int(w * dpi), height=int(h * dpi))
        _leg = self.ax.get_legend()
        if _leg is not None: _leg.set_visible(False)
        self.fig.tight_layout(pad=0.5)
        self.fig.set_layout_engine('none')
        if _leg is not None: _leg.set_visible(True)
        self.canvas.draw_idle()
        self._plot_sc.after(
            50, lambda: self._plot_sc.configure(
                scrollregion=self._plot_sc.bbox("all")))

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
        ax.figure.set_layout_engine('none')
        if _leg is not None:
            _leg.set_visible(True)
        canvas.draw()
        if kb:
            for lbl in ax.get_xticklabels() + ax.get_yticklabels():
                lbl.set_fontweight('bold')
            if _leg is not None:
                _leg.set_visible(False)
            ax.figure.tight_layout()
            ax.figure.set_layout_engine('none')
            if _leg is not None:
                _leg.set_visible(True)
            canvas.draw()

    def _plot(self):
        if not self.active_file:
            return
        entry = self.files.get(self.active_file)
        if entry is None:
            return

        xcol   = self.x_var.get()
        ycol   = self.y_var.get()
        if not xcol or not ycol:
            return

        x_unit = self.x_unit_var.get()
        y_unit = self.y_unit_var.get()
        x_scale, x_label = self._get_unit_scale(xcol, x_unit)
        y_scale, y_label = self._get_unit_scale(ycol, y_unit)

        show_markers   = self.show_markers_var.get()
        connect_lines  = self.connect_lines_var.get()
        show_leg       = self.legend_show_var.get()
        leg_frame      = self.legend_frame_var.get()
        leg_loc        = self.legend_loc_var.get() or "best"
        try:
            leg_size = float(self.legend_size_var.get())
        except ValueError:
            leg_size = 8.0

        # Save dragged legend position before clearing
        if self._legend_obj is not None:
            _loc = getattr(self._legend_obj, '_loc', None)
            if isinstance(_loc, (tuple, list)):
                self._legend_manual_pos = tuple(_loc)
        self._legend_obj = None
        self._clear_annotation(redraw=False)
        self.ax.clear()

        try:
            _lw = float(self.linewidth_var.get())
        except (ValueError, TypeError):
            _lw = 1.5

        has_data = False
        self._legend_stable_keys = []   # reset for this replot
        # Rank 1 (index 0, top of file list) drawn last → appears in front
        _visible_items = [(s, e) for s, e in self.files.items()
                          if not e.get("hidden", False)]
        for short, fentry in reversed(_visible_items):
            df = fentry["df"]
            if xcol not in df.columns or ycol not in df.columns:
                continue
            base_color   = fentry.get("color",  "#1f77b4")
            _ls_style, _mk_style, _ms_style = _PLOT_STYLES.get(
                fentry.get("plot_style", "Line+Circle"), ("-", "o", 5))
            file_marker  = (_mk_style or "o") if show_markers else ""
            file_line    = "-" if connect_lines else ""
            if not file_marker and not file_line:
                file_marker = "o"
            self.ax.plot(
                df[xcol] * x_scale,
                df[ycol] * y_scale,
                color=base_color,
                marker=file_marker or None,
                markersize=_ms_style if file_marker else 0,
                linestyle=file_line,
                linewidth=_lw,
                label=short,
            )
            self._legend_stable_keys.append(short)
            has_data = True

        try: _lpad = float(self.label_pad_var.get())
        except Exception: _lpad = 4.0
        try: _tpad = float(self.title_pad_var.get())
        except Exception: _tpad = 6.0
        self.ax.set_xlabel(x_label, labelpad=_lpad)
        self.ax.set_ylabel(y_label, labelpad=_lpad)
        self.ax.set_title("Nyquist Plot", pad=_tpad)

        # Draw once to settle constrained_layout before capturing auto limits
        self.canvas.draw()
        self._auto_xlim = self.ax.get_xlim()
        self._auto_ylim = self.ax.get_ylim()

        self._apply_range()

        # Reference lines from the active file only
        draw_reflines(self.ax, entry.get("reflines", []))

        apply_grid(
            self.ax,
            self.x_grid_var.get(), self.y_grid_var.get(),
            self.x_grid_int_var.get(), self.y_grid_int_var.get(),
            self.grid_style_var.get(),
            linewidth=self.grid_linewidth_var.get(),
            color=self.grid_color_var.get(),
        )

        if show_leg and has_data:
            self._legend_obj = self.ax.legend(fontsize=leg_size, loc=leg_loc)
            self._legend_obj.set_draggable(True)
            self._legend_obj.get_frame().set_visible(leg_frame)
            self._current_legend_size = leg_size
            # Capture auto-labels, then apply custom labels by stable key
            self._legend_auto_labels = [t.get_text() for t in self._legend_obj.get_texts()]
            for i, text_obj in enumerate(self._legend_obj.get_texts()):
                if i < len(self._legend_stable_keys):
                    custom = self._legend_stable_map.get(self._legend_stable_keys[i])
                    if custom:
                        text_obj.set_text(custom)
            # Restore dragged position
            if self._legend_manual_pos is not None:
                self._legend_obj._loc = self._legend_manual_pos

        self._apply_font_to_ax(self.ax, self.canvas)

    def _apply_range(self):
        """Apply manual axis limits from the range entry fields."""
        changed = False
        for setter, kwarg, var in (
            (self.ax.set_xlim, "left",   self.x_min_var),
            (self.ax.set_xlim, "right",  self.x_max_var),
            (self.ax.set_ylim, "bottom", self.y_min_var),
            (self.ax.set_ylim, "top",    self.y_max_var),
        ):
            try:
                setter(**{kwarg: float(var.get())})
                changed = True
            except (ValueError, TypeError):
                pass

        # Flip axes if requested
        xl = self.ax.get_xlim()
        if self.x_flip_var.get() != (xl[0] > xl[1]):
            self.ax.set_xlim(xl[1], xl[0])
            changed = True
        yl = self.ax.get_ylim()
        if self.y_flip_var.get() != (yl[0] > yl[1]):
            self.ax.set_ylim(yl[1], yl[0])
            changed = True

        if changed:
            self.canvas.draw_idle()

    def _schedule_range_replot(self):
        """Debounce range-entry changes (400 ms) to avoid replotting on every keystroke."""
        if self._range_replot_id is not None:
            self.after_cancel(self._range_replot_id)
        self._range_replot_id = self.after(400, self._do_range_replot)

    def _do_range_replot(self):
        self._range_replot_id = None
        self._auto_replot()

    def _clear_plot(self):
        """Called when all files are removed."""
        self._legend_obj = None
        self._clear_annotation(redraw=False)
        self.ax.clear()
        self.ax.set_title("")
        self.canvas.draw()
        self._auto_xlim = None
        self._auto_ylim = None

    def _apply_plot_height(self):
        """Constrain the canvas widget height, or restore auto-fill if blank."""
        widget = self.canvas.get_tk_widget()
        val = self._plot_height_var.get().strip()
        if val:
            try:
                h_px = int(val)
                if h_px > 0:
                    widget.pack_forget()
                    widget.pack(fill=tk.X, expand=False)
                    widget.config(height=h_px)
                    self.canvas.draw_idle()
                    return
            except ValueError:
                pass
        # Blank or invalid → auto-fill
        widget.pack_forget()
        widget.pack(fill=tk.BOTH, expand=True)
        self.canvas.draw_idle()

    def _reset_view(self):
        """Restore the auto-scaled limits from the last _plot() call (Home button)."""
        if self._auto_xlim is not None:
            self.ax.set_xlim(self._auto_xlim)
            self.ax.set_ylim(self._auto_ylim)
            self.canvas.draw_idle()

    # ════════════════════════════════════════════════════════════════
    # Mouse interactions (scroll-zoom, pan, legend resize)
    # ════════════════════════════════════════════════════════════════
    def _on_scroll(self, event):
        if event.inaxes is not self.ax:
            return
        scale = 0.8 if event.step > 0 else 1.25
        xl, yl = self.ax.get_xlim(), self.ax.get_ylim()
        xd, yd = event.xdata, event.ydata
        xf = (xd - xl[0]) / (xl[1] - xl[0]) if xl[1] != xl[0] else 0.5
        yf = (yd - yl[0]) / (yl[1] - yl[0]) if yl[1] != yl[0] else 0.5
        nxr = (xl[1] - xl[0]) * scale
        nyr = (yl[1] - yl[0]) * scale
        self.ax.set_xlim(xd - nxr * xf, xd + nxr * (1 - xf))
        self.ax.set_ylim(yd - nyr * yf, yd + nyr * (1 - yf))
        self.canvas.draw_idle()

    def _event_on_legend(self, event):
        if self._legend_obj is None:
            return False
        try:
            renderer = self.fig.canvas.get_renderer()
            return self._legend_obj.get_window_extent(renderer).contains(event.x, event.y)
        except Exception:
            return False

    def _on_press(self, event):
        # Handle dblclick on title strip even when the click is outside the axes proper
        if event.button == 1 and getattr(event, "dblclick", False):
            try:
                r        = self.canvas.get_renderer()
                ax_bbox  = self.ax.get_window_extent(r)
                fig_bbox = self.ax.get_figure().get_window_extent(r)
                t_bbox   = self.ax.title.get_window_extent(r)
                on_title = (
                    (t_bbox.width > 2 and t_bbox.contains(event.x, event.y))
                    or (ax_bbox.x0 <= event.x <= ax_bbox.x1
                        and ax_bbox.y1 <= event.y <= fig_bbox.y1)
                )
                if on_title:
                    self._edit_plot_title()
                    return
            except Exception:
                pass

        if event.inaxes is not self.ax:
            return
        on_leg = self._event_on_legend(event)
        if event.button == 1:
            self._pan_moved = False
            if on_leg:
                if getattr(event, "dblclick", False):
                    self._edit_legend_labels()
            else:
                if getattr(event, "dblclick", False):
                    return   # dblclick inside axes but not on legend/title — ignore
                self._panning   = True
                self._pan_start = (event.xdata, event.ydata)
        elif event.button == 3 and on_leg:
            self._legend_resizing = True
            self._resize_start_y  = event.y
            self._resize_start_sz = self._current_legend_size

    def _on_release(self, event):
        self._panning = False
        was_resizing  = self._legend_resizing
        self._legend_resizing = False
        if was_resizing:
            self.legend_size_var.set(str(int(round(self._current_legend_size))))
            return

        on_leg = self._event_on_legend(event)
        if (event.button == 1
                and not self._pan_moved
                and event.inaxes is self.ax
                and not on_leg):
            self._handle_click_annotate(event)
        elif event.button == 3 and event.inaxes is self.ax and not on_leg:
            self._clear_annotation()

    def _on_motion(self, event):
        if self._panning and event.inaxes is self.ax and event.xdata is not None:
            self._pan_moved = True
            dx = self._pan_start[0] - event.xdata
            dy = self._pan_start[1] - event.ydata
            self.ax.set_xlim(self.ax.get_xlim()[0] + dx, self.ax.get_xlim()[1] + dx)
            self.ax.set_ylim(self.ax.get_ylim()[0] + dy, self.ax.get_ylim()[1] + dy)
            self.canvas.draw_idle()
            return

        if self._legend_resizing and self._legend_obj is not None:
            dy     = event.y - self._resize_start_y
            new_sz = max(4.0, min(30.0, self._resize_start_sz + dy / 5.0))
            self._current_legend_size = new_sz
            for t in self._legend_obj.get_texts():
                t.set_fontsize(new_sz)
            tt = self._legend_obj.get_title()
            if tt:
                tt.set_fontsize(new_sz)
            self.canvas.draw()

    # ════════════════════════════════════════════════════════════════
    # Click annotation (left-click → nearest point coords; right-click → dismiss)
    # ════════════════════════════════════════════════════════════════
    def _handle_click_annotate(self, event):
        import numpy as np
        _CLICK_CYCLE_PX = 8

        lines = [ln for ln in self.ax.lines
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
            disp  = self.ax.transData.transform(
                np.column_stack([xd[mask], yd[mask]]))
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
            self._click_candidate_idx = (
                (self._click_candidate_idx + 1) % len(candidates))
        else:
            self._click_candidate_idx = 0
        self._last_click_pos = (event.x, event.y)

        idx = self._click_candidate_idx
        n   = len(candidates)
        _, ln, x, y = candidates[idx]
        label = ln.get_label() or "?"

        xlim  = self.ax.get_xlim()
        ylim  = self.ax.get_ylim()
        xf    = (x - xlim[0]) / (xlim[1] - xlim[0]) if xlim[1] != xlim[0] else 0.5
        yf    = (y - ylim[0]) / (ylim[1] - ylim[0]) if ylim[1] != ylim[0] else 0.5
        xoff  = -95 if xf > 0.65 else 15
        yoff  = -60 if yf > 0.65 else 15

        order_hint = f"  [{idx + 1}/{n}]" if n > 1 else ""
        text = f"x = {x:.4g}\ny = {y:.4g}\n{label}{order_hint}"
        if n > 1 and idx == 0:
            text += "\n↻ click again to cycle"

        self._clear_annotation(redraw=False)
        self._ann = self.ax.annotate(
            text,
            xy=(x, y),
            xytext=(xoff, yoff),
            textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.4", fc="lightyellow", ec="gray", alpha=0.92),
            arrowprops=dict(arrowstyle="->", color="gray", lw=1.2),
            fontsize=8,
            zorder=10,
        )
        self._ann_dot, = self.ax.plot(
            x, y, "o",
            color=ln.get_color(),
            markersize=7,
            zorder=11,
            label="_ann_dot",
        )
        self.canvas.draw_idle()

    def _clear_annotation(self, redraw=True):
        for attr in ("_ann", "_ann_dot"):
            artist = getattr(self, attr, None)
            if artist is not None:
                try:
                    artist.remove()
                except Exception:
                    pass
                setattr(self, attr, None)
        self._last_click_pos      = None
        self._click_candidate_idx = 0
        if redraw:
            self.canvas.draw_idle()

    # ════════════════════════════════════════════════════════════════
    # Legend helpers
    # ════════════════════════════════════════════════════════════════
    def _toggle_legend(self):
        self._auto_replot()

    def _toggle_legend_frame(self):
        if self._legend_obj is not None:
            self._legend_obj.get_frame().set_visible(self.legend_frame_var.get())
            self.canvas.draw()
        else:
            self._auto_replot()

    def _on_legend_size_change(self):
        if self._legend_obj is not None:
            try:
                sz = float(self.legend_size_var.get())
            except ValueError:
                return
            self._current_legend_size = sz
            for t in self._legend_obj.get_texts():
                t.set_fontsize(sz)
            self.canvas.draw()
        else:
            self._auto_replot()

    def _edit_legend_labels(self):
        if self._legend_obj is None:
            from tkinter import messagebox
            messagebox.showinfo("Info", "Plot data first to create a legend.")
            return
        self._legend_obj.set_draggable(False)
        self._legend_obj = open_legend_editor(
            self, self._legend_obj, self.canvas, self._current_legend_size)
        if self._legend_obj is not None:
            self._legend_obj.set_draggable(True)
            # Persist labels by stable key so they survive file show/hide transitions.
            new_texts = [t.get_text() for t in self._legend_obj.get_texts()]
            for i, (key, new_text) in enumerate(
                    zip(self._legend_stable_keys, new_texts)):
                auto = (self._legend_auto_labels[i]
                        if i < len(self._legend_auto_labels) else new_text)
                if new_text and new_text != auto:
                    self._legend_stable_map[key] = new_text
                else:
                    self._legend_stable_map.pop(key, None)

    def _edit_plot_title(self):
        """Prompt the user to edit the Nyquist plot title (double-click on title area)."""
        from tkinter.simpledialog import askstring
        current = self.ax.title.get_text()
        new_title = askstring("Edit Title", "Plot title:", initialvalue=current, parent=self)
        if new_title is not None:
            self.ax.set_title(new_title)
            self.canvas.draw_idle()

    # ════════════════════════════════════════════════════════════════
    # Reference line helpers  (per active file)
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
        self.files[self.active_file].setdefault("reflines", []).append(
            ("x", v, style, color, lw))
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
        self.files[self.active_file].setdefault("reflines", []).append(
            ("y", v, style, color, lw))
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
        for axis, val, *_ in (self.files.get(self.active_file, {})
                               .get("reflines", [])):
            self._reflines_lb.insert(
                tk.END, f"{'X' if axis == 'x' else 'Y'} = {val:.4g}")

    def _on_refline_select(self):
        sel = self._reflines_lb.curselection()
        if not sel or not self.active_file:
            return
        reflines = self.files.get(self.active_file, {}).get("reflines", [])
        if sel[0] >= len(reflines):
            return
        entry = reflines[sel[0]]
        self._refline_style_var.set(entry[2])
        self._refline_color_var.set(entry[3])
        self._refline_linewidth_var.set(str(entry[4]) if len(entry) > 4 else "1.0")

    def _on_refline_style_color_change(self):
        sel = self._reflines_lb.curselection()
        if not sel or not self.active_file:
            return
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
