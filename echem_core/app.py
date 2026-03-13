"""Main application window – assembles all mixins and builds the UI."""

from collections import OrderedDict

import tkinter as tk
from tkinter import ttk
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from .legend_editor import open_legend_editor
from .checklist import CheckableListbox
from .file_manager import FileManagerMixin, _COLOR_NAMES, _COLOR_HEX, _is_impedance_col, _PLOT_STYLES, _PLOT_STYLE_NAMES
from .correction import CorrectionMixin
from .plotting import PlottingMixin, copy_figure_to_clipboard
from .ecsa import ECSAMixin
from .export import ExportMixin
from .ecsa_panel import ECSAPanel
from .multi_echem_panel import MultiEchemPanel
from .multi_echem2_panel import MultiEchem2Panel
from .eis_panel import EISPanel

_CYCLE_BG = "#e8f0fe"   # light blue for the cycle checkbox area
_CYCLE_ACTIVE_BG = "#cce0ff"


class EchemPanel(
    FileManagerMixin,
    CorrectionMixin,
    PlottingMixin,
    ECSAMixin,
    ExportMixin,
    ttk.Frame,
):
    """A self-contained electrochemistry panel (left controls + right plot).

    Each instance has its own file state and matplotlib figure, so two panels
    placed in separate notebook tabs operate completely independently.

    Parameters
    ----------
    master     : parent widget (e.g. a ttk.Frame tab)
    show_ecsa  : include the ECSA Calc section in the left panel
    show_log   : include the Log section at the bottom of the left panel
    """

    def __init__(self, master, *, show_ecsa=False, show_log=False):
        ttk.Frame.__init__(self, master)
        # Independent state for every panel instance
        self.files = OrderedDict()
        self.active_file = None
        self._suppress_replot = False
        self._loading_files = False
        self._build_panel(show_ecsa=show_ecsa, show_log=show_log)

    # ── Panel construction ───────────────────────────────────────────
    def _build_panel(self, show_ecsa=False, show_log=False):
        body = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True)

        # ── Scrollable left panel ───────────────────────────────────
        left_outer = ttk.Frame(body, width=290)
        body.add(left_outer, weight=0)

        _left_canvas = tk.Canvas(left_outer, highlightthickness=0)
        _left_scroll = ttk.Scrollbar(left_outer, orient=tk.VERTICAL, command=_left_canvas.yview)
        _left_canvas.configure(yscrollcommand=_left_scroll.set)
        _left_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        _left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        left = tk.Frame(_left_canvas)
        _left_win = _left_canvas.create_window((0, 0), window=left, anchor=tk.NW)

        def _on_left_cfg(e):
            _left_canvas.configure(scrollregion=_left_canvas.bbox("all"))
        left.bind("<Configure>", _on_left_cfg)

        def _on_left_canvas_cfg(e):
            _left_canvas.itemconfig(_left_win, width=e.width)
        _left_canvas.bind("<Configure>", _on_left_canvas_cfg)

        def _on_left_wheel(e):
            _left_canvas.yview_scroll(-1 * (e.delta // 120), "units")
        _left_canvas.bind("<MouseWheel>", _on_left_wheel)

        # ── File list ───────────────────────────────────────────────
        ttk.Label(left, text="Files:", font=("", 9, "bold")).pack(anchor=tk.W, padx=4, pady=(6, 0))
        file_btn_row = ttk.Frame(left)
        file_btn_row.pack(fill=tk.X, padx=4)
        ttk.Button(file_btn_row, text="Load File(s)", command=self._load_files).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(file_btn_row, text="Remove", command=self._remove_file).pack(side=tk.LEFT)

        file_list_frame = ttk.Frame(left)
        file_list_frame.pack(fill=tk.X, padx=4, pady=2)
        self.file_listbox = CheckableListbox(file_list_frame, height=5,
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
        self.linewidth_var = tk.StringVar(value="3")
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

        # ── Axis selectors + unit conversion ────────────────────────
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
        _ALL_UNITS = ["(auto)",
                      "A", "mA", "µA", "nA",
                      "V", "mV", "µV", "nV",
                      "s", "ms", "µs", "min", "h"]

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

        def _refresh_unit_after_select(unit_var, unit_cb):
            chosen = unit_var.get()
            if chosen and chosen != "(auto)":
                if chosen.endswith("/cm²"):
                    opts = list(_DIM_OPTS["J"])
                else:
                    dim = _UNIT_DIMS.get(chosen)
                    opts = list(_DIM_OPTS.get(dim, _ALL_UNITS))
                unit_cb["values"] = opts
            self._auto_replot()

        def _refresh_j_in_combos():
            """Add or remove 'J' from column comboboxes based on area."""
            has_j = self._all_files_have_area()
            for combo, var in ((self.x_combo, self.x_var),
                               (self.y_combo, self.y_var)):
                vals = list(combo["values"])
                if has_j and "J" not in vals:
                    vals.append("J")
                    combo["values"] = vals
                elif not has_j and "J" in vals:
                    vals.remove("J")
                    combo["values"] = vals
                    if var.get() == "J":
                        # Fall back to first column
                        if vals:
                            var.set(vals[0])

        # X-axis: defaults to voltage (V)
        ttk.Label(left, text="X-axis:").pack(anchor=tk.W, padx=4, pady=(6, 0))
        x_axis_row = ttk.Frame(left)
        x_axis_row.pack(fill=tk.X, padx=4, pady=2)
        self.x_var = tk.StringVar()
        self.x_combo = ttk.Combobox(x_axis_row, textvariable=self.x_var, state="readonly", width=16)
        self.x_combo.pack(side=tk.LEFT)
        self.x_unit_var = tk.StringVar(value="V")
        x_unit_cb = ttk.Combobox(x_axis_row, textvariable=self.x_unit_var,
                                  values=_DIM_OPTS["E"], state="readonly", width=8)
        x_unit_cb.pack(side=tk.LEFT, padx=(4, 0))
        self.x_combo.bind("<<ComboboxSelected>>",
                          lambda e: _refresh_unit_opts(self.x_var, self.x_unit_var, x_unit_cb))
        x_unit_cb.bind("<<ComboboxSelected>>",
                       lambda e: _refresh_unit_after_select(self.x_unit_var, x_unit_cb))

        # Y-axis: defaults to current (mA)
        ttk.Label(left, text="Y-axis:").pack(anchor=tk.W, padx=4, pady=(6, 0))
        y_axis_row = ttk.Frame(left)
        y_axis_row.pack(fill=tk.X, padx=4, pady=2)
        self.y_var = tk.StringVar()
        self.y_combo = ttk.Combobox(y_axis_row, textvariable=self.y_var, state="readonly", width=16)
        self.y_combo.pack(side=tk.LEFT)
        self.y_unit_var = tk.StringVar(value="mA")
        y_unit_cb = ttk.Combobox(y_axis_row, textvariable=self.y_unit_var,
                                  values=_DIM_OPTS["I"], state="readonly", width=8)
        y_unit_cb.pack(side=tk.LEFT, padx=(4, 0))
        self.y_combo.bind("<<ComboboxSelected>>",
                          lambda e: _refresh_unit_opts(self.y_var, self.y_unit_var, y_unit_cb))
        y_unit_cb.bind("<<ComboboxSelected>>",
                       lambda e: _refresh_unit_after_select(self.y_unit_var, y_unit_cb))

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

        # Current density (per-file electrode area — unlocks "J" in column combos)
        area_row = ttk.Frame(left)
        area_row.pack(fill=tk.X, padx=4, pady=(2, 0))
        ttk.Label(area_row, text="Area (cm²):").pack(side=tk.LEFT)
        self.area_var = tk.StringVar()
        _area_e = ttk.Entry(area_row, textvariable=self.area_var, width=8)
        _area_e.pack(side=tk.LEFT, padx=(4, 0))

        def _on_area_change(e=None):
            # Persist area to the active file entry immediately
            if self.active_file and self.active_file in self.files:
                self.files[self.active_file]["area"] = self.area_var.get()
            # Update "J" availability in column combos, then refresh unit combos
            _refresh_j_in_combos()
            self._suppress_replot = True
            _refresh_unit_opts(self.x_var, self.x_unit_var, x_unit_cb)
            self._suppress_replot = False
            _refresh_unit_opts(self.y_var, self.y_unit_var, y_unit_cb)

        _area_e.bind("<Return>",   _on_area_change)
        _area_e.bind("<FocusOut>", _on_area_change)
        ttk.Label(area_row, text="all files need area for J",
                  foreground="gray", font=("", 8)).pack(side=tk.LEFT, padx=6)

        # Reference electrode
        ttk.Label(left, text="Reference Electrode:").pack(anchor=tk.W, padx=4, pady=(6, 0))
        self.ref_electrode_var = tk.StringVar(value="Ag/AgCl")
        ref_combo = ttk.Combobox(
            left, textvariable=self.ref_electrode_var,
            values=[
                "Ag/AgCl", "SCE", "SHE", "NHE", "RHE",
                "Hg/HgO", "Hg/HgSO4 (MSE)", "Fc/Fc+", "Ag/Ag+", "Li/Li+",
            ],
            state="readonly", width=24,
        )
        ref_combo.pack(padx=4, pady=2)
        ref_combo.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())

        # Plot range
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Plot Range", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        def _bind_range(entry):
            entry.bind("<Return>", lambda e: self._auto_replot())
            entry.bind("<FocusOut>", lambda e: self._auto_replot())

        xrange_frame = ttk.Frame(left)
        xrange_frame.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(xrange_frame, text="X min:").pack(side=tk.LEFT)
        self.x_min_var = tk.StringVar()
        _xmin = ttk.Entry(xrange_frame, textvariable=self.x_min_var, width=7)
        _xmin.pack(side=tk.LEFT, padx=(2, 4))
        _bind_range(_xmin)
        ttk.Label(xrange_frame, text="X max:").pack(side=tk.LEFT)
        self.x_max_var = tk.StringVar()
        _xmax = ttk.Entry(xrange_frame, textvariable=self.x_max_var, width=7)
        _xmax.pack(side=tk.LEFT, padx=(2, 4))
        _bind_range(_xmax)
        ttk.Label(xrange_frame, text="Int:").pack(side=tk.LEFT)
        self.x_grid_int_var = tk.StringVar(value="0")
        _xgi = ttk.Entry(xrange_frame, textvariable=self.x_grid_int_var, width=5)
        _xgi.pack(side=tk.LEFT, padx=(2, 0))
        _xgi.bind("<Return>",   lambda e: self._auto_replot())
        _xgi.bind("<FocusOut>", lambda e: self._auto_replot())

        yrange_frame = ttk.Frame(left)
        yrange_frame.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(yrange_frame, text="Y min:").pack(side=tk.LEFT)
        self.y_min_var = tk.StringVar()
        _ymin = ttk.Entry(yrange_frame, textvariable=self.y_min_var, width=7)
        _ymin.pack(side=tk.LEFT, padx=(2, 4))
        _bind_range(_ymin)
        ttk.Label(yrange_frame, text="Y max:").pack(side=tk.LEFT)
        self.y_max_var = tk.StringVar()
        _ymax = ttk.Entry(yrange_frame, textvariable=self.y_max_var, width=7)
        _ymax.pack(side=tk.LEFT, padx=(2, 4))
        _bind_range(_ymax)
        ttk.Label(yrange_frame, text="Int:").pack(side=tk.LEFT)
        self.y_grid_int_var = tk.StringVar(value="0")
        _ygi = ttk.Entry(yrange_frame, textvariable=self.y_grid_int_var, width=5)
        _ygi.pack(side=tk.LEFT, padx=(2, 0))
        _ygi.bind("<Return>",   lambda e: self._auto_replot())
        _ygi.bind("<FocusOut>", lambda e: self._auto_replot())

        ttk.Label(left, text="(leave blank for auto)", foreground="gray").pack(anchor=tk.W, padx=4)

        flip_row = ttk.Frame(left)
        flip_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        self.x_flip_var = tk.BooleanVar(value=False)
        self.y_flip_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(flip_row, text="Flip X", variable=self.x_flip_var,
                        command=self._auto_replot).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Checkbutton(flip_row, text="Flip Y", variable=self.y_flip_var,
                        command=self._auto_replot).pack(side=tk.LEFT)

        # Plot title
        title_row = ttk.Frame(left)
        title_row.pack(fill=tk.X, padx=4, pady=(2, 2))
        ttk.Label(title_row, text="Title:").pack(side=tk.LEFT)
        self.plot_title_var = tk.StringVar(value="Title")
        _title_entry = ttk.Entry(title_row, textvariable=self.plot_title_var)
        _title_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

        def _on_title_change(e=None):
            self.ax.set_title(self.plot_title_var.get())
            _fn = getattr(self, '_apply_font_to_ax', None)
            if _fn is not None:
                _fn(self.ax, self.canvas)
            else:
                self.canvas.draw_idle()

        _title_entry.bind("<Return>",   _on_title_change)
        _title_entry.bind("<FocusOut>", _on_title_change)

        # Legend
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Legend", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        leg_toggle_row = ttk.Frame(left)
        leg_toggle_row.pack(fill=tk.X, padx=4, pady=(0, 2))
        self.legend_show_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            leg_toggle_row, text="Show Legend", variable=self.legend_show_var,
            command=self._auto_replot,
        ).pack(side=tk.LEFT)
        self.legend_frame_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            leg_toggle_row, text="Show Frame", variable=self.legend_frame_var,
            command=self._auto_replot,
        ).pack(side=tk.LEFT, padx=(8, 0))

        leg_row = ttk.Frame(left)
        leg_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(leg_row, text="Size:").pack(side=tk.LEFT)
        self.legend_size_var = tk.StringVar(value="20")
        _leg_size_e = ttk.Entry(leg_row, textvariable=self.legend_size_var, width=5)
        _leg_size_e.pack(side=tk.LEFT, padx=(2, 8))
        _leg_size_e.bind("<Return>",   lambda e: self._auto_replot())
        _leg_size_e.bind("<FocusOut>", lambda e: self._auto_replot())
        ttk.Label(leg_row, text="Location:").pack(side=tk.LEFT)
        self.legend_loc_var = tk.StringVar(value="best")
        _leg_loc_cb = ttk.Combobox(
            leg_row, textvariable=self.legend_loc_var,
            values=[
                "best", "upper right", "upper left", "lower left", "lower right",
                "right", "center left", "center right", "lower center",
                "upper center", "center",
            ],
            state="readonly", width=12,
        )
        _leg_loc_cb.pack(side=tk.LEFT, padx=2)
        def _on_leg_loc_select(e=None):
            self._legend_manual_pos = None  # user explicitly chose a location
            # Also neutralise the live legend's _loc so _plot() doesn't
            # re-capture the old tuple before ax.clear() runs.
            if self._legend_obj is not None:
                self._legend_obj._loc = 0
            self._auto_replot()
        _leg_loc_cb.bind("<<ComboboxSelected>>", _on_leg_loc_select)

        ttk.Button(left, text="Edit Labels", command=self._edit_legend_labels).pack(anchor=tk.W, padx=4, pady=2)
        ttk.Label(left, text="(left-drag legend to move, right-drag to resize; dbl-click to edit)",
                  foreground="gray").pack(anchor=tk.W, padx=4)

        # Grid
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Grid", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        grid_xy_row = ttk.Frame(left)
        grid_xy_row.pack(fill=tk.X, padx=4, pady=2)
        self.x_grid_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(grid_xy_row, text="X", variable=self.x_grid_var,
                        command=self._auto_replot).pack(side=tk.LEFT)
        self.y_grid_var = tk.BooleanVar(value=True)
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
        self.grid_color_var = tk.StringVar(value="black")
        _gcol_cb = ttk.Combobox(grid_style_row, textvariable=self.grid_color_var,
                                values=["gray", "black", "red", "blue", "green",
                                        "orange", "purple", "crimson", "royalblue",
                                        "darkorange", "teal"],
                                state="readonly", width=9)
        _gcol_cb.pack(side=tk.LEFT, padx=(2, 6))
        _gcol_cb.bind("<<ComboboxSelected>>", lambda e: self._auto_replot())
        ttk.Label(grid_style_row, text="Width:").pack(side=tk.LEFT)
        self.grid_linewidth_var = tk.StringVar(value="2")
        _glw = ttk.Entry(grid_style_row, textvariable=self.grid_linewidth_var, width=4)
        _glw.pack(side=tk.LEFT, padx=(2, 0))
        _glw.bind("<Return>",   lambda e: self._auto_replot())
        _glw.bind("<FocusOut>", lambda e: self._auto_replot())

        # Font
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Font", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
        self.font_title_size_var = tk.StringVar(value="40")
        self.font_title_bold_var = tk.BooleanVar(value=False)
        self.font_label_size_var = tk.StringVar(value="30")
        self.font_label_bold_var = tk.BooleanVar(value=False)
        self.font_tick_size_var  = tk.StringVar(value="20")
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
        self.title_pad_var = tk.StringVar(value="20")
        _tpad_e = ttk.Entry(_spacing_row, textvariable=self.title_pad_var, width=4)
        _tpad_e.pack(side=tk.LEFT, padx=(2, 6))
        _tpad_e.bind("<Return>",   lambda e: self._auto_replot())
        _tpad_e.bind("<FocusOut>", lambda e: self._auto_replot())
        ttk.Label(_spacing_row, text="Label").pack(side=tk.LEFT)
        self.label_pad_var = tk.StringVar(value="20")
        _lpad_e = ttk.Entry(_spacing_row, textvariable=self.label_pad_var, width=4)
        _lpad_e.pack(side=tk.LEFT, padx=(2, 0))
        _lpad_e.bind("<Return>",   lambda e: self._auto_replot())
        _lpad_e.bind("<FocusOut>", lambda e: self._auto_replot())

        # Reference Lines
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Reference Lines", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
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
        self._reflines = []

        # IR / RHE correction
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="IR / RHE Correction", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        r_frame = ttk.Frame(left)
        r_frame.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(r_frame, text="R_sol (Ohm):").pack(side=tk.LEFT)
        self.r_sol_var = tk.StringVar(value="0")
        _rsol_e = ttk.Entry(r_frame, textvariable=self.r_sol_var, width=10)
        _rsol_e.pack(side=tk.LEFT, padx=4)
        _rsol_e.bind("<Return>",   lambda e: self._apply_correction())
        _rsol_e.bind("<FocusOut>", lambda e: self._apply_correction())

        e_frame = ttk.Frame(left)
        e_frame.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(e_frame, text="E_ref (V vs RHE):").pack(side=tk.LEFT)
        self.e_ref_var = tk.StringVar(value="0")
        _eref_e = ttk.Entry(e_frame, textvariable=self.e_ref_var, width=10)
        _eref_e.pack(side=tk.LEFT, padx=4)
        _eref_e.bind("<Return>",   lambda e: self._apply_correction())
        _eref_e.bind("<FocusOut>", lambda e: self._apply_correction())

        ttk.Label(left, text="(auto-applied on Enter / focus change)",
                  foreground="gray", font=("", 8)).pack(anchor=tk.W, padx=4)
        ttk.Button(left, text="Reset Correction",
                   command=self._reset_correction).pack(anchor=tk.W, padx=4, pady=(2, 0))

        # Cycle selector (checkboxes)
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Cycles:").pack(anchor=tk.W, padx=4)
        btn_row = ttk.Frame(left)
        btn_row.pack(fill=tk.X, padx=4)
        ttk.Button(btn_row, text="Select All", command=self._select_all).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_row, text="Deselect All", command=self._deselect_all).pack(side=tk.LEFT)

        cycle_outer = ttk.Frame(left)
        cycle_outer.pack(fill=tk.X, padx=4, pady=2)
        cycle_canvas = tk.Canvas(cycle_outer, background=_CYCLE_BG, highlightthickness=0, height=150)
        cyc_vscroll = ttk.Scrollbar(cycle_outer, orient=tk.VERTICAL, command=cycle_canvas.yview)
        cyc_hscroll = ttk.Scrollbar(cycle_outer, orient=tk.HORIZONTAL, command=cycle_canvas.xview)
        cycle_canvas.configure(yscrollcommand=cyc_vscroll.set, xscrollcommand=cyc_hscroll.set)
        cyc_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        cyc_hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        cycle_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._cycle_inner = tk.Frame(cycle_canvas, background=_CYCLE_BG)
        self._cycle_canvas_win = cycle_canvas.create_window((0, 0), window=self._cycle_inner, anchor=tk.NW)
        self._cycle_canvas = cycle_canvas

        def _on_inner_configure(e):
            cycle_canvas.configure(scrollregion=cycle_canvas.bbox("all"))
        self._cycle_inner.bind("<Configure>", _on_inner_configure)

        def _on_cycle_mousewheel(e):
            cycle_canvas.yview_scroll(-1 * (e.delta // 120), "units")
            return "break"
        cycle_canvas.bind("<MouseWheel>", _on_cycle_mousewheel)
        self._cycle_inner.bind("<MouseWheel>", _on_cycle_mousewheel)

        self._cycle_vars = {}

        plot_btn_row = ttk.Frame(left)
        plot_btn_row.pack(fill=tk.X, padx=4, pady=6)
        ttk.Button(plot_btn_row, text="Plot", command=self._plot).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(plot_btn_row, text="Export Excel", command=self._export_excel).pack(side=tk.LEFT)

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

        # ── Plot height ──────────────────────────────────────────────
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

        # ── Optional ECSA section ────────────────────────────────────
        if show_ecsa:
            ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
            ttk.Label(left, text="ECSA Calc", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

            sr_frame = ttk.Frame(left)
            sr_frame.pack(fill=tk.X, padx=4, pady=2)
            ttk.Label(sr_frame, text="Scan rate (mV/s):").pack(side=tk.LEFT)
            self.scan_rate_var = tk.StringVar(value="50")
            ttk.Entry(sr_frame, textvariable=self.scan_rate_var, width=8).pack(side=tk.LEFT, padx=4)

            ttk.Button(left, text="Calculate ECSA", command=self._calc_ecsa).pack(padx=4, pady=4)
            self.ecsa_label = ttk.Label(left, text="", wraplength=230)
            self.ecsa_label.pack(anchor=tk.W, padx=4, pady=2)
        else:
            # Keep vars alive so ECSAMixin._calc_ecsa never raises AttributeError
            self.scan_rate_var = tk.StringVar(value="50")
            self.ecsa_label = ttk.Label(left, text="")

        # ── Optional log section ─────────────────────────────────────
        if show_log:
            ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
            ttk.Label(left, text="Log", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)
            log_frame = ttk.Frame(left)
            log_frame.pack(fill=tk.X, padx=4, pady=2)
            self.log_text = tk.Text(log_frame, height=6, state=tk.DISABLED,
                                    wrap=tk.WORD, relief=tk.SUNKEN, borderwidth=1)
            log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
            self.log_text.configure(yscrollcommand=log_scroll.set)
            log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            self.log_text.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # ── Right panel – matplotlib ────────────────────────────────
        right = ttk.Frame(body)
        body.add(right, weight=1)

        self.fig = Figure(figsize=(6, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=right)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        toolbar_frame = ttk.Frame(right)
        toolbar_frame.pack(fill=tk.X, side=tk.BOTTOM)

        class _Toolbar(NavigationToolbar2Tk):
            def home(tb_self, *args):
                self._reset_view()
        toolbar = _Toolbar(self.canvas, toolbar_frame, pack_toolbar=False)
        toolbar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        toolbar.update()
        tk.Button(
            toolbar_frame, text="Copy",
            command=lambda: copy_figure_to_clipboard(self.fig),
            relief=tk.RAISED, borderwidth=1, padx=6,
        ).pack(side=tk.LEFT, padx=(4, 2), pady=1)

        self._init_plot_interactions()

    # ── Log helper ──────────────────────────────────────────────────
    def _log(self, message: str):
        """Append a line to the log widget (no-op if this panel has no log)."""
        log = getattr(self, "log_text", None)
        if log is None:
            return
        log.configure(state=tk.NORMAL)
        log.insert(tk.END, message + "\n")
        log.see(tk.END)
        log.configure(state=tk.DISABLED)

    # ── Cycle helpers ───────────────────────────────────────────────
    def _populate_cycle_checkboxes(self, cycles, selected):
        """Rebuild the checkbox list for the given cycles."""
        for w in self._cycle_inner.winfo_children():
            w.destroy()
        self._cycle_vars.clear()

        selected_set = set(selected)
        ncols = 9

        def _cb_wheel(e):
            self._cycle_canvas.yview_scroll(-1 * (e.delta // 120), "units")
            return "break"

        for i, c in enumerate(cycles):
            row_idx, col_idx = divmod(i, ncols)
            var = tk.BooleanVar(value=(c in selected_set))
            var.trace_add("write", self._on_cycle_toggle)
            cb = tk.Checkbutton(
                self._cycle_inner,
                text=f"C{c}",
                variable=var,
                background=_CYCLE_BG,
                activebackground=_CYCLE_ACTIVE_BG,
                selectcolor=_CYCLE_BG,
                anchor=tk.W,
            )
            cb.grid(row=row_idx, column=col_idx, sticky=tk.W, padx=2, pady=1)
            cb.bind("<MouseWheel>", _cb_wheel)
            self._cycle_vars[c] = var

        self._cycle_inner.update_idletasks()
        self._cycle_canvas.configure(scrollregion=self._cycle_canvas.bbox("all"))
        self._cycle_canvas.yview_moveto(0)

    def _on_cycle_toggle(self, *_args):
        if not self._suppress_replot:
            self._auto_replot()

    def _select_all(self):
        self._suppress_replot = True
        for var in self._cycle_vars.values():
            var.set(True)
        self._suppress_replot = False
        self._auto_replot()

    def _deselect_all(self):
        self._suppress_replot = True
        for var in self._cycle_vars.values():
            var.set(False)
        self._suppress_replot = False
        self._auto_replot()

    def _selected_cycles(self):
        return [c for c, var in self._cycle_vars.items() if var.get()]

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
            # Persist labels by stable key (file:cycle) so they survive
            # single↔multi-file display format changes.
            new_texts = [t.get_text() for t in self._legend_obj.get_texts()]
            for i, (key, new_text) in enumerate(
                    zip(self._legend_stable_keys, new_texts)):
                auto = (self._legend_auto_labels[i]
                        if i < len(self._legend_auto_labels) else new_text)
                if new_text and new_text != auto:
                    self._legend_stable_map[key] = new_text
                else:
                    self._legend_stable_map.pop(key, None)

    # ── Reference line helpers ───────────────────────────────────────
    def _add_xrefline(self):
        try:
            v = float(self._ref_x_var.get())
        except ValueError:
            return
        style = self._refline_style_var.get()
        color = self._refline_color_var.get()
        lw    = self._refline_linewidth_var.get()
        self._reflines.append(('x', v, style, color, lw))
        self._reflines_lb.insert(tk.END, f"X = {v:.4g}")
        self._auto_replot()

    def _add_yrefline(self):
        try:
            v = float(self._ref_y_var.get())
        except ValueError:
            return
        style = self._refline_style_var.get()
        color = self._refline_color_var.get()
        lw    = self._refline_linewidth_var.get()
        self._reflines.append(('y', v, style, color, lw))
        self._reflines_lb.insert(tk.END, f"Y = {v:.4g}")
        self._auto_replot()

    def _on_refline_select(self):
        """Populate the style/color/width widgets from the selected line's settings."""
        sel = self._reflines_lb.curselection()
        if not sel:
            return
        entry = self._reflines[sel[0]]
        self._refline_style_var.set(entry[2])
        self._refline_color_var.set(entry[3])
        self._refline_linewidth_var.set(entry[4] if len(entry) > 4 else "1.0")

    def _on_refline_style_color_change(self):
        """Apply new style/color/width to the currently selected reference line."""
        sel = self._reflines_lb.curselection()
        if not sel:
            return  # nothing selected — widgets just set defaults for next line
        idx = sel[0]
        axis, val = self._reflines[idx][:2]
        self._reflines[idx] = (axis, val,
                               self._refline_style_var.get(),
                               self._refline_color_var.get(),
                               self._refline_linewidth_var.get())
        self._auto_replot()

    def _remove_refline(self):
        sel = self._reflines_lb.curselection()
        if not sel:
            return
        idx = sel[0]
        self._reflines.pop(idx)
        self._reflines_lb.delete(idx)
        self._auto_replot()

    # ── Plot height helper ────────────────────────────────────────────
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

    # ── Line width helper ─────────────────────────────────────────────
    def _on_linewidth_change(self):
        """Persist line width to the active file's entry, then replot."""
        if self.active_file and self.active_file in self.files:
            self.files[self.active_file]["linewidth"] = self.linewidth_var.get()
        self._auto_replot()

    def _on_plot_style_change(self):
        """Persist plot shape/style to the active file's entry, then replot."""
        if self.active_file and self.active_file in self.files:
            self.files[self.active_file]["plot_style"] = self.plot_style_var.get()
        self._auto_replot()

    # ── Gradient helper ──────────────────────────────────────────────
    def _on_gradient_change(self):
        """Persist gradient settings to the active file's entry, then replot."""
        if self.active_file and self.active_file in self.files:
            self.files[self.active_file]["cycle_gradient"] = self.cycle_gradient_var.get()
            self.files[self.active_file]["cycle_reverse"]  = self.cycle_reverse_var.get()
            self.files[self.active_file]["lightness_step"] = self.lightness_step_var.get()
        self._auto_replot()

    # ── Font helpers ─────────────────────────────────────────────────
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

    # ── File color helper ────────────────────────────────────────────
    def _on_file_color_change(self, event=None):
        if not self.active_file:
            return
        self.files[self.active_file]["color"] = _COLOR_HEX.get(
            self.file_color_var.get(), "#1f77b4")
        self._auto_replot()

    # ── J / area helpers ─────────────────────────────────────────────
    def _all_files_have_area(self):
        """True only when every loaded file has a positive electrode area."""
        if not self.files:
            return False
        for short, entry in self.files.items():
            if short == self.active_file:
                try:
                    if float(self.area_var.get()) <= 0:
                        return False
                except (ValueError, TypeError):
                    return False
            else:
                try:
                    if float(entry.get("area", "") or 0) <= 0:
                        return False
                except (ValueError, TypeError):
                    return False
        return True

    def _get_column_list(self, df):
        """Append virtual 'J' column when all files have area set (non-EIS only)."""
        cols = super()._get_column_list(df)
        if self._all_files_have_area() and not any(_is_impedance_col(c) for c in cols):
            cols.append("J")
        return cols

    def _clear_plot(self):
        """Clear the plot when all files are removed."""
        self._clear_annotation(redraw=False)
        self._legend_obj = None
        self.ax.clear()
        self.canvas.draw()

    def _sync_file_selection_from_line(self, ln):
        """Select the file matching the clicked line and update the left panel.

        Suppresses replot (keeps annotation alive) and preserves the current
        view (avoids jarring pan/zoom jumps on annotation click).
        """
        label = ln.get_label() or ""
        if not label or label.startswith("_"):
            return
        for short in self.files:
            if label == short or label.startswith(short + " "):
                keys = list(self.files.keys())
                idx = keys.index(short)
                self._loading_files = True
                try:
                    self.file_listbox.selection_clear(0, tk.END)
                    self.file_listbox.selection_set(idx)
                    self.file_listbox.see(idx)
                finally:
                    self._loading_files = False
                if short != self.active_file:
                    self._save_active_state()
                    # Capture view so the file switch doesn't pan/zoom the shared axes
                    xlim = self.ax.get_xlim()
                    ylim = self.ax.get_ylim()
                    old = self._suppress_replot
                    self._suppress_replot = True
                    self._switch_active_file(short)
                    self._suppress_replot = old
                    # Restore the view that was in place before the switch
                    self.ax.set_xlim(xlim)
                    self.ax.set_ylim(ylim)
                break

    def _on_columns_changed(self):
        """Refresh unit combobox options whenever x_var/y_var are updated."""
        fn = getattr(self, '_do_refresh_unit_combos', None)
        if fn:
            fn()

    def _save_active_state(self):
        """Extend base save to preserve the current plot view and gradient settings per file."""
        if self.active_file and self.active_file in self.files:
            self.files[self.active_file]["view_xlim"]       = self.ax.get_xlim()
            self.files[self.active_file]["view_ylim"]       = self.ax.get_ylim()
            self.files[self.active_file]["cycle_gradient"]  = self.cycle_gradient_var.get()
            self.files[self.active_file]["cycle_reverse"]   = self.cycle_reverse_var.get()
            self.files[self.active_file]["lightness_step"]  = self.lightness_step_var.get()
            self.files[self.active_file]["linewidth"]       = self.linewidth_var.get()
            self.files[self.active_file]["plot_style"]      = self.plot_style_var.get()
        super()._save_active_state()

    def _switch_active_file(self, short):
        """Extend base switch to restore per-file zoom/pan state after replot."""
        # Restore per-file UI vars BEFORE super() calls _auto_replot → _plot →
        # _save_active_state.  At that point self.active_file is already 'short',
        # so the UI must already reflect this file's settings or _save_active_state
        # will overwrite the new file's entry with the old file's values.
        entry = self.files.get(short, {})
        color = entry.get("color", "#1f77b4")
        name  = next((n for n, h in _COLOR_HEX.items() if h == color), "Blue")
        self.file_color_var.set(name)
        self.cycle_gradient_var.set(entry.get("cycle_gradient", True))
        self.cycle_reverse_var.set(entry.get("cycle_reverse", False))
        self.lightness_step_var.set(entry.get("lightness_step", "0.15"))
        self.linewidth_var.set(entry.get("linewidth", "3"))
        self.plot_style_var.set(entry.get("plot_style", "Line"))

        super()._switch_active_file(short)

        # Restore zoom/pan view after the replot
        entry = self.files.get(short)
        if entry and "view_xlim" in entry:
            self.ax.set_xlim(entry["view_xlim"])
            self.ax.set_ylim(entry["view_ylim"])
            self.canvas.draw_idle()


# ── Main application window ──────────────────────────────────────────
class EchemGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Electrochemistry Analysis")
        self.geometry("1100x750")
        self._build_ui()

    def _build_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        gen_tab = ttk.Frame(notebook)
        notebook.add(gen_tab, text="General E.Chem")
        EchemPanel(gen_tab, show_ecsa=False, show_log=True).pack(fill=tk.BOTH, expand=True)

        multi_tab = ttk.Frame(notebook)
        notebook.add(multi_tab, text="Multi E.Chem")
        MultiEchemPanel(multi_tab).pack(fill=tk.BOTH, expand=True)

        multi2_tab = ttk.Frame(notebook)
        notebook.add(multi2_tab, text="Multi E.Chem 2")
        MultiEchem2Panel(multi2_tab).pack(fill=tk.BOTH, expand=True)

        ecsa_tab = ttk.Frame(notebook)
        notebook.add(ecsa_tab, text="ECSA Calc")
        ECSAPanel(ecsa_tab).pack(fill=tk.BOTH, expand=True)

        eis_tab = ttk.Frame(notebook)
        notebook.add(eis_tab, text="Nyquist Plot")
        EISPanel(eis_tab).pack(fill=tk.BOTH, expand=True)
