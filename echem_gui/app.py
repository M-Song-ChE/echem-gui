"""Main application window – assembles all mixins and builds the UI."""

from collections import OrderedDict

import tkinter as tk
from tkinter import ttk
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from .legend_editor import open_legend_editor
from .file_manager import FileManagerMixin
from .correction import CorrectionMixin
from .plotting import PlottingMixin
from .ecsa import ECSAMixin
from .export import ExportMixin
from .ecsa_panel import ECSAPanel

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
        self.file_listbox = tk.Listbox(file_list_frame, height=5, selectmode=tk.BROWSE, exportselection=False)
        fl_scroll = ttk.Scrollbar(file_list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=fl_scroll.set)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        fl_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.bind("<<ListboxSelect>>", self._on_file_select)

        # ── Axis selectors + unit conversion ────────────────────────
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
        _ALL_UNITS = ["(auto)",
                      "A", "mA", "µA", "nA",
                      "V", "mV", "µV", "nV",
                      "s", "ms", "µs", "min", "h"]

        def _refresh_unit_opts(col_var, unit_var, unit_cb):
            col = col_var.get()
            raw_unit = col.rsplit("/", 1)[-1].strip() if "/" in col else ""
            dim  = _UNIT_DIMS.get(raw_unit)
            opts = _DIM_OPTS.get(dim, _ALL_UNITS)
            unit_cb["values"] = opts
            if unit_var.get() not in opts:
                unit_var.set("(auto)")
            self._auto_replot()

        def _refresh_unit_after_select(unit_var, unit_cb):
            chosen = unit_var.get()
            if chosen and chosen != "(auto)":
                dim  = _UNIT_DIMS.get(chosen)
                opts = _DIM_OPTS.get(dim, _ALL_UNITS)
                unit_cb["values"] = opts
            self._auto_replot()

        # X-axis: defaults to voltage (V)
        ttk.Label(left, text="X-axis:").pack(anchor=tk.W, padx=4, pady=(6, 0))
        x_axis_row = ttk.Frame(left)
        x_axis_row.pack(fill=tk.X, padx=4, pady=2)
        self.x_var = tk.StringVar()
        self.x_combo = ttk.Combobox(x_axis_row, textvariable=self.x_var, state="readonly", width=16)
        self.x_combo.pack(side=tk.LEFT)
        self.x_unit_var = tk.StringVar(value="V")
        x_unit_cb = ttk.Combobox(x_axis_row, textvariable=self.x_unit_var,
                                  values=_DIM_OPTS["E"], state="readonly", width=6)
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
                                  values=_DIM_OPTS["I"], state="readonly", width=6)
        y_unit_cb.pack(side=tk.LEFT, padx=(4, 0))
        self.y_combo.bind("<<ComboboxSelected>>",
                          lambda e: _refresh_unit_opts(self.y_var, self.y_unit_var, y_unit_cb))
        y_unit_cb.bind("<<ComboboxSelected>>",
                       lambda e: _refresh_unit_after_select(self.y_unit_var, y_unit_cb))

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
        _xmin = ttk.Entry(xrange_frame, textvariable=self.x_min_var, width=8)
        _xmin.pack(side=tk.LEFT, padx=(2, 6))
        _bind_range(_xmin)
        ttk.Label(xrange_frame, text="X max:").pack(side=tk.LEFT)
        self.x_max_var = tk.StringVar()
        _xmax = ttk.Entry(xrange_frame, textvariable=self.x_max_var, width=8)
        _xmax.pack(side=tk.LEFT, padx=2)
        _bind_range(_xmax)

        yrange_frame = ttk.Frame(left)
        yrange_frame.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(yrange_frame, text="Y min:").pack(side=tk.LEFT)
        self.y_min_var = tk.StringVar()
        _ymin = ttk.Entry(yrange_frame, textvariable=self.y_min_var, width=8)
        _ymin.pack(side=tk.LEFT, padx=(2, 6))
        _bind_range(_ymin)
        ttk.Label(yrange_frame, text="Y max:").pack(side=tk.LEFT)
        self.y_max_var = tk.StringVar()
        _ymax = ttk.Entry(yrange_frame, textvariable=self.y_max_var, width=8)
        _ymax.pack(side=tk.LEFT, padx=2)
        _bind_range(_ymax)

        ttk.Label(left, text="(leave blank for auto)", foreground="gray").pack(anchor=tk.W, padx=4)

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
        self.legend_size_var = tk.StringVar(value="8")
        ttk.Entry(leg_row, textvariable=self.legend_size_var, width=5).pack(side=tk.LEFT, padx=(2, 8))
        ttk.Label(leg_row, text="Location:").pack(side=tk.LEFT)
        self.legend_loc_var = tk.StringVar(value="best")
        ttk.Combobox(
            leg_row, textvariable=self.legend_loc_var,
            values=[
                "best", "upper right", "upper left", "lower left", "lower right",
                "right", "center left", "center right", "lower center",
                "upper center", "center",
            ],
            state="readonly", width=12,
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(left, text="Edit Labels", command=self._edit_legend_labels).pack(anchor=tk.W, padx=4, pady=2)
        ttk.Label(left, text="(left-drag legend to move, right-drag to resize)",
                  foreground="gray").pack(anchor=tk.W, padx=4)

        # IR / RHE correction
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="IR / RHE Correction", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        r_frame = ttk.Frame(left)
        r_frame.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(r_frame, text="R_sol (Ohm):").pack(side=tk.LEFT)
        self.r_sol_var = tk.StringVar(value="0")
        ttk.Entry(r_frame, textvariable=self.r_sol_var, width=10).pack(side=tk.LEFT, padx=4)

        e_frame = ttk.Frame(left)
        e_frame.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(e_frame, text="E_ref (V vs RHE):").pack(side=tk.LEFT)
        self.e_ref_var = tk.StringVar(value="0")
        ttk.Entry(e_frame, textvariable=self.e_ref_var, width=10).pack(side=tk.LEFT, padx=4)

        corr_btn_row = ttk.Frame(left)
        corr_btn_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Button(corr_btn_row, text="Apply Correction",
                   command=self._apply_correction).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(corr_btn_row, text="Reset", command=self._reset_correction).pack(side=tk.LEFT)

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
        toolbar = _Toolbar(self.canvas, toolbar_frame)
        toolbar.update()

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
        open_legend_editor(self, self._legend_obj, self.canvas, self._current_legend_size)


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

        ecsa_tab = ttk.Frame(notebook)
        notebook.add(ecsa_tab, text="ECSA Calc")
        ECSAPanel(ecsa_tab).pack(fill=tk.BOTH, expand=True)
