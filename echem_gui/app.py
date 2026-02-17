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


class EchemGUI(
    FileManagerMixin,
    CorrectionMixin,
    PlottingMixin,
    ECSAMixin,
    ExportMixin,
    tk.Tk,
):
    def __init__(self):
        super().__init__()
        self.title("Electrochemistry Analysis")
        self.geometry("1100x750")
        self.files = OrderedDict()
        self.active_file = None
        self._suppress_replot = False
        self._build_ui()

    # ── UI construction ─────────────────────────────────────────────
    def _build_ui(self):
        body = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        left = ttk.Frame(body, width=280)
        body.add(left, weight=0)

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

        # ── Axis selectors ──────────────────────────────────────────
        ttk.Label(left, text="X-axis:").pack(anchor=tk.W, padx=4, pady=(6, 0))
        self.x_var = tk.StringVar()
        self.x_combo = ttk.Combobox(left, textvariable=self.x_var, state="readonly", width=24)
        self.x_combo.pack(padx=4, pady=2)

        ttk.Label(left, text="Y-axis:").pack(anchor=tk.W, padx=4, pady=(6, 0))
        self.y_var = tk.StringVar()
        self.y_combo = ttk.Combobox(left, textvariable=self.y_var, state="readonly", width=24)
        self.y_combo.pack(padx=4, pady=2)

        # ── Plot range ──────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Plot Range", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        xrange_frame = ttk.Frame(left)
        xrange_frame.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(xrange_frame, text="X min:").pack(side=tk.LEFT)
        self.x_min_var = tk.StringVar()
        ttk.Entry(xrange_frame, textvariable=self.x_min_var, width=8).pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(xrange_frame, text="X max:").pack(side=tk.LEFT)
        self.x_max_var = tk.StringVar()
        ttk.Entry(xrange_frame, textvariable=self.x_max_var, width=8).pack(side=tk.LEFT, padx=2)

        yrange_frame = ttk.Frame(left)
        yrange_frame.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(yrange_frame, text="Y min:").pack(side=tk.LEFT)
        self.y_min_var = tk.StringVar()
        ttk.Entry(yrange_frame, textvariable=self.y_min_var, width=8).pack(side=tk.LEFT, padx=(2, 6))
        ttk.Label(yrange_frame, text="Y max:").pack(side=tk.LEFT)
        self.y_max_var = tk.StringVar()
        ttk.Entry(yrange_frame, textvariable=self.y_max_var, width=8).pack(side=tk.LEFT, padx=2)

        ttk.Label(left, text="(leave blank for auto)", foreground="gray").pack(anchor=tk.W, padx=4)

        # ── Legend ──────────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Legend", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

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

        leg_hint = ttk.Label(left, text="(left-drag legend to move, right-drag to resize)", foreground="gray")
        leg_hint.pack(anchor=tk.W, padx=4)

        # ── IR / RHE correction ─────────────────────────────────────
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
        ttk.Button(corr_btn_row, text="Apply Correction", command=self._apply_correction).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(corr_btn_row, text="Reset", command=self._reset_correction).pack(side=tk.LEFT)

        # ── Cycle selector (checkboxes) ────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Cycles:").pack(anchor=tk.W, padx=4)
        btn_row = ttk.Frame(left)
        btn_row.pack(fill=tk.X, padx=4)
        ttk.Button(btn_row, text="Select All", command=self._select_all).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(btn_row, text="Deselect All", command=self._deselect_all).pack(side=tk.LEFT)

        # Scrollable frame for cycle checkboxes
        cycle_outer = ttk.Frame(left)
        cycle_outer.pack(fill=tk.BOTH, expand=True, padx=4, pady=2)
        cycle_canvas = tk.Canvas(cycle_outer, highlightthickness=0)
        cyc_scroll = ttk.Scrollbar(cycle_outer, orient=tk.VERTICAL, command=cycle_canvas.yview)
        cycle_canvas.configure(yscrollcommand=cyc_scroll.set)
        cyc_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        cycle_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._cycle_inner = ttk.Frame(cycle_canvas)
        self._cycle_canvas_win = cycle_canvas.create_window((0, 0), window=self._cycle_inner, anchor=tk.NW)
        self._cycle_canvas = cycle_canvas

        # Keep the inner frame width in sync and update scrollregion
        def _on_inner_configure(e):
            cycle_canvas.configure(scrollregion=cycle_canvas.bbox("all"))
        self._cycle_inner.bind("<Configure>", _on_inner_configure)
        def _on_canvas_configure(e):
            cycle_canvas.itemconfig(self._cycle_canvas_win, width=e.width)
        cycle_canvas.bind("<Configure>", _on_canvas_configure)

        # Mouse wheel scroll inside the checkbox area
        def _on_cycle_mousewheel(e):
            cycle_canvas.yview_scroll(-1 * (e.delta // 120), "units")
        cycle_canvas.bind("<MouseWheel>", _on_cycle_mousewheel)
        self._cycle_inner.bind("<MouseWheel>", _on_cycle_mousewheel)

        # cycle_num → BooleanVar; populated by _populate_cycle_checkboxes
        self._cycle_vars = {}  # {int: tk.BooleanVar}

        plot_btn_row = ttk.Frame(left)
        plot_btn_row.pack(fill=tk.X, padx=4, pady=6)
        ttk.Button(plot_btn_row, text="Plot", command=self._plot).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(plot_btn_row, text="Export Excel", command=self._export_excel).pack(side=tk.LEFT)

        # ── ECSA section ────────────────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="ECSA (prototype)", font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        sr_frame = ttk.Frame(left)
        sr_frame.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(sr_frame, text="Scan rate (mV/s):").pack(side=tk.LEFT)
        self.scan_rate_var = tk.StringVar(value="50")
        ttk.Entry(sr_frame, textvariable=self.scan_rate_var, width=8).pack(side=tk.LEFT, padx=4)

        ttk.Button(left, text="Calculate ECSA", command=self._calc_ecsa).pack(padx=4, pady=4)
        self.ecsa_label = ttk.Label(left, text="", wraplength=230)
        self.ecsa_label.pack(anchor=tk.W, padx=4, pady=2)

        # ── Right panel – matplotlib ────────────────────────────────
        right = ttk.Frame(body)
        body.add(right, weight=1)

        self.fig = Figure(figsize=(6, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=right)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # Toolbar at the bottom of the plot panel
        toolbar_frame = ttk.Frame(right)
        toolbar_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # Custom toolbar: Home button calls our _reset_view
        class _Toolbar(NavigationToolbar2Tk):
            def home(tb_self, *args):
                self._reset_view()
        toolbar = _Toolbar(self.canvas, toolbar_frame)
        toolbar.update()

        # Set up all interactive mouse events (zoom, pan, legend drag)
        self._init_plot_interactions()

    # ── Cycle helpers ───────────────────────────────────────────────
    def _populate_cycle_checkboxes(self, cycles, selected):
        """Rebuild the checkbox list for the given cycles.

        Args:
            cycles:   sorted list of cycle numbers to show
            selected: list of cycle numbers that should be checked
        """
        # Clear existing checkboxes
        for w in self._cycle_inner.winfo_children():
            w.destroy()
        self._cycle_vars.clear()

        selected_set = set(selected)
        for c in cycles:
            var = tk.BooleanVar(value=(c in selected_set))
            var.trace_add("write", self._on_cycle_toggle)
            cb = ttk.Checkbutton(self._cycle_inner, text=f"Cycle {c}", variable=var)
            cb.pack(anchor=tk.W)
            # Forward mouse wheel from checkbuttons to the canvas
            cb.bind("<MouseWheel>", lambda e: self._cycle_canvas.yview_scroll(
                -1 * (e.delta // 120), "units"))
            self._cycle_vars[c] = var

    def _on_cycle_toggle(self, *_args):
        """Called when any cycle checkbox is toggled."""
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
        """Open a dialog to manually edit legend labels."""
        if self._legend_obj is None:
            from tkinter import messagebox
            messagebox.showinfo("Info", "Plot data first to create a legend.")
            return
        open_legend_editor(self, self._legend_obj, self.canvas, self._current_legend_size)
