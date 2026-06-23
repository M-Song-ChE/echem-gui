"""OCV/Ru Extractor Panel.

Replaces the legacy Nyquist Plot tab. Files are grouped into samples by shared
filename substring; each sample bundle holds 1-2 OCV files and any number of
EIS files. The panel auto-extracts:
  • OCV: last voltage value from the time-series
  • Ru:  Re(Z) at the row where |Im(Z)| is minimum (restricted to Re(Z) > 0)

UI:
  Left  — Load button, sample table (Name | OCV | EIS1 Ru | EIS2 Ru | …),
          per-sample color/show toggle, file breakdown listbox
  Right — OCV (time vs E) | Nyquist (Re vs -Im) side-by-side
"""

import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from collections import OrderedDict
from difflib import SequenceMatcher

import numpy as np
import pandas as pd

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from .file_manager import _read_mpr, _COLOR_NAMES, _COLOR_HEX
from .plotting import copy_figure_to_clipboard
from .checklist import CheckableListbox
from . import session_manager as _sm


# ── Default rotating palette for new samples ─────────────────────────────────
_DEFAULT_COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
]

# Markers used when a sample has multiple EIS files (cycled per-file)
_EIS_MARKERS = ["o", "s", "^", "D", "v", "P", "X", "*"]


# ════════════════════════════════════════════════════════════════════════════
# File classification / sample-name derivation helpers
# ════════════════════════════════════════════════════════════════════════════
def _classify_file(filename: str) -> str:
    """Return 'ocv', 'eis', or 'unknown' based on filename keywords."""
    base = os.path.splitext(os.path.basename(filename))[0].lower()
    # EIS variants: PEIS, GEIS, SEIS, EIS, optionally followed by digits
    if re.search(r"(?<![a-z])(p?eis|geis|seis)\d*(?![a-z])", base):
        return "eis"
    if re.search(r"(?<![a-z])ocv\d*(?![a-z])", base):
        return "ocv"
    return "unknown"


def _longest_common_substring(strings: list[str]) -> str:
    """Iteratively narrow the common substring across all *strings*."""
    if not strings:
        return ""
    common = strings[0]
    for s in strings[1:]:
        sm = SequenceMatcher(None, common, s, autojunk=False)
        m = sm.find_longest_match(0, len(common), 0, len(s))
        common = common[m.a:m.a + m.size]
        if not common:
            break
    return common


_NAME_CUT_RE = re.compile(
    r"\s+vs\s+|_RE\d|_CE\d|_\d+_PEIS|_\d+_GEIS|_\d+_SEIS|_C\d+$|\s+\d+_PEIS",
    re.IGNORECASE,
)


def _name_from_one(base: str) -> str | None:
    """Extract the sample name from one Bio-Logic-style filename basename.

    Pattern:  ``Pn_TYPE_<name>_<electrode/electrolyte/...>``
    Returns the ``<name>`` portion, trimmed at electrode/electrolyte markers,
    or None if the filename doesn't match the Pn_TYPE_ pattern.
    """
    m = re.match(r"^[Pp]\d+_[A-Za-z]+\d*_(.+)$", base)
    if not m:
        return None
    rest = m.group(1)
    cut = _NAME_CUT_RE.split(rest, maxsplit=1)[0]
    return cut.rstrip("_- ").strip() or None


def _derive_sample_name(filenames: list[str]) -> str:
    """Choose a sample name from a list of file paths.

    Prefers the ``Pn_TYPE_<name>_…`` Bio-Logic convention; falls back to longest
    common substring of basenames if the pattern isn't matched.
    """
    if not filenames:
        return "Sample"
    bases = [os.path.splitext(os.path.basename(f))[0] for f in filenames]

    # Try Pn_TYPE_<name>_… extraction per file; all should agree.
    candidates = [n for n in (_name_from_one(b) for b in bases) if n]
    if candidates:
        unique = list(dict.fromkeys(candidates))
        if len(unique) == 1:
            return unique[0]
        # Multiple distinct candidates → take their longest common substring
        common = _longest_common_substring(unique).strip("_- ")
        return common or unique[0]

    # Fallback: LCS of full basenames with OCV/EIS markers stripped
    common = bases[0] if len(bases) == 1 else _longest_common_substring(bases)
    if not common.strip():
        common = bases[0]
    cleaned = re.sub(r"(?i)(p?eis|geis|seis)\d*", "", common)
    cleaned = re.sub(r"(?i)ocv\d*", "", cleaned)
    cleaned = re.sub(r"^[Pp]\d+[_\s\-.]*", "", cleaned)
    cleaned = re.sub(r"^[_\W\s]+|[_\W\s]+$", "", cleaned)
    cleaned = re.sub(r"[_\s]{2,}", "_", cleaned)
    return cleaned.strip() or bases[0]


# ════════════════════════════════════════════════════════════════════════════
# Column discovery + value extraction
# ════════════════════════════════════════════════════════════════════════════
def _read_file_df(path: str) -> pd.DataFrame:
    """Read .mpr (galvani) or .txt (tab-separated) into a DataFrame."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".mpr":
        return _read_mpr(path)
    df = pd.read_csv(path, sep="\t")
    df.columns = [c.strip() for c in df.columns]
    return df


def _find_voltage_col(df: pd.DataFrame) -> str | None:
    for c in df.columns:
        cl = c.lower()
        if "ewe" in cl or cl.startswith("e/v") or "potential" in cl or "voltage" in cl:
            return c
    for c in df.columns:
        cl = c.lower()
        if cl.endswith("/v") and not cl.endswith(("mv", "uv", "nv", "µv")):
            return c
    return None


def _find_time_col(df: pd.DataFrame) -> str | None:
    for c in df.columns:
        if c.lower().startswith("time"):
            return c
    return None


def _find_re_z_col(df: pd.DataFrame) -> str | None:
    for c in df.columns:
        if "re(z)" in c.lower():
            return c
    return None


def _find_im_z_col(df: pd.DataFrame) -> str | None:
    """Prefer -Im(Z); fall back to Im(Z)."""
    for c in df.columns:
        if "-im(z)" in c.lower():
            return c
    for c in df.columns:
        if "im(z)" in c.lower():
            return c
    return None


def _extract_ocv_value(df: pd.DataFrame) -> float | None:
    """Return the last voltage value (the stable OCV)."""
    vcol = _find_voltage_col(df)
    if vcol is None or df.empty:
        return None
    try:
        return float(df[vcol].dropna().iloc[-1])
    except Exception:
        return None


def _extract_ru_value(df: pd.DataFrame) -> tuple[float | None, float | None]:
    """Find Ru = Re(Z) where |Im(Z)| is minimum (Re(Z) > 0).

    Returns (Ru, y_at_Ru) where y_at_Ru is the corresponding -Im(Z) value
    (used to annotate the Nyquist plot).
    """
    rec = _find_re_z_col(df)
    imc = _find_im_z_col(df)
    if rec is None or imc is None or df.empty:
        return None, None
    sub = df[[rec, imc]].dropna()
    sub = sub[sub[rec] > 0]
    if sub.empty:
        return None, None
    idx = sub[imc].abs().idxmin()
    try:
        ru = float(sub.loc[idx, rec])
        y_raw = float(sub.loc[idx, imc])
        # Convert to -Im(Z) for plotting
        y_neg_im = y_raw if "-im" in imc.lower() else -y_raw
        return ru, y_neg_im
    except Exception:
        return None, None


# ════════════════════════════════════════════════════════════════════════════
# Panel
# ════════════════════════════════════════════════════════════════════════════
class OcvRuPanel(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)

        self.samples: "OrderedDict[str, dict]" = OrderedDict()
        self.active_sample: str | None = None
        self._suppress_replot = False
        self._next_color_idx = 0

        # Plot size — width and height are PER-PLOT dimensions (each sample has
        # an OCV plot + a Nyquist plot side-by-side at this size each).
        self.plot_w_var = tk.StringVar(value="7.0")
        self.plot_h_var = tk.StringVar(value="4.5")

        # Per-sample frame registry (sample_name → dict with frame+axes+interactions)
        self._sample_frames: dict = {}
        # Drag-to-reorder state for plot-frame headers (right panel)
        self._frame_drag: dict | None = None

        self._build_panel()
        self.after(500, self._apply_plot_size)

    # ─────────────────────────────────────────────────────────────────
    # UI construction
    # ─────────────────────────────────────────────────────────────────
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
        _lc.bind("<MouseWheel>",
                 lambda e: _lc.yview_scroll(-1 * (e.delta // 120), "units"))

        # ── Sample buttons ───────────────────────────────────────
        ttk.Label(left, text="Samples", font=("", 9, "bold")).pack(
            anchor=tk.W, padx=4, pady=(6, 0))

        bb1 = ttk.Frame(left); bb1.pack(fill=tk.X, padx=4, pady=2)
        ttk.Button(bb1, text="Load New Sample…",
                   command=self._load_new_sample).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(bb1, text="Add to Selected…",
                   command=self._add_to_active_sample).pack(side=tk.LEFT)

        bb2 = ttk.Frame(left); bb2.pack(fill=tk.X, padx=4, pady=2)
        ttk.Button(bb2, text="Remove",
                   command=self._remove_sample).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(bb2, text="Rename…",
                   command=self._rename_sample).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(bb2, text="Export CSV",
                   command=self._export_csv).pack(side=tk.LEFT)

        ttk.Label(left,
                  text="(drag ⠿ to reorder · checkbox = show/hide · "
                       "click row to select)",
                  foreground="gray").pack(anchor=tk.W, padx=4, pady=(2, 0))

        # ── Sample list (drag-reorder + show/hide via CheckableListbox) ───
        self.sample_lb = CheckableListbox(left, height=8,
                                          on_check=self._on_sample_check,
                                          on_reorder=self._on_sample_reorder)
        self.sample_lb.pack(fill=tk.X, padx=4, pady=(2, 4))
        self.sample_lb.bind("<<ListboxSelect>>", self._on_listbox_select)

        # ── Extracted values table ────────────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(
            fill=tk.X, padx=4, pady=4)
        ttk.Label(left, text="Extracted values",
                  font=("", 9, "bold")).pack(anchor=tk.W, padx=4)

        vt_frame = ttk.Frame(left)
        vt_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=(2, 4))
        self.values_tv = ttk.Treeview(vt_frame, columns=("ocv",),
                                      show="tree headings", height=8)
        self.values_tv.heading("#0", text="Sample")
        self.values_tv.heading("ocv", text="OCV (V)")
        self.values_tv.column("#0", width=160, anchor=tk.W, stretch=True)
        self.values_tv.column("ocv", width=72, anchor=tk.E, stretch=False)
        self.values_tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vt_vs = ttk.Scrollbar(vt_frame, orient=tk.VERTICAL,
                              command=self.values_tv.yview)
        vt_vs.pack(side=tk.RIGHT, fill=tk.Y)
        self.values_tv.configure(yscrollcommand=vt_vs.set)
        self.values_tv.bind("<<TreeviewSelect>>", self._on_values_table_select)

        # ── Selected-sample detail panel ──────────────────────────
        ttk.Separator(left, orient=tk.HORIZONTAL).pack(
            fill=tk.X, padx=4, pady=6)
        ttk.Label(left, text="Selected sample", font=("", 9, "bold")).pack(
            anchor=tk.W, padx=4)

        cr = ttk.Frame(left); cr.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(cr, text="Color:").pack(side=tk.LEFT)
        self.color_var = tk.StringVar(value="Blue")
        _cb = ttk.Combobox(cr, textvariable=self.color_var,
                           values=_COLOR_NAMES, state="readonly", width=14)
        _cb.pack(side=tk.LEFT, padx=(4, 0))
        _cb.bind("<<ComboboxSelected>>", self._on_color_change)

        ttk.Label(left, text="Files:").pack(anchor=tk.W, padx=4, pady=(4, 0))
        self.files_lb = tk.Listbox(left, height=6)
        self.files_lb.pack(fill=tk.X, padx=4, pady=(0, 6))

        # ── Right panel ───────────────────────────────────────────
        right = ttk.Frame(body)
        body.add(right, weight=1)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        _size_bar = ttk.Frame(right)
        _size_bar.grid(row=0, column=0, sticky="ew", padx=4, pady=2)
        ttk.Label(_size_bar, text="Plot size (in):").pack(
            side=tk.LEFT, padx=(4, 2))
        ttk.Label(_size_bar, text="W").pack(side=tk.LEFT)
        _pw_e = ttk.Entry(_size_bar, textvariable=self.plot_w_var, width=5)
        _pw_e.pack(side=tk.LEFT, padx=(1, 6))
        ttk.Label(_size_bar, text="H").pack(side=tk.LEFT)
        _ph_e = ttk.Entry(_size_bar, textvariable=self.plot_h_var, width=5)
        _ph_e.pack(side=tk.LEFT, padx=(1, 0))
        for _e in (_pw_e, _ph_e):
            _e.bind("<Return>",   lambda ev: self._apply_plot_size())
            _e.bind("<FocusOut>", lambda ev: self._apply_plot_size())

        _right_inner = ttk.Frame(right)
        _right_inner.grid(row=1, column=0, sticky="nsew")
        _right_inner.rowconfigure(0, weight=1)
        _right_inner.columnconfigure(0, weight=1)
        _plot_sc = tk.Canvas(_right_inner, highlightthickness=0)
        _vs = ttk.Scrollbar(_right_inner, orient=tk.VERTICAL,
                            command=_plot_sc.yview)
        _hs = ttk.Scrollbar(_right_inner, orient=tk.HORIZONTAL,
                            command=_plot_sc.xview)
        _plot_sc.configure(yscrollcommand=_vs.set, xscrollcommand=_hs.set)
        _vs.grid(row=0, column=1, sticky="ns")
        _hs.grid(row=1, column=0, sticky="ew")
        _plot_sc.grid(row=0, column=0, sticky="nsew")
        _plot_sc.bind("<MouseWheel>",
                      lambda e: _plot_sc.yview_scroll(-1*(e.delta//120), "units"))
        _plot_sc.bind("<Shift-MouseWheel>",
                      lambda e: _plot_sc.xview_scroll(-1*(e.delta//120), "units"))
        _plots_frame = ttk.Frame(_plot_sc)
        _plot_sc.create_window((0, 0), window=_plots_frame, anchor=tk.NW)
        _plots_frame.bind("<Configure>",
                          lambda e: _plot_sc.configure(scrollregion=_plot_sc.bbox("all")))
        self._plot_sc = _plot_sc
        self._plots_frame = _plots_frame

        # Empty-state placeholder shown while no samples are loaded
        self._empty_label = ttk.Label(
            _plots_frame,
            text="(Load a sample to see its OCV + Nyquist plots here)",
            foreground="gray")
        self._empty_label.pack(padx=20, pady=20)

        # Drop indicator line shown while a header is being dragged
        self._drop_line = tk.Frame(_plots_frame, bg="#1a73e8", height=3)

        self._plot()

    # ─────────────────────────────────────────────────────────────────
    # Plot size (applied uniformly to every sample's OCV+Nyquist figures)
    # ─────────────────────────────────────────────────────────────────
    def _apply_plot_size(self, event=None):
        try:
            w = float(self.plot_w_var.get())
            h = float(self.plot_h_var.get())
        except ValueError:
            return
        w = max(2.0, min(50.0, w))
        h = max(2.0, min(50.0, h))
        for sf in self._sample_frames.values():
            for key in ("ocv", "eis"):
                ax_state = sf.get(key)
                if ax_state is None:
                    continue
                ax_state["fig"].set_size_inches(w, h)
                ax_state["canvas"].get_tk_widget().config(
                    width=int(w * 100), height=int(h * 100))
                try:
                    ax_state["fig"].tight_layout()
                except Exception:
                    pass
                ax_state["canvas"].draw_idle()
        self._plot_sc.after(50, lambda: self._plot_sc.configure(
            scrollregion=self._plot_sc.bbox("all")))

    # ─────────────────────────────────────────────────────────────────
    # Sample loading
    # ─────────────────────────────────────────────────────────────────
    def _load_new_sample(self):
        paths = filedialog.askopenfilenames(
            title="Select all OCV/EIS files for ONE sample",
            filetypes=[("Data files", "*.mpr *.txt"),
                       ("MPR files", "*.mpr"),
                       ("Text files", "*.txt"),
                       ("All files", "*.*")],
        )
        if not paths:
            return
        self._add_sample_from_paths(list(paths))

    def _add_to_active_sample(self):
        if not self.active_sample:
            messagebox.showinfo("No sample selected",
                                "Select a sample first, or use 'Load New Sample…'.")
            return
        paths = filedialog.askopenfilenames(
            title=f"Add OCV/EIS files to '{self.active_sample}'",
            filetypes=[("Data files", "*.mpr *.txt"),
                       ("MPR files", "*.mpr"),
                       ("Text files", "*.txt"),
                       ("All files", "*.*")],
        )
        if not paths:
            return
        ocv_new, eis_new, skipped = self._build_file_entries(list(paths))
        if not ocv_new and not eis_new:
            messagebox.showwarning("Nothing added",
                                   "No OCV or EIS files were recognised in the selection.")
            return
        s = self.samples[self.active_sample]
        s["ocv_files"].extend(ocv_new)
        s["eis_files"].extend(eis_new)
        if skipped:
            messagebox.showinfo("Skipped",
                                "Skipped (not OCV/EIS):\n" + "\n".join(skipped))
        self._refresh_sample_list()
        self._select_sample(self.active_sample)
        self._plot()

    def _build_file_entries(self, paths):
        """Read files and classify them; return (ocv_list, eis_list, skipped_names)."""
        ocv_files, eis_files, skipped = [], [], []
        for p in paths:
            kind = _classify_file(p)
            try:
                df = _read_file_df(p)
            except Exception as exc:
                messagebox.showerror(
                    "Read error",
                    f"Failed to read {os.path.basename(p)}:\n{exc}")
                continue
            entry = {"path": p, "filename": os.path.basename(p), "df": df}
            if kind == "ocv":
                entry["ocv_value"] = _extract_ocv_value(df)
                ocv_files.append(entry)
            elif kind == "eis":
                ru, y = _extract_ru_value(df)
                entry["ru_value"] = ru
                entry["ru_y"] = y
                eis_files.append(entry)
            else:
                skipped.append(entry["filename"])
        return ocv_files, eis_files, skipped

    def _add_sample_from_paths(self, paths):
        ocv_files, eis_files, skipped = self._build_file_entries(paths)
        if not ocv_files and not eis_files:
            messagebox.showwarning("No data",
                                   "No OCV or EIS files were recognised.")
            return
        if skipped:
            messagebox.showinfo("Skipped",
                                "Skipped (not OCV/EIS):\n" + "\n".join(skipped))

        all_paths = [f["path"] for f in (*ocv_files, *eis_files)]
        name = _derive_sample_name(all_paths)
        # Ensure unique
        original = name
        i = 2
        while name in self.samples:
            name = f"{original} ({i})"
            i += 1

        color = _DEFAULT_COLORS[self._next_color_idx % len(_DEFAULT_COLORS)]
        self._next_color_idx += 1

        self.samples[name] = {
            "name": name,
            "ocv_files": ocv_files,
            "eis_files": eis_files,
            "color": color,
            "hidden": False,
        }
        self._refresh_sample_list()
        self._select_sample(name)
        self._plot()

    # ─────────────────────────────────────────────────────────────────
    # Sample list (CheckableListbox) + values table (Treeview)
    # ─────────────────────────────────────────────────────────────────
    def _refresh_sample_list(self):
        """Rebuild the sample listbox + values table to match self.samples."""
        prev_active = self.active_sample
        # Listbox: sample name as row text (clean, stable identifier)
        self.sample_lb.clear()
        for name in self.samples:
            self.sample_lb.insert(
                tk.END, name,
                checked=not self.samples[name].get("hidden", False))
        if prev_active and prev_active in self.samples:
            self._set_lb_selection(prev_active)
        self._refresh_values_table()

    def _refresh_values_table(self):
        """Rebuild the read-only Treeview that shows OCV / Ru columns per sample."""
        max_eis = max((len(s["eis_files"]) for s in self.samples.values()),
                      default=0)
        cols = ["ocv"] + [f"ru{i+1}" for i in range(max_eis)]
        self.values_tv.delete(*self.values_tv.get_children())
        self.values_tv["columns"] = tuple(cols)
        self.values_tv.heading("#0", text="Sample")
        self.values_tv.column("#0", width=160, anchor=tk.W, stretch=True)
        self.values_tv.heading("ocv", text="OCV (V)")
        self.values_tv.column("ocv", width=72, anchor=tk.E, stretch=False)
        for i in range(max_eis):
            cid = f"ru{i+1}"
            self.values_tv.heading(cid, text=f"Ru{i+1} (Ω)")
            self.values_tv.column(cid, width=78, anchor=tk.E, stretch=False)

        for name, s in self.samples.items():
            ocv_vals = [f.get("ocv_value") for f in s["ocv_files"]
                        if f.get("ocv_value") is not None]
            if not ocv_vals:
                ocv_cell = ""
            elif len(ocv_vals) == 1:
                ocv_cell = f"{ocv_vals[0]:.4f}"
            else:
                ocv_cell = " / ".join(f"{v:.4f}" for v in ocv_vals)
            row = [ocv_cell]
            for i in range(max_eis):
                if i < len(s["eis_files"]):
                    rv = s["eis_files"][i].get("ru_value")
                    row.append(f"{rv:.3f}" if rv is not None else "—")
                else:
                    row.append("")
            self.values_tv.insert("", tk.END, iid=name, text=name,
                                  values=tuple(row))
        if self.active_sample and self.active_sample in self.samples:
            try:
                self.values_tv.selection_set(self.active_sample)
            except Exception:
                pass

    def _set_lb_selection(self, name: str):
        """Set the listbox selection to row matching *name* (no event)."""
        try:
            idx = next(i for i in range(self.sample_lb.size())
                       if self.sample_lb.get(i) == name)
        except StopIteration:
            return
        self.sample_lb._set_selection(idx, fire_event=False)
        self.sample_lb.see(idx)

    def _on_listbox_select(self, event=None):
        sel = self.sample_lb.curselection()
        if not sel:
            return
        name = self.sample_lb.get(sel[0])
        if name in self.samples:
            self._select_sample(name)

    def _on_values_table_select(self, event=None):
        sel = self.values_tv.selection()
        if sel and sel[0] in self.samples:
            self._select_sample(sel[0])

    def _on_sample_check(self, text, visible):
        """Listbox checkbox toggled → show/hide that sample."""
        if text not in self.samples:
            return
        self.samples[text]["hidden"] = not visible
        self._plot()

    def _on_sample_reorder(self, new_texts):
        """Drag-handle reordered the listbox → reorder samples to match."""
        new_names = [t for t in new_texts if t in self.samples]
        for n in self.samples:
            if n not in new_names:
                new_names.append(n)
        self.samples = OrderedDict((n, self.samples[n]) for n in new_names)
        self._refresh_values_table()
        self._plot()

    def _select_sample(self, name):
        if name not in self.samples:
            return
        prev = self.active_sample
        self.active_sample = name
        s = self.samples[name]
        # Reflect color in combobox (match hex → name; fall back to literal hex)
        match = next((cn for cn, hx in _COLOR_HEX.items() if hx == s["color"]),
                     None)
        if match:
            self.color_var.set(match)

        self.files_lb.delete(0, tk.END)
        for f in s["ocv_files"]:
            v = f.get("ocv_value")
            tag = f"OCV: {f['filename']}"
            if v is not None:
                tag += f"   →  {v:.4f} V"
            self.files_lb.insert(tk.END, tag)
        for i, f in enumerate(s["eis_files"]):
            rv = f.get("ru_value")
            tag = f"EIS{i+1}: {f['filename']}"
            if rv is not None:
                tag += f"   →  Ru = {rv:.3f} Ω"
            self.files_lb.insert(tk.END, tag)

        # Sync listbox + values-table highlight (without firing select events back)
        self._set_lb_selection(name)
        try:
            if name in self.values_tv.get_children() and \
                    self.values_tv.selection() != (name,):
                self.values_tv.selection_set(name)
                self.values_tv.see(name)
        except Exception:
            pass

        # Update active-row border highlight (cheap; avoids full replot)
        if prev and prev in self._sample_frames and prev != name:
            self._refresh_sample_frame_chrome(prev)
        if name in self._sample_frames:
            self._refresh_sample_frame_chrome(name)
            # Scroll the row into view
            try:
                self._plot_sc.update_idletasks()
                row = self._sample_frames[name]["row"]
                y_top = row.winfo_y()
                total = max(1, self._plots_frame.winfo_height())
                self._plot_sc.yview_moveto(max(0.0, y_top / total))
            except Exception:
                pass

    def _on_color_change(self, event=None):
        if not self.active_sample:
            return
        hx = _COLOR_HEX.get(self.color_var.get(), "#1f77b4")
        self.samples[self.active_sample]["color"] = hx
        self._plot()

    def _remove_sample(self):
        if not self.active_sample:
            return
        name = self.active_sample
        if name in self.samples:
            del self.samples[name]
        self.active_sample = None
        self.files_lb.delete(0, tk.END)
        self._refresh_sample_list()
        self._plot()

    def _rename_sample(self):
        if not self.active_sample:
            return
        old = self.active_sample
        new = simpledialog.askstring("Rename sample",
                                     "New sample name:",
                                     initialvalue=old, parent=self)
        if not new or new == old:
            return
        if new in self.samples:
            messagebox.showwarning("Duplicate",
                                   f"Sample '{new}' already exists.")
            return
        # Preserve order
        new_dict: "OrderedDict[str, dict]" = OrderedDict()
        for k, v in self.samples.items():
            if k == old:
                v["name"] = new
                new_dict[new] = v
            else:
                new_dict[k] = v
        self.samples = new_dict
        self.active_sample = new
        self._refresh_sample_list()
        self._select_sample(new)
        self._plot()

    # ─────────────────────────────────────────────────────────────────
    # Export
    # ─────────────────────────────────────────────────────────────────
    def _export_csv(self):
        if not self.samples:
            messagebox.showinfo("Nothing to export", "No samples loaded.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            title="Export extracted OCV / Ru values",
        )
        if not path:
            return
        max_eis = max((len(s["eis_files"]) for s in self.samples.values()),
                      default=0)
        rows = []
        for name, s in self.samples.items():
            ocv_vals = [f.get("ocv_value")
                        for f in s["ocv_files"] if f.get("ocv_value") is not None]
            if ocv_vals:
                ocv_avg = sum(ocv_vals) / len(ocv_vals)
                ocv_cell = f"{ocv_avg:.6f}"
            else:
                ocv_cell = ""
            row = {"Sample": name, "OCV (V)": ocv_cell}
            for i in range(max_eis):
                key = f"EIS{i+1} Ru (Ohm)"
                if i < len(s["eis_files"]):
                    rv = s["eis_files"][i].get("ru_value")
                    row[key] = "" if rv is None else f"{rv:.4f}"
                else:
                    row[key] = ""
            rows.append(row)
        try:
            pd.DataFrame(rows).to_csv(path, index=False)
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))
            return
        messagebox.showinfo("Export complete",
                            f"Saved {len(rows)} samples to:\n{path}")

    # ─────────────────────────────────────────────────────────────────
    # Per-sample plot frames
    # ─────────────────────────────────────────────────────────────────
    _ACTIVE_BORDER  = "#1a73e8"
    _INACTIVE_BORDER = "#cccccc"

    def _create_sample_frame(self, name: str):
        """Build the row UI (header + OCV figure + Nyquist figure) for *name*."""
        try:
            w = float(self.plot_w_var.get())
            h = float(self.plot_h_var.get())
        except ValueError:
            w, h = 7.0, 4.5

        # Outer row frame with coloured border to indicate active sample
        row = tk.Frame(self._plots_frame, bd=2, relief=tk.SOLID,
                       highlightthickness=0,
                       background=self._INACTIVE_BORDER)
        row.pack(fill=tk.X, padx=4, pady=4, anchor="nw")

        # Header strip — click activates sample; drag (vertical) reorders rows.
        hdr = tk.Frame(row, bg="#f0f0f0", cursor="fleur")
        hdr.pack(fill=tk.X)
        handle = tk.Label(hdr, text="⠿", bg="#f0f0f0",
                          font=("", 12), cursor="fleur")
        handle.pack(side=tk.LEFT, padx=(6, 0), pady=4)
        color_box = tk.Frame(hdr, width=14, height=14,
                             background=self.samples[name].get("color", "#1f77b4"),
                             cursor="fleur")
        color_box.pack(side=tk.LEFT, padx=(6, 6), pady=4)
        color_box.pack_propagate(False)
        title_lbl = tk.Label(hdr, text=name, bg="#f0f0f0",
                             font=("", 10, "bold"), anchor="w",
                             cursor="fleur")
        title_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        plots_row = tk.Frame(row, bg="white")
        plots_row.pack(fill=tk.X)

        # Build one matplotlib canvas per plot type
        def _build_axis(parent, kind: str, side: str):
            fig = Figure(figsize=(w, h), dpi=100)
            ax = fig.add_subplot(111)
            inner = tk.Frame(parent, bg="white")
            inner.pack(side=side, padx=2, pady=2)
            canvas = FigureCanvasTkAgg(fig, master=inner)
            canvas.get_tk_widget().pack()
            canvas.get_tk_widget().config(
                width=int(w * 100), height=int(h * 100))
            tb_frame = ttk.Frame(inner)
            tb_frame.pack(fill=tk.X)
            tb = NavigationToolbar2Tk(canvas, tb_frame, pack_toolbar=False)
            tb.pack(side=tk.LEFT, fill=tk.X, expand=True)
            tb.update()
            tk.Button(tb_frame, text="Copy",
                      command=lambda f=fig: copy_figure_to_clipboard(f),
                      relief=tk.RAISED, borderwidth=1, padx=6,
                      ).pack(side=tk.LEFT, padx=(4, 2), pady=1)
            return {
                "fig": fig, "ax": ax, "canvas": canvas, "inner": inner,
                "kind": kind,
                # Interaction state
                "panning": False, "pan_start": None, "pan_moved": False,
                "ann": None, "ann_dot": None,
                "ann_last": None, "ann_idx": 0,
                "auto_xlim": None, "auto_ylim": None,
            }

        ocv_state = _build_axis(plots_row, "ocv", tk.LEFT)
        eis_state = _build_axis(plots_row, "eis", tk.LEFT)

        # Click on the plot area (not the draggable header) → activate sample
        for w_ in (row, plots_row):
            w_.bind("<Button-1>", lambda e, n=name: self._select_sample(n),
                    add="+")
        # Header strip: press/drag/release for reorder; click (no drag) activates
        for w_ in (hdr, handle, color_box, title_lbl):
            w_.bind("<ButtonPress-1>",
                    lambda e, n=name: self._on_header_press(e, n))
            w_.bind("<B1-Motion>",
                    lambda e, n=name: self._on_header_drag(e, n))
            w_.bind("<ButtonRelease-1>",
                    lambda e, n=name: self._on_header_release(e, n))

        # Wire matplotlib interactions per axis
        for st in (ocv_state, eis_state):
            cv = st["canvas"]
            cv.mpl_connect("scroll_event",
                           lambda ev, n=name, k=st["kind"]: self._on_scroll(ev, n, k))
            cv.mpl_connect("button_press_event",
                           lambda ev, n=name, k=st["kind"]: self._on_press(ev, n, k))
            cv.mpl_connect("button_release_event",
                           lambda ev, n=name, k=st["kind"]: self._on_release(ev, n, k))
            cv.mpl_connect("motion_notify_event",
                           lambda ev, n=name, k=st["kind"]: self._on_motion(ev, n, k))

        self._sample_frames[name] = {
            "row": row, "hdr": hdr, "title": title_lbl,
            "color_box": color_box, "handle": handle,
            "ocv": ocv_state, "eis": eis_state,
        }

    def _destroy_sample_frame(self, name: str):
        sf = self._sample_frames.pop(name, None)
        if not sf:
            return
        try:
            sf["row"].destroy()
        except Exception:
            pass

    def _refresh_sample_frame_chrome(self, name: str):
        """Update the row's header text + colour box + border to match sample state."""
        sf = self._sample_frames.get(name)
        s = self.samples.get(name)
        if not sf or not s:
            return
        sf["title"].config(text=name)
        sf["color_box"].config(background=s.get("color", "#1f77b4"))
        is_active = (name == self.active_sample)
        sf["row"].config(background=self._ACTIVE_BORDER if is_active
                          else self._INACTIVE_BORDER)

    # ─────────────────────────────────────────────────────────────────
    # Drag-to-reorder on plot frame headers (pattern from multi_echem2)
    # ─────────────────────────────────────────────────────────────────
    def _on_header_press(self, event, name):
        self._frame_drag = {
            "name": name,
            "start_x": event.x_root, "start_y": event.y_root,
            "active": False, "target": None, "target_top": True,
        }

    def _on_header_drag(self, event, name):
        drag = self._frame_drag
        if drag is None or drag["name"] != name:
            return
        if not drag["active"]:
            if (abs(event.x_root - drag["start_x"])
                    + abs(event.y_root - drag["start_y"])) < 6:
                return
            drag["active"] = True

        # Find the visible sample frame the cursor is over
        target = None
        target_top = True
        for tn, sf in self._sample_frames.items():
            if tn == name:
                continue
            pf = sf.get("row")
            if pf is None:
                continue
            x0 = pf.winfo_rootx()
            y0 = pf.winfo_rooty()
            w = pf.winfo_width()
            h = pf.winfo_height()
            if x0 <= event.x_root <= x0 + w and y0 <= event.y_root <= y0 + h:
                target = tn
                target_top = (event.y_root - y0) < h / 2
                break
        drag["target"] = target
        drag["target_top"] = target_top

        # Move / show the drop indicator line
        if target is not None:
            pf = self._sample_frames[target]["row"]
            rx = pf.winfo_x()
            ry = pf.winfo_y()
            rw = pf.winfo_width()
            rh = pf.winfo_height()
            line_y = ry if target_top else ry + rh - 3
            self._drop_line.place(x=rx, y=line_y, width=rw, height=3)
            self._drop_line.lift()
        else:
            self._drop_line.place_forget()

    def _on_header_release(self, event, name):
        drag = self._frame_drag
        self._frame_drag = None
        self._drop_line.place_forget()
        if drag is None:
            return
        # No drag movement → treat as plain click (activate the sample)
        if not drag["active"]:
            self._select_sample(name)
            return
        target = drag.get("target")
        if target is None or target == name:
            return
        self._reorder_samples_by_drag(name, target, before=drag.get("target_top", True))

    def _reorder_samples_by_drag(self, from_name, to_name, *, before=True):
        """Move *from_name* just before (or after) *to_name* in self.samples."""
        keys = list(self.samples.keys())
        if from_name not in keys or to_name not in keys:
            return
        keys.remove(from_name)
        to_idx = keys.index(to_name)
        keys.insert(to_idx if before else to_idx + 1, from_name)
        self.samples = OrderedDict((k, self.samples[k]) for k in keys)
        self._refresh_sample_list()
        self._plot()

    # ─────────────────────────────────────────────────────────────────
    # Plotting coordinator
    # ─────────────────────────────────────────────────────────────────
    def _plot(self):
        if self._suppress_replot:
            return

        # Drop frames for samples that no longer exist
        for n in list(self._sample_frames.keys()):
            if n not in self.samples:
                self._destroy_sample_frame(n)

        any_visible = False
        # Ensure frames exist for current samples in correct order
        for n in self.samples:
            s = self.samples[n]
            if s.get("hidden"):
                # If hidden, destroy frame so it doesn't take space
                if n in self._sample_frames:
                    self._destroy_sample_frame(n)
                continue
            if n not in self._sample_frames:
                self._create_sample_frame(n)
            else:
                # Re-pack in current order
                self._sample_frames[n]["row"].pack_forget()
                self._sample_frames[n]["row"].pack(
                    fill=tk.X, padx=4, pady=4, anchor="nw")
            self._refresh_sample_frame_chrome(n)
            self._plot_sample(n)
            any_visible = True

        # Empty-state placeholder
        if any_visible:
            self._empty_label.pack_forget()
        else:
            self._empty_label.pack(padx=20, pady=20)

        # Update scroll region after layout settles
        self._plot_sc.after(50, lambda: self._plot_sc.configure(
            scrollregion=self._plot_sc.bbox("all")))

    def _plot_sample(self, name: str):
        """Draw OCV and Nyquist axes for one sample."""
        sf = self._sample_frames.get(name)
        s = self.samples.get(name)
        if not sf or not s:
            return
        color = s["color"]

        # ── OCV axis ───────────────────────────────────────────────
        ax_ocv = sf["ocv"]["ax"]
        self._clear_annotation(name, "ocv", redraw=False)
        ax_ocv.clear()
        ax_ocv.set_title(f"{name} — OCV", fontsize=10)
        ax_ocv.set_xlabel("Time (s)")
        ax_ocv.set_ylabel("E (V)")
        ax_ocv.grid(True, linestyle="--", alpha=0.3)
        any_ocv = False
        for i, f in enumerate(s["ocv_files"]):
            df = f["df"]
            tcol = _find_time_col(df)
            vcol = _find_voltage_col(df)
            if tcol is None or vcol is None or df.empty:
                continue
            label = f"OCV{i+1}" if len(s["ocv_files"]) > 1 else "OCV"
            ls = "-" if i == 0 else "--"
            ax_ocv.plot(df[tcol], df[vcol], color=color, linestyle=ls,
                        linewidth=1.3, label=label)
            # Annotate last-point OCV value
            v = f.get("ocv_value")
            if v is not None:
                try:
                    last_t = float(df[tcol].dropna().iloc[-1])
                    ax_ocv.scatter([last_t], [v], color=color,
                                   edgecolor="black", s=60, marker="*",
                                   zorder=5)
                except Exception:
                    pass
            any_ocv = True
        if any_ocv:
            ax_ocv.legend(loc="best", fontsize=8)
        else:
            ax_ocv.text(0.5, 0.5, "No OCV data",
                        ha="center", va="center",
                        transform=ax_ocv.transAxes,
                        color="gray", fontsize=10)

        # ── Nyquist axis ───────────────────────────────────────────
        ax_eis = sf["eis"]["ax"]
        self._clear_annotation(name, "eis", redraw=False)
        ax_eis.clear()
        ax_eis.set_title(f"{name} — Nyquist", fontsize=10)
        ax_eis.set_xlabel("Re(Z) (Ω)")
        ax_eis.set_ylabel("-Im(Z) (Ω)")
        ax_eis.grid(True, linestyle="--", alpha=0.3)
        any_eis = False
        for i, f in enumerate(s["eis_files"]):
            df = f["df"]
            rec = _find_re_z_col(df)
            imc = _find_im_z_col(df)
            if rec is None or imc is None or df.empty:
                continue
            label = f"EIS{i+1}" if len(s["eis_files"]) > 1 else "EIS"
            ydata = df[imc] if "-im" in imc.lower() else -df[imc]
            marker = _EIS_MARKERS[i % len(_EIS_MARKERS)]
            ax_eis.plot(df[rec], ydata, color=color, linestyle="-",
                        marker=marker, markersize=4, linewidth=1.0,
                        label=label)
            rv = f.get("ru_value")
            ry = f.get("ru_y")
            if rv is not None and ry is not None:
                ax_eis.scatter([rv], [ry], color=color,
                               edgecolor="black", s=80,
                               marker="*", zorder=5)
            any_eis = True
        if any_eis:
            ax_eis.legend(loc="best", fontsize=8)
        else:
            ax_eis.text(0.5, 0.5, "No EIS data",
                        ha="center", va="center",
                        transform=ax_eis.transAxes,
                        color="gray", fontsize=10)

        for key, ax in (("ocv", ax_ocv), ("eis", ax_eis)):
            try:
                sf[key]["fig"].tight_layout()
            except Exception:
                pass
            sf[key]["auto_xlim"] = ax.get_xlim()
            sf[key]["auto_ylim"] = ax.get_ylim()
            sf[key]["canvas"].draw_idle()

    # ─────────────────────────────────────────────────────────────────
    # Mouse interactions: scroll-zoom, drag-pan, click-annotate
    # ─────────────────────────────────────────────────────────────────
    def _ax_state(self, sample_name: str, kind: str):
        sf = self._sample_frames.get(sample_name)
        return sf.get(kind) if sf else None

    def _on_scroll(self, event, sample_name, kind):
        st = self._ax_state(sample_name, kind)
        if not st or event.inaxes is not st["ax"]:
            return
        ax = st["ax"]
        scale = 0.8 if event.step > 0 else 1.25
        xl, yl = ax.get_xlim(), ax.get_ylim()
        xd, yd = event.xdata, event.ydata
        if xd is None or yd is None:
            return
        xf = (xd - xl[0]) / (xl[1] - xl[0]) if xl[1] != xl[0] else 0.5
        yf = (yd - yl[0]) / (yl[1] - yl[0]) if yl[1] != yl[0] else 0.5
        nxr = (xl[1] - xl[0]) * scale
        nyr = (yl[1] - yl[0]) * scale
        ax.set_xlim(xd - nxr * xf, xd + nxr * (1 - xf))
        ax.set_ylim(yd - nyr * yf, yd + nyr * (1 - yf))
        st["canvas"].draw_idle()

    def _on_press(self, event, sample_name, kind):
        # Any click activates the sample
        self._select_sample(sample_name)
        st = self._ax_state(sample_name, kind)
        if not st or event.inaxes is not st["ax"]:
            return
        if event.button == 1:
            st["panning"] = True
            st["pan_start"] = (event.xdata, event.ydata)
            st["pan_moved"] = False
        elif event.button == 3:
            # Right-click → clear annotation
            self._clear_annotation(sample_name, kind, redraw=True)

    def _on_motion(self, event, sample_name, kind):
        st = self._ax_state(sample_name, kind)
        if not st or not st.get("panning"):
            return
        if event.inaxes is not st["ax"] or event.xdata is None:
            return
        ax = st["ax"]
        x0, y0 = st["pan_start"]
        dx = x0 - event.xdata
        dy = y0 - event.ydata
        if abs(dx) > 1e-12 or abs(dy) > 1e-12:
            st["pan_moved"] = True
        xl = ax.get_xlim(); yl = ax.get_ylim()
        ax.set_xlim(xl[0] + dx, xl[1] + dx)
        ax.set_ylim(yl[0] + dy, yl[1] + dy)
        st["canvas"].draw_idle()

    def _on_release(self, event, sample_name, kind):
        st = self._ax_state(sample_name, kind)
        if not st:
            return
        was_panning = st.get("panning")
        st["panning"] = False
        if (was_panning and not st.get("pan_moved")
                and event.button == 1
                and event.inaxes is st["ax"]):
            self._handle_click_annotate(event, sample_name, kind)

    def _handle_click_annotate(self, event, sample_name, kind):
        """Annotate the curve point nearest the click; cycle on repeat click."""
        st = self._ax_state(sample_name, kind)
        if not st:
            return
        ax = st["ax"]
        _CYCLE_PX = 8

        lines = [ln for ln in ax.lines
                 if len(ln.get_xdata()) > 0 and ln.get_visible()
                 and not ln.get_label().startswith("_")]
        if not lines:
            return

        candidates = []
        for ln in lines:
            xd = np.asarray(ln.get_xdata(), dtype=float)
            yd = np.asarray(ln.get_ydata(), dtype=float)
            mask = np.isfinite(xd) & np.isfinite(yd)
            if not mask.any():
                continue
            disp = ax.transData.transform(
                np.column_stack([xd[mask], yd[mask]]))
            dists = np.hypot(disp[:, 0] - event.x, disp[:, 1] - event.y)
            best = int(np.argmin(dists))
            candidates.append((float(dists[best]), ln,
                               float(xd[mask][best]), float(yd[mask][best])))
        if not candidates:
            return
        candidates.sort(key=lambda t: t[0])

        last = st.get("ann_last")
        if (last is not None
                and abs(event.x - last[0]) <= _CYCLE_PX
                and abs(event.y - last[1]) <= _CYCLE_PX):
            st["ann_idx"] = (st["ann_idx"] + 1) % len(candidates)
        else:
            st["ann_idx"] = 0
        st["ann_last"] = (event.x, event.y)

        idx = st["ann_idx"]
        n = len(candidates)
        _, ln, x, y = candidates[idx]
        label = ln.get_label() or "?"

        xl = ax.get_xlim(); yl = ax.get_ylim()
        xf = (x - xl[0]) / (xl[1] - xl[0]) if xl[1] != xl[0] else 0.5
        yf = (y - yl[0]) / (yl[1] - yl[0]) if yl[1] != yl[0] else 0.5
        xoff = -95 if xf > 0.65 else 15
        yoff = -60 if yf > 0.65 else 15

        order_hint = f"  [{idx + 1}/{n}]" if n > 1 else ""
        text = f"x = {x:.4g}\ny = {y:.4g}\n{label}{order_hint}"
        if n > 1 and idx == 0:
            text += "\n↻ click again to cycle"

        self._clear_annotation(sample_name, kind, redraw=False)
        st["ann"] = ax.annotate(
            text, xy=(x, y), xytext=(xoff, yoff),
            textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.4", fc="lightyellow",
                      ec="gray", alpha=0.92),
            arrowprops=dict(arrowstyle="->", color="gray", lw=1.2),
            fontsize=8, zorder=10,
        )
        st["ann"].set_in_layout(False)
        st["ann_dot"], = ax.plot(
            x, y, "o", color=ln.get_color(), markersize=7,
            zorder=11, label="_ann_dot")
        st["canvas"].draw_idle()

    def _clear_annotation(self, sample_name, kind, redraw=True):
        st = self._ax_state(sample_name, kind)
        if not st:
            return
        for attr in ("ann", "ann_dot"):
            artist = st.get(attr)
            if artist is not None:
                try:
                    artist.remove()
                except Exception:
                    pass
                st[attr] = None
        st["ann_last"] = None
        st["ann_idx"] = 0
        if redraw:
            st["canvas"].draw_idle()

    # ─────────────────────────────────────────────────────────────────
    # Session save / restore
    # ─────────────────────────────────────────────────────────────────
    def get_session_state(self, data_store: dict) -> dict:
        samples_serial = []
        for name, s in self.samples.items():
            srec = {
                "name": name,
                "color": s["color"],
                "hidden": s.get("hidden", False),
                "ocv_files": [],
                "eis_files": [],
            }
            for f in s["ocv_files"]:
                df = f.get("df")
                if df is None:
                    continue
                h = _sm.df_hash(df)
                data_store[h] = df
                srec["ocv_files"].append({
                    "path": f.get("path", ""),
                    "filename": f.get("filename", ""),
                    "data_hash": h,
                    "ocv_value": f.get("ocv_value"),
                })
            for f in s["eis_files"]:
                df = f.get("df")
                if df is None:
                    continue
                h = _sm.df_hash(df)
                data_store[h] = df
                srec["eis_files"].append({
                    "path": f.get("path", ""),
                    "filename": f.get("filename", ""),
                    "data_hash": h,
                    "ru_value": f.get("ru_value"),
                    "ru_y": f.get("ru_y"),
                })
            samples_serial.append(srec)
        return {
            "active_sample": self.active_sample,
            "plot_w_var": self.plot_w_var.get(),
            "plot_h_var": self.plot_h_var.get(),
            "samples": samples_serial,
        }

    def restore_session_state(self, state: dict, data_store: dict) -> None:
        old = self._suppress_replot
        self._suppress_replot = True
        try:
            self.samples.clear()
            self.active_sample = None
            try:
                self.plot_w_var.set(state.get("plot_w_var", "16.0"))
                self.plot_h_var.set(state.get("plot_h_var", "6.5"))
            except Exception:
                pass

            for srec in state.get("samples", []):
                name = srec.get("name")
                if not name:
                    continue
                ocv_files = []
                for f in srec.get("ocv_files", []):
                    df = data_store.get(f.get("data_hash", ""))
                    if df is None:
                        continue
                    ocv_files.append({
                        "path": f.get("path", ""),
                        "filename": f.get("filename", ""),
                        "df": df.copy(),
                        "ocv_value": f.get("ocv_value"),
                    })
                eis_files = []
                for f in srec.get("eis_files", []):
                    df = data_store.get(f.get("data_hash", ""))
                    if df is None:
                        continue
                    eis_files.append({
                        "path": f.get("path", ""),
                        "filename": f.get("filename", ""),
                        "df": df.copy(),
                        "ru_value": f.get("ru_value"),
                        "ru_y": f.get("ru_y"),
                    })
                self.samples[name] = {
                    "name": name,
                    "color": srec.get("color", "#1f77b4"),
                    "hidden": srec.get("hidden", False),
                    "ocv_files": ocv_files,
                    "eis_files": eis_files,
                }
        finally:
            self._suppress_replot = old

        self._refresh_sample_list()
        active = state.get("active_sample")
        if active and active in self.samples:
            self._select_sample(active)
        self._apply_plot_size()
        self._plot()
