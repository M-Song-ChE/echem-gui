"""File loading, removal, and switching logic."""

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import OrderedDict
import pandas as pd


def _is_voltage_col(c):
    lo = c.lower()
    return ("ewe" in lo or "ece" in lo or "potential" in lo or "voltage" in lo
            or lo.startswith("e/")
            or (lo.endswith("/v") and not lo.endswith(("mv", "µv", "nv"))))


def _is_current_col(c):
    lo = c.lower()
    return (lo.startswith("i/") or "i/ma" in lo or "i/a" in lo
            or "i/µa" in lo or "current" in lo)


def _is_time_col(c):
    return c.lower().startswith("time/")


def _is_impedance_col(c):
    lo = c.lower()
    return ("re(z)" in lo or "im(z)" in lo or lo.startswith("freq/")
            or "|z|" in lo or "phase(z)" in lo)


def _default_xcol(cols):
    """Return the best X column based on detected data type (EIS / OCV / CV)."""
    has_impedance = any(_is_impedance_col(c) for c in cols)
    has_current   = any(_is_current_col(c)   for c in cols)
    has_time      = any(_is_time_col(c)       for c in cols)

    # EIS data: prefer Re(Z) on X axis for Nyquist plot
    if has_impedance:
        for c in cols:
            if "re(z)" in c.lower():
                return c
        for c in cols:
            if _is_impedance_col(c):
                return c

    # CV / LSV: voltage + current present → voltage on X
    if has_current:
        for c in cols:
            if _is_voltage_col(c):
                return c

    # OCV / time-series: no current, time present → time on X
    if has_time:
        for c in cols:
            if _is_time_col(c):
                return c

    # Generic fallback: first voltage-like, then second column
    for c in cols:
        if _is_voltage_col(c):
            return c
    return cols[1] if len(cols) > 1 else cols[0]


def _default_ycol(cols, x_col=""):
    """Return the best Y column based on detected data type (EIS / OCV / CV)."""
    has_impedance = any(_is_impedance_col(c) for c in cols)
    has_current   = any(_is_current_col(c)   for c in cols)

    # EIS data: prefer -Im(Z) on Y axis
    if has_impedance:
        for c in cols:
            if c == x_col:
                continue
            if "-im(z)" in c.lower():
                return c
        for c in cols:
            if c != x_col and _is_impedance_col(c):
                return c

    # CV / LSV: current on Y
    if has_current:
        for c in cols:
            if c == x_col:
                continue
            if _is_current_col(c):
                return c

    # OCV / time-series: voltage on Y (X is time)
    for c in cols:
        if c == x_col:
            continue
        if _is_voltage_col(c):
            return c

    # Final fallback: first column that differs from x_col
    for c in cols:
        if c != x_col:
            return c
    return cols[0]


# EC-Lab column names to extract from .mpr files (after angle-bracket cleanup)
_MPR_DESIRED = frozenset({
    "time/s", "Ewe/V", "I/mA", "cycle number",
    "Re(Z)/Ohm", "-Im(Z)/Ohm", "freq/Hz", "Phase(Z)/deg",
})


def _read_mpr(path: str) -> "pd.DataFrame":
    """Read a BioLogic .mpr binary and return a DataFrame of desired columns only.

    Uses galvani (lazily imported).  Raises ImportError with install hint if
    galvani is missing, ValueError if no recognized columns are found.

    When galvani encounters a column ID it doesn't recognise (newer EC-Lab
    firmware), this function injects a float32 placeholder into galvani's
    column-ID map and retries.  Unknown columns are silently skipped because
    they won't appear in _MPR_DESIRED.
    """
    try:
        from galvani import BioLogic
    except ImportError:
        raise ImportError(
            "galvani is required to load .mpr files.\n"
            "Install it with:  pip install galvani"
        )

    # Galvani raises NotImplementedError for column IDs it doesn't know.
    # We inject float32 placeholders and retry.  The placeholder byte-width
    # must match the actual column size in the binary file; try the four
    # most common EC-Lab element sizes until np.frombuffer succeeds.
    _DTYPES = ["<f4", "<f8", "<u4", "<u2"]   # 4, 8, 4, 2 bytes
    _injected: dict[int, str] = {}            # col_id → current dtype string

    mpr = None
    for elem_dtype in _DTYPES:
        # Update any already-injected placeholders to the new candidate size
        for cid in _injected:
            BioLogic.VMPdata_colID_dtype_map[cid] = (f"_unknown_{cid}", elem_dtype)

        for _attempt in range(30):
            try:
                mpr = BioLogic.MPRfile(path)
                break   # success
            except NotImplementedError as exc:
                m = re.search(r"Column ID (\d+)", str(exc))
                if not m:
                    raise
                cid = int(m.group(1))
                BioLogic.VMPdata_colID_dtype_map[cid] = (f"_unknown_{cid}", elem_dtype)
                _injected[cid] = elem_dtype
            except (ValueError, AssertionError) as exc:
                if "buffer size must be a multiple" in str(exc) or isinstance(exc, AssertionError):
                    break   # wrong element size — try next candidate
                raise

        if mpr is not None:
            break
    else:
        raise ValueError(
            "Cannot load .mpr file: unrecognised column layout "
            "(tried element sizes 4, 8, 2 bytes — file may require a newer galvani)."
        )

    df  = pd.DataFrame(mpr.data)
    # Apply the same column cleanup used for .txt files
    df.columns = [c.strip().replace("<", "").replace(">", "") for c in df.columns]
    keep = [c for c in df.columns if c in _MPR_DESIRED]
    if not keep:
        raise ValueError(
            "No recognized columns found in the .mpr file.\n"
            f"Columns present: {list(df.columns)}"
        )
    return df[keep].reset_index(drop=True)


_COLOR_NAMES = ["Blue", "Orange", "Green", "Red", "Purple",
                "Brown", "Pink", "Gray", "Olive", "Cyan"]
_COLOR_HEX = {
    "Blue":   "#1f77b4", "Orange": "#ff7f0e", "Green":  "#2ca02c",
    "Red":    "#d62728", "Purple": "#9467bd", "Brown":  "#8c564b",
    "Pink":   "#e377c2", "Gray":   "#7f7f7f", "Olive":  "#bcbd22",
    "Cyan":   "#17becf",
}
_PALETTE = [_COLOR_HEX[n] for n in _COLOR_NAMES]
_MARKERS = ["o", "s", "^", "D", "v", "P", "*", "X", "h", "p"]


class FileManagerMixin:
    """Mixin that provides file load / remove / switch behaviour.

    Expects the host class to have:
        self.files, self.active_file, self._suppress_replot
        self.file_listbox
        self.x_combo, self.y_combo, self.x_var, self.y_var
        self.r_sol_var, self.e_ref_var
        self._populate_cycle_checkboxes(), self._selected_cycles()
        self._auto_replot()
    """

    def _load_files(self):
        paths = filedialog.askopenfilenames(
            filetypes=[
                ("EC-Lab / Text files", "*.mpr *.txt"),
                ("BioLogic MPR",        "*.mpr"),
                ("Text files",          "*.txt"),
                ("All files",           "*.*"),
            ]
        )
        if not paths:
            return
        for path in paths:
            short = os.path.basename(path)
            base_short = short
            counter = 1
            while short in self.files:
                short = f"{base_short} ({counter})"
                counter += 1
            try:
                if path.lower().endswith(".mpr"):
                    df_raw = _read_mpr(path)
                else:
                    df_raw = pd.read_csv(path, sep="\t")
                    # Strip whitespace and remove angle-bracket wrappers (e.g. <I>/mA → I/mA)
                    df_raw.columns = [c.strip().replace("<", "").replace(">", "")
                                       for c in df_raw.columns]
                    # Drop blank "Unnamed: N" columns produced by trailing tab separators
                    df_raw = df_raw.loc[:, ~df_raw.columns.str.match(r"^Unnamed")]
            except Exception as exc:
                messagebox.showerror("Load error", f"{base_short}: {exc}")
                continue

            color_idx = len(self.files)
            # When only one unique cycle value exists (e.g. single EIS sweep tagged
            # cycle 0), pre-select it so the plot loop doesn't skip the file.
            if ("cycle number" in df_raw.columns
                    and df_raw["cycle number"].nunique() <= 1):
                _auto_cycles = sorted(int(c) for c in df_raw["cycle number"].unique())
            else:
                _auto_cycles = []

            self.files[short] = {
                "path": path,
                "df_raw": df_raw,
                "df": df_raw.copy(),
                "selected_cycles": _auto_cycles,
                "r_sol": 0.0,
                "e_ref": 0.0,
                "area": "",
                "color":          _PALETTE[color_idx % len(_PALETTE)],
                "marker":         _MARKERS[color_idx % len(_MARKERS)],
                "cycle_gradient": True,
                "cycle_reverse":  False,
                "lightness_step": "0.08",
                "hidden":         False,
            }
            self.file_listbox.insert(tk.END, short)

        # Save current file's state, then switch to the newly loaded file.
        # Guard flag prevents <<ListboxSelect>> (fired by selection_set on Windows)
        # from calling _on_file_select prematurely before we do the explicit switch.
        # (do NOT replot — the user's current plot is preserved until they click Plot)
        if not self.files:
            return  # all files failed to load — nothing to switch to
        last_idx = self.file_listbox.size() - 1
        self.file_listbox.selection_clear(0, tk.END)
        self._loading_files = True
        try:
            self.file_listbox.selection_set(last_idx)
        finally:
            self._loading_files = False
        self._save_active_state()
        self._suppress_replot = True
        self._switch_active_file(list(self.files.keys())[last_idx])
        self._suppress_replot = False
        # Auto-display when no meaningful cycle selection is required:
        #   • no "cycle number" column at all (e.g. OCV, plain time-series), OR
        #   • only one unique cycle value (e.g. single EIS sweep labelled cycle 0)
        entry = self.files.get(self.active_file)
        if entry is not None:
            _df = entry["df"]
            if ("cycle number" not in _df.columns
                    or _df["cycle number"].nunique() <= 1):
                self._auto_replot()

    def _remove_file(self):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        short = self.file_listbox.get(idx)
        self.file_listbox.delete(idx)
        del self.files[short]

        if self.active_file == short:
            self.active_file = None
            self._suppress_replot = True
            self._populate_cycle_checkboxes([], [])
            self._suppress_replot = False
            self.x_combo["values"] = []
            self.y_combo["values"] = []
            self.x_var.set("")
            self.y_var.set("")
            self.r_sol_var.set("0")
            self.e_ref_var.set("0")
            if self.files:
                self.file_listbox.selection_set(0)
                self._switch_active_file(list(self.files.keys())[0])
                return
            self._clear_plot()
            return

        self._auto_replot()

    def _on_file_visibility_change(self, short, visible):
        """Called when a file's checkbox is toggled in the CheckableListbox."""
        if short not in self.files:
            return
        self.files[short]["hidden"] = not visible
        self._auto_replot()

    def _on_file_reorder(self, new_order):
        """Called when the file list is drag-reordered. Rebuilds self.files in new order."""
        new_files = OrderedDict()
        for name in new_order:
            if name in self.files:
                new_files[name] = self.files[name]
        # Keep any entries not in new_order (safety)
        for name, entry in self.files.items():
            if name not in new_files:
                new_files[name] = entry
        self.files = new_files
        self._auto_replot()

    def _on_file_select(self, event):
        if getattr(self, "_loading_files", False):
            return  # programmatic selection during _load_files — ignore
        sel = self.file_listbox.curselection()
        if not sel:
            return
        short = self.file_listbox.get(sel[0])
        if short != self.active_file:
            self._save_active_state()
            self._switch_active_file(short)

    def _save_active_state(self):
        """Save cycle selection and correction values for the current file."""
        if self.active_file and self.active_file in self.files:
            entry = self.files[self.active_file]
            entry["selected_cycles"] = self._selected_cycles()
            try:
                entry["r_sol"] = float(self.r_sol_var.get())
            except ValueError:
                pass
            try:
                entry["e_ref"] = float(self.e_ref_var.get())
            except ValueError:
                pass
            area_var = getattr(self, "area_var", None)
            if area_var is not None:
                entry["area"] = area_var.get()

    def _clear_plot(self):
        """Called when all files are removed. Override in subclasses to clear the plot."""
        pass

    def _get_column_list(self, df):
        """Return column names for axis comboboxes. Override to add virtual columns.

        For EIS data (impedance columns detected), only the EIS-relevant columns
        (Re(Z), Im(Z), freq, |Z|, Phase) are exposed so that time/voltage/current/
        cycle-number metadata don't clutter the axis selectors.
        """
        all_cols = list(df.columns)
        eis_cols = [c for c in all_cols if _is_impedance_col(c)]
        if eis_cols:
            return eis_cols
        return all_cols

    def _switch_active_file(self, short):
        """Switch the UI to display the given file's data."""
        self.active_file = short
        entry = self.files[short]
        df = entry["df"]

        # Restore area first so _get_column_list() (which may check area) sees
        # the correct value for the incoming file.
        self.r_sol_var.set(str(entry["r_sol"]))
        self.e_ref_var.set(str(entry["e_ref"]))
        area_var = getattr(self, "area_var", None)
        if area_var is not None:
            area_var.set(entry.get("area", ""))

        cols = self._get_column_list(df)
        self.x_combo["values"] = cols
        self.y_combo["values"] = cols
        if not self.x_var.get() or self.x_var.get() not in cols:
            self.x_var.set(_default_xcol(cols))
        if not self.y_var.get() or self.y_var.get() not in cols:
            self.y_var.set(_default_ycol(cols, self.x_var.get()))

        # Rebuild cycle checkboxes, suppressing auto-replot during update
        old_suppress = self._suppress_replot
        self._suppress_replot = True
        if "cycle number" in df.columns:
            cycles = sorted(int(c) for c in df["cycle number"].unique())
            saved = entry["selected_cycles"]
            self._populate_cycle_checkboxes(cycles, saved)
        else:
            self._populate_cycle_checkboxes([], [])
        self._suppress_replot = old_suppress

        self._on_columns_changed()
        self._auto_replot()

    def _on_columns_changed(self):
        """Called after x_var/y_var are set. Override to refresh unit comboboxes."""
        pass
