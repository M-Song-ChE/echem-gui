"""File loading, removal, and switching logic."""

import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


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
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
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
            self.files[short] = {
                "path": path,
                "df_raw": df_raw,
                "df": df_raw.copy(),
                "selected_cycles": [],   # none pre-selected; user picks manually
                "r_sol": 0.0,
                "e_ref": 0.0,
                "area": "",
                "color":          _PALETTE[color_idx % len(_PALETTE)],
                "marker":         _MARKERS[color_idx % len(_MARKERS)],
                "cycle_gradient": True,
                "cycle_reverse":  False,
                "lightness_step": "0.08",
            }
            self.file_listbox.insert(tk.END, short)

        # Save current file's state, then switch to the newly loaded file.
        # Guard flag prevents <<ListboxSelect>> (fired by selection_set on Windows)
        # from calling _on_file_select prematurely before we do the explicit switch.
        # (do NOT replot — the user's current plot is preserved until they click Plot)
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
        """Return column names for axis comboboxes. Override to add virtual columns."""
        return list(df.columns)

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
            self.x_var.set(cols[1] if len(cols) > 1 else cols[0])
        if not self.y_var.get() or self.y_var.get() not in cols:
            self.y_var.set(cols[2] if len(cols) > 2 else cols[0])

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

        self._auto_replot()
