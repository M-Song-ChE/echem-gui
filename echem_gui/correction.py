"""IR compensation and RHE conversion logic."""

from tkinter import messagebox


class CorrectionMixin:
    """Mixin that provides IR / RHE correction behaviour.

    Expects the host class to have:
        self.active_file, self.files
        self.r_sol_var, self.e_ref_var
        self._auto_replot()
    """

    def _apply_correction(self):
        if not self.active_file:
            messagebox.showinfo("Info", "Load a file first.")
            return
        entry = self.files[self.active_file]
        df_raw = entry["df_raw"]

        if "Ewe/V" not in df_raw.columns:
            messagebox.showerror("Error", "Column 'Ewe/V' not found.")
            return

        try:
            r_sol = float(self.r_sol_var.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid R_sol value.")
            return
        try:
            e_ref = float(self.e_ref_var.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid E_ref value.")
            return

        df = df_raw.copy()
        corrected = df["Ewe/V"].copy()

        if r_sol != 0:
            if "<I>/mA" not in df.columns:
                messagebox.showerror("Error", "Column '<I>/mA' not found for IR correction.")
                return
            corrected = corrected - (df["<I>/mA"] / 1000.0) * r_sol

        if e_ref != 0:
            corrected = corrected + e_ref

        df["Ewe/V"] = corrected
        entry["df"] = df
        entry["r_sol"] = r_sol
        entry["e_ref"] = e_ref

        self._auto_replot()

    def _reset_correction(self):
        if not self.active_file:
            return
        entry = self.files[self.active_file]
        entry["df"] = entry["df_raw"].copy()
        entry["r_sol"] = 0.0
        entry["e_ref"] = 0.0
        self.r_sol_var.set("0")
        self.e_ref_var.set("0")

        self._auto_replot()
