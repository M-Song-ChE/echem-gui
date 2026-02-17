"""ECSA (Electrochemical Surface Area) estimation."""

import numpy as np
from tkinter import messagebox


class ECSAMixin:
    """Mixin that provides ECSA calculation.

    Expects the host class to have:
        self.active_file, self.files
        self.scan_rate_var, self.ecsa_label
        self._selected_cycles()
    """

    def _calc_ecsa(self):
        if not self.active_file:
            messagebox.showinfo("Info", "Load a file first.")
            return
        df = self.files[self.active_file]["df"]

        if "cycle number" not in df.columns or "<I>/mA" not in df.columns or "Ewe/V" not in df.columns:
            self.ecsa_label.config(text="Required columns (Ewe/V, <I>/mA, cycle number) not found.")
            return

        cycles = self._selected_cycles()
        if len(cycles) < 2:
            self.ecsa_label.config(text="Select at least 2 cycles for ECSA estimation.")
            return

        try:
            scan_rate = float(self.scan_rate_var.get())
        except ValueError:
            self.ecsa_label.config(text="Invalid scan rate.")
            return

        delta_i_values = []
        for c in cycles:
            sub = df[df["cycle number"] == c]
            ewe = sub["Ewe/V"]
            mid_e = (ewe.min() + ewe.max()) / 2.0
            i_vals = sub.loc[sub["Ewe/V"].between(mid_e - 0.005, mid_e + 0.005), "<I>/mA"]
            if len(i_vals) >= 2:
                delta_i = (i_vals.max() - i_vals.min()) / 2.0
            else:
                delta_i = np.nan
            delta_i_values.append(delta_i)

        delta_i_arr = np.array(delta_i_values)
        valid = ~np.isnan(delta_i_arr)
        if valid.sum() < 2:
            self.ecsa_label.config(text="Not enough valid cycles to estimate Cdl.")
            return

        mean_delta_i = np.nanmean(delta_i_arr)
        cdl = mean_delta_i / (scan_rate / 1000.0)
        cs = 0.040  # mF/cm^2
        ecsa = cdl / cs

        self.ecsa_label.config(
            text=f"Cdl ≈ {cdl:.4f} mF\nECSA ≈ {ecsa:.2f} cm² (Cs = {cs} mF/cm²)\n(prototype estimate)"
        )
