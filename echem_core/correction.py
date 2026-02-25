"""IR compensation and RHE conversion logic."""


class CorrectionMixin:
    """Mixin that provides IR / RHE correction behaviour.

    Expects the host class to have:
        self.active_file, self.files
        self.r_sol_var, self.e_ref_var
        self._auto_replot()
    """

    def _apply_correction(self):
        """Apply IR compensation and RHE shift to the active file.

        Always re-derives from df_raw so it is safe to call repeatedly.
        Invalid / blank inputs are silently treated as zero.
        """
        if not self.active_file or self.active_file not in self.files:
            return
        entry = self.files[self.active_file]
        df_raw = entry["df_raw"]

        try:
            r_sol = float(self.r_sol_var.get())
        except ValueError:
            r_sol = 0.0
        try:
            e_ref = float(self.e_ref_var.get())
        except ValueError:
            e_ref = 0.0

        try:
            df = df_raw.copy()
            if r_sol != 0.0 and "Ewe/V" in df.columns and "I/mA" in df.columns:
                df["Ewe/V"] = df["Ewe/V"] - (df["I/mA"] / 1000.0) * r_sol
            if e_ref != 0.0 and "Ewe/V" in df.columns:
                df["Ewe/V"] = df["Ewe/V"] + e_ref
        except Exception:
            return  # silently abort if any column issue occurs

        entry["df"]   = df
        entry["r_sol"] = r_sol
        entry["e_ref"] = e_ref
        self._auto_replot()

    def _reset_correction(self):
        if not self.active_file:
            return
        entry = self.files[self.active_file]
        entry["df"]    = entry["df_raw"].copy()
        entry["r_sol"] = 0.0
        entry["e_ref"] = 0.0
        self.r_sol_var.set("0")
        self.e_ref_var.set("0")
        self._auto_replot()
