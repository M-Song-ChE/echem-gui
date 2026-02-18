"""Excel export – exports the currently selected file only.

Layout:
  Sheet "Raw"       — selected cycles side-by-side with a blank column gap
  Sheet "Corrected" — same layout but with IR/RHE correction applied
"""

import os
from tkinter import filedialog, messagebox
import pandas as pd


def _build_side_by_side(df, cycles):
    """Place each cycle's data block side-by-side, separated by one empty column.

    Returns a single DataFrame ready to write to one sheet.
    """
    blocks = []
    for c in cycles:
        sub = df[df["cycle number"] == c].reset_index(drop=True)
        renamed = sub.rename(columns=lambda col: f"C{c} {col}")
        blocks.append(renamed)

    if not blocks:
        return pd.DataFrame()

    combined = blocks[0]
    for blk in blocks[1:]:
        spacer = pd.DataFrame({"": [""] * max(len(combined), len(blk))})
        combined = pd.concat([combined, spacer, blk], axis=1)

    return combined


class ExportMixin:
    """Mixin that provides Excel export.

    Expects the host class to have:
        self.files, self.active_file
        self._save_active_state()
    """

    def _export_excel(self):
        if not self.active_file:
            messagebox.showinfo("Info", "Select a file first.")
            return

        self._save_active_state()

        entry = self.files[self.active_file]
        df_raw = entry["df_raw"]
        df_corr = entry["df"]
        cycles = entry["selected_cycles"]

        if not cycles and "cycle number" in df_raw.columns:
            messagebox.showinfo("Info", "Select at least one cycle to export.")
            return

        base = os.path.splitext(self.active_file)[0]
        out_path = filedialog.asksaveasfilename(
            initialfile=f"{base}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Excel file",
        )
        if not out_path:
            return

        try:
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                if "cycle number" in df_raw.columns and cycles:
                    raw_sheet = _build_side_by_side(df_raw, cycles)
                    corr_sheet = _build_side_by_side(df_corr, cycles)
                else:
                    raw_sheet = df_raw.copy()
                    corr_sheet = df_corr.copy()

                raw_sheet.to_excel(writer, sheet_name="Raw", index=False)
                corr_sheet.to_excel(writer, sheet_name="Corrected", index=False)

            messagebox.showinfo("Export complete", f"Saved to:\n{out_path}")
        except Exception as exc:
            messagebox.showerror("Export error", str(exc))
