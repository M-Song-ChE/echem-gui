"""Dialog for editing legend labels in-place."""

import tkinter as tk
from tkinter import ttk


def open_legend_editor(parent, legend_obj, canvas, font_size):
    """Open a popup dialog listing all legend labels as editable text fields.

    On OK, update the legend texts and redraw.
    """
    texts = legend_obj.get_texts()
    if not texts:
        return

    dlg = tk.Toplevel(parent)
    dlg.title("Edit Legend Labels")
    dlg.resizable(False, False)
    dlg.grab_set()

    ttk.Label(dlg, text="Edit each legend entry:", font=("", 9, "bold")).pack(
        anchor=tk.W, padx=10, pady=(10, 4)
    )

    entries = []
    for t in texts:
        row = ttk.Frame(dlg)
        row.pack(fill=tk.X, padx=10, pady=2)
        var = tk.StringVar(value=t.get_text())
        ent = ttk.Entry(row, textvariable=var, width=40)
        ent.pack(fill=tk.X)
        entries.append((t, var))

    def _apply():
        for text_obj, var in entries:
            text_obj.set_text(var.get())
            text_obj.set_fontsize(font_size)
        canvas.draw()
        dlg.destroy()

    def _cancel():
        dlg.destroy()

    btn_row = ttk.Frame(dlg)
    btn_row.pack(fill=tk.X, padx=10, pady=(8, 10))
    ttk.Button(btn_row, text="OK", command=_apply).pack(side=tk.RIGHT, padx=(4, 0))
    ttk.Button(btn_row, text="Cancel", command=_cancel).pack(side=tk.RIGHT)

    # Focus the first entry
    if entries:
        entries[0][1].set(entries[0][1].get())  # trigger var
        dlg.after(50, lambda: list(dlg.children.values())[1].winfo_children()[0].focus_set())

    # Center dialog on parent window
    dlg.update_idletasks()
    pw = parent.winfo_width()
    ph = parent.winfo_height()
    px = parent.winfo_x()
    py = parent.winfo_y()
    dw = dlg.winfo_width()
    dh = dlg.winfo_height()
    dlg.geometry(f"+{px + (pw - dw) // 2}+{py + (ph - dh) // 2}")
