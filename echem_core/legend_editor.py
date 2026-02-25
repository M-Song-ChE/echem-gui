"""Dialog for editing legend labels in-place."""

import tkinter as tk
from tkinter import ttk

_MAX_ENTRIES_HEIGHT = 380   # px — dialog entries area is capped at this height


def open_legend_editor(parent, legend_obj, canvas, font_size):
    """Open a popup dialog listing all legend labels as editable text fields.

    On OK, update the legend texts and redraw.
    """
    texts = legend_obj.get_texts()
    if not texts:
        return

    dlg = tk.Toplevel(parent)
    dlg.title("Edit Legend Labels")
    dlg.resizable(True, True)
    dlg.grab_set()

    ttk.Label(dlg, text="Edit each legend entry:", font=("", 9, "bold")).pack(
        anchor=tk.W, padx=10, pady=(10, 4)
    )

    # ── Scrollable entries area ──────────────────────────────────────
    ent_outer = ttk.Frame(dlg)
    ent_outer.pack(fill=tk.BOTH, expand=True, padx=10, pady=2)

    ent_canvas = tk.Canvas(ent_outer, width=360, height=200, highlightthickness=0)
    ent_scroll = ttk.Scrollbar(ent_outer, orient=tk.VERTICAL, command=ent_canvas.yview)
    ent_canvas.configure(yscrollcommand=ent_scroll.set)

    inner = ttk.Frame(ent_canvas)
    inner_win = ent_canvas.create_window((0, 0), window=inner, anchor=tk.NW)

    def _on_inner_cfg(e):
        ent_canvas.configure(scrollregion=ent_canvas.bbox("all"))
        # Grow canvas height to fit content, but cap at _MAX_ENTRIES_HEIGHT
        h = min(inner.winfo_reqheight(), _MAX_ENTRIES_HEIGHT)
        ent_canvas.configure(height=h)
    inner.bind("<Configure>", _on_inner_cfg)

    def _on_canvas_cfg(e):
        ent_canvas.itemconfig(inner_win, width=e.width)
    ent_canvas.bind("<Configure>", _on_canvas_cfg)

    def _on_wheel(e):
        ent_canvas.yview_scroll(-1 * (e.delta // 120), "units")
    ent_canvas.bind("<MouseWheel>", _on_wheel)
    inner.bind("<MouseWheel>", _on_wheel)

    ent_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    ent_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    entries = []
    for t in texts:
        row = ttk.Frame(inner)
        row.pack(fill=tk.X, pady=2)
        var = tk.StringVar(value=t.get_text())
        ent = ttk.Entry(row, textvariable=var, width=40)
        ent.pack(fill=tk.X)
        ent.bind("<MouseWheel>", _on_wheel)
        entries.append((t, var))

    # ── Buttons ──────────────────────────────────────────────────────
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

    # Center dialog on parent, but clamp to screen so it's never off-screen
    dlg.update_idletasks()
    sw = dlg.winfo_screenwidth()
    sh = dlg.winfo_screenheight()
    dw = dlg.winfo_width()
    dh = dlg.winfo_height()
    cx = parent.winfo_x() + (parent.winfo_width() - dw) // 2
    cy = parent.winfo_y() + (parent.winfo_height() - dh) // 2
    cx = max(0, min(cx, sw - dw))
    cy = max(0, min(cy, sh - dh))
    dlg.geometry(f"+{cx}+{cy}")
    parent.wait_window(dlg)
