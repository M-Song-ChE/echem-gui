"""Dialog for editing legend labels and order in-place."""

import tkinter as tk
from tkinter import ttk

_MAX_ENTRIES_HEIGHT = 380   # px — dialog entries area is capped at this height


def open_legend_editor(parent, legend_obj, canvas, font_size):
    """Open a popup dialog to rename and reorder legend entries.

    Each row shows ↑/↓ buttons to swap order and an editable text field.

    On OK:
    - If the entry order changed, the legend is recreated via ax.legend(handles, labels, ...)
      so the new order is reflected in the plot.
    - If only text was edited, the Text objects are updated in-place, preserving any dragged
      legend position.

    Returns the legend object (a new one when recreated, the original when text-only).
    """
    ax           = legend_obj.axes
    handles_orig = list(legend_obj.legend_handles)
    texts_orig   = legend_obj.get_texts()
    if not texts_orig:
        return legend_obj

    # Capture settings to replicate if we recreate the legend
    try:
        loc_code = legend_obj._loc
    except AttributeError:
        loc_code = 0
    frame_visible = legend_obj.get_frame().get_visible()

    dlg = tk.Toplevel(parent)
    dlg.title("Edit Legend Labels")
    dlg.resizable(True, True)
    dlg.grab_set()

    ttk.Label(dlg, text="Edit labels · ↑/↓ to reorder:", font=("", 9, "bold")).pack(
        anchor=tk.W, padx=10, pady=(10, 4)
    )

    # ── Scrollable entries area ──────────────────────────────────────
    ent_outer = ttk.Frame(dlg)
    ent_outer.pack(fill=tk.BOTH, expand=True, padx=10, pady=2)

    ent_canvas = tk.Canvas(ent_outer, width=400, height=200, highlightthickness=0)
    ent_scroll = ttk.Scrollbar(ent_outer, orient=tk.VERTICAL, command=ent_canvas.yview)
    ent_canvas.configure(yscrollcommand=ent_scroll.set)

    inner = ttk.Frame(ent_canvas)
    inner_win = ent_canvas.create_window((0, 0), window=inner, anchor=tk.NW)

    def _on_inner_cfg(e):
        ent_canvas.configure(scrollregion=ent_canvas.bbox("all"))
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

    # Mutable list of [handle, StringVar] — pairs are swapped together for reordering
    items = [[h, tk.StringVar(value=t.get_text())]
             for h, t in zip(handles_orig, texts_orig)]

    def _rebuild_rows():
        for w in inner.winfo_children():
            w.destroy()
        for idx in range(len(items)):
            row = ttk.Frame(inner)
            row.pack(fill=tk.X, pady=1)
            ttk.Button(row, text="↑", width=2,
                       command=lambda i=idx: _move(i, -1)).pack(side=tk.LEFT, padx=(0, 1))
            ttk.Button(row, text="↓", width=2,
                       command=lambda i=idx: _move(i, +1)).pack(side=tk.LEFT, padx=(0, 4))
            ent = ttk.Entry(row, textvariable=items[idx][1], width=36)
            ent.pack(side=tk.LEFT, fill=tk.X, expand=True)
            ent.bind("<MouseWheel>", _on_wheel)

    def _move(idx, direction):
        target = idx + direction
        if target < 0 or target >= len(items):
            return
        items[idx], items[target] = items[target], items[idx]
        _rebuild_rows()

    _rebuild_rows()

    # ── Buttons ──────────────────────────────────────────────────────
    result = [legend_obj]   # mutable container so _apply can write back

    def _apply():
        new_handles   = [item[0] for item in items]
        new_labels    = [item[1].get() for item in items]
        order_changed = (new_handles != handles_orig)

        if order_changed:
            # Recreate the legend with the new handle/label order
            new_leg = ax.legend(
                handles=new_handles, labels=new_labels,
                fontsize=font_size,
                loc=loc_code,
                frameon=frame_visible,
            )
            new_leg.set_draggable(True)
            new_leg.get_frame().set_visible(frame_visible)
            canvas.draw()
            result[0] = new_leg
        else:
            # Text-only change: update Text objects in-place to preserve drag position
            for text_obj, item in zip(texts_orig, items):
                text_obj.set_text(item[1].get())
                text_obj.set_fontsize(font_size)
            canvas.draw()
            # result[0] stays as legend_obj (original reference)

        dlg.destroy()

    def _cancel():
        dlg.destroy()

    btn_row = ttk.Frame(dlg)
    btn_row.pack(fill=tk.X, padx=10, pady=(8, 10))
    ttk.Button(btn_row, text="OK", command=_apply).pack(side=tk.RIGHT, padx=(4, 0))
    ttk.Button(btn_row, text="Cancel", command=_cancel).pack(side=tk.RIGHT)

    # Center dialog on parent, clamped to screen
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

    return result[0]
