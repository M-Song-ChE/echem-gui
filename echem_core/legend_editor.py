"""Dialog for editing legend labels and order in-place."""

import tkinter as tk
from tkinter import ttk

_MAX_ENTRIES_HEIGHT = 380   # px — dialog entries area is capped at this height
_SEL_BG  = "#cce8ff"
_NORM_BG = "#f0f0f0"


def open_legend_editor(parent, legend_obj, canvas, font_size):
    """Open a popup dialog to rename and reorder legend entries.

    Each row shows a ⠿ drag handle to reorder and an editable text field.

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

    ttk.Label(dlg, text="Drag ⠿ to reorder · click label to edit:",
              font=("", 9, "bold")).pack(anchor=tk.W, padx=10, pady=(10, 4))

    # ── Scrollable entries area ──────────────────────────────────────
    ent_outer = ttk.Frame(dlg)
    ent_outer.pack(fill=tk.BOTH, expand=True, padx=10, pady=2)

    ent_canvas = tk.Canvas(ent_outer, width=400, height=200, highlightthickness=0)
    ent_scroll = ttk.Scrollbar(ent_outer, orient=tk.VERTICAL, command=ent_canvas.yview)
    ent_canvas.configure(yscrollcommand=ent_scroll.set)

    inner = tk.Frame(ent_canvas, background=_NORM_BG)
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

    # ── Drag-to-reorder state ────────────────────────────────────────
    drag = {"idx": None, "start_y": 0, "active": False,
            "target_idx": None, "target_top": True}

    drop_line = tk.Frame(inner, bg="#1a73e8", height=2)

    def _rebuild_rows():
        for w in inner.winfo_children():
            if w is not drop_line:
                w.destroy()
        inner.columnconfigure(0, weight=1)
        for idx in range(len(items)):
            row = tk.Frame(inner, background=_NORM_BG, cursor="arrow")
            row.grid(row=idx, column=0, sticky="ew", pady=1)

            handle = tk.Label(row, text="⠿", background=_NORM_BG,
                              cursor="fleur", font=("", 11))
            handle.pack(side=tk.LEFT, padx=(4, 2))

            ent = ttk.Entry(row, textvariable=items[idx][1], width=36)
            ent.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
            ent.bind("<MouseWheel>", _on_wheel)

            # Bind drag to the handle
            handle.bind("<Button-1>",        lambda e, i=idx: _on_press(e, i))
            handle.bind("<B1-Motion>",       lambda e, i=idx: _on_drag(e, i))
            handle.bind("<ButtonRelease-1>", lambda e, i=idx: _on_release(e, i))
            handle.bind("<MouseWheel>", _on_wheel)
            row.bind("<MouseWheel>", _on_wheel)

    def _on_press(event, idx):
        drag["idx"]     = idx
        drag["start_y"] = event.y_root
        drag["active"]  = False
        drag["target_idx"] = None
        drag["target_top"] = True

    def _on_drag(event, idx):
        if drag["idx"] != idx:
            return
        if not drag["active"]:
            if abs(event.y_root - drag["start_y"]) < 5:
                return
            drag["active"] = True

        # Identify target row
        rows = [w for w in inner.winfo_children() if isinstance(w, tk.Frame) and w is not drop_line]
        target_idx = None
        target_top = True
        for i, row in enumerate(rows):
            y0 = row.winfo_rooty()
            h  = row.winfo_height()
            if h > 0 and y0 <= event.y_root <= y0 + h:
                target_idx = i
                target_top = (event.y_root - y0) < h / 2
                break

        drag["target_idx"] = target_idx
        drag["target_top"] = target_top

        if target_idx is not None and target_idx != idx:
            row = rows[target_idx]
            ry  = row.winfo_y()
            rh  = row.winfo_height()
            rw  = inner.winfo_width()
            line_y = ry if target_top else ry + rh - 2
            drop_line.place(x=0, y=line_y, width=rw, height=2)
            drop_line.lift()
        else:
            drop_line.place_forget()

    def _on_release(event, idx):
        drop_line.place_forget()
        if not drag["active"]:
            drag["idx"] = None
            return
        from_idx   = drag["idx"]
        target_idx = drag["target_idx"]
        target_top = drag["target_top"]
        drag["idx"] = None
        if from_idx is None or target_idx is None or target_idx == from_idx:
            return
        # Reorder items list
        item = items.pop(from_idx)
        ti = target_idx
        if from_idx < ti:
            ti -= 1
        to_idx = ti if target_top else ti + 1
        items.insert(to_idx, item)
        _rebuild_rows()

    _rebuild_rows()

    # ── Buttons ──────────────────────────────────────────────────────
    # result[0] = legend object, result[1] = permutation list
    # permutation[new_pos] = orig_pos (index into handles_orig / texts_orig)
    result = [legend_obj, list(range(len(items)))]

    def _apply():
        new_handles   = [item[0] for item in items]
        new_labels    = [item[1].get() for item in items]
        order_changed = (new_handles != handles_orig)
        # Permutation: new position i came from original position j
        result[1] = [handles_orig.index(item[0]) for item in items]

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

    return result[0], result[1]
