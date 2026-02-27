"""CheckableListbox — a Listbox-compatible widget with per-row checkboxes.

Each row shows  [checkbox]  [⠿ handle]  [filename label].
Clicking the checkbox toggles visibility and fires on_check(text, visible).
Clicking/dragging the handle or label selects the file and fires <<ListboxSelect>>;
dragging also reorders rows and fires on_reorder(new_texts_list).

Public API (mirrors tk.Listbox):
    insert(END, text, *, checked=True)
    delete(idx)
    clear()
    selection_clear(start, end)
    selection_set(idx)       → also fires <<ListboxSelect>>
    curselection()           → (idx,) or ()
    get(idx)                 → text at idx
    size()                   → number of rows
    see(idx)                 → scroll row into view
"""

import tkinter as tk
from tkinter import ttk

_SEL_BG    = "#cce8ff"   # selected-row highlight
_NORM_BG   = "white"     # normal row background
_HIDDEN_BG = "#f0f0f0"   # dimmed background for hidden files


class CheckableListbox(tk.Frame):
    """Listbox-compatible widget with a visibility checkbox and drag-to-reorder on each row."""

    def __init__(self, master, *, height=5, on_check=None, on_reorder=None,
                 show_checkboxes=True, **kw):
        """
        Parameters
        ----------
        height          : visible rows (used to set the canvas height in pixels)
        on_check        : callable(text, visible) — called when a checkbox changes
        on_reorder      : callable(new_texts_list) — called after rows are drag-reordered
        show_checkboxes : if False, the checkbox column is hidden (drag-to-reorder only)
        """
        super().__init__(master, **kw)
        self._on_check        = on_check
        self._on_reorder      = on_reorder
        self._show_checkboxes = show_checkboxes
        self._rows            = []   # list of dicts: {text, var, frame, handle, cb, label}
        self._selected_idx = None
        self._rdrag      = None   # drag-to-reorder state

        # ── Internal canvas + scrollbar ──────────────────────────────
        self._canvas = tk.Canvas(self, background=_NORM_BG,
                                 highlightthickness=1, highlightbackground="#aaa",
                                 height=height * 20)
        self._scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL,
                                        command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._scrollbar.set)
        self._scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ── Inner frame that holds all rows ─────────────────────────
        self._inner = tk.Frame(self._canvas, background=_NORM_BG)
        self._win = self._canvas.create_window((0, 0), window=self._inner, anchor=tk.NW)

        self._inner.bind("<Configure>", self._on_inner_configure)
        self._canvas.bind("<Configure>", self._on_canvas_configure)
        self._canvas.bind("<MouseWheel>", self._on_wheel)

        # Drop indicator line (placed on top of the grid with place())
        self._drop_line = tk.Frame(self._inner, bg="#1a73e8", height=2)

    # ── Layout callbacks ─────────────────────────────────────────────
    def _on_inner_configure(self, _e):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_configure(self, e):
        self._canvas.itemconfig(self._win, width=e.width)

    def _on_wheel(self, e):
        self._canvas.yview_scroll(-1 * (e.delta // 120), "units")

    # ── Row building ─────────────────────────────────────────────────
    def _build_row(self, idx, text, *, checked=True):
        """Create the widgets for one row at position *idx* and return the row dict."""
        init_bg = _NORM_BG if checked else _HIDDEN_BG
        row_frame = tk.Frame(self._inner, background=init_bg, cursor="arrow")
        row_frame.grid(row=idx, column=0, sticky="ew", padx=1, pady=0)
        self._inner.columnconfigure(0, weight=1)

        # Set value BEFORE adding the trace so the trace doesn't fire on construction
        var = tk.BooleanVar(value=checked)

        cb = ttk.Checkbutton(row_frame, variable=var)
        if self._show_checkboxes:
            cb.pack(side=tk.LEFT, padx=(2, 0))

        # ⠿ drag handle — cursor="fleur" only here; dragging restricted to this widget
        handle = tk.Label(row_frame, text="⠿", background=init_bg,
                          cursor="fleur", font=("", 11))
        handle.pack(side=tk.LEFT, padx=(3, 0))

        label = tk.Label(row_frame, text=text, anchor=tk.W,
                         background=init_bg, cursor="arrow")
        label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 4))

        row = {"text": text, "var": var, "frame": row_frame,
               "handle": handle, "cb": cb, "label": label}

        # Checkbox toggle fires on_check
        var.trace_add("write", lambda *_: self._on_var_write(text, var))

        # Clicking the label → select only (no drag)
        label.bind("<Button-1>", lambda e, t=text: self._select_by_text(t))

        # Drag handle → select on press + drag-to-reorder on motion
        handle.bind("<Button-1>",        lambda e, t=text: self._on_row_press(e, t))
        handle.bind("<B1-Motion>",       lambda e, t=text: self._on_row_drag(e, t))
        handle.bind("<ButtonRelease-1>", lambda e, t=text: self._on_row_release(e, t))

        # Forward wheel events from every row widget
        for widget in (cb, handle, label, row_frame):
            widget.bind("<MouseWheel>", self._on_wheel)

        return row

    # ── Background helpers ────────────────────────────────────────────
    def _set_row_bg(self, row, bg):
        row["frame"].configure(background=bg)
        row["handle"].configure(background=bg)
        row["label"].configure(background=bg)

    # ── Checkbox callback ─────────────────────────────────────────────
    def _on_var_write(self, text, var):
        visible = var.get()
        if self._on_check is not None:
            self._on_check(text, visible)
        row = next((r for r in self._rows if r["text"] == text), None)
        if row is None:
            return
        is_sel = (self._rows.index(row) == self._selected_idx)
        bg = _SEL_BG if is_sel else (_NORM_BG if visible else _HIDDEN_BG)
        self._set_row_bg(row, bg)

    # ── Selection ─────────────────────────────────────────────────────
    def _select_by_text(self, text):
        idx = next((i for i, r in enumerate(self._rows) if r["text"] == text), None)
        if idx is None:
            return
        self._set_selection(idx, fire_event=True)

    def _set_selection(self, idx, *, fire_event=False):
        if self._selected_idx is not None and self._selected_idx < len(self._rows):
            old = self._rows[self._selected_idx]
            vis = old["var"].get()
            self._set_row_bg(old, _NORM_BG if vis else _HIDDEN_BG)

        self._selected_idx = idx

        if idx is not None and idx < len(self._rows):
            self._set_row_bg(self._rows[idx], _SEL_BG)

        if fire_event:
            self.event_generate("<<ListboxSelect>>")

    # ── Drag-to-reorder ───────────────────────────────────────────────
    def _on_row_press(self, event, text):
        """Button press on a row: select immediately + begin tracking drag."""
        self._select_by_text(text)
        idx = next((i for i, r in enumerate(self._rows) if r["text"] == text), None)
        self._rdrag = {"idx": idx, "text": text,
                       "start_y": event.y_root, "active": False,
                       "target_idx": None, "target_top": True}

    def _on_row_drag(self, event, text):
        d = self._rdrag
        if d is None or d["text"] != text:
            return

        if not d["active"]:
            if abs(event.y_root - d["start_y"]) < 4:
                return
            d["active"] = True

        # Find which row the cursor is over
        target_idx = None
        target_top = True
        for i, r in enumerate(self._rows):
            y0 = r["frame"].winfo_rooty()
            h  = r["frame"].winfo_height()
            if h > 0 and y0 <= event.y_root <= y0 + h:
                target_idx = i
                target_top = (event.y_root - y0) < h / 2
                break

        d["target_idx"] = target_idx
        d["target_top"] = target_top

        if target_idx is not None and target_idx != d["idx"]:
            r = self._rows[target_idx]
            ry = r["frame"].winfo_y()
            rh = r["frame"].winfo_height()
            line_y = ry if target_top else ry + rh - 2
            # relwidth=1.0 fills the parent (_inner) width robustly
            self._drop_line.place(x=0, y=line_y, relwidth=1.0, height=2)
            self._drop_line.lift()
        else:
            self._drop_line.place_forget()

    def _on_row_release(self, event, text):
        self._drop_line.place_forget()
        d = self._rdrag
        self._rdrag = None
        if d is None or not d["active"]:
            return
        from_idx   = d["idx"]
        target_idx = d["target_idx"]
        target_top = d["target_top"]
        if target_idx is None or target_idx == from_idx:
            return
        self._do_reorder_rows(from_idx, target_idx, target_top)

    def _do_reorder_rows(self, from_idx, target_idx, target_top):
        # Note selected row text before reordering
        sel_text = (self._rows[self._selected_idx]["text"]
                    if self._selected_idx is not None and self._selected_idx < len(self._rows)
                    else None)

        row = self._rows.pop(from_idx)
        if from_idx < target_idx:
            target_idx -= 1
        to_idx = target_idx if target_top else target_idx + 1
        self._rows.insert(to_idx, row)

        # Re-grid all rows
        for i, r in enumerate(self._rows):
            r["frame"].grid(row=i, column=0, sticky="ew", padx=1, pady=0)

        # Restore selected_idx to follow the selected row
        if sel_text is not None:
            self._selected_idx = next(
                (i for i, r in enumerate(self._rows) if r["text"] == sel_text), None)

        if self._on_reorder is not None:
            self._on_reorder([r["text"] for r in self._rows])

    # ── Public Listbox-compatible API ────────────────────────────────
    def clear(self):
        """Remove all rows without firing any callbacks."""
        for row in self._rows:
            row["frame"].destroy()
        self._rows.clear()
        self._selected_idx = None

    def insert(self, _pos, text, *, checked=True):
        """Append a new row (END is the only supported position)."""
        idx = len(self._rows)
        row = self._build_row(idx, text, checked=checked)
        self._rows.append(row)

    def delete(self, idx):
        """Remove the row at *idx*."""
        if idx < 0 or idx >= len(self._rows):
            return
        row = self._rows.pop(idx)
        row["frame"].destroy()
        if self._selected_idx is not None:
            if self._selected_idx == idx:
                self._selected_idx = None
            elif self._selected_idx > idx:
                self._selected_idx -= 1
        for i, r in enumerate(self._rows):
            r["frame"].grid(row=i, column=0, sticky="ew", padx=1, pady=0)

    def selection_clear(self, start, end):  # noqa: ARG002
        """Clear the current selection (start/end ignored; only single-select)."""
        self._set_selection(None)

    def selection_set(self, idx):
        """Highlight row *idx* and fire <<ListboxSelect>>."""
        if 0 <= idx < len(self._rows):
            self._set_selection(idx, fire_event=True)

    def curselection(self):
        """Return (idx,) if a row is selected, else ()."""
        if self._selected_idx is not None and self._selected_idx < len(self._rows):
            return (self._selected_idx,)
        return ()

    def get(self, idx):
        """Return the text of row *idx*."""
        if 0 <= idx < len(self._rows):
            return self._rows[idx]["text"]
        return ""

    def size(self):
        """Return the number of rows."""
        return len(self._rows)

    def see(self, idx):
        """Scroll so that row *idx* is visible."""
        if not self._rows or idx < 0 or idx >= len(self._rows):
            return
        total = len(self._rows)
        self._canvas.yview_moveto(idx / total)
