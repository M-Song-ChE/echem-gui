"""CheckableListbox — a Listbox-compatible widget with per-row checkboxes.

Each row shows  [checkbox]  [filename label].
Clicking the checkbox toggles visibility and fires on_check(text, visible).
Clicking the label selects/activates the file and fires <<ListboxSelect>>.

Public API (mirrors tk.Listbox):
    insert(END, text)
    delete(idx)
    selection_clear(start, end)
    selection_set(idx)       → also fires <<ListboxSelect>>
    curselection()           → (idx,) or ()
    get(idx)                 → text at idx
    size()                   → number of rows
    see(idx)                 → scroll row into view
"""

import tkinter as tk
from tkinter import ttk

_SEL_BG   = "#cce8ff"   # selected-row highlight
_NORM_BG  = "white"     # normal row background


class CheckableListbox(tk.Frame):
    """Listbox-compatible widget with a visibility checkbox on each row."""

    def __init__(self, master, *, height=5, on_check=None, **kw):
        """
        Parameters
        ----------
        height   : visible rows (used to set the canvas height in pixels)
        on_check : callable(text: str, visible: bool) — called when a checkbox changes
        """
        super().__init__(master, **kw)
        self._on_check = on_check
        self._rows = []           # list of dicts: {text, var, frame, cb, label}
        self._selected_idx = None

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

    # ── Layout callbacks ─────────────────────────────────────────────
    def _on_inner_configure(self, _e):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_configure(self, e):
        self._canvas.itemconfig(self._win, width=e.width)

    def _on_wheel(self, e):
        self._canvas.yview_scroll(-1 * (e.delta // 120), "units")

    # ── Row building ─────────────────────────────────────────────────
    def _build_row(self, idx, text):
        """Create the widgets for one row at position *idx* and return the row dict."""
        row_frame = tk.Frame(self._inner, background=_NORM_BG, cursor="arrow")
        row_frame.grid(row=idx, column=0, sticky="ew", padx=1, pady=0)
        self._inner.columnconfigure(0, weight=1)

        var = tk.BooleanVar(value=True)

        cb = ttk.Checkbutton(row_frame, variable=var)
        cb.pack(side=tk.LEFT, padx=(2, 0))

        label = tk.Label(row_frame, text=text, anchor=tk.W,
                         background=_NORM_BG, cursor="arrow")
        label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(2, 4))

        row = {"text": text, "var": var, "frame": row_frame, "cb": cb, "label": label}

        # Checkbox toggle fires on_check
        var.trace_add("write", lambda *_: self._on_var_write(text, var))

        # Clicking the label selects the file
        for widget in (label, row_frame):
            widget.bind("<Button-1>", lambda e, t=text: self._select_by_text(t))

        # Forward wheel events from every row widget
        for widget in (cb, label, row_frame):
            widget.bind("<MouseWheel>", self._on_wheel)

        return row

    def _on_var_write(self, text, var):
        visible = var.get()
        if self._on_check is not None:
            self._on_check(text, visible)
        # Update row background to dim hidden files slightly
        row = next((r for r in self._rows if r["text"] == text), None)
        if row is None:
            return
        is_sel = (self._rows.index(row) == self._selected_idx)
        bg = _SEL_BG if is_sel else (_NORM_BG if visible else "#f0f0f0")
        row["frame"].configure(background=bg)
        row["label"].configure(background=bg)

    def _select_by_text(self, text):
        idx = next((i for i, r in enumerate(self._rows) if r["text"] == text), None)
        if idx is None:
            return
        self._set_selection(idx, fire_event=True)

    def _set_selection(self, idx, *, fire_event=False):
        # Clear old highlight
        if self._selected_idx is not None and self._selected_idx < len(self._rows):
            old = self._rows[self._selected_idx]
            vis = old["var"].get()
            old["frame"].configure(background=_NORM_BG if vis else "#f0f0f0")
            old["label"].configure(background=_NORM_BG if vis else "#f0f0f0")

        self._selected_idx = idx

        # Apply new highlight
        if idx is not None and idx < len(self._rows):
            row = self._rows[idx]
            row["frame"].configure(background=_SEL_BG)
            row["label"].configure(background=_SEL_BG)

        if fire_event:
            self.event_generate("<<ListboxSelect>>")

    # ── Public Listbox-compatible API ────────────────────────────────
    def insert(self, _pos, text):
        """Append a new row (END is the only supported position)."""
        idx = len(self._rows)
        row = self._build_row(idx, text)
        self._rows.append(row)

    def delete(self, idx):
        """Remove the row at *idx*."""
        if idx < 0 or idx >= len(self._rows):
            return
        row = self._rows.pop(idx)
        row["frame"].destroy()
        # Adjust selected index
        if self._selected_idx is not None:
            if self._selected_idx == idx:
                self._selected_idx = None
            elif self._selected_idx > idx:
                self._selected_idx -= 1
        # Re-grid remaining rows
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
        # Scroll the canvas so the row is visible
        self._canvas.yview_moveto(idx / total)
