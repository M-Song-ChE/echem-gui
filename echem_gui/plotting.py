"""Plotting, mouse-wheel zoom, pan-drag, legend drag/resize, axis-range, reset view."""

from tkinter import messagebox


class PlottingMixin:
    """Mixin that provides plotting behaviour.

    Expects the host class to have:
        self.files, self.ax, self.canvas, self.fig
        self.x_var, self.y_var
        self.x_min_var, self.x_max_var, self.y_min_var, self.y_max_var
        self.legend_size_var, self.legend_loc_var
        self._save_active_state(), self._selected_cycles()
    """

    # ── Initialisation (called once from _build_ui) ─────────────────
    def _init_plot_interactions(self):
        """Bind all matplotlib canvas events for zoom, pan, legend resize."""
        self._legend_obj = None
        self._current_legend_size = 8.0
        self._auto_xlim = None
        self._auto_ylim = None

        # Interaction state
        self._panning = False
        self._pan_start = None
        self._legend_resizing = False
        self._resize_start_y = None
        self._resize_start_size = None

        self.canvas.mpl_connect("scroll_event", self._on_scroll)
        self.canvas.mpl_connect("button_press_event", self._on_press)
        self.canvas.mpl_connect("button_release_event", self._on_release)
        self.canvas.mpl_connect("motion_notify_event", self._on_motion)

    # ── Hit-testing ─────────────────────────────────────────────────
    def _event_on_legend(self, event):
        """Return True if matplotlib event is inside the legend bbox."""
        if self._legend_obj is None:
            return False
        try:
            # Need a fresh render so window_extent is accurate
            renderer = self.fig.canvas.get_renderer()
            bbox = self._legend_obj.get_window_extent(renderer)
            return bbox.contains(event.x, event.y)
        except Exception:
            return False

    # ── Mouse-wheel zoom (centred on cursor) ────────────────────────
    def _on_scroll(self, event):
        if event.inaxes != self.ax:
            return
        scale = 0.8 if event.step > 0 else 1.25
        xlim = self.ax.get_xlim()
        ylim = self.ax.get_ylim()
        xd, yd = event.xdata, event.ydata

        new_xr = (xlim[1] - xlim[0]) * scale
        new_yr = (ylim[1] - ylim[0]) * scale
        xf = (xd - xlim[0]) / (xlim[1] - xlim[0]) if xlim[1] != xlim[0] else 0.5
        yf = (yd - ylim[0]) / (ylim[1] - ylim[0]) if ylim[1] != ylim[0] else 0.5

        self.ax.set_xlim(xd - new_xr * xf, xd + new_xr * (1 - xf))
        self.ax.set_ylim(yd - new_yr * yf, yd + new_yr * (1 - yf))
        self.canvas.draw_idle()

    # ── Press / release / motion ────────────────────────────────────
    def _on_press(self, event):
        if event.inaxes != self.ax:
            return

        on_legend = self._event_on_legend(event)

        if event.button == 1:  # left click
            if on_legend:
                # Legend position dragging is handled by set_draggable(True)
                # — do nothing here, just don't start panning
                pass
            else:
                self._panning = True
                self._pan_start = (event.xdata, event.ydata)

        elif event.button == 3:  # right click
            if on_legend:
                self._legend_resizing = True
                self._resize_start_y = event.y
                self._resize_start_size = self._current_legend_size

    def _on_release(self, event):
        self._panning = False
        if self._legend_resizing:
            self._legend_resizing = False
            self.legend_size_var.set(str(int(round(self._current_legend_size))))

    def _on_motion(self, event):
        # ── Panning ─────────────────────────────────────────────────
        if self._panning:
            if event.inaxes != self.ax or event.xdata is None:
                return
            dx = self._pan_start[0] - event.xdata
            dy = self._pan_start[1] - event.ydata
            xlim = self.ax.get_xlim()
            ylim = self.ax.get_ylim()
            self.ax.set_xlim(xlim[0] + dx, xlim[1] + dx)
            self.ax.set_ylim(ylim[0] + dy, ylim[1] + dy)
            self.canvas.draw_idle()
            return

        # ── Legend font-size resize (right-drag) ────────────────────
        if self._legend_resizing and self._legend_obj is not None:
            dy = event.y - self._resize_start_y  # pixels up = larger
            new_size = self._resize_start_size + dy / 5.0
            new_size = max(4.0, min(30.0, new_size))
            self._current_legend_size = new_size
            # Update all text items in the legend
            for text in self._legend_obj.get_texts():
                text.set_fontsize(new_size)
            # Also update the legend title if present
            title = self._legend_obj.get_title()
            if title:
                title.set_fontsize(new_size)
            # Force a full redraw so the legend frame resizes
            self.canvas.draw()

    # ── Reset view ──────────────────────────────────────────────────
    def _reset_view(self):
        """Restore the auto-scaled axis limits from the last _plot() call."""
        if self._auto_xlim is not None:
            self.ax.set_xlim(self._auto_xlim)
            self.ax.set_ylim(self._auto_ylim)
            self.canvas.draw_idle()

    # ── Auto-replot helper ──────────────────────────────────────────
    def _auto_replot(self):
        """Silently re-plot if data and axes are available."""
        if self._suppress_replot:
            return
        if not self.files:
            return
        if not self.x_var.get() or not self.y_var.get():
            return
        self._plot()

    # ── Main plot routine ───────────────────────────────────────────
    def _plot(self):
        if not self.files:
            messagebox.showinfo("Info", "Load a file first.")
            return
        xcol = self.x_var.get()
        ycol = self.y_var.get()
        if not xcol or not ycol:
            messagebox.showinfo("Info", "Select X and Y axes.")
            return

        self._save_active_state()
        self.ax.clear()
        multi = len(self.files) > 1
        has_legend = False

        for short, entry in self.files.items():
            df = entry["df"]
            if xcol not in df.columns or ycol not in df.columns:
                continue
            cycles = entry["selected_cycles"]

            if "cycle number" in df.columns:
                if not cycles:
                    continue
                for c in cycles:
                    sub = df[df["cycle number"] == c]
                    label = f"{short} C{c}" if multi else f"Cycle {c}"
                    self.ax.plot(sub[xcol], sub[ycol], label=label)
                has_legend = True
            else:
                label = short if multi else None
                self.ax.plot(df[xcol], df[ycol], label=label)
                if label:
                    has_legend = True

        self.ax.set_xlabel(xcol)
        self.ax.set_ylabel(ycol)

        # Store auto-scaled limits before user overrides
        self.fig.tight_layout()
        self.canvas.draw()
        self._auto_xlim = self.ax.get_xlim()
        self._auto_ylim = self.ax.get_ylim()

        # Apply manual axis range if specified
        self._apply_axis_range()

        # Legend — use set_draggable(True) for reliable position dragging
        self._legend_obj = None
        if has_legend:
            try:
                legend_size = float(self.legend_size_var.get())
            except ValueError:
                legend_size = 8.0
            self._current_legend_size = legend_size
            legend_loc = self.legend_loc_var.get() or "best"
            self._legend_obj = self.ax.legend(fontsize=legend_size, loc=legend_loc)
            self._legend_obj.set_draggable(True)

        self.canvas.draw()

    # ── Axis range helper ───────────────────────────────────────────
    def _apply_axis_range(self):
        """Set axis limits from the range entry fields. Blank = auto."""
        try:
            x_min = float(self.x_min_var.get())
        except ValueError:
            x_min = None
        try:
            x_max = float(self.x_max_var.get())
        except ValueError:
            x_max = None
        try:
            y_min = float(self.y_min_var.get())
        except ValueError:
            y_min = None
        try:
            y_max = float(self.y_max_var.get())
        except ValueError:
            y_max = None

        if x_min is not None or x_max is not None:
            cur = self.ax.get_xlim()
            self.ax.set_xlim(
                x_min if x_min is not None else cur[0],
                x_max if x_max is not None else cur[1],
            )
        if y_min is not None or y_max is not None:
            cur = self.ax.get_ylim()
            self.ax.set_ylim(
                y_min if y_min is not None else cur[0],
                y_max if y_max is not None else cur[1],
            )
