"""Plotting, mouse-wheel zoom, pan-drag, legend drag/resize, axis-range, reset view,
and click-to-annotate with overlap cycling."""

import colorsys
import io
import math
import numpy as np
import matplotlib.colors as _mcolors
from tkinter import messagebox


def copy_figure_to_clipboard(fig):
    """Copy *fig* to the Windows clipboard as a high-resolution PNG image.

    Renders the figure at 150 DPI with tight bounding box, converts to the
    CF_DIB clipboard format (standard Windows bitmap), and places it on the
    clipboard so the user can paste directly into Word / PowerPoint / etc.

    Requires Pillow. Falls back to an error dialog if Pillow is absent or the
    copy fails for any reason.
    """
    try:
        import ctypes
        from PIL import Image

        buf = io.BytesIO()
        fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
        buf.seek(0)
        img = Image.open(buf).convert('RGB')

        # Convert to BMP DIB (Windows clipboard native image format).
        # The BMP file header is 14 bytes; strip it to obtain raw DIB data.
        bmp_buf = io.BytesIO()
        img.save(bmp_buf, 'BMP')
        dib_data = bmp_buf.getvalue()[14:]

        GMEM_MOVEABLE = 0x0002
        CF_DIB = 8
        kernel32 = ctypes.windll.kernel32
        user32   = ctypes.windll.user32

        # Declare proper 64-bit return types; without this, ctypes defaults to
        # c_int (32-bit) which truncates HGLOBAL handles on 64-bit Windows,
        # causing the "access violation writing 0x..." crash.
        kernel32.GlobalAlloc.argtypes  = [ctypes.c_uint, ctypes.c_size_t]
        kernel32.GlobalAlloc.restype   = ctypes.c_void_p
        kernel32.GlobalLock.argtypes   = [ctypes.c_void_p]
        kernel32.GlobalLock.restype    = ctypes.c_void_p
        kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
        user32.OpenClipboard.argtypes  = [ctypes.c_void_p]
        user32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
        user32.SetClipboardData.restype  = ctypes.c_void_p

        hMem = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(dib_data))
        if not hMem:
            raise RuntimeError("GlobalAlloc failed")
        pMem = kernel32.GlobalLock(hMem)
        if not pMem:
            raise RuntimeError("GlobalLock failed")
        ctypes.memmove(pMem, dib_data, len(dib_data))
        kernel32.GlobalUnlock(hMem)

        if not user32.OpenClipboard(None):
            raise RuntimeError("OpenClipboard failed")
        user32.EmptyClipboard()
        user32.SetClipboardData(CF_DIB, ctypes.c_void_p(hMem))
        user32.CloseClipboard()

    except ImportError:
        messagebox.showwarning(
            "Copy to Clipboard",
            "Pillow is required for clipboard copy.\n"
            "Install with:  pip install Pillow"
        )
    except Exception as exc:
        messagebox.showerror("Copy to Clipboard", f"Failed to copy plot: {exc}")

# Label prefix for internal artists (hidden from legend, excluded from click picking)
_ANN_DOT_LABEL = "_click_dot"


# Maps J density range units to the underlying current base unit
_J_TO_BASE = {"A/cm²": "A", "mA/cm²": "mA", "µA/cm²": "µA", "nA/cm²": "nA"}

# Physical-dimension unit sets used for "vs Ref" axis-label guarding
_VOLTAGE_UNITS = frozenset({"V", "mV", "µV", "nV"})
_CURRENT_UNITS = frozenset({"A", "mA", "µA", "nA"})

# Pixel radius: if a new click lands within this many pixels of the previous one,
# treat it as a "same spot" repeated click and advance to the next overlapping line.
_CLICK_CYCLE_PX = 8

_GRID_STYLE_MAP = {"solid": "-", "dashed": "--", "dotted": ":", "dash-dot": "-."}


def _cycle_colors(base_color, n, step=0.08, reverse=False):
    """Return n colors with linearly varying lightness around base_color.

    reverse=False → first=lightest, last=darkest (final cycle most visible).
    reverse=True  → first=darkest, last=lightest.
    """
    if n <= 0:
        return []
    r, g, b = _mcolors.to_rgb(base_color)
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    colors = []
    for i in range(n):
        offset = ((n - 1) / 2 - i) * step   # positive for i=0, negative for i=n-1
        if reverse:
            offset = -offset
        new_l = min(0.85, max(0.15, l + offset))
        colors.append(colorsys.hls_to_rgb(h, new_l, s))
    return colors


def apply_grid(ax, x_grid, y_grid, x_interval, y_interval, style="dashed",
               linewidth=0.8, color="gray"):
    """Apply X/Y grid lines to *ax* with optional fixed tick interval.

    Parameters
    ----------
    ax         : matplotlib Axes
    x_grid     : bool – show X grid
    y_grid     : bool – show Y grid
    x_interval : str or float – tick spacing for X; 0 or blank = auto
    y_interval : str or float – tick spacing for Y; 0 or blank = auto
    style      : one of "solid", "dashed", "dotted", "dash-dot"
    linewidth  : grid line width (default 0.8)
    color      : grid line color (default "gray")
    """
    from matplotlib.ticker import MultipleLocator, AutoLocator
    ls = _GRID_STYLE_MAP.get(style, "--")
    try:
        lw = float(linewidth)
    except (ValueError, TypeError):
        lw = 0.8
    ax.xaxis.set_major_locator(AutoLocator())
    ax.yaxis.set_major_locator(AutoLocator())
    # Always apply interval as tick spacing (even when grid is off)
    try:
        xi = float(x_interval)
        if xi > 0:
            ax.xaxis.set_major_locator(MultipleLocator(xi))
    except (ValueError, TypeError):
        pass
    try:
        yi = float(y_interval)
        if yi > 0:
            ax.yaxis.set_major_locator(MultipleLocator(yi))
    except (ValueError, TypeError):
        pass
    ax.grid(False)
    if x_grid:
        ax.grid(True, axis="x", which="major", linestyle=ls,
                alpha=0.4, color=color, linewidth=lw)
    if y_grid:
        ax.grid(True, axis="y", which="major", linestyle=ls,
                alpha=0.4, color=color, linewidth=lw)


def draw_reflines(ax, reflines):
    """Draw vertical (X) and horizontal (Y) reference lines on ax.

    reflines = list of ('x'|'y', float, style, color[, linewidth]) tuples.
    linewidth is optional (defaults to 1.0 for backward compatibility).
    Labels start with '_' so they are excluded from the legend automatically.
    """
    for entry in reflines:
        axis, val, style, color = entry[:4]
        try:
            lw = float(entry[4]) if len(entry) > 4 else 1.0
        except (ValueError, TypeError):
            lw = 1.0
        ls = _GRID_STYLE_MAP.get(style, '--')
        if axis == 'x':
            ax.axvline(val, color=color, linestyle=ls,
                       linewidth=lw, alpha=0.7, label='_xref')
        else:
            ax.axhline(val, color=color, linestyle=ls,
                       linewidth=lw, alpha=0.7, label='_yref')


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
        """Bind all matplotlib canvas events for zoom, pan, legend resize, click-annotate."""
        self._legend_obj = None
        self._current_legend_size = 8.0
        self._legend_stable_map  = {}    # {stable_key: custom_label}
        self._legend_stable_keys = []    # stable keys emitted during the last _plot()
        self._legend_auto_labels = []    # auto-labels from the last _plot() (for edit diffing)
        self._custom_xlabel      = None  # user-set axis label override (double-click)
        self._custom_ylabel      = None
        self._legend_manual_pos = None   # saved dragged position (axes-fraction tuple)
        self._auto_xlim = None
        self._auto_ylim = None

        # Pan state
        self._panning = False
        self._pan_start = None
        self._pan_moved = False       # True only when the mouse actually moves during a press

        # Legend resize state
        self._legend_resizing = False
        self._resize_start_y = None
        self._resize_start_size = None

        # Click-annotate state
        self._ann = None              # current matplotlib Annotation artist
        self._ann_dot = None          # highlight marker on the selected data point
        self._last_click_pos = None   # (x, y) in display pixels of last annotation click
        self._click_candidate_idx = 0 # which candidate is currently shown (for cycling)

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

    # ── Title-area hit test (shared by press handlers) ──────────────
    @staticmethod
    def _hit_title_area(event, ax, fig):
        """Return True if the event is on the title text or in the strip above the axes."""
        try:
            renderer = event.canvas.get_renderer()
            t_bbox  = ax.title.get_window_extent(renderer)
            ax_bbox = ax.get_window_extent(renderer)
            fig_bbox = fig.get_window_extent(renderer)
            # Exact title text hit
            if t_bbox.width > 2 and t_bbox.contains(event.x, event.y):
                return True
            # Title strip: above top of axes, within axes x-span, within figure
            if (ax_bbox.x0 <= event.x <= ax_bbox.x1
                    and ax_bbox.y1 <= event.y <= fig_bbox.y1):
                return True
        except Exception:
            pass
        return False

    # ── Press / release / motion ────────────────────────────────────
    def _on_press(self, event):
        # Handle dblclick on title / axis labels (may be outside axes bounding box)
        if event.button == 1 and getattr(event, 'dblclick', False):
            if self._hit_title_area(event, self.ax, self.fig):
                self._edit_plot_title()
                return
            try:
                renderer = event.canvas.get_renderer()
                xl = self.ax.xaxis.label
                if xl.get_text() and xl.get_window_extent(renderer).contains(event.x, event.y):
                    self._edit_axis_label('x')
                    return
                yl = self.ax.yaxis.label
                if yl.get_text() and yl.get_window_extent(renderer).contains(event.x, event.y):
                    self._edit_axis_label('y')
                    return
            except Exception:
                pass

        if event.inaxes != self.ax:
            return

        on_legend = self._event_on_legend(event)

        if event.button == 1:  # left button
            self._pan_moved = False   # reset drag-detection on each new press
            if on_legend:
                # Legend dragging handled by set_draggable(True); don't start panning
                if getattr(event, 'dblclick', False):
                    self._edit_legend_labels()
                    return
            else:
                if getattr(event, 'dblclick', False):
                    # dblclick inside axes but not on legend → check title then ignore
                    if self._hit_title_area(event, self.ax, self.fig):
                        self._edit_plot_title()
                    return
                self._panning = True
                self._pan_start = (event.xdata, event.ydata)

        elif event.button == 3:  # right button
            if on_legend:
                self._legend_resizing = True
                self._resize_start_y = event.y
                self._resize_start_size = self._current_legend_size

    def _on_release(self, event):
        self._panning = False

        # Save whether we were resizing before clearing the flag
        was_resizing = self._legend_resizing
        self._legend_resizing = False
        if was_resizing:
            self.legend_size_var.set(str(int(round(self._current_legend_size))))
            return

        on_legend = self._event_on_legend(event)

        if (event.button == 1
                and not self._pan_moved
                and event.inaxes == self.ax
                and not on_legend):
            # True left-click (no drag) inside axes → annotate nearest point
            self._handle_click_annotate(event)

        elif event.button == 3 and not on_legend:
            # Right-click outside legend → dismiss annotation
            self._clear_annotation()

    def _on_motion(self, event):
        # ── Panning ─────────────────────────────────────────────────
        if self._panning:
            if event.inaxes != self.ax or event.xdata is None:
                return
            self._pan_moved = True    # actual mouse movement detected → this is a drag
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
            for text in self._legend_obj.get_texts():
                text.set_fontsize(new_size)
            title = self._legend_obj.get_title()
            if title:
                title.set_fontsize(new_size)
            # Force full redraw so the legend frame resizes visually
            self.canvas.draw()

    # ── Click-annotate: nearest point + overlap cycling ─────────────
    def _handle_click_annotate(self, event):
        """Find the data point closest (in pixel space) to the click and annotate it.

        Overlap handling: if the user clicks the same screen spot repeatedly
        (within _CLICK_CYCLE_PX pixels), each click advances to the next-nearest
        line, cycling through all plotted lines ranked by pixel distance.
        The annotation shows [current/total] so the user always knows how many
        overlapping lines are available.
        """
        # Collect real data lines, excluding internal annotation markers
        lines = [
            ln for ln in self.ax.lines
            if len(ln.get_xdata()) > 0
            and ln.get_visible()
            and not ln.get_label().startswith("_")
        ]
        if not lines:
            return

        # For every line, find the single point nearest the click in PIXEL space.
        # Pixel-space distance is used because it matches what the user sees on screen,
        # regardless of the axis units or scale.
        candidates = []
        for ln in lines:
            xdata = np.asarray(ln.get_xdata(), dtype=float)
            ydata = np.asarray(ln.get_ydata(), dtype=float)
            mask = np.isfinite(xdata) & np.isfinite(ydata)
            if not mask.any():
                continue
            xd, yd = xdata[mask], ydata[mask]
            # ax.transData.transform: data coords → display pixels (origin = fig bottom-left)
            disp = self.ax.transData.transform(np.column_stack([xd, yd]))
            dists = np.hypot(disp[:, 0] - event.x, disp[:, 1] - event.y)
            best = int(np.argmin(dists))
            candidates.append((float(dists[best]), ln, float(xd[best]), float(yd[best])))

        if not candidates:
            return

        # Sort all candidates by ascending pixel distance so [0] is always closest
        candidates.sort(key=lambda t: t[0])

        # ── Cycling logic ────────────────────────────────────────────
        # Same-spot repeated click → advance to next candidate
        if (self._last_click_pos is not None
                and abs(event.x - self._last_click_pos[0]) <= _CLICK_CYCLE_PX
                and abs(event.y - self._last_click_pos[1]) <= _CLICK_CYCLE_PX):
            self._click_candidate_idx = (self._click_candidate_idx + 1) % len(candidates)
        else:
            # New click location → start from the nearest line
            self._click_candidate_idx = 0

        self._last_click_pos = (event.x, event.y)

        idx = self._click_candidate_idx
        n = len(candidates)
        dist_px, ln, x, y = candidates[idx]

        label = ln.get_label() or "?"

        # ── Build annotation text ────────────────────────────────────
        order_hint = f"  [{idx + 1}/{n}]" if n > 1 else ""
        text = f"x = {x:.4g}\ny = {y:.4g}\n{label}{order_hint}"
        if n > 1 and idx == 0:
            text += "\n↻ click again to cycle"

        # ── Smart offset: push text box away from the axes edge ─────
        xlim = self.ax.get_xlim()
        ylim = self.ax.get_ylim()
        xspan = xlim[1] - xlim[0] or 1.0
        yspan = ylim[1] - ylim[0] or 1.0
        x_frac = (x - xlim[0]) / xspan
        y_frac = (y - ylim[0]) / yspan
        xoff = -95 if x_frac > 0.65 else 15
        yoff = -60 if y_frac > 0.65 else 15

        # ── Remove old annotation before drawing new one ─────────────
        self._clear_annotation(redraw=False)

        self._ann = self.ax.annotate(
            text,
            xy=(x, y),
            xytext=(xoff, yoff),
            textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.4", fc="lightyellow", ec="gray", alpha=0.92),
            arrowprops=dict(arrowstyle="->", color="gray", lw=1.2),
            fontsize=8,
            zorder=10,
        )

        # Highlight dot on the selected data point (labeled with "_" → excluded from legend)
        self._ann_dot, = self.ax.plot(
            x, y, "o",
            color=ln.get_color(),
            markersize=7,
            zorder=11,
            label=_ANN_DOT_LABEL,
        )

        self.canvas.draw_idle()
        getattr(self, '_sync_file_selection_from_line', lambda _ln: None)(ln)

    def _clear_annotation(self, redraw=True):
        """Remove the current click annotation and its highlight dot."""
        if self._ann is not None:
            try:
                self._ann.remove()
            except Exception:
                pass
            self._ann = None
        if self._ann_dot is not None:
            try:
                self._ann_dot.remove()
            except Exception:
                pass
            self._ann_dot = None
        # Reset cycling state so the next click starts fresh
        self._last_click_pos = None
        self._click_candidate_idx = 0
        if redraw:
            self.canvas.draw_idle()

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

        # ── Resolve "J" virtual column to the real current column ──────────
        _x_is_J = (xcol == "J")
        _y_is_J = (ycol == "J")

        # Find the actual current column from the first file that has one
        _real_xcol = xcol
        _real_ycol = ycol
        if _x_is_J or _y_is_J:
            for entry in self.files.values():
                for c in entry["df"].columns:
                    if "/" in c:
                        u = c.rsplit("/", 1)[-1].strip()
                        if u in _CURRENT_UNITS:
                            if _x_is_J:
                                _real_xcol = c
                            if _y_is_J:
                                _real_ycol = c
                            break
                break

        # ── Resolve axis scales ────────────────────────────────────────────
        x_unit_var = getattr(self, "x_unit_var", None)
        y_unit_var = getattr(self, "y_unit_var", None)
        _x_unit_str = x_unit_var.get() if x_unit_var else "(auto)"
        _y_unit_str = y_unit_var.get() if y_unit_var else "(auto)"

        # X-axis scale
        if _x_is_J:
            _xbase = _J_TO_BASE.get(_x_unit_str)
            if _xbase:
                x_scale_base, _ = self._get_axis_unit_scale(_real_xcol, _xbase)
                x_label = f"J ({_x_unit_str})"
            else:
                x_scale_base = 1.0
                _src = _real_xcol.rsplit("/", 1)[-1].strip() if "/" in _real_xcol else "?"
                x_label = f"J ({_src}/cm²)"
        else:
            x_scale_base, x_label = self._get_axis_unit_scale(_real_xcol, _x_unit_str)

        # Y-axis scale
        if _y_is_J:
            _ybase = _J_TO_BASE.get(_y_unit_str)
            if _ybase:
                y_scale_base, _ = self._get_axis_unit_scale(_real_ycol, _ybase)
                y_label = f"J ({_y_unit_str})"
            else:
                y_scale_base = 1.0
                _src = _real_ycol.rsplit("/", 1)[-1].strip() if "/" in _real_ycol else "?"
                y_label = f"J ({_src}/cm²)"
        else:
            y_scale_base, y_label = self._get_axis_unit_scale(_real_ycol, _y_unit_str)

        # Clear annotation BEFORE ax.clear() so .remove() still works on live artists
        self._clear_annotation(redraw=False)
        # Save dragged legend position before axes are cleared
        if self._legend_obj is not None:
            _loc = getattr(self._legend_obj, '_loc', None)
            if isinstance(_loc, (tuple, list)):
                self._legend_manual_pos = tuple(_loc)
        self.ax.clear()

        multi = len(self.files) > 1
        has_legend = False
        self._legend_stable_keys = []   # reset for this replot

        _lw_v = getattr(self, 'linewidth_var', None)
        try:
            _lw = float(_lw_v.get()) if _lw_v else 1.5
        except (ValueError, TypeError):
            _lw = 1.5

        for short, entry in self.files.items():
            if entry.get("hidden", False):
                continue
            df = entry["df"]
            if _real_xcol not in df.columns or _real_ycol not in df.columns:
                continue
            cycles = entry["selected_cycles"]

            # Per-file area for J density conversion
            try:
                _farea = float(entry.get("area", "") or 0)
            except (ValueError, TypeError):
                _farea = 0.0

            x_scale = (x_scale_base / _farea if (_x_is_J and _farea > 0)
                       else x_scale_base)
            y_scale = (y_scale_base / _farea if (_y_is_J and _farea > 0)
                       else y_scale_base)

            _grad = entry.get("cycle_gradient", True)
            _rev  = entry.get("cycle_reverse",  False)
            try:    _step = float(entry.get("lightness_step", "0.08"))
            except: _step = 0.08
            base_color = entry.get("color", "#1f77b4")
            if "cycle number" in df.columns:
                if not cycles:
                    continue
                cycle_cols = (_cycle_colors(base_color, len(cycles), _step, _rev)
                              if _grad else [base_color] * len(cycles))
                for i, c in enumerate(cycles):
                    sub   = df[df["cycle number"] == c]
                    label = f"{short} C{c}" if multi else f"Cycle {c}"
                    self.ax.plot(sub[_real_xcol] * x_scale,
                                 sub[_real_ycol] * y_scale,
                                 color=cycle_cols[i], label=label, linewidth=_lw)
                    # Stable key always includes filename → survives single↔multi transitions
                    self._legend_stable_keys.append(f"{short}:C{c}")
                has_legend = True
            else:
                label = short if multi else None
                self.ax.plot(df[_real_xcol] * x_scale,
                            df[_real_ycol] * y_scale,
                            color=base_color, label=label, linewidth=_lw)
                if label:
                    self._legend_stable_keys.append(short)
                    has_legend = True

        # ── Axis labels: append "(vs Ref)" only for voltage-type axes ──────────
        ref = getattr(self, "ref_electrode_var", None)
        ref_text = ref.get().strip() if ref else ""

        # J is never voltage → no "(vs Ref)"
        if _x_is_J:
            _x_is_V = False
        else:
            _x_src = _real_xcol.rsplit("/", 1)[-1].strip() if "/" in _real_xcol else ""
            _x_is_V = (_x_unit_str in _VOLTAGE_UNITS if _x_unit_str != "(auto)"
                       else _x_src in _VOLTAGE_UNITS)
        try:
            _lpad = float(getattr(self, 'label_pad_var', None).get())
        except Exception:
            _lpad = 4.0
        try:
            _tpad = float(getattr(self, 'title_pad_var', None).get())
        except Exception:
            _tpad = 6.0

        xlabel = f"{x_label}  (vs {ref_text})" if (ref_text and _x_is_V) else x_label
        self.ax.set_xlabel(
            self._custom_xlabel if self._custom_xlabel is not None else xlabel,
            labelpad=_lpad)

        if _y_is_J:
            _y_is_V = False
        else:
            _y_src = _real_ycol.rsplit("/", 1)[-1].strip() if "/" in _real_ycol else ""
            _y_is_V = (_y_unit_str in _VOLTAGE_UNITS if _y_unit_str != "(auto)"
                       else _y_src in _VOLTAGE_UNITS)
        ylabel = f"{y_label}  (vs {ref_text})" if (ref_text and _y_is_V) else y_label
        self.ax.set_ylabel(
            self._custom_ylabel if self._custom_ylabel is not None else ylabel,
            labelpad=_lpad)

        # Restore user-set plot title (persists across ax.clear() calls)
        _title_var = getattr(self, "plot_title_var", None)
        if _title_var is not None:
            self.ax.set_title(_title_var.get(), pad=_tpad)

        # Store auto-scaled limits before user overrides
        self.fig.tight_layout()
        self.canvas.draw()
        self._auto_xlim = self.ax.get_xlim()
        self._auto_ylim = self.ax.get_ylim()

        # Apply manual axis range if specified
        self._apply_axis_range()

        # Reference lines — each tuple carries its own style and color
        draw_reflines(self.ax, getattr(self, '_reflines', []))

        # Legend — use set_draggable(True) for reliable position dragging
        self._legend_obj = None
        if has_legend and self.legend_show_var.get():
            try:
                legend_size = float(self.legend_size_var.get())
            except ValueError:
                legend_size = 8.0
            self._current_legend_size = legend_size
            legend_loc = self.legend_loc_var.get() or "best"
            self._legend_obj = self.ax.legend(fontsize=legend_size, loc=legend_loc)
            self._legend_obj.set_draggable(True)
            frame_visible = getattr(self, "legend_frame_var", None)
            self._legend_obj.get_frame().set_visible(
                frame_visible.get() if frame_visible is not None else True
            )
            # Capture auto-labels (used in edit dialog for change-detection)
            self._legend_auto_labels = [t.get_text() for t in self._legend_obj.get_texts()]
            # Apply custom labels by stable key (survives single↔multi-file transitions)
            for i, text_obj in enumerate(self._legend_obj.get_texts()):
                if i < len(self._legend_stable_keys):
                    custom = self._legend_stable_map.get(self._legend_stable_keys[i])
                    if custom:
                        text_obj.set_text(custom)
            # Restore dragged position (overrides loc= argument)
            if self._legend_manual_pos is not None:
                self._legend_obj._loc = self._legend_manual_pos

        _xgv = getattr(self, 'x_grid_var', None)
        if _xgv is not None:
            _glw_v = getattr(self, 'grid_linewidth_var', None)
            _gcol_v = getattr(self, 'grid_color_var', None)
            apply_grid(
                self.ax,
                _xgv.get(),
                getattr(self, 'y_grid_var').get(),
                getattr(self, 'x_grid_int_var').get(),
                getattr(self, 'y_grid_int_var').get(),
                getattr(self, 'grid_style_var').get(),
                linewidth=_glw_v.get() if _glw_v else "0.8",
                color=_gcol_v.get() if _gcol_v else "gray",
            )

        _font_fn = getattr(self, '_apply_font_to_ax', None)
        if _font_fn is not None:
            _font_fn(self.ax, self.canvas)
        else:
            self.canvas.draw()

    # ── Unit conversion helper ───────────────────────────────────────
    def _get_axis_unit_scale(self, col, target_unit):
        """Return (scale_factor, display_label) for *col* converted to *target_unit*.

        Scale factor is applied to the raw data column before plotting.
        Display label replaces the axis label on the chart.

        Rules:
        - target_unit == "(auto)" → no conversion, label = raw column name
        - source and target in the same physical dimension → compute factor
        - unknown / mismatched dimension → scale=1, label uses the new unit string
        """
        if not target_unit or target_unit == "(auto)":
            if "/" in col:
                _cb, _cu = col.rsplit("/", 1)
                return 1.0, f"{_cb.strip()} ({_cu.strip()})"
            return 1.0, col

        # SI scale factors for every supported unit
        _FACTORS = {
            "A":   1.0,    "mA":  1e-3,   "µA":  1e-6,   "nA":  1e-9,
            "V":   1.0,    "mV":  1e-3,   "µV":  1e-6,   "nV":  1e-9,
            "s":   1.0,    "ms":  1e-3,   "µs":  1e-6,
            "min": 60.0,   "h":   3600.0,
            # Impedance (base = Ω)
            "Ohm": 1.0,    "Ω":   1.0,    "mΩ":  1e-3,
            "kΩ":  1e3,    "MΩ":  1e6,
            # Frequency (base = Hz)
            "Hz":  1.0,    "kHz": 1e3,    "MHz": 1e6,
            # Phase angle (base = rad)
            "rad": 1.0,    "deg": math.pi / 180.0,
        }
        # Physical dimension tags (conversion only allowed within same tag)
        _DIMS = {
            "A": "I", "mA": "I", "µA": "I", "nA": "I",
            "V": "E", "mV": "E", "µV": "E", "nV": "E",
            "s": "t", "ms": "t", "µs": "t", "min": "t", "h": "t",
            "Ohm": "Z", "Ω": "Z", "mΩ": "Z", "kΩ": "Z", "MΩ": "Z",
            "Hz": "f", "kHz": "f", "MHz": "f",
            "rad": "φ", "deg": "φ",
        }

        # Extract source unit and base name from column name (e.g. "Ewe/V" → "V", "Ewe")
        if "/" in col:
            col_base, source_unit = col.rsplit("/", 1)
            col_base   = col_base.strip()
            source_unit = source_unit.strip()
        else:
            col_base    = col
            source_unit = None

        display_label = f"{col_base} ({target_unit})"

        src_f = _FACTORS.get(source_unit)
        tgt_f = _FACTORS.get(target_unit)
        if (src_f is not None and tgt_f is not None
                and _DIMS.get(source_unit) == _DIMS.get(target_unit)):
            return src_f / tgt_f, display_label

        # Can't determine conversion — show data unchanged, update label only
        return 1.0, display_label

    # ── Plot title editor ────────────────────────────────────────────
    def _edit_plot_title(self):
        """Prompt the user to edit the main plot title (double-click on title area)."""
        from tkinter.simpledialog import askstring
        current = self.ax.title.get_text()
        new_title = askstring("Edit Title", "Plot title:", initialvalue=current, parent=self)
        if new_title is not None:
            self.ax.set_title(new_title)
            self.canvas.draw_idle()
            # Sync back to the left-panel title entry if present
            _title_var = getattr(self, "plot_title_var", None)
            if _title_var is not None:
                _title_var.set(new_title)

    def _edit_axis_label(self, which):
        """Edit X or Y axis label by double-clicking it. Empty string reverts to auto."""
        from tkinter.simpledialog import askstring
        if which == 'x':
            current = self.ax.get_xlabel()
            new = askstring("Edit X Label", "X axis label\n(blank to revert to auto):",
                            initialvalue=current, parent=self)
            if new is not None:
                self._custom_xlabel = new if new.strip() else None
                self.ax.set_xlabel(
                    new if new.strip() else current)
                self.canvas.draw_idle()
        else:
            current = self.ax.get_ylabel()
            new = askstring("Edit Y Label", "Y axis label\n(blank to revert to auto):",
                            initialvalue=current, parent=self)
            if new is not None:
                self._custom_ylabel = new if new.strip() else None
                self.ax.set_ylabel(
                    new if new.strip() else current)
                self.canvas.draw_idle()

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

        # Flip axes if requested (toggle direction without changing numeric extent)
        xl = self.ax.get_xlim()
        if getattr(self, "x_flip_var", None) and self.x_flip_var.get() != (xl[0] > xl[1]):
            self.ax.set_xlim(xl[1], xl[0])
        yl = self.ax.get_ylim()
        if getattr(self, "y_flip_var", None) and self.y_flip_var.get() != (yl[0] > yl[1]):
            self.ax.set_ylim(yl[1], yl[0])
