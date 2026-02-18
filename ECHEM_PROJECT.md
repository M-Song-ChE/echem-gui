# Echem GUI — Project Memory

## Project Overview
Electrochemistry Analysis GUI built with Python/tkinter + matplotlib.
**Location:** `C:\Users\Mefford\PycharmProjects\echem_gui\echem-gui\`
**Launch:** `python run_echem.py` or `python -m echem_gui`

## Package Structure
```
echem-gui/
  run_echem.py              ← entry point (thin launcher)
  echem_gui/
    __init__.py             ← exports EchemGUI
    __main__.py             ← python -m echem_gui support
    app.py                  ← EchemGUI (tk.Tk window), EchemPanel class
    ecsa_panel.py           ← ECSAPanel class (dedicated ECSA Calc tab)
    file_manager.py         ← FileManagerMixin: load/remove/switch files
    correction.py           ← CorrectionMixin: IR compensation + RHE conversion
    plotting.py             ← PlottingMixin: plot, zoom, pan, legend drag/resize, reset view, click-annotate
    ecsa.py                 ← ECSAMixin: legacy ECSA calc (used only by General E.Chem tab)
    export.py               ← ExportMixin: Excel export (raw + corrected sheets)
    legend_editor.py        ← open_legend_editor(): dialog for renaming legend labels
```

## Architecture
The app uses a **two-tab Notebook** at the top level:
- **General E.Chem tab** → `EchemPanel(ttk.Frame + all mixins)`, `show_log=True`
- **ECSA Calc tab** → `ECSAPanel(ttk.Frame + FileManagerMixin + CorrectionMixin)`

Each panel is fully **independent**: its own `files` dict, `active_file`, figures, and canvases. Switching tabs never affects the other tab's data or plots.

### EchemPanel (General E.Chem)
- Inherits: `FileManagerMixin, CorrectionMixin, PlottingMixin, ECSAMixin, ExportMixin, ttk.Frame`
- Left panel: scrollable canvas with all controls
- Right panel: single matplotlib `Figure` + `NavigationToolbar2Tk` (with custom Home → `_reset_view()`)
- Optional sections: `show_ecsa=True` adds legacy ECSA Calc block; `show_log=True` adds Log widget
- Constructed by `EchemGUI._build_ui()` via `EchemPanel(gen_tab, show_ecsa=False, show_log=True)`

### ECSAPanel (ECSA Calc)
- Inherits: `FileManagerMixin, CorrectionMixin, ttk.Frame`
- Left panel: scrollable canvas (files, axis+unit selectors, IR correction, cycles, scan-rate table, ECSA params, buttons, legend-frame toggle, log)
- Right panel: **two independent figures** stacked vertically, each with its own `NavigationToolbar2Tk`:
  - `fig_cv / ax_cv / canvas_cv` — CV curves (upper)
  - `fig_cdl / ax_cdl / canvas_cdl` — Cdl extraction scatter + linear fit (lower)
- Interactions (zoom/pan/annotate/legend drag) registered on **both** canvases; `_get_canvas(ax)` routes draw calls to the correct one

### EchemGUI (main window)
- Inherits only `tk.Tk`
- Creates `ttk.Notebook`, adds `gen_tab` and `ecsa_tab` frames, instantiates one panel per tab

### Data model (per panel instance)
- `self.files = OrderedDict[str, dict]` keyed by short filename
  - Each entry: `{"path", "df_raw", "df", "selected_cycles", "r_sol", "e_ref"}`
  - `df_raw` = original parsed data; `df` = corrected working copy
  - `selected_cycles` is always `[]` on first load; user picks manually
- `self.active_file`: currently selected filename
- `self._suppress_replot`: prevents cascading auto-replots during bulk UI updates
- `self._loading_files`: blocks `<<ListboxSelect>>` during programmatic `selection_set()`

## Key Features

### General E.Chem tab
1. **Multi-file support** — load multiple `.txt` files, manage in listbox, overlay on plot
2. **Axis selectors + unit dropdowns** — X and Y each have a column selector and a unit combobox with dimension-aware filtering (I/E/t families)
3. **Unit conversion** — `_get_axis_unit_scale(col, target)` returns `(scale_factor, display_label)`; scale applied to data before plotting
4. **IR correction** — `E_corrected = E_raw − (I_mA / 1000) × R_sol`
5. **RHE conversion** — `E_RHE = E_measured + E_ref_vs_RHE`
6. **Reference electrode selector** — appended to x-axis label as `(vs Ag/AgCl)` etc.
7. **Auto-replot** — updates on cycle selection, correction, file switch; suppressed during load
8. **Plot range** — blank entry = auto; triggers replot on Return / FocusOut
9. **Cycle checkboxes** — 9-column grid, scrollable (vertical + horizontal), "Select All / Deselect All"
10. **Legend controls** — show/hide, frame toggle, font size, location, "Edit Labels" dialog, drag to move, right-drag to resize
11. **Mouse interactions** (via `PlottingMixin`):
    - Scroll wheel = zoom centred on cursor
    - Left-drag = pan
    - Left-click = annotate nearest point (pixel-space distance, overlap cycling through lines)
    - Right-click = dismiss annotation
12. **Excel export** — active file only; "Raw" and "Corrected" sheets, cycles side-by-side
13. **Log widget** — scrollable read-only text area at the bottom of the left panel
14. **Legacy ECSA section** (hidden from General tab currently) — prototype Cdl estimate via `ECSAMixin`

### ECSA Calc tab
1. **Independent file state** — loads its own files, does not share with General tab
2. **Axis selectors + unit dropdowns** — same dimension-aware unit filtering as General tab; units applied for display only (extraction uses raw column values)
3. **IR / RHE correction** — same as General tab
4. **Cycle checkboxes** — 9-column grid, same UX as General tab
5. **Scan-rate per cycle table** — 8-column grid, dynamically rebuilt when cycle selection changes; each entry has a `trace_add("write", …)` that triggers a debounced (300 ms) CV replot so legend updates as you type
6. **E_std entry** — vertical red dashed line drawn on CV; Return/FocusOut triggers replot
7. **Cs entry** — specific capacitance (default 0.040 mF/cm²)
8. **Plot CV button** — explicit replot of selected cycles
9. **Extract Cdl & ECSA button** — runs extraction, updates lower plot, logs results
10. **CV Legend Frame toggle** — "Show CV Legend Frame" checkbox
11. **Deselect All** — clears CV plot to placeholder (does NOT fall back to all cycles)
12. **File switch** — clears Cdl plot and result label automatically
13. **Two independent toolbars** — each matplotlib toolbar controls only its own plot (Home, Back, Forward, Zoom, Pan, Configure, Save all work per-plot)
14. **Log widget** — same as General tab

## ECSA Physics (Cdl extraction)
```
For each selected cycle (one cycle = one scan rate):
  1. Split CV at argmax(E) → anodic branch (E↑) / cathodic branch (E↓)
  2. Interpolate ja at E_std on anodic branch
  3. Interpolate jc at E_std on cathodic branch
  4. Δj/2 = (ja − jc) / 2

Linear fit:  scan_rate (mV/s)  vs  Δj/2 (mA)
  slope [mA/(mV/s)] = [10⁻³ A / (10⁻³ V/s)] = F

  cdl_mF = slope × 1000     [mF]
  ECSA   = cdl_mF / Cs      [cm²]   (Cs default 0.040 mF/cm²)
```
- Extraction always uses raw column data regardless of display unit selection.
- E_std is in the same units as the raw x-column (typically V).
- Result logged per file, allowing multiple catalysts to be processed in sequence.

## Dependencies
- Python standard: `tkinter`, `collections`
- Third-party: `pandas`, `numpy`, `matplotlib`, `openpyxl`

## Important Design Decisions
- File loading does NOT auto-replot (preserves current plot)
- Newly loaded files have `selected_cycles = []` — no cycles pre-checked
- Plot skips files with `cycle number` column but no selected cycles
- Export only exports the active file; blocks if no cycles selected
- `constrained_layout=True` on Figure prevents subplot title/label overlap
- Two separate Figure objects in ECSAPanel (not subplots) gives each plot a fully functional independent toolbar including Home/Zoom history

## Known Patterns / Gotchas
- **`_suppress_replot` must save/restore** — use `old = self._suppress_replot` pattern, not hard-set `False`
- **`_loading_files` guard** — `selection_set()` fires `<<ListboxSelect>>` synchronously on Windows; wrap with `_loading_files = True/False` and early-return in `_on_file_select`
- `canvas.draw()` (not `draw_idle()`) needed for legend resize to show frame changes in real-time
- `set_draggable(True)` called after every `ax.legend(...)` call; old legend ref becomes stale after `ax.clear()` so reset to `None` before clearing
- Toolbar Home button override requires subclassing `NavigationToolbar2Tk` (attribute assignment does not work — command is bound at init time)
- Tab-separated `.txt` files expected; column names normalized on load: whitespace stripped, `<`/`>` removed (e.g. `<Ewe>/V` → `Ewe/V`)
- `_clear_annotation(redraw=False)` must be called **before** `ax.clear()` so `artist.remove()` runs on live axes
- Annotation highlight dot uses label `"_click_dot"` (or `"_ann_dot"`); `_`-prefix hides from legend and excludes from pick candidates via `not ln.get_label().startswith("_")`
- `_pan_moved` reset to `False` on every press, set `True` on actual motion; gates annotation on release
- Scan rate `StringVar` traces accumulate if not removed — `_rebuild_sr_table` calls `var.trace_remove()` for all previous trace IDs before rebuilding
- Cycle column count: **9** for both EchemPanel and ECSAPanel
- Scan-rate table column count: **8** for ECSAPanel
