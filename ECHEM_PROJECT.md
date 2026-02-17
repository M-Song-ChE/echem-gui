# Echem GUI — Project Memory

## Project Overview
Electrochemistry Analysis GUI built with Python/tkinter + matplotlib.
**Location:** `C:\Users\thsrk\pycharm\`
**Launch:** `python run_echem.py` or `python -m echem_gui`

## Package Structure
```
pycharm/
  run_echem.py              ← entry point (thin launcher)
  echem_gui/                ← main package
    __init__.py             ← exports EchemGUI
    __main__.py             ← python -m echem_gui support
    app.py                  ← main EchemGUI class, UI layout, mixin assembly
    file_manager.py         ← FileManagerMixin: load/remove/switch files
    correction.py           ← CorrectionMixin: IR compensation + RHE conversion
    plotting.py             ← PlottingMixin: plot, zoom, pan, legend drag/resize, reset view
    ecsa.py                 ← ECSAMixin: electrochemical surface area estimation
    export.py               ← ExportMixin: Excel export (raw + corrected sheets)
    legend_editor.py        ← open_legend_editor(): dialog for renaming legend labels
```

## Architecture
- **Mixin pattern**: each feature is a mixin class; `EchemGUI` inherits all + `tk.Tk`
- **Data model**: `self.files = OrderedDict[str, dict]` keyed by short filename
  - Each entry: `{"path", "df_raw", "df", "selected_cycles", "r_sol", "e_ref"}`
  - `df_raw` = original parsed data; `df` = corrected working copy
- **`self.active_file`**: currently selected filename in UI
- **`self._suppress_replot`**: flag to prevent cascading auto-replots during bulk UI updates
- **Custom toolbar** subclass (`_Toolbar` in app.py) overrides `home()` → `_reset_view()`

## Key Features Implemented
1. **Multi-file support** — load multiple `.txt` files, manage in listbox, overlay on plot
2. **IR correction** — `E_corrected = E_raw - (I_mA / 1000) * R_sol`
3. **RHE conversion** — `E_RHE = E_measured + E_ref_vs_RHE`
4. **Auto-replot** — plot updates on cycle selection, correction apply/reset, file switch
   - Suppressed during file loading (preserves current plot)
5. **Mouse interactions** (matplotlib event system):
   - Scroll wheel = zoom (centered on cursor)
   - Left-drag on plot = pan
   - Left-drag on legend = move (`set_draggable(True)`)
   - Right-drag on legend = resize font (4–30 pt, live update)
6. **Adjustable X/Y range** — entry fields, blank = auto
7. **Legend controls** — font size + location dropdown + "Edit Labels" button
8. **Edit Legend Labels** — dialog (legend_editor.py) with text fields for each entry
9. **Reset View** — Home button (house icon) in toolbar subclass calls `_reset_view()`
10. **Excel export** — active file only, requires cycle selection
    - Sheet "Raw": all selected cycles side-by-side, blank column gap, headers prefixed `C{n}`
    - Sheet "Corrected": same layout with IR/RHE applied
    - Uses `asksaveasfilename` dialog
11. **ECSA calculation** — operates on active file, prototype Cdl estimation

## Dependencies
- Python standard: `tkinter`, `collections`
- Third-party: `pandas`, `numpy`, `matplotlib`, `openpyxl`

## Important Design Decisions
- File loading does NOT auto-replot (user's current plot is preserved)
- Plot skips files with `cycle number` column but no selected cycles
- Export only exports the active (selected) file, blocks if no cycles selected
- Legend hit-testing uses fresh renderer for accurate bbox
- Pan uses data-coordinate invariant: clicked data point stays under cursor

## Known Patterns / Gotchas
- **`_suppress_replot` must save/restore** (not hard-set False) — `_switch_active_file` uses
  `old_suppress = self._suppress_replot` pattern to avoid canceling outer suppression
- `canvas.draw()` (not `draw_idle()`) needed for legend resize to show frame changes in real-time
- `set_bbox_to_anchor` breaks after first use if coordinates are mixed — use `set_draggable(True)`
- Toolbar Home button override requires subclassing `NavigationToolbar2Tk`, not attribute assignment
  (button command is bound at init time, `toolbar.home = ...` does NOT update it)
- Tab-separated `.txt` files expected as input (`pd.read_csv(path, sep="\t")`)
- Column names are stripped of trailing whitespace on load
