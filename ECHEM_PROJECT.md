# Echem GUI — Project Memory

## Project Overview
Electrochemistry GUI built with Python/tkinter + matplotlib.
**Repo:** https://github.com/M-Song-ChE/echem-gui
**Location:** `C:\Users\Mefford\PycharmProjects\echem_gui\`
**Launch:** `python run_echem.py` or `python -m echem_core`

## Package Structure
```
echem_gui/
  run_echem.py              ← entry point (thin launcher)
  echem_core/
    __init__.py             ← exports EchemGUI
    __main__.py             ← python -m echem_core support
    app.py                  ← EchemGUI (tk.Tk window), EchemPanel class
    multi_echem_panel.py    ← MultiEchemPanel class (Multi E.Chem tab)
    multi_echem2_panel.py   ← MultiEchem2Panel class (Multi E.Chem 2 tab — group-based overlay)
    ecsa_panel.py           ← ECSAPanel class (dedicated ECSA Calc tab)
    eis_panel.py            ← EISPanel class (Nyquist Plot tab)
    file_manager.py         ← FileManagerMixin: load/remove/switch files; data-type-aware _default_xcol/_default_ycol; column-type predicates (_is_voltage_col, _is_current_col, _is_time_col, _is_impedance_col); _on_file_visibility_change; _on_file_reorder; _MPR_DESIRED frozenset; _read_mpr(path) with retry loop for unknown galvani column IDs (tries <f4>/<f8>/<u4>/<u2> until buffer size matches)
    correction.py           ← CorrectionMixin: IR compensation + RHE conversion
    plotting.py             ← PlottingMixin: plot, zoom, pan, legend drag/resize, reset view, click-annotate; draw_reflines() helper
    ecsa.py                 ← ECSAMixin: legacy ECSA calc (used only by General E.Chem tab)
    export.py               ← ExportMixin: Excel export (raw + corrected sheets)
    legend_editor.py        ← open_legend_editor(): blocking dialog — rename + ⠿ drag-handle reorder legend entries; returns new legend (recreated when order changes, original when text-only)
    checklist.py            ← CheckableListbox: tk.Frame subclass; [checkbox][⠿ handle][label] rows; Listbox-compatible API (insert/delete/selection_set/curselection/get/size/see); fires <<ListboxSelect>> on label click, on_check(text, visible) on checkbox toggle, and on_reorder(new_texts_list) after drag-to-reorder; ⠿ handle has cursor="fleur" + all drag bindings; row_frame and label have cursor="arrow" + selection only; internal Canvas+Scrollbar; used by General/Multi/EIS tabs (ECSA uses plain tk.Listbox)
```

## Architecture
The app uses a **five-tab Notebook** at the top level (in this order):
- **General E.Chem tab** → `EchemPanel(ttk.Frame + all mixins)`, `show_log=True`
- **Multi E.Chem tab** → `MultiEchemPanel(ttk.Frame + FileManagerMixin + CorrectionMixin)`
- **Multi E.Chem 2 tab** → `MultiEchem2Panel(ttk.Frame + FileManagerMixin + CorrectionMixin)` — group-based overlay; each group has its own figure; files assigned to groups; active group drives left-panel controls
- **ECSA Calc tab** → `ECSAPanel(ttk.Frame + FileManagerMixin + CorrectionMixin)`
- **Nyquist Plot tab** → `EISPanel(ttk.Frame + FileManagerMixin)`

Each panel is fully **independent**: its own `files` dict, `active_file`, figures, and canvases. Switching tabs never affects the other tab's data or plots.

### EchemPanel (General E.Chem)
- Inherits: `FileManagerMixin, CorrectionMixin, PlottingMixin, ECSAMixin, ExportMixin, ttk.Frame`
- Left panel: scrollable canvas with all controls
- Right panel: single matplotlib `Figure` + `NavigationToolbar2Tk` (custom Home → `_reset_view()`)
- Optional sections: `show_ecsa=True` adds legacy ECSA Calc block; `show_log=True` adds Log widget
- **J virtual column**: when all loaded files have a positive electrode area, a "J" entry appears in both column comboboxes; selecting it computes current density (raw current / area) per file at plot time. Unit range combobox shows A/cm², mA/cm², µA/cm², nA/cm²
- **Per-file view preservation**: `_plot()` captures `ax.get_xlim()/get_ylim()` into `_prev_view` before `ax.clear()` (only when `auto_xlim` is not None); restores them after auto-scale and before `_apply_axis_range()` so manual-range inputs still override
- **Click-to-switch**: `_sync_file_selection_from_line` (called on annotate click) saves current state, suppresses replot, switches active file, and restores xlim/ylim so clicking any plot line updates the full left-panel to that file's settings
- **Highlight (Origin-style)**: `_plot_highlight` bool (default `False`); set True on listbox select or line click, reset on right-click; `_apply_highlight_to_axes()` dims unselected lines to alpha=0.55, raises active file to zorder=3, draws a glow shadow (`linewidth×2.5, alpha=0.18, label='_glow'`); called after legend so legend handles keep alpha=1.0; `_active_cycle` (int|None) narrows highlight to a single cycle when set by clicking a cycle line
- **Cycle-specific highlight**: clicking a cycle line (`f"{short} C{c}"`) in `_sync_file_selection_from_line` sets `self._active_cycle = c`; `_apply_highlight_to_axes` checks `_active_cycle` to glow only that cycle; listbox select resets `_active_cycle = None` (whole-file highlight)
- **Reference lines**: panel-level `self._reflines` (list of `('x'|'y', float, style, color)` 4-tuples); persists across file adds/removes; drawn on the shared overlay plot via `draw_reflines()`
- **Overrides** `_save_active_state`, `_switch_active_file`, `_get_column_list`, `_clear_plot` from mixins

### MultiEchemPanel (Multi E.Chem)
- Inherits: `FileManagerMixin, CorrectionMixin, ttk.Frame`
- Left panel: shared axis/unit controls, plot range, reference electrode, IR/RHE correction, cycle checkboxes, legend options — all apply to the **active file only**
- Right panel: scrollable canvas holding one `tk.Frame + Figure` per loaded file, arranged in a 2-column grid; each frame has a blue-gray `⠿  {short}` header strip (`cursor="fleur"`) that serves as both a visual label and a drag handle for reordering; figures remain visible simultaneously
- **Click-to-select**: clicking anywhere on a file's plot or toolbar frame calls `_activate_file(short)`, which updates the listbox selection and switches the left panel to that file
- **J virtual column**: per-file check — "J" appears in column comboboxes only when the active file has area > 0; each file's J is computed with its own stored area
- **Per-file view preservation**: `_plot_file(short)` captures `ax.get_xlim()/get_ylim()` into `_prev_view` before `ax.clear()` (only when `auto_xlim` is not None); restores them after auto-scale and before `_apply_range()` so manual inputs still override; `view_xlim`/`view_ylim` in entry dict still used by `_on_file_visibility_change` for hide/unhide snapshots
- **Reference lines**: per-file `entry["reflines"]` (list of `('x'|'y', float, style, color)` 4-tuples); listbox refreshes on `_switch_active_file`; drawn in `_plot_file` via `draw_reflines()`
- **Highlight (gold header)**: `_highlight_active_header()` sets active file's header strip to `#ffd54f` (gold), others to `#c0cfe4` (default blue-gray); called from `_switch_active_file`; each file's `hdr_frame`/`hdr_label` stored in its entry dict at figure-creation time
- **Scroll to active**: `_scroll_to_active_file(short)` scrolls `_right_canvas` so the active file's frame is visible; called from `_switch_active_file` via `after_idle`
- **Key methods**: `_create_file_figure`, `_relayout_figures`, `_plot_file(short)`, `_activate_file(short)`, `_reset_file_view(short)`, `_highlight_active_header()`, `_scroll_to_active_file(short)`

### ECSAPanel (ECSA Calc)
- Inherits: `FileManagerMixin, CorrectionMixin, ttk.Frame`
- IR/RHE correction UI **removed** (correction not needed for ECSA); `r_sol_var` and `e_ref_var` kept as hidden `StringVar` so `FileManagerMixin` still works
- Left panel: scrollable canvas (files, axis+unit selectors, CV plot range, reference electrode, cycles, scan-rate table, ECSA params, buttons, legend-frame toggles, result label, log)
- Right panel: **two independent figures** stacked vertically, each with its own `NavigationToolbar2Tk`:
  - `fig_cv / ax_cv / canvas_cv` — CV curves (upper)
  - `fig_cdl / ax_cdl / canvas_cdl` — Cdl extraction scatter + linear fit (lower); legend shows equation, Cdl, R², and ECSA
- Interactions (zoom/pan/annotate/legend drag) registered on **both** canvases; `_get_canvas(ax)` routes draw calls to the correct one
- **Per-file view preservation**: saves/restores xlim/ylim for both `ax_cv` and `ax_cdl` independently; Cdl view restored right after `_replot_cdl`, CV view restored after the direct `_plot_cv()` call in `_switch_active_file`
- **Per-file isolation**: E_std, Cs, scan-rate data, axis selections, column/unit choices, plot ranges, legend settings, Cdl data and result text all saved/restored per file
- **CV and Cdl plot titles** include the active filename (e.g. `"sample.txt  —  CV Curves  (non-Faradaic region)"`) to prevent confusion when multiple files are loaded
- **Reference lines**: separate `entry["cv_reflines"]` and `entry["cdl_reflines"]` per file (list of `('x'|'y', float, style, color)` 4-tuples); two independent UI sections in the left panel; listboxes refresh on `_switch_active_file`; drawn in `_plot_cv`, `_replot_cdl`, and `_extract_cdl_ecsa` via `draw_reflines()`
- **Overrides** from mixins:
  - `_clear_plot` — clears both plots and result label when all files removed
  - `_save_active_state` — saves all per-file ECSA state + view limits
  - `_switch_active_file` — restores full per-file state, Cdl plot, view limits
  - `_auto_replot` — delegates to `_plot_cv()` only
  - `_plot` — also delegates to `_plot_cv()`
  - `_edit_legend_labels(leg, is_cv)` — disables legend draggable, opens editor for the specified legend (CV or Cdl), updates `_legend_cv` / `_legend_cdl` with the returned legend object, re-enables draggable; also triggered by double-click on either legend via `_ei_press`

### EchemGUI (main window)
- Inherits only `tk.Tk`
- Creates `ttk.Notebook`, adds `gen_tab`, `multi_tab`, `ecsa_tab`, `eis_tab` frames in that order
- Each tab instantiates its panel directly; no shared state

### Data model (per panel instance)
- `self.files = OrderedDict[str, dict]` keyed by short filename
  - Base fields set on load: `{"path", "df_raw", "df", "selected_cycles", "r_sol", "e_ref", "area", "hidden"}`
  - `df_raw` = original parsed data; `df` = corrected working copy
  - `selected_cycles` is always `[]` on first load; user picks manually
  - `area` = electrode area string (cm²); used for J density calculation
  - `hidden` = bool (default `False`); set by `_on_file_visibility_change`; plot loops skip entries where `hidden=True`
  - `view_xlim`, `view_ylim` = axis limit snapshot used by `MultiEchemPanel._on_file_visibility_change` for hide/unhide; general zoom preservation is now handled internally by `_plot_file`/`_plot_group` via `_prev_view` local variable
  - **Color/marker fields** (set in `file_manager._load_files`): `"color"` (hex string from Tab10-like palette), `"marker"` (matplotlib marker string); palette cycles through 10 named colors so successive files auto-differentiate
  - **Per-file gradient fields** (defaulted in `_switch_active_file`): `"cycle_gradient"` (bool, default `True`), `"cycle_reverse"` (bool, default `False`), `"lightness_step"` (str float, default `"0.15"`); saved by `_save_active_state` / `_on_gradient_change`
- `self.active_file`: currently selected filename
- `self._suppress_replot`: prevents cascading auto-replots during bulk UI updates
- `self._loading_files`: blocks `<<ListboxSelect>>` during programmatic `selection_set()`

**MultiEchemPanel additional per-file field:**
- `custom_title` — user-edited subplot title string (default `""`); set via the Title entry in the left panel or by double-clicking the title strip; used in `_plot_file` (active file reads `plot_title_var`, non-active reads `entry["custom_title"]`)

**MultiEchem2Panel additional per-group field:**
- `custom_title` — user-edited group plot title (default `""`); set via the Title entry or by double-clicking the title strip; saved in `_save_active_group_state()`, restored in `_switch_active_group()`; `_plot_group()` uses `plot_title_var.get()` for active group, `gentry.get("custom_title", "")` for others
- `hdr_frame`, `hdr_label` — header strip `tk.Frame` and `tk.Label` widgets stored in gentry at creation time; used by `_highlight_active_headers()` to set gold/green background
- `line_to_file` — `{Line2D: filename}` dict rebuilt on every `_plot_group()`; used for click-to-select
- `line_to_cycle` — `{Line2D: cycle_int}` dict rebuilt on every `_plot_group()`; used for cycle-specific highlight
- `legend_labels` — `dict` keyed by `"{fname}:C{c}"` (or `fname` for non-cycle files) mapping to custom label strings; replaces the old positional list so labels survive cycle add/remove and file reordering

**ECSAPanel additional per-file fields** (set by `setdefault` in `_switch_active_file`):
- `sr_data`, `e_std`, `cs`, `x_col`, `y_col`, `x_unit`, `y_unit`
- `x_min`, `x_max`, `y_min`, `y_max`, `ref_electrode`
- `legend_frame_cv`, `legend_frame_cdl`, `cdl_data`, `result_text`
- `cv_x_grid`, `cv_y_grid`, `cv_x_grid_int`, `cv_y_grid_int`, `cv_grid_style`
- `cdl_x_grid`, `cdl_y_grid`, `cdl_x_grid_int`, `cdl_y_grid_int`, `cdl_grid_style`
- `cv_reflines`, `cdl_reflines` — lists of `('x'|'y', float, style, color)` 4-tuples
- `view_xlim_cv`, `view_ylim_cv`, `view_xlim_cdl`, `view_ylim_cdl`

**ECSAPanel panel-level state:**
- `self._sr_vars = {cycle_num: StringVar}` — scan rate entry per cycle
- `self._sr_traces = {cycle_num: (var, trace_id)}` — trace IDs for cleanup before rebuild
- `self._cv_redraw_id` — `after()` ID for debounced CV replot (300 ms)
- `self.e_std_rec_var` — `StringVar` for the green "Rec:" label; updated by `_update_e_std_rec()` on every `_plot_cv()` call

## Key Features

### General E.Chem tab
1. **Multi-file support** — load multiple `.txt`/`.mpr` files, manage in CheckableListbox, overlay on single plot; checkbox hides/shows a file's contribution without losing any settings (cycles, corrections, zoom, colors, etc.); ⠿ drag handle in each row reorders files (fires `_on_file_reorder`)
1b. **Auto-merge sequential CV files** — `_load_files` groups selected files by EC-Lab CVA pattern `_(\d{2,3})_([A-Za-z]+)_(C\d+)\.(mpr|txt)$`; groups with ≥2 files whose method is in `_MERGE_METHODS = {CV, CVA, LSV, DPV, NPV, SWV}` are automatically merged into one entry with consecutively renumbered `cycle number`; `time/s` preserved verbatim (EC-Lab records absolute time); CA/OCV/EIS etc. always loaded individually; post-load dialog lists merged groups; helpers: `_read_one_df`, `_make_file_entry`, `_merge_dfs`, `_unique_short`
2. **Axis selectors + unit dropdowns** — X and Y each have a column selector and a unit combobox with dimension-aware filtering (I/E/t/J families)
3. **J (current density) column** — virtual column; requires all files to have area > 0; computes I/area per file at plot time; density unit range: A/cm², mA/cm², µA/cm², nA/cm²
4. **Unit conversion** — `_get_axis_unit_scale(col, target)` returns `(scale_factor, display_label)`; label format is `col (unit)` e.g. `I (mA)`, `Ewe (V)`
5. **"(vs Ref)" logic** — appended to axis label only when the column/unit is a voltage type (V/mV/µV/nV); applies to both X and Y axes
6. **IR correction** — `E_corrected = E_raw − (I_mA / 1000) × R_sol`
7. **RHE conversion** — `E_RHE = E_measured + E_ref_vs_RHE`
8. **Reference electrode selector** — appended to axis label as `(vs Ag/AgCl)` etc.
9. **Auto-replot** — updates on cycle selection, correction, file switch; suppressed during load
10. **Plot range** — blank = auto; triggers replot on Return / FocusOut
11. **Cycle checkboxes** — 9-column grid, scrollable, "Select All / Deselect All"
12. **Legend controls** — show/hide, frame toggle, font size, location, "Edit Labels" dialog (blocking; drag disabled during edit), drag to move, right-drag to resize; **double-click on legend** opens the same editor; dialog supports both rename and ⠿ drag-handle reorder; text-only edits preserve drag position, reorder recreates the legend
12b. **Legend resize — full live scaling** — `_scale_legend_spacing(leg, ratio)` in `plotting.py` (module-level, imported by all panels); walks the legend's internal box tree; scales `DrawingArea.width/height` (handle icon containers) and the artists inside them (`Line2D` xdata/ydata/markersize, `Rectangle` geometry), plus `sep`/`pad` on all `VPacker`/`HPacker` nodes; called on each `_on_motion` event with ratio = `new_sz / prev_sz`; result: text, handle shapes, and spacing all scale proportionally in real time so the legend box size matches what `ax.legend(fontsize=new_sz)` would produce — no jump on next replot
12c. **Legend size preservation** — in Multi E.Chem panels, `_on_release` (when `was_resizing=True`) syncs the new `leg_size` back to `legend_size_var` for the active file/group so the next `_plot_file`/`_plot_group` call reads the updated size instead of the stale UI value
13. **Per-file zoom/pan preservation** — switching files restores each file's last view state
20. **Flip X / Flip Y** — `x_flip_var` / `y_flip_var` BooleanVars; applied at end of `_apply_axis_range()` via `ax.set_xlim/ylim` reversal; fire `_auto_replot` on toggle
21. **Swap X↔Y** — `_swap_xy` closure swaps x_var↔y_var, x_unit_var↔y_unit_var, x_min/max↔y_min/max, x_flip↔y_flip; suppresses first `_refresh_unit_opts` replot to avoid double replot
14. **Mouse interactions** (PlottingMixin): scroll = zoom, left-drag = pan, left-click = annotate (switches active file), right-click = dismiss
15. **Reference lines** — add X (vertical) or Y (horizontal) dashed guide lines at typed values; each line has its own style (dashed/dotted/solid/dash-dot) and color; managed via listbox + Remove; selecting a line loads its style/color into the dropdowns for individual editing; panel-level (shared across all overlaid files)
16. **Excel export** — active file only; "Raw" and "Corrected" sheets, cycles side-by-side
17. **File colors** — each file auto-assigned a distinct base color from a Tab10-like palette on load; overridable via "Color:" combobox in the left panel; stored in `entry["color"]`
17b. **Per-file line width** (General tab only) — "Width:" entry in the Color row; changing and pressing Return/FocusOut saves to `entry["linewidth"]` via `_on_linewidth_change()` and replots; `_switch_active_file` restores value; `plotting.py` reads `entry.get("linewidth", "3")` per file inside the plot loop
17c. **Per-file plot shape** (all tabs) — "Shape:" combobox in the Color/Width row; 13 options: Line, Line+Dot, Line+Circle, Line+Star, Line+Square, Line+Triangle, Line+Diamond, Dot, Circle, Star, Square, Triangle, Diamond; stored as `entry["plot_style"]`; maps to `(linestyle, marker, markersize)` via `_PLOT_STYLES` dict in `file_manager.py`; `_on_plot_style_change()` saves immediately + replots; EIS default = "Line+Circle", others = "Line"
18. **Cycle color gradient** — when Gradient is checked, cycles within a file are tinted from lightest (first) to darkest (last) so evolution is easy to track; "Reverse" flips the order; "Step" spinbox (0.01–0.30) controls lightness delta per cycle; each file stores its own gradient settings independently
19. **Editable plot title** — default text is **blank**; type in the left-panel title entry or double-click anywhere in the title strip above the plot; both update in sync; persists across replots
20. **Editable axis labels** (General tab) — double-click the X or Y axis label text on the plot to rename it via a dialog; entering a blank string reverts to the auto-generated label (column + unit); custom label persists across parameter changes until explicitly cleared; stored as `_custom_xlabel` / `_custom_ylabel` on the panel instance
21. **Label/title spacing** — "Spacing (pt): Title [__] Label [__]" row in the Font section (all tabs); Title pad controls gap between axes frame and title (default 6 pt); Label pad controls gap between tick numbers and axis labels (default 4 pt); stored as `title_pad_var` / `label_pad_var`; applied via `ax.set_title(..., pad=N)` and `ax.set_xlabel(..., labelpad=N)`
22. **Plot size controls (all tabs)** — **W [__] H [__] inches** fields in the Font/Size section; resize the matplotlib figure and the canvas widget (`fig.set_size_inches(w, h)` + `canvas.get_tk_widget().config(width=int(w*dpi), height=int(h*dpi))`); maximum 50 inches in either direction; right panel is a scrollable `tk.Canvas` (`_plot_sc`) with both horizontal and vertical `ttk.Scrollbar`s; `_apply_plot_size()` updates `scrollregion` via `after(50, ...)` after `draw_idle()`; defaults: General E.Chem W=21.0/H=12.5, ECSA W=21.0/H=6.0, Nyquist W=21.0/H=12.5, Multi E.Chem 1&2 W=10.5/H=5.5; ECSA `_apply_plot_size()` iterates over both `(fig_cv, canvas_cv)` and `(fig_cdl, canvas_cdl)`; reference stored as `self._plot_sc`
23. **Copy to clipboard** — "Copy" text button added next to each tab's toolbar (all tabs); copies the current figure to the Windows clipboard as CF_DIB so it can be pasted into Word/PowerPoint; requires Pillow; previous clipboard crash fixed by setting correct `restype=ctypes.c_void_p` on all 64-bit Win32 API calls

### Multi E.Chem tab
1. **One plot per file** — each loaded file gets its own figure with a blue-gray ⠿ drag-handle header strip; CheckableListbox checkbox hides/shows a file's subplot in the grid without losing settings; zoom/pan state is preserved across hide/unhide; both the file list ⠿ handle and the subplot header strip support drag-to-reorder (file list fires `_on_file_reorder` which calls `_relayout_figures`; subplot drag calls `_on_frame_press/drag/release` directly)
2. **Per-file settings** — axis columns, units, plot range, reference electrode, IR/RHE correction, cycle selection, legend options all independent per file
3. **J (current density) column** — per-file check: "J" added to combos only when active file has area > 0
4. **Click-to-select** — clicking any plot or its toolbar selects that file in the listbox and updates left controls
5. **Per-file zoom/pan preservation** — `_plot_file` captures view before `ax.clear()` and restores after auto-scale; first plot still auto-scales; Home toolbar resets to `auto_xlim`/`auto_ylim`
5b. **Gold header on selection** — active file's header turns gold (`#ffd54f`); others stay blue-gray; `_highlight_active_header()` called on `_switch_active_file`; auto-scrolls to the selected file's subplot
5c. **Legend label stable-key system** — `entry["legend_labels"]` is now a `dict` keyed by `"{short}:C{c}"` (or `short` for non-cycle files); labels survive cycle add/remove without positional mismatch
6. **Shared axis/unit UI** — left controls show settings of the currently active (selected) file only
7. **Legend controls** — per-file show/hide, frame, size, location; legend draggable and font-resizable; "Edit Labels" / double-click-on-legend dialog (blocking; drag disabled during edit); supports rename + ⠿ drag-handle reorder
12. **Flip X / Flip Y** — per-file `x_flip`/`y_flip` saved in entry dict; active file uses vars, non-active files read from `entry.get("x_flip/y_flip", False)` during full relayout
13. **Swap X↔Y** — same `_swap_xy` pattern as General tab; per-file state saved/restored
8. **Reference lines** — per-file X/Y guide lines; each line carries its own style and color; listbox refreshes when switching files
9. **File colors** — per-file base color from palette; overridable via "Color:" combobox
10. **Cycle color gradient** — per-file gradient/reverse/step settings; same UX as General tab
11. **Subplot zoom** — double-click any subplot to expand it to fill the full panel; a "← Back to Grid" bar appears at the top; clicking it restores the 2-column grid; figures remain live (no recreation)
12. **Editable subplot titles** — **Title entry** in the left panel sets `entry["custom_title"]` for the active file; default is blank `""`; also editable by double-clicking the title strip on any subplot; saved in `_save_active_state()`, restored in `_switch_active_file()` via `entry.setdefault("custom_title", "")` + `plot_title_var.set(...)`; `_plot_file()` uses `plot_title_var.get()` for the active file and `entry.get("custom_title", "")` for non-active files; double-clicking in a zoomed view is the same gesture but zoom-toggle takes priority unless the cursor is in the title strip

### Multi E.Chem 2 tab
1. **Group-based overlay** — files assigned to named groups; each group has its own subplot in a 2-column grid; multiple files overlaid per subplot
2. **Drag-to-reorder groups** — group header `cursor="fleur"` + `_on_frame_press/drag/release` → `_reorder_groups()`; `_drop_line` blue bar shows drop target
3. **Per-group settings** — axis columns/units, plot range, legend, grid, font, reference lines, title — all independent per group; active group drives left-panel controls
4. **Per-file settings (within groups)** — color, linewidth, style, area, cycles, IR/RHE correction independent per file; cycle checkboxes update on active-file switch
5. **Gold header on selection** — when a file is selected (listbox or line click), ALL group headers containing that file turn gold (`#ffd54f`); others stay green (`#c8e6c9`); `_highlight_active_headers()` called from `_apply_highlight_to_group()`
6. **Highlight (Origin-style)** — `_plot_highlight` bool + `_active_cycle` int|None; set on file/line click, reset on right-click; `_apply_highlight_to_group()` dims unselected lines, raises active file/cycle with glow shadow; called after legend build so legend handles keep alpha=1.0
7. **Cycle-specific highlight** — clicking a specific cycle line sets `_active_cycle`; only that `(file, cycle)` combination glows; other cycles from the same file are dimmed; listbox select resets `_active_cycle=None` (whole-file)
8. **Per-group zoom/pan preservation** — `_plot_group` captures view before `ax.clear()` and restores after auto-scale; first plot auto-scales; Home toolbar resets to `auto_xlim`/`auto_ylim`
9. **Legend label stable-key system** — `gentry["legend_labels"]` is a `dict` keyed by `"{fname}:C{c}"` (or `fname` for non-cycle); survives file add/remove and cycle selection changes
10. **`line_to_file` + `line_to_cycle`** — dicts rebuilt on every `_plot_group()`; used for click-to-select (file switch) and cycle-specific highlight respectively
11. **Sync listboxes** — clicking a line or selecting from any listbox syncs both the main file listbox and the group-files listbox
12. **Reference lines** — per-group; style/color per line; drawn via `draw_reflines()` after `_apply_group_range()`

### ECSA Calc tab
1. **Independent file state** — fully separate from other tabs
2. **No IR/RHE correction** — section removed; not needed for double-layer capacitance extraction
3. **Axis selectors + unit dropdowns** — dimension-aware; display only; extraction uses raw column values
4. **Cycle checkboxes** — 9-column grid, same UX as General tab
5. **Scan-rate per cycle table** — 8-column grid; each entry has `trace_add("write", …)` triggering debounced (300 ms) CV replot so legend updates as you type
6. **E_std entry** — red dashed vertical line on CV; immediate save to entry dict + replot on Return/FocusOut; **Rec: label** shown in green next to the entry — auto-computed midpoint `(E_max + E_min) / 2` of the actual plotted X-column data for the selected cycles; updates on every `_plot_cv()` call
7. **Cs entry** — specific capacitance (default 0.040 mF/cm²); immediate save to entry dict on Return/FocusOut
8. **Extract Cdl & ECSA** — runs extraction, updates Cdl plot; legend shows fit equation + Cdl + R² + ECSA; results persisted per file for restore on file switch
9. **Per-file zoom/pan preservation** — CV and Cdl views independently saved/restored per file
10. **Two independent toolbars** — each toolbar controls only its own plot; Home button restores auto-scaled limits from the last draw
11. **CV and Cdl plot titles** — include the active filename for easy identification during multi-file analysis
12. **Reference lines** — separate CV and Cdl sections; per-file, per-line style and color; both listboxes refresh on file switch
13. **File colors** — per-file base color from palette; overridable via "Color:" combobox; applies to CV cycles
14. **Cycle color gradient** — per-file gradient/reverse/step settings; gradient applied to CV cycles only; Cdl scatter/fit plot remains fixed colors (steelblue/tomato)
15. **Editable plot titles** — double-click the title strip on either the CV or Cdl plot to rename it
16. **Flip X / Flip Y** — `x_flip_var` / `y_flip_var` on CV plot; applied in `_apply_cv_range()`; fire `_plot_cv()` on toggle
17. **Swap X↔Y** — `_swap_xy` closure swaps all CV axis state; calls `_refresh_unit_opts` + `_plot_cv()`

### Nyquist Plot tab
1. **EIS / impedance data** — loads tab-separated `.txt` files with Re(Z) and -Im(Z) columns; CheckableListbox checkbox hides/shows individual file traces; ⠿ drag handle reorders files
2. **Axis selectors + unit dropdowns** — X and Y each independently configurable
3. **Multi-file overlay** — all loaded files shown on a single Nyquist plot; each file uses its auto-assigned palette color and unique marker shape from `entry["color"]` / `entry["marker"]`
4. **Connect lines toggle** — show/hide connecting line between data points
5. **Show markers toggle** — show/hide point markers
6. **Per-file zoom/pan preservation** — same mechanism as other tabs
7. **Editable plot title** — double-click the title strip to rename
8. **Flip X / Flip Y** — `x_flip_var` / `y_flip_var`; applied in `_apply_range()`; fire `self._plot()` on toggle
9. **Swap X↔Y** — `_swap_xy` closure; no unit-refresh needed (both axes use ohm-family units); calls `self._plot()` directly

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
- Result logged per file; ECSA shown in Cdl plot legend.

## Development Workflow

### First Time Setup (New PC)
```bash
# 1. Clone the repo into your projects folder
cd C:\Users\YourName\PycharmProjects
git clone https://github.com/M-Song-ChE/echem-gui.git

# 2. Create a virtual environment inside the cloned folder
cd echem-gui
python -m venv .venv

# 3. Activate it (Windows)
.venv\Scripts\activate

# 4. Install dependencies
pip install -r requirements.txt
```
Then in PyCharm: `File → Open` → select the `echem-gui` folder (use Open, not New Project).
Set interpreter to `.venv\Scripts\python.exe` via `Settings → Project → Python Interpreter`.

### Resuming Work (After Initial Setup)
```bash
cd echem-gui
.venv\Scripts\activate
git pull origin main        # always pull before starting
```

### Daily Routine
```bash
# Start of session
git pull origin main

# End of session
git add <changed files>
git commit -m "brief description"
git push origin main
```

### Updating Requirements
When you add a new package:
```bash
pip install some-package
pip freeze > requirements.txt   # regenerate the full list
```
Or manually add it to `requirements.txt` with a minimum version, e.g. `some-package>=1.0.0`.
On the other PC after pulling: `pip install -r requirements.txt` picks up the new package.

### .gitignore Notes
The project `.gitignore` already excludes:
- `.venv/` — each PC creates its own; never commit it
- `__pycache__/` and `*.pyc` — auto-generated bytecode
- `.idea/` — PyCharm machine-specific settings
- `.claude/` — Claude Code local session data
- `*.xlsx` — output files; regenerate from the app as needed
- `Thumbs.db`, `.DS_Store` — OS thumbnails/metadata

If you ever accidentally stage a file that should be ignored:
```bash
git rm --cached <file>      # unstage without deleting the local file
```

## Dependencies
- Python standard: `tkinter`, `collections`
- Third-party: `pandas`, `numpy`, `matplotlib`, `openpyxl`, `galvani` (for `.mpr` binary loading; lazily imported — app launches without it, errors only when an `.mpr` file is selected)

## Important Design Decisions
- File loading does NOT auto-replot (preserves current plot)
- Newly loaded files have `selected_cycles = []` — no cycles pre-checked
- **ECSA Calc uses a plain `tk.Listbox`** (no hide/show) — the tab shows only the active file's CV so hiding is not meaningful; the plain listbox avoids zoom-state complications from the shared `ax_cv` axes
- Plot skips files with `cycle number` column but no selected cycles
- Export only exports the active file; blocks if no cycles selected
- `constrained_layout=True` on Figure prevents subplot title/label overlap
- Two separate Figure objects in ECSAPanel (not subplots) gives each plot a fully functional independent toolbar
- J column is a UI-level virtual column, not in the DataFrame; resolved to the actual current column at plot time by searching for a column whose unit suffix is in `_CURRENT_UNITS = {"A","mA","µA","nA"}`

## Known Patterns / Gotchas
- **`_suppress_replot` must save/restore** — use `old = self._suppress_replot` pattern, not hard-set `False`
- **`_loading_files` guard** — `selection_set()` fires `<<ListboxSelect>>` synchronously on Windows; wrap with `_loading_files = True/False` and early-return in `_on_file_select`
- **`_clear_plot` hook** — `FileManagerMixin._remove_file` calls `self._clear_plot()` when the last file is removed; base implementation is a no-op; override in each panel to clear canvases
- `canvas.draw()` (not `draw_idle()`) needed for legend resize to show frame changes in real-time
- `set_draggable(True)` called after every `ax.legend(...)` call; old legend ref becomes stale after `ax.clear()` so reset to `None` before clearing
- **Legend must not affect axes layout** — `_apply_font_to_ax` is called after the legend is placed; it calls `tight_layout()` which would otherwise shrink the axes to fit an overflowing legend. Fix: temporarily hide the legend (`_leg.set_visible(False)`) around every `tight_layout()` call in `_apply_font_to_ax`, then restore it. Applied in all four panels (`app.py`, `eis_panel.py`, `ecsa_panel.py`, `multi_echem_panel.py`).
- Toolbar Home button override requires subclassing `NavigationToolbar2Tk` (attribute assignment does not work — command is bound at init time)
- **Two file formats supported**: BioLogic `.mpr` binary (via `galvani`) and tab-separated `.txt` (via `pd.read_csv`); both can be mixed in the same session. Column names normalized on load: whitespace stripped, `<`/`>` removed (e.g. `<Ewe>/V` → `Ewe/V`). `.mpr` loader additionally filters to `_MPR_DESIRED` columns only.
- **`_clear_annotation` naming differs by panel** — EchemPanel (PlottingMixin) uses `_clear_annotation(redraw=False)`; ECSAPanel uses `_ei_clear_ann(redraw=False)`. Both must be called **before** `ax.clear()`.
- **Axis label format** — `col (unit)` e.g. `I (mA)`, `time (ms)`; auto case converts the column name's own `/` separator to the same format (e.g. `I/mA` → `I (mA)`)
- **`_pan_moved`** reset to `False` on every press, set `True` on actual motion; gates annotation on release
- Scan rate `StringVar` traces accumulate if not removed — `_rebuild_sr_table` calls `var.trace_remove()` for all previous trace IDs before rebuilding
- **Unit scale method name differs** — PlottingMixin: `_get_axis_unit_scale(col, target)`; ECSAPanel and MultiEchemPanel: `_get_unit_scale(col, target_unit)` (same logic, different name)
- **View preservation timing** — in ECSAPanel, Cdl view is restored immediately after `_replot_cdl()`, CV view after `_auto_replot()`; order matters since both draw to canvas
- **Area var in file_manager** — `_save_active_state` and `_switch_active_file` handle `area_var` via `getattr(self, "area_var", None)` so panels without it are unaffected
- **CorrectionMixin column names** — `_apply_correction` looks for `"Ewe/V"` and `"I/mA"`; silently no-ops if columns absent (since recent cleanup of correction.py)
- **`open_legend_editor` returns the legend object** — if the entry order changed, the legend is recreated via `ax.legend(handles, labels, ...)` and the new object is returned; if only text changed, the original object is returned (Text objects updated in-place to preserve drag position). All callers must assign the return value: `self._legend_obj = open_legend_editor(...)`. The old legend ref is stale after recreation. The dialog uses ⠿ drag handles (not ↑/↓ buttons) — `_on_press/_on_drag/_on_release` handlers, `drop_line` indicator, same pattern as `CheckableListbox`.
- **`open_legend_editor` must be blocking** — uses `dlg.grab_set()` + `parent.wait_window(dlg)`; without `wait_window`, the function returns immediately and matplotlib's `DraggableLegend` handler stays in "dragging" state (never receives button_release); always call `legend.set_draggable(False)` before opening and re-enable after
- **Legend label stable-key system** — General tab: `_legend_stable_map = {"{short}:C{c}": custom_label}`; `_legend_stable_keys = []` rebuilt on every `_plot()`; keys always file-qualified so they survive single↔multi-file format changes. Multi E.Chem 1: `entry["legend_labels"]` is a `dict` keyed by `"{short}:C{c}"` (or `short`); `_edit_legend_labels` reconstructs keys from `entry["selected_cycles"]` + `self.active_file`. Multi E.Chem 2: `gentry["legend_labels"]` is a `dict` keyed by `"{fname}:C{c}"` (or `fname`); save/restore iterates group files in plot-loop order, skipping empty sub-DataFrames. Do NOT use positional lists — they break when cycles are added/removed or when two files share cycle numbers. `_legend_auto_labels` is captured after `ax.legend()` for change-detection in the General tab edit dialog.
- **Legend location dropdown override** — `_on_leg_loc_select` must set `self._legend_obj._loc = 0` (non-tuple) in addition to clearing `_legend_manual_pos = None`; otherwise `_plot()` reads the still-alive tuple from the old legend object before `ax.clear()` and re-saves it, defeating the dropdown selection
- **`draw_reflines` tuple format** — each entry is a 4-tuple `('x'|'y', float, style, color)`; style is a key into `_GRID_STYLE_MAP`; labels start with `'_'` so they are excluded from the legend automatically; call after `_apply_axis_range()` / `_apply_range()` so reflines don't perturb autoscaling; call before `apply_grid()` / `canvas.draw()`
- **ECSAPanel `_auto_xlim_cdl` / `_auto_ylim_cdl`** — only set after `canvas_cdl.draw()` completes inside `_replot_cdl` and `_extract_cdl_ecsa`; if either function crashes before that point, the reset-view button will silently do nothing (value stays `None`)
- **Zoom preservation in `_plot_file`/`_plot_group`** — capture `_prev_view = (ax.get_xlim(), ax.get_ylim())` before `ax.clear()` (skip if `auto_xlim is None` = first plot); restore after `canvas.draw()` saves `auto_xlim`; then call `_apply_range()` which overrides with any user-entered min/max values. This order means: toolbar Home → `auto_xlim` restored → becomes next `_prev_view`; user-entered range → applied after restore → overrides zoom. Do NOT save/restore `view_xlim` in `_save_active_state` for these panels; the plot function handles it directly.
- **Highlight legend integrity** — all lines drawn at `alpha=1.0` in insertion order; `ax.legend()` called → proxy handles created at alpha=1.0; THEN `_apply_highlight_to_axes()`/`_apply_highlight_to_group()` dims non-active lines. This ensures legend symbols always stay at full alpha regardless of highlight state.
- **Highlight glow pattern** — draw `ax.plot(..., linewidth×2.5, alpha=0.18, label='_glow', zorder=1.9)` for each highlighted line; `label='_glow'` (starts with `_`) auto-excludes from legend; remove all `label='_glow'` lines before redrawing to avoid stale glows after file switch.
- **`_active_cycle` scope** — shared panel-level attribute; reset to `None` on listbox select, right-click, or group-file-listbox select; set to `int` when clicking a cycle line; used by both `_apply_highlight_to_axes` (General) and `_apply_highlight_to_group` (Multi 2) to narrow glow to a single `(file, cycle)` pair.
- **`_switch_active_file` UI-restore ordering (critical)** — `FileManagerMixin._switch_active_file` sets `self.active_file = short` and then calls `_auto_replot()`; `_auto_replot` → `_plot/_plot_file` → `_save_active_state` will immediately write the current UI var values into `self.files[short]` (the new file). If UI vars still hold the old file's values at that moment, the new file's settings are clobbered. **Fix:** always restore per-file UI vars (color, gradient, etc.) **before** calling `super()._switch_active_file()` in any panel override.
- **`_cycle_colors(base_color, n, step, reverse)`** (module-level in `plotting.py`) — converts the base color to HLS, offsets lightness linearly across `n` cycles. `reverse=False` → first cycle lightest, last darkest (most recently evolved = most visible). `reverse=True` flips. Clamps lightness to [0.15, 0.85]. Uses `colorsys` + `matplotlib.colors`; returns a list of `(r, g, b)` tuples.
- **`file_manager` palette constants** — `_COLOR_NAMES`, `_COLOR_HEX`, `_PALETTE`, `_MARKERS`, `_PLOT_STYLES`, `_PLOT_STYLE_NAMES` defined at module level; imported by all panel files. `_PLOT_STYLES` maps style name → `(linestyle, marker, markersize)` for the 13 plot shape options.
- **Per-file plot style pattern** — `entry["plot_style"]` stores the style name; `_on_plot_style_change()` saves immediately + replots; `_save_active_state` also saves; `_switch_active_file` restores via `entry.get("plot_style", "Line")`; `plotting.py` reads `entry.get("plot_style", "Line")` inside the file loop and unpacks `_PLOT_STYLES[name]` to `(_ls, _mk, _ms)`; marker arg uses `_mk or None` so empty string → None (no marker)
- **Per-file linewidth (General tab)** — `entry["linewidth"]` per file; `_on_linewidth_change()` saves immediately; `plotting.py` reads `entry.get("linewidth", "3")` inside the file loop (not before it); EIS also saves/restores per-file linewidth via `_on_linewidth_change()` / `_switch_active_file`
- **General tab defaults** — linewidth=3, legend_size=20, title_size=40, label_size=30, tick_size=20, spacing=20, grid X/Y=True, grid color=black, grid width=2, cycle step=0.15 (all tabs)
- **Cycle gradient defaults (all tabs)** — `cycle_reverse=False`, `lightness_step="0.15"` — set in `file_manager.py` on load and as fallback defaults throughout all panels
- **`_default_xcol(cols)` / `_default_ycol(cols, x_col)`** — module-level helpers in `file_manager.py`; data-type-aware: detect **EIS** (impedance columns present) → X=Re(Z), Y=-Im(Z); detect **OCV/time-series** (time present, no current) → X=time, Y=voltage; detect **CV/LSV** (current present) → X=voltage, Y=current; generic fallback to second column. Used by base `_switch_active_file` and imported by Multi/ECSA panels. Companion predicates: `_is_impedance_col`, `_is_voltage_col`, `_is_current_col`, `_is_time_col`.
- **`FileManagerMixin._get_column_list(df)`** — returns only EIS columns (`_is_impedance_col`) when EIS data detected, hiding time/voltage/current/cycle-number metadata from the axis comboboxes. For non-EIS data returns all columns. `EchemPanel` overrides this (calls `super()`) to append the virtual "J" column for non-EIS data only.
- **`_UNIT_DIMS` / `_DIM_OPTS` EIS extensions** — added to both `app.py` and `multi_echem_panel.py`: `"Ohm"/"Ω"/"mΩ"/"kΩ"/"MΩ" → "Z"`, `"Hz"/"kHz"/"MHz" → "f"`, `"deg"/"rad" → "φ"`; `_DIM_OPTS["Z"]` = mΩ/Ω/kΩ/MΩ, `"f"` = Hz/kHz/MHz, `"φ"` = deg/rad. Same factors added to `_FACTORS`/`_DIMS` inside `_get_axis_unit_scale` (plotting.py) and `_get_unit_scale` (multi_echem_panel.py).
- **`_read_mpr` unknown column retry loop** — galvani raises `NotImplementedError("Column ID N after column X is unknown")` for column IDs added in newer EC-Lab firmware. `_read_mpr` catches this, extracts the column ID via regex, injects a placeholder entry into `BioLogic.VMPdata_colID_dtype_map`, and retries. Tries element sizes `<f4`, `<f8`, `<u4`, `<u2` in outer loop; catches `ValueError("buffer size must be a multiple…")` or `AssertionError` to detect wrong size and advance to next. Placeholder column names (`_unknown_N`) are absent from `_MPR_DESIRED` so they are silently filtered out.
- **`CheckableListbox` hide/show guard** — `_on_file_visibility_change` is defined in `FileManagerMixin` (calls `_auto_replot`) and overridden in `MultiEchemPanel` (also snapshots/restores per-file zoom); ECSA Calc does NOT override it and does NOT use `CheckableListbox`, so `hidden` flag is never set in that panel
- **`_on_file_reorder` pattern** — base in `FileManagerMixin`: rebuild `self.files` as a new `OrderedDict` in new order (iterate `new_order`, fallback for any missing names), then call `_auto_replot()`. Override in `MultiEchemPanel`: same rebuild but call `_relayout_figures()` instead (figure objects already exist; no need to replot all data).
- **Flip X/Y implementation** — after `ax.set_xlim/ylim(lo, hi)`, check `xl = ax.get_xlim(); xl[0] > xl[1]` to detect current direction; toggle with `ax.set_xlim(xl[1], xl[0])` only when `flip_var.get() != (xl[0] > xl[1])`. In Multi E.Chem, non-active files read from `entry.get("x_flip/y_flip", False)` instead of the UI var.
- **Swap X↔Y double-replot prevention** — `_refresh_unit_opts` calls `_auto_replot` internally; set `_suppress_replot = True` before first call and `False` before second so only the second triggers a replot
- **Multi E.Chem zoom preservation on hide/unhide** — snapshot is taken from `entry["ax"].get_xlim/ylim()` only when `not ae.get("hidden")` to avoid clobbering the saved zoom with the 0–1 range that appears when axes are cleared
- **Title dblclick detection** — `PlottingMixin._hit_title_area(event, ax, fig)` static method checks both: (a) `ax.title.get_window_extent(renderer).contains(event.x, event.y)` for when title text is visible, and (b) the horizontal strip `ax_bbox.y1 ≤ event.y ≤ fig_bbox.y1` for when title is empty; check is performed **before** the `event.inaxes` guard since the title strip is outside the axes bounding box
- **Multi E.Chem zoom bar placement** — `right_outer` uses `grid` manager (not `pack`) so the zoom bar row reliably collapses to zero height via `grid_remove()` and appears at the top before the canvas row; mixing `pack` and `grid` on siblings of the same parent is an error in tkinter
- **Multi E.Chem zoom `columnspan` reset** — `grid(columnspan=2)` during zoom mode persists until explicitly overridden; `_relayout_figures` must pass `columnspan=1` when restoring the normal 2-column grid or files appear merged
- **ECSAPanel `_switch_active_file` calls `_plot_cv()` directly** — `_load_files` sets `_suppress_replot=True` before calling `_switch_active_file`, so calling `_auto_replot()` inside `_switch_active_file` would silently skip the CV redraw (old curves would remain). The Cdl plot was already handled by direct `ax_cdl.clear()` calls. CV now also uses a direct `_plot_cv()` call at the end of `_switch_active_file`, bypassing the suppress guard. This is intentional: switching to a new file must always refresh both plots regardless of suppress state.
