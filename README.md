# Echem GUI — User Manual

## Table of Contents
1. [Getting Started](#1-getting-started)
2. [Interface Overview](#2-interface-overview)
3. [General E.Chem Tab](#3-general-echem-tab)
4. [Multi E.Chem Tab](#4-multi-echem-tab)
5. [Multi E.Chem 2 Tab](#5-multi-echem-2-tab)
6. [ECSA Calc Tab](#6-ecsa-calc-tab)
7. [OCV/Ru Extractor Tab](#7-ocvru-extractor-tab)
8. [ORR Analysis Tab](#8-orr-analysis-tab)
9. [Hupd Calc Tab](#9-hupd-calc-tab)
10. [Common Controls (All Tabs)](#10-common-controls-all-tabs)
11. [Tips and Shortcuts](#11-tips-and-shortcuts)
12. [Session Save & Restore](#12-session-save--restore)

---

## 1. Getting Started

### Running the App
- **From the exe:** Double-click `EchemGUI.exe` inside the `EchemGUI` folder.
- **From Python:** Run `python run_echem.py` in the project folder.

### Supported File Formats
The app reads two file types:
- **BioLogic `.mpr` binary files** — loaded directly from EC-Lab without any manual export step. Requires the `galvani` package (`pip install galvani`). Files created with newer EC-Lab firmware versions that contain unrecognised column types are handled automatically — those extra columns are skipped and the recognised data loads normally.
- **Tab-separated `.txt` files** — as exported by EC-Lab (BioLogic) and similar potentiostats.

Column names are recognized automatically (e.g. `Ewe/V`, `I/mA`, `time/s`, `cycle number`, `Re(Z)/Ohm`, `-Im(Z)/Ohm`). Both file types can be mixed within the same tab and session.

---

## 2. Interface Overview

The app has six tabs at the top:

| Tab | Purpose |
|-----|---------|
| **General E.Chem** | Overlay multiple files on one plot; full correction and export |
| **Multi E.Chem** | View one plot per file simultaneously in a 2-column grid |
| **Multi E.Chem 2** | Group files into named groups; overlay all groups on one plot each |
| **ECSA Calc** | Extract electrochemical surface area (ECSA) from CV data |
| **OCV/Ru Extractor** | Group OCV + EIS files per sample; auto-extract stable OCV and Ru; one Nyquist + one OCV plot per sample |
| **ORR Analysis** | Background-subtracted RDE polarization curves (N2/O2, per RPM, per sample) |
| **Hupd Calc** | Hupd-based ECSA calculation from CV data — Q_H integration, ECSA, and roughness factor |

Each tab is fully **independent** — files loaded in one tab are not shared with others.

The layout in every tab is:
- **Left panel** — scrollable controls (file list, axis settings, correction, cycles, etc.)
- **Right panel** — the plot area (scrollable; size controlled by W/H fields)

---

## 3. General E.Chem Tab

Use this tab to load one or more data files and overlay them all on a single plot.

### 3.1 Loading Files
1. Click **Load File(s)** to open a file browser. Select one or more `.mpr` or `.txt` files.
2. Loaded files appear in the **file list**. Each row has a **checkbox**, a **⠿ drag handle**, and a filename label.
   - **Click the filename** to make it the active file — all left-panel controls reflect that file's settings.
   - **Uncheck the checkbox** to hide that file's curves from the plot without removing it. All settings (cycles, corrections, zoom, colors, etc.) are fully preserved. Re-check to bring it back instantly.
   - **Drag the ⠿ handle** up or down to reorder files in the list. The plot updates to reflect the new order.
3. To remove a file, select it and click **Remove**.

#### Auto-merging sequential CV files
When you select multiple files at once, the app automatically detects EC-Lab CVA sequence naming and merges matching groups for you:

- **Pattern recognised:** `..._NN_METHOD_Cxx.mpr / .txt`
  e.g. `sample_05_CV_C01.mpr`, `sample_07_CV_C01.mpr`, `sample_09_CV_C01.mpr`
- Files sharing the same base name, method, and channel but differing only in the sequence number are merged into a single entry: `sample_05-09_CV_C01_merged.mpr`
- **Cycle numbers** are renumbered consecutively (cycles 1–2 from the first file stay 1–2; cycles 1–2 from the second become 3–4; etc.).
- **time/s** is kept exactly as recorded by EC-Lab — no modification.
- A dialog appears after loading, listing every group that was auto-merged.
- Only voltammetry methods are merged automatically (**CV, CVA, LSV, DPV, NPV, SWV**). CA, OCV, EIS, and other techniques are always loaded as individual files.

### 3.2 Axis Settings
- **X / Y column selectors** — choose which data column to plot on each axis. The available columns are filtered by data type: if the file contains EIS impedance data, only the impedance-related columns (`Re(Z)/Ohm`, `-Im(Z)/Ohm`, `freq/Hz`, `Phase(Z)/deg`) are shown; for CV/OCV files all columns are shown.
- **Smart defaults** — when a file is first loaded the app automatically selects sensible defaults based on data type:
  - **EIS** (impedance data): X = `Re(Z)/Ohm`, Y = `-Im(Z)/Ohm`
  - **OCV / time-series** (voltage + time, no current): X = `time/s`, Y = `Ewe/V`
  - **CV / LSV** (voltage + current): X = `Ewe/V`, Y = `I/mA`
- **Unit dropdowns** — select the display unit. The available units match the column type:
  - Voltage: V, mV, µV, nV
  - Current: A, mA, µA, nA
  - Time: s, ms, µs, min, h
  - Impedance: mΩ, Ω, kΩ, MΩ
  - Frequency: Hz, kHz, MHz
  - Phase: deg, rad
- **J (current density)** — if you enter a positive electrode area (cm²) for each file, a virtual "J" column appears in the column dropdowns. Selecting it plots current density (I ÷ area).

### 3.3 Corrections
- **Reference electrode** — select your reference (RHE, Ag/AgCl, SCE, NHE, MSE). The axis label updates to show `(vs Ref)`.
- **E_ref (V)** — potential offset for RHE conversion: `E_RHE = E_measured + E_ref`.
- **R_sol (Ω)** — uncompensated resistance. IR correction: `E_corr = E − (I_mA / 1000) × R_sol`.
- Click **Apply Correction** to apply; the plot updates immediately.

### 3.4 Cycle Selection
Files with a `cycle number` column show a grid of cycle checkboxes.
- Check individual cycles to include them in the plot.
- Use **Select All / Deselect All** for quick selection.
- Files with no cycles selected are skipped on the plot (no error).

### 3.5 Plot Range
- Enter values in the **X min / X max / Y min / Y max** boxes to fix the axis range.
- Leave blank for automatic scaling.
- Press **Enter** or click away to apply.

### 3.6 Axis Orientation
- **Flip X / Flip Y** — check to invert the direction of that axis (e.g. display potential decreasing right-to-left, or current density going downward).
- **⇄ Swap X↔Y** — swaps the X and Y axes in one click: column, unit, range limits, and flip state all exchange simultaneously.

### 3.7 File Colors and Cycle Gradient
- **Color combobox** — change the base color of the active file. Each file is auto-assigned a distinct color on load.
- **Gradient checkbox** — when checked, cycles within a file are drawn with a lightness gradient: the first cycle is the lightest and the last cycle is the darkest (most evolved cycle = most visible).
- **Reverse checkbox** — flips the gradient direction (first cycle darkest, last lightest).
- **Step spinbox** — controls how much the lightness changes between consecutive cycles (range 0.01–0.30; default 0.15).
- These settings are saved **per file** — switching files restores each file's individual settings.

### 3.8 Legend
- **Show Legend checkbox** — toggles the legend on/off.
- **Frame checkbox** — adds a border around the legend box.
- **Font size** — adjust legend text size.
- **Location** — choose legend anchor position, or drag the legend freely with the mouse.
- **Edit Labels** — opens a dialog to **rename and reorder** legend entries. Drag the **⠿** handle on any row to reorder. Renaming without reordering preserves any drag position; reordering recreates the legend in the new order.
- **Double-click the legend** directly on the plot to open the same edit dialog instantly.
- **Legend order** — entries follow file rank order (rank-1 file first) with cycles in ascending numeric order (1, 2, 3…) within each file. Reordering files in the file list resets the legend to the new rank order. Custom order set in the editor is preserved until the file rank changes.

### 3.9 Reference Lines
Add horizontal or vertical guide lines to the plot:
1. Type a value in the **X =** or **Y =** entry box.
2. Choose a **line style** (dashed, dotted, solid, dash-dot) and **color**.
3. Click **Add X Line** or **Add Y Line**.
4. To edit or remove a line, select it in the listbox — its style/color loads into the dropdowns. Click **Remove** to delete it.

Reference lines in this tab are shared across all overlaid files.

### 3.10 Excel Export
Click **Export to Excel** to save the active file's data to an `.xlsx` file with two sheets:
- **Raw** — original data, one column group per selected cycle.
- **Corrected** — IR/RHE corrected data.

### 3.11 Plot Title and Axis Labels
- **Plot title** — the title entry in the left panel is blank by default. Edit it directly in the entry field; the plot updates on Enter or focus change. You can also **double-click** anywhere in the title strip above the plot.
- **Axis labels** — double-click the **X axis label** or **Y axis label** text on the plot to rename it. Entering a blank string reverts to the auto-generated label. Custom labels persist across unit/column changes until explicitly cleared.

### 3.12 Font Spacing and Plot Size
In the Font/Size section of the left panel:
- **Spacing (pt): Title [__] Label [__]** — adjust the gap (in points) between the top of the axes and the title (default 6), and between the tick numbers and the axis labels (default 4). Press Enter or click away to apply.
- **W [__] H [__] inches** — set the figure width and height in inches. The plot canvas resizes immediately and a horizontal scrollbar appears if the figure is wider than the panel. Default is W=21.0, H=12.5. Maximum is 50 inches in either direction.

---

## 4. Multi E.Chem Tab

Use this tab to view all loaded files **side by side** in a 2-column grid, each with its own independent plot.

### 4.1 Loading and Selecting Files
- Load files the same way as in the General tab.
- Each row in the file list has a **checkbox** and a **⠿ drag handle**.
  - **Uncheck** to collapse that file's subplot from the grid; re-check to restore it. All per-file settings and zoom state are preserved.
  - **Drag the ⠿ handle** to reorder files; the subplot grid rearranges to match.
- Click any subplot (or its toolbar) to make that file **active** — the left panel updates to show that file's settings.
- Each subplot has a **⠿ header strip** at the top showing the filename. Drag this strip to reorder subplots within the grid.

### 4.2 Per-File Settings
Every control in the left panel (axis columns, units, range, cycles, correction, colors, gradient, legend, reference lines) applies **only to the active file**. Each file remembers its own settings independently.

### 4.3 Plot Title
- The **Title** entry in the left panel sets the title of the active file's subplot. Default is blank.
- You can also **double-click** the title strip on any subplot to rename it inline.
- The custom title is saved per file and persists across replots and file switches.

### 4.4 Subplot Zoom
- **Double-click the ⠿ header strip** of any subplot to expand it to fill the entire right panel.
- Double-click the header strip again, or click the **← Back to Grid** button at the top, to return to the 2-column grid.
- Figures remain live during zoom — no data is lost.

### 4.5 Axis Orientation
- **Flip X / Flip Y** — invert the direction of either axis.
- **⇄ Swap X↔Y** — swap the X and Y axes (column, unit, range, and flip state) in one click.
- These settings are saved **per file**.

### 4.6 Colors and Cycle Gradient
Same controls as the General tab, but settings are per-file. Switching files restores that file's color and gradient settings.

### 4.7 Plot Size
- **W [__] H [__] inches** — set the figure width and height for all subplots. Default is W=10.5, H=5.5. Maximum is 50 inches.
- The scrollable right panel adjusts automatically; use the horizontal and vertical scrollbars to navigate when the grid is larger than the visible area.

---

## 5. Multi E.Chem 2 Tab

Use this tab to organise files into named **groups**, where each group produces one overlay plot (all files in the group drawn on the same axes). Useful for comparing treatment conditions, electrode sets, or replicates side by side.

### 5.1 Group Management
- **Create a group** — type a group name in the entry box at the top and click **New Group** (or press Enter).
- **Add files to a group** — load files first, select one or more in the Loaded Files list, then click **↓ Add Selected Files to Group**.
- **Remove files from a group** — select a file in the Files in Selected Group list and click **↑ Remove Selected from Group**.
- **Rename / remove a group** — select the group name and use the rename/remove controls.
- Groups are shown in a **CheckableListbox** on the left:
  - **Uncheck** a group to hide its subplot from the grid without deleting it. Re-check to restore.
  - **Drag the ⠿ handle** to reorder groups; the subplot grid rearranges to match.
  - **Click the group name** to make it the active group — the left panel updates to show that group's settings.
- The **Files in Selected Group** list also uses checkboxes and drag handles:
  - **Uncheck** a file to hide it from that group's plot only. It remains in other groups unaffected.
  - **Drag the ⠿ handle** to reorder files within the group (controls overlay draw order and legend rank order; legend resets to the new rank when file order changes).

### 5.2 Per-Group Settings
Every control in the left panel (axis columns, units, range, reference electrode, legend, reference lines, font sizes) applies **only to the active group**. Each group remembers its own settings independently.

### 5.3 Per-File Settings Within a Group
Cycle selection, IR compensation (R_sol), and RHE conversion (E_ref) are stored **independently per file per group**. Adding the same file to two groups and changing its cycles in group 1 does not affect group 2. Color, line style, and gradient are global per file (same appearance in all groups).

### 5.4 Plot Title
- The **Title** entry in the left panel sets the title of the active group's plot. Default is blank.
- You can also **double-click** the title strip on any group's subplot to rename it inline.
- The custom title is saved per group and persists across replots.

### 5.5 Plot Size
- **W [__] H [__] inches** — set the figure size for all group plots. Default is W=10.5, H=5.5. Maximum is 50 inches.

---

## 6. ECSA Calc Tab

Use this tab to extract the **electrochemical surface area (ECSA)** from cyclic voltammetry data using the double-layer capacitance (Cdl) method.

### 6.1 Overview of the Layout
The right panel has **two stacked plots**:
- **Upper plot (CV)** — the cyclic voltammetry curves for selected cycles.
- **Lower plot (Cdl)** — the linear fit of scan rate vs. Δj/2 used to extract Cdl and ECSA.

### 6.2 Loading Files and Selecting Cycles
Load files and select cycles the same way as in the General tab. Each cycle should correspond to a different scan rate.

### 6.3 Scan Rate Table
After selecting cycles, a scan rate input table appears:
- Enter the scan rate (mV/s) for each cycle.
- The CV legend updates as you type (no need to press Enter).

### 6.4 Setting ECSA Parameters
- **E_std (V)** — the potential at which Δj is measured (the standard potential in the non-Faradaic region). A red dashed vertical line marks this position on the CV plot.
  - **Rec:** label shown in green next to the entry field — this is the **recommended E_std**, automatically computed as the midpoint of the actual potential range of the currently plotted data: `(E_max + E_min) / 2`. It updates whenever the plot refreshes.
- **Cs (mF/cm²)** — specific capacitance of the material (default: 0.040 mF/cm² for standard electrolyte). Used to convert Cdl to ECSA.

### 6.5 Extracting ECSA
1. Select the cycles, enter scan rates, and set E_std and Cs.
2. Click **Extract Cdl & ECSA**.
3. The Cdl plot shows a scatter of (scan rate, Δj/2) points with a linear fit. The legend shows:
   - The fit equation
   - Cdl (mF)
   - R² (goodness of fit)
   - ECSA (cm²)

### 6.6 ECSA Physics
```
For each selected cycle (one cycle = one scan rate):
  1. Split the CV at the vertex potential into anodic (↑) and cathodic (↓) branches
  2. Interpolate ja at E_std on the anodic branch
  3. Interpolate jc at E_std on the cathodic branch
  4. Δj/2 = (ja − jc) / 2

Linear fit:  scan rate (mV/s)  vs  Δj/2 (mA)
  slope = Cdl [F] → cdl_mF = slope × 1000
  ECSA  = cdl_mF / Cs  [cm²]
```

### 6.7 Axis Orientation (CV Plot)
- **Flip X / Flip Y** — invert the direction of either axis on the CV plot.
- **⇄ Swap X↔Y** — swap the X and Y axes (column, unit, range, and flip state) in one click.

### 6.8 Colors and Cycle Gradient
- The **Color combobox** changes the base color used for CV cycles.
- The **Gradient / Reverse / Step** controls work the same as in the General tab and are per-file.
- The Cdl scatter/fit plot always uses fixed colors (steelblue dots, tomato fit line).

### 6.9 Reference Lines
Separate reference line sections are provided for the CV plot and the Cdl plot. Each has its own listbox and add/remove controls.

### 6.10 Editable Plot Titles
Double-click the title strip on either the CV or Cdl plot to rename it.

### 6.11 Plot Size
- **W [__] H [__] inches** — set the figure size applied to both plots. Default is W=21.0, H=6.0. Maximum is 50 inches.

---

## 7. OCV/Ru Extractor Tab

Use this tab to **bundle each sample's OCV and EIS files together** and let the app automatically extract two numbers per sample:
- **OCV (V)** — the last voltage value from the OCV time-series (stable open-circuit potential).
- **Ru (Ω)** — the real-impedance value at the row where `|Im(Z)|` is closest to zero (restricted to Re(Z) > 0) — i.e. where the high-frequency end of the Nyquist curve crosses the real axis.

### 7.1 Loading Files
1. Click **Load New Sample…** and multi-select **all** OCV and EIS files for one sample (1 or 2 OCV files + any number of EIS files).
2. Files are auto-classified by filename: `OCV` → OCV file, `EIS` / `PEIS` / `GEIS` / `SEIS` → EIS file. Unrecognised files are skipped with a dialog.
3. The app derives the **sample name** from the Bio-Logic-style filename pattern `Pn_TYPE_<name>_<electrode/electrolyte…>`. Example: `P1_OCV1_LTS-BDRDE_04(Pt)_RE3_vs_RHE2_…` → sample name `LTS-BDRDE_04(Pt)`.
4. Use **Add to Selected…** to append more files to the currently active sample.

### 7.2 Loaded Samples List
Each loaded sample appears as a row in the **Samples** list with a checkbox and a `⠿` drag handle:
- **Checkbox** — show/hide that sample's plots (the row disappears from the right panel without losing data).
- **⠿ drag handle** — drag up/down to reorder samples.
- **Click a row** — activates that sample (highlights its plot row, scrolls into view).

### 7.3 Extracted Values Table
Below the sample list, a read-only table lists every sample with columns:

| Sample | OCV (V) | Ru1 (Ω) | Ru2 (Ω) | … |
|--------|---------|---------|---------|---|
| LTS-BDRDE_04(Pt) | 0.3616 | 40.857 | 40.084 | … |

`Ru1` / `Ru2` columns are added dynamically based on the largest EIS-file count across samples. Click a row to activate that sample. **Export CSV** dumps this whole table.

### 7.4 Per-Sample Plots (Right Panel)
Each visible sample gets its **own row** in the right panel containing two independent matplotlib figures side by side:
- **Left** — OCV (Time vs E). A ⭐ marker indicates the last-row OCV value.
- **Right** — Nyquist (Re(Z) vs −Im(Z)). A ⭐ marker indicates the Ru point.

Each axis has its own matplotlib toolbar + **Copy** button.

### 7.5 Plot Interactions (per axis)
- **Scroll wheel** — zoom in / out around the cursor.
- **Left-drag** — pan.
- **Left-click (no drag)** — annotate the curve point nearest the click with `x`, `y`, and label. Click again within a few pixels to cycle through overlapping candidates.
- **Right-click** — clear the annotation.
- **Click anywhere on a sample's row** — that sample becomes active (also reorders the active border highlight and syncs the left-panel selection).

### 7.6 Reordering Plot Rows
Each sample row has a **⠿ + colour box + title** header strip you can drag vertically to reorder rows. A blue drop indicator shows where the row will land. The Samples list, the values table, and the plot order all stay in sync.

### 7.7 Per-Sample Controls
The **Selected sample** section in the left panel lets you change the active sample's **colour** (rotating default palette assigned on load) and shows a file breakdown listbox with each file's filename + extracted value.

Use **Rename…** to give a sample any name you like; use **Remove** to delete it; use **Export CSV** to save the entire values table to disk.

### 7.8 Plot Size
- **W [__] H [__] inches** — figure size **per plot** (each row has an OCV figure W×H and a Nyquist figure W×H side by side). Default is W=7.0, H=4.5. Maximum is 50 inches.

---

## 8. ORR Analysis Tab

Use this tab to compare rotating disk electrode (RDE) oxygen reduction reaction (ORR) performance across multiple samples. Each sample groups N2 (background) and O2 (signal) CV file pairs by RPM. Background subtraction, IR correction, and RHE conversion are applied automatically.

### 8.1 Loading Files
1. Click **Load Files** to select `.mpr` or `.txt` CV files (N2 and O2, any number of RPM values).
2. N2 vs O2 is auto-detected from the filename (looks for `n2`/`o2` as word fragments).
3. The RPM index is auto-extracted from the filename pattern `_NN_CV_` (e.g. `_04_CV_`, `_1600_CV_`).
4. Files appear tagged `(N2)`, `(O2)`, or `(??)` with **color coding**: N2 rows are light blue, O2 rows are light orange — no auto-merge is performed.
5. Use **Sel N2** / **Sel O2** buttons above the list to instantly select all N2 or all O2 files for quick sample assignment.

### 8.2 Creating Samples and Pairing Files
1. Click **New Sample** and enter a name.
2. Select one or more N2 and/or O2 files in the Loaded Files list and click **↓ Add Selected Files to Sample**.
3. The app pairs files by RPM index automatically and shows them in the **RPM Pairs** table.
   - Files selected **in the same batch** that share the same catalyst and RPM are paired together first, preventing cross-contamination when the same catalyst/RPM exists in different experiments.
   - Lone files (single gas in the batch) merge into an existing incomplete pair if one matches.
4. Edit the **RPM** field in each row to enter the actual rotation speed from your lab notes (e.g. `400`, `900`, `1600`). Press Enter to save.

### 8.3 RPM Pair Table
The **RPM Pairs** section lists all matched N2/O2 pairs for the active sample:
- Pairs are **grouped by catalyst** with a purple separator header between each catalyst group.
- **Plot checkbox** — include or exclude this pair from the plot without deleting it. Unchecked pairs are also excluded from Tafel and KL analyses. The enabled state is saved with the session.
- **Catalyst** field — editable; renaming a catalyst here updates all pairs and correction entries sharing the same label in this sample.
- **RPM** field — editable; actual rotation speed from your lab notes.
- **N2 file / O2 file** — displayed at full width (green = found, red = missing). Resize the panel or window to see more of long filenames.
- **✕** button — remove a pair permanently.

### 8.4 Correction (active sample)
- **R_sol N2 (Ω)** — uncompensated resistance for the N2 session: `E = Ewe/V − (I/1000) × R_sol`.
- **R_sol O2 (Ω)** — uncompensated resistance for the O2 session (applied independently).
- **E_ref (V vs RHE)** — shared RHE offset: `E_RHE = E_corr + E_ref`.
- **Area (cm²)** — leave blank for I (mA); enter a value for J (mA cm⁻²).
- **ECSA_Hupd (cm²)** — electrochemical surface area from the Hupd Calc tab. Enter to enable specific activity (SA) calculation in the Extract Report window. Stored per catalyst; also auto-fills the ECSA input in the Specific Activity and Comparison windows.

### 8.5 Processing Pipeline (per pair)
1. Extract the **last cycle** from both N2 and O2.
2. Apply separate IR correction, then shared RHE conversion.
3. Extract the **anodic scan** (cathodic-vertex upward, sorted ascending).
4. Restrict to the overlapping E range and interpolate N2 → O2 grid.
5. Subtract: `I_net = I_O2 − I_N2_interp`. Divide by area if provided.

### 8.6 Plot Interactions
| Action | Effect |
|--------|--------|
| **Scroll up** | Zoom in (centered on cursor) |
| **Scroll down** | Zoom out |
| **Left-drag** | Pan |
| **Left-click** on a curve | Highlight all lines for that catalyst (others dimmed); bring them to front; switch Catalyst selector to match |
| **Right-click** | Clear highlight and annotation |
| **Left-drag legend** | Move legend |
| **Right-drag legend** | Resize legend font |
| **Double-click header** | Zoom sample full-panel |
| **Double-click axis label** | Edit axis label inline |

### 8.7 Multiple Samples
Each sample has its own subplot. Use **Cols** to control the grid width. Double-click a header strip to zoom; drag to reorder; uncheck to hide.

### 8.8 Plot Size
**W [__] H [__] inches** — size for all sample subplots; scrollbars appear automatically.

### 8.9 Extract Report
Click **Extract Report** in the Analysis section to open a report window that tabulates key ORR metrics for all currently visible (plotted) samples.

**Controls:**
- **E value (V vs Ref)** — target potential for the J@E and SA@E columns (default 0.90 V). Press Enter or click Compute to update.
- **Compute** — fills the table for all visible samples at RPMs 400, 900, 1600, 2500.
- **Copy TSV (→ Excel)** — copies the table as tab-separated values to the clipboard for direct paste into Excel.

**Columns extracted (per RPM):**
| Column | Description |
|--------|-------------|
| **I at E (mA)** | Current at the target potential × electrode area (if Area is set) |
| **SA at E (mA/cm²_ECSA)** | Kinetic current density normalised by ECSA_Hupd: `j_k = J@E × J_lim / (J_lim − J@E)`, then `SA = |j_k| / ECSA_Hupd`. Requires ECSA_Hupd to be entered in the correction panel. |
| **JL (mA/cm²)** | Limiting (diffusion) current — minimum current density on the curve |

The closest available RPM within ±50 rpm of each target is used; cells show blank or "N/A" when data or ECSA is missing.

---

## 9. Hupd Calc Tab

Use this tab to calculate ECSA and roughness factor (RF) from the hydrogen underpotential deposition (Hupd) region of a CV.

### 9.1 Loading Files

Click **Load Files** to load one or more `.mpr` or `.txt` CV files. Use the **Cycle** combobox to select which cycle to analyse (default: last cycle). Each file has independent **R_sol** (IR correction) and **E_ref** (RHE offset) fields.

### 9.2 Parameters

| Parameter | Description |
|-----------|-------------|
| **Scan Rate (mV/s)** | Potential sweep rate |
| **DL Region (V)** | Double-layer region for two-point baseline (Lo / Hi) |
| **Hupd Range (V)** | Integration window E1 – E2 (e.g. 0.05 – 0.40 V vs. RHE) |
| **Q_ref (µC/cm²)** | Reference charge for H monolayer (210 µC/cm² for Pt) |
| **Geo Area (cm²)** | Geometric electrode area for RF calculation |

### 9.3 Results

Click **Compute All**. For each file the app isolates the anodic half-cycle, fits a two-point linear baseline in the DL region, and integrates only the area **above** the baseline:

Q_H [µC] = (1/v) × ∫_{E1}^{E2} max(I_meas − I_baseline, 0) dE

then reports **Q_H**, **ECSA** (= Q_H / Q_ref), and **RF** (= ECSA / Geo Area) in the results table.

### 9.4 Plot

- Gray — full corrected cycle; blue — anodic half-cycle
- Orange band + dashed edge lines — DL region; green dashed verticals — E1/E2
- Black dashed — two-point baseline; green fill — integration area above baseline
- Light blue annotation box — Q_H/ECSA/RF (appears after Compute All)

### 9.5 Draggable Elements

Drag the four dashed boundary lines on the plot to adjust DL Lo/Hi and Hupd E1/E2; input fields update in real time. After Compute All, the result annotation box and the plot legend are also draggable and their positions persist across replots.

---

## 10. Common Controls (All Tabs)

### Mouse Interactions on the Plot
| Action | Effect |
|--------|--------|
| **Scroll wheel** | Zoom in/out around the cursor |
| **Left-drag** | Pan the plot |
| **Left-click** (on a data point) | Annotate that point with its coordinates |
| **Right-click** | Dismiss the annotation |
| **Double-click** (legend) | Open the legend label editor (rename + reorder) |
| **Double-click** (title strip) | Rename the plot title |
| **Double-click** (axis label) | Rename the X or Y axis label (General E.Chem tab) |
| **Right-drag** (legend) | Resize legend font size live — text, handle shapes, and entry spacing all scale together |

### Navigation Toolbar
Each plot has a toolbar below it:
- **Home** — reset the view to the auto-scaled limits from the last draw
- **Back / Forward** — navigate view history
- **Pan / Zoom** — standard matplotlib pan and zoom tools
- **Save** — save the current plot as an image file
- **Copy** — copy the current plot image to the Windows clipboard; paste directly into Word, PowerPoint, etc.

### Zoom/Pan Preservation
Each file remembers its last zoom/pan state. Switching to another file and back restores the view exactly where you left it.

### Legend Resize Behaviour
Right-drag on the legend to resize it. All three visual components scale together in real time:
- **Label text** — font size increases/decreases smoothly.
- **Handle icons** — the colored line or marker shape on the left of each entry grows/shrinks to match.
- **Entry spacing** — the gap between rows and the border padding scale proportionally.

The legend size is preserved when any other plot change is made (cycle selection, axis range, correction, etc.) — it does not reset to the default.

---

## 11. Tips and Shortcuts

- **Per-file independence** — every control in the left panel saves its value to the currently active file. Switch files freely; settings are never mixed up between files. In Multi E.Chem 2, cycle selection and corrections are also independent per file *per group* — the same file can have different cycles selected in different groups.
- **Plot highlight** — clicking a file in the list or clicking a line on the plot activates highlight mode: the selected line glows and others are dimmed. Right-click anywhere on the plot to clear the highlight. Highlight is never activated automatically when loading files.
- **Hide without losing settings** — unchecking a file in the file list removes it from the plot instantly. All cycles, corrections, zoom state, colors, and gradient settings are preserved. Re-check to restore the exact same view. *(Not available in the ECSA Calc tab.)*
- **Drag to reorder** — grab the **⠿** handle in the file list to drag files up or down. In Multi E.Chem you can also drag the **⠿ header strip** on each subplot to reorder the grid. The legend editor uses the same drag-handle pattern.
- **Axis swap shortcut** — use **⇄ Swap X↔Y** to instantly swap axes when you want to flip between, e.g., E vs. I and I vs. E without manually changing each dropdown.
- **Auto-merge sequential CV files** — when loading multiple EC-Lab CVA files at once (e.g. `sample_05_CV_C01.mpr`, `sample_07_CV_C01.mpr`, `sample_09_CV_C01.mpr`), the app detects the sequence pattern and automatically merges them into one entry with consecutively renumbered cycles. CA, OCV, EIS, and other non-voltammetry files in the same selection are loaded individually. A dialog confirms what was merged.
- **Smart axis defaults** — when loading a file for the first time, the app detects its data type and picks appropriate column defaults: **EIS** files → `Re(Z)` vs `-Im(Z)`; **OCV/time-series** files → `time/s` vs `Ewe/V`; **CV/LSV** files → `Ewe/V` vs `I/mA`. EIS files also filter the column dropdowns to show only impedance columns, keeping the selector clean.
- **EIS units** — when plotting EIS data in the General or Multi E.Chem tab, the unit dropdowns offer impedance units (mΩ, Ω, kΩ, MΩ), frequency units (Hz, kHz, MHz), and phase units (deg, rad). Data is scaled automatically when you switch.
- **Newer EC-Lab firmware files** — `.mpr` files that contain column types not yet known to galvani now load automatically; unrecognised columns are silently skipped.
- **Gradient for tracking evolution** — turn on Gradient to see how your CV cycles evolve: lightest = earliest, darkest = latest (most evolved).
- **Cycle order** — cycles are always plotted and listed in the legend in ascending numeric order (1, 2, 3…), so the gradient naturally reflects temporal evolution from lightest (earliest) to darkest (latest).
- **E_std (Rec)** — the green **Rec:** value shown next to the E_std field is the midpoint of your data's actual potential range. This is a reliable starting point for the non-Faradaic region.
- **Cs value** — the default 0.040 mF/cm² is appropriate for Pt in 0.1 M HClO₄. Adjust for your material and electrolyte.
- **Multi E.Chem zoom** — double-click the **⠿ header strip** of any subplot to zoom it to full panel; double-click the header again (or click ← Back to Grid) to return to the grid.
- **Plot size for publication** — use the W/H fields to set an exact figure size in inches. The scrollable canvas accommodates very wide or tall figures (up to 50 inches). Default sizes are chosen to match common publication column widths.
- **Blank titles by default** — all tabs default to no title. Type in the Title field or double-click the title strip to add one only when needed.
- **Legend resize** — right-drag the legend to resize it. The entire legend box (text, handle icons, and spacing) scales live as you drag. The size is remembered and not reset by subsequent plot changes.
- **Sharing results** — use the toolbar's Save button to export any plot as PNG/PDF, or use Export to Excel (General tab) for the raw and corrected data.
- **Copy to clipboard** — click the **Copy** button next to any toolbar to put the current plot on the clipboard and paste it directly into Word or PowerPoint at full resolution.
- **Rebuilding the exe** — if you update the code and want a fresh exe, run `pyinstaller EchemGUI.spec` from the project folder (with PyInstaller installed).

---

## 12. Session Save & Restore

The app can save the complete state of all five tabs — loaded files, groups, axis settings, corrections, cycle selections, colors, legend positions, and plot sizes — into a single `.echemsession` file. Raw data is embedded inside the file, so sessions can be shared or moved to another computer without bringing the original data files along.

### 10.1 Saving a Session
- **File → Save Session** (or **Ctrl+S**) — save the current session. On first save, a file browser opens; subsequent saves overwrite the same file.
- **File → Save Session As** — always opens a file browser so you can choose a new name or location.
- The `.echemsession` file is a ZIP archive containing:
  - `preview.png` — a thumbnail of the General E.Chem tab's current plot.
  - `data/{hash}.csv` — all loaded DataFrames, deduplicated by content: if the same source file is loaded in multiple tabs, it is stored only once.
  - `{tab}_state.json` — the full state of each of the five tabs.

### 10.2 Loading a Session
- **File → Load Session** (or **Ctrl+O**) — open a file browser, select a `.echemsession` file, and restore the full session.
- All five tabs are restored to exactly the state they were in when the session was saved.

### 10.3 Auto-Save on Close
- When you close the app, the current session is **automatically saved** to a hidden autosave file (`~/.echem_sessions/autosave.echemsession`).
- The next time you open the app, a dialog appears asking if you want to restore the last session. The dialog shows when the autosave was created.
  - Click **Yes** to restore it.
  - Click **No** to start a fresh session.
- This works even if you closed without manually saving.

### 10.4 What Is Saved Per Tab
| Tab | What is saved |
|-----|--------------|
| **General E.Chem** | All loaded files, per-file axis/unit/correction/cycle/color/gradient/legend/refline settings, zoom state, plot size, custom labels, legend order and positions |
| **Multi E.Chem** | All loaded files with the same per-file settings as above; grid column count; active file |
| **Multi E.Chem 2** | All loaded files and group definitions; per-group axis/legend/refline settings; per-(group, file) cycle and correction state; grid column count; active group and file |
| **ECSA Calc** | All loaded files with per-file scan rate tables, E_std, Cs, extracted Cdl/ECSA results, CV/Cdl reference lines, and zoom state for both plots |
| **OCV/Ru Extractor** | All loaded samples (each bundles its OCV + EIS files), per-sample color, show/hide state, sample order, extracted OCV / Ru values, and plot size |
| **ORR Analysis** | All loaded files, sample definitions, N2/O2 pair tables with RPM values, per-sample correction (R_sol N2/O2, E_ref, area) and all plot settings |
