# Echem GUI — User Manual

## Table of Contents
1. [Getting Started](#1-getting-started)
2. [Interface Overview](#2-interface-overview)
3. [General E.Chem Tab](#3-general-echem-tab)
4. [Multi E.Chem Tab](#4-multi-echem-tab)
5. [Multi E.Chem 2 Tab](#5-multi-echem-2-tab)
6. [ECSA Calc Tab](#6-ecsa-calc-tab)
7. [Nyquist Plot Tab](#7-nyquist-plot-tab)
8. [Common Controls (All Tabs)](#8-common-controls-all-tabs)
9. [Tips and Shortcuts](#9-tips-and-shortcuts)

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

The app has five tabs at the top:

| Tab | Purpose |
|-----|---------|
| **General E.Chem** | Overlay multiple files on one plot; full correction and export |
| **Multi E.Chem** | View one plot per file simultaneously in a 2-column grid |
| **Multi E.Chem 2** | Group files into named groups; overlay all files in each group on one plot |
| **ECSA Calc** | Extract electrochemical surface area (ECSA) from CV data |
| **Nyquist Plot** | Plot EIS impedance data as a Nyquist diagram |

Each tab is fully **independent** — files loaded in one tab are not shared with others.

The layout in every tab is:
- **Left panel** — scrollable controls (file list, axis settings, correction, cycles, etc.)
- **Right panel** — scrollable plot area; width and height are adjustable via W/H fields

---

## 3. General E.Chem Tab

Use this tab to load one or more data files and overlay them all on a single plot.

### 3.1 Loading Files
1. Click **Load File(s)** to open a file browser. Select one or more `.mpr` or `.txt` files.
2. Loaded files appear in the **file list**. Each row has a **checkbox**, a **⠿ drag handle**, and a filename label.
   - **Click the filename** to make it the active file.
   - **Uncheck the checkbox** to hide that file's curves without removing it. All settings are preserved.
   - **Drag the ⠿ handle** to reorder files.
3. To remove a file, select it and click **Remove**.

#### Auto-merging sequential CV files
When you select multiple files at once, the app automatically detects EC-Lab CVA sequence naming and merges matching groups:
- Files sharing the same base name, method, and channel but differing only in the sequence number (e.g. `sample_05_CV_C01.mpr`, `sample_07_CV_C01.mpr`) are merged into a single entry with consecutively renumbered cycle numbers.
- Only voltammetry methods are merged (**CV, CVA, LSV, DPV, NPV, SWV**). CA, OCV, EIS are always loaded individually.
- A dialog appears after loading, listing every group that was auto-merged.

### 3.2 Axis Settings
- **X / Y column selectors** — choose which data column to plot on each axis. EIS files show only impedance columns; CV/OCV files show all columns.
- **Smart defaults** — EIS → `Re(Z)` vs `-Im(Z)`; OCV/time-series → `time/s` vs `Ewe/V`; CV/LSV → `Ewe/V` vs `I/mA`.
- **Unit dropdowns** — dimension-aware: Voltage (V/mV/µV/nV), Current (A/mA/µA/nA), Time (s/ms/µs/min/h), Impedance (mΩ/Ω/kΩ/MΩ), Frequency (Hz/kHz/MHz), Phase (deg/rad).
- **J (current density)** — enter a positive electrode area (cm²) per file to enable a virtual "J" column for current density.

### 3.3 Corrections
- **Reference electrode** — select your reference (RHE, Ag/AgCl, SCE, NHE, MSE).
- **E_ref (V)** — `E_RHE = E_measured + E_ref`.
- **R_sol (Ω)** — `E_corr = E − (I_mA / 1000) × R_sol`.
- Click **Apply Correction** to apply.

### 3.4 Cycle Selection
Files with a `cycle number` column show a grid of cycle checkboxes.
- Use **Select All / Deselect All** for quick selection.

### 3.5 Plot Range
Enter values in the **X min / X max / Y min / Y max** boxes to fix the axis range. Leave blank for automatic scaling.

### 3.6 Axis Orientation
- **Flip X / Flip Y** — invert the direction of that axis.
- **⇄ Swap X↔Y** — swap the X and Y axes (column, unit, range, and flip state) in one click.

### 3.7 File Colors and Cycle Gradient
- **Color combobox** — base color per file; auto-assigned distinct colors on load.
- **Gradient / Reverse / Step** — lightness gradient across cycles; per-file; default step 0.15.

### 3.8 Legend
- **Show / Frame / Font size / Location** controls.
- **Edit Labels** — rename and reorder legend entries via ⠿ drag handles.
- **Double-click the legend** on the plot to open the same editor.
- **Right-drag the legend** to resize it live — text, handle icons, and spacing all scale together; size is preserved across replots.

### 3.9 Reference Lines
Add X (vertical) or Y (horizontal) guide lines with individual style and color.

### 3.10 Excel Export
**Export to Excel** saves the active file to `.xlsx` with Raw and Corrected sheets.

### 3.11 Plot Title and Axis Labels
- **Title** entry in the left panel (blank by default); also double-click the title strip on the plot.
- **Axis labels** — double-click the axis label on the plot to rename; blank reverts to auto-generated label.

### 3.12 Font Spacing and Plot Size
- **Spacing (pt): Title [__] Label [__]** — gap between axes and title, and between ticks and axis labels.
- **W [__] H [__] inches** — figure size in inches. Default W=21.0, H=12.5. Maximum 50 inches. A scrollbar appears when the figure exceeds the panel width.

---

## 4. Multi E.Chem Tab

Use this tab to view all loaded files **side by side** in a 2-column grid, each with its own independent plot.

### 4.1 Loading and Selecting Files
- Each file row has a **checkbox** (hide/show subplot) and a **⠿ drag handle** (reorder).
- Click any subplot or the filename in the list to make it the **active file**.
- Each subplot has a **⠿ header strip** — drag it to reorder within the grid.

### 4.2 Per-File Settings
All left-panel controls apply **only to the active file**. Each file remembers its own settings independently.

### 4.3 Plot Title
- The **Title** entry in the left panel sets the active file's subplot title. Default is blank.
- Also double-click the title strip on any subplot to rename it inline.
- Per-file; persists across replots.

### 4.4 Subplot Zoom
Double-click any subplot to expand it. A **← Back to Grid** button restores the grid.

### 4.5 Axis Orientation
**Flip X / Flip Y** and **⇄ Swap X↔Y** — per-file settings.

### 4.6 Colors and Cycle Gradient
Same controls as General tab, per-file.

### 4.7 Plot Size
- **W [__] H [__] inches** — figure size for all subplots. Default W=10.5, H=5.5. Maximum 50 inches.

---

## 5. Multi E.Chem 2 Tab

Use this tab to organise files into named **groups**. Each group produces one overlay plot with all files in the group drawn on the same axes. Useful for comparing treatment conditions, replicates, or electrode sets.

### 5.1 Group Management
- **Create a group** — type a name and click **New Group**.
- **Add files to a group** — load files first, then assign them to groups.
- Click a group name in the listbox to make it the **active group**.

### 5.2 Per-Group Settings
All left-panel controls apply only to the active group. Each group remembers its settings independently.

### 5.3 Plot Title
- The **Title** entry in the left panel sets the active group's plot title. Default is blank.
- Also double-click the title strip on any group subplot.

### 5.4 File Order within a Group
Drag the **⠿ handle** next to each file in the group to reorder the overlay draw order.

### 5.5 Plot Size
- **W [__] H [__] inches** — figure size for all group plots. Default W=10.5, H=5.5. Maximum 50 inches.

---

## 6. ECSA Calc Tab

Use this tab to extract ECSA from CV data using the double-layer capacitance (Cdl) method.

### 6.1 Layout
Right panel has two stacked plots: **CV** (upper) and **Cdl** (lower).

### 6.2 Scan Rate Table
After selecting cycles, enter the scan rate (mV/s) per cycle. The CV legend updates as you type.

### 6.3 ECSA Parameters
- **E_std (V)** — potential for Δj measurement; red dashed line marks it on the CV. The green **Rec:** label shows the recommended midpoint `(E_max + E_min) / 2`.
- **Cs (mF/cm²)** — specific capacitance; default 0.040 mF/cm².

### 6.4 Extracting ECSA
Click **Extract Cdl & ECSA**. The Cdl plot shows the scatter + linear fit; legend shows fit equation, Cdl, R², and ECSA.

### 6.5 ECSA Physics
```
Δj/2 = (ja − jc) / 2  at E_std for each cycle (scan rate)
Linear fit: scan_rate vs Δj/2  →  slope = Cdl [F]
cdl_mF = slope × 1000;  ECSA = cdl_mF / Cs  [cm²]
```

### 6.6 Plot Size
- **W [__] H [__] inches** — applied to both CV and Cdl plots. Default W=21.0, H=6.0. Maximum 50 inches.

---

## 7. Nyquist Plot Tab

Use this tab to visualize EIS data as a Nyquist diagram (Re(Z) vs. −Im(Z)).

### 7.1 Loading and Display
Load `.mpr` or `.txt` files with impedance columns. All files are overlaid on one plot, each with a distinct color and marker.

### 7.2 Options
- **Connect lines / Show markers** — toggle connecting lines and markers.
- **Flip X / Flip Y** and **⇄ Swap X↔Y** — axis orientation controls.
- **Unit dropdowns** — Ω, kΩ, MΩ.

### 7.3 Plot Size
- **W [__] H [__] inches** — Default W=21.0, H=12.5. Maximum 50 inches.

---

## 8. Common Controls (All Tabs)

### Mouse Interactions on the Plot
| Action | Effect |
|--------|--------|
| **Scroll wheel** | Zoom in/out around the cursor |
| **Left-drag** | Pan the plot |
| **Left-click** (on a data point) | Annotate with coordinates |
| **Right-click** | Dismiss annotation |
| **Double-click** (legend) | Open legend label editor |
| **Double-click** (title strip) | Rename plot title |
| **Double-click** (axis label) | Rename axis label (General E.Chem tab) |
| **Right-drag** (legend) | Resize legend live — text, handle shapes, and spacing all scale together |

### Navigation Toolbar
- **Home** — reset view to last auto-scaled limits
- **Back / Forward** — navigate view history
- **Pan / Zoom** — standard tools
- **Save** — export plot as image
- **Copy** — copy to Windows clipboard (paste into Word/PowerPoint)

### Legend Resize Behaviour
Right-drag the legend to resize. All visual components scale proportionally in real time:
- Label text font size
- Handle icon size (the colored line/marker shape left of each label)
- Row spacing and border padding

The legend size is **preserved** when any other plot change is made.

### Zoom/Pan Preservation
Each file remembers its last zoom/pan state. Switching files and back restores the view exactly.

---

## 9. Tips and Shortcuts

- **Hide without losing settings** — uncheck a file to remove it from the plot. All settings, corrections, and zoom state are preserved. Re-check to restore instantly.
- **Drag to reorder** — use ⠿ handles in file lists, subplot headers, and the legend editor.
- **Blank titles by default** — type in the Title field or double-click the title strip only when a title is needed.
- **Legend resize** — right-drag resizes the full legend (text + handle icons + spacing) live. The size is preserved across replots.
- **Plot size controls** — available in all tabs via W/H inch fields. Scrollbars appear automatically for oversized figures. Useful for setting exact publication figure dimensions.
- **Auto-merge** — loading multiple EC-Lab CVA sequence files at once auto-merges them with renumbered cycles. A confirmation dialog lists what was merged.
- **Smart defaults** — the app detects EIS vs. CV vs. OCV data type on load and picks appropriate columns and units automatically.
- **Copy to clipboard** — the **Copy** button on every toolbar copies the figure for direct pasting into Word or PowerPoint.
- **Export to Excel** — available in the General E.Chem tab; exports Raw and Corrected data for the active file.
- **Rebuilding the exe** — run `pyinstaller EchemGUI.spec` from the project folder.
