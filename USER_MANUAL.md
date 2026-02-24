# Echem GUI — User Manual

## Table of Contents
1. [Getting Started](#1-getting-started)
2. [Interface Overview](#2-interface-overview)
3. [General E.Chem Tab](#3-general-echem-tab)
4. [Multi E.Chem Tab](#4-multi-echem-tab)
5. [ECSA Calc Tab](#5-ecsa-calc-tab)
6. [Nyquist Plot Tab](#6-nyquist-plot-tab)
7. [Common Controls (All Tabs)](#7-common-controls-all-tabs)
8. [Tips and Shortcuts](#8-tips-and-shortcuts)

---

## 1. Getting Started

### Running the App
- **From the exe:** Double-click `EchemGUI.exe` inside the `EchemGUI` folder.
- **From Python:** Run `python run_echem.py` in the project folder.

### Supported File Format
The app reads **tab-separated `.txt` files** as exported by EC-Lab (BioLogic) and similar potentiostats. Column names are recognized automatically (e.g. `Ewe/V`, `I/mA`, `time/s`, `cycle number`, `Re(Z)/Ohm`, `-Im(Z)/Ohm`).

---

## 2. Interface Overview

The app has four tabs at the top:

| Tab | Purpose |
|-----|---------|
| **General E.Chem** | Overlay multiple files on one plot; full correction and export |
| **Multi E.Chem** | View one plot per file simultaneously in a grid |
| **ECSA Calc** | Extract electrochemical surface area (ECSA) from CV data |
| **Nyquist Plot** | Plot EIS impedance data as a Nyquist diagram |

Each tab is fully **independent** — files loaded in one tab are not shared with others.

The layout in every tab is:
- **Left panel** — scrollable controls (file list, axis settings, correction, cycles, etc.)
- **Right panel** — the plot area

---

## 3. General E.Chem Tab

Use this tab to load one or more data files and overlay them all on a single plot.

### 3.1 Loading Files
1. Click **Load File(s)** to open a file browser. Select one or more `.txt` files.
2. Loaded files appear in the **file listbox**. Click a filename to make it the active file — all left-panel controls reflect that file's settings.
3. To remove a file, select it and click **Remove File**.

### 3.2 Axis Settings
- **X / Y column selectors** — choose which data column to plot on each axis.
- **Unit dropdowns** — select the display unit (e.g. V, mV, mA, µA, A/cm²). The data is scaled automatically.
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

### 3.6 File Colors and Cycle Gradient
- **Color combobox** — change the base color of the active file. Each file is auto-assigned a distinct color on load.
- **Gradient checkbox** — when checked, cycles within a file are drawn with a lightness gradient: the first cycle is the lightest and the last cycle is the darkest (most evolved cycle = most visible).
- **Reverse checkbox** — flips the gradient direction (first cycle darkest, last lightest).
- **Step spinbox** — controls how much the lightness changes between consecutive cycles (range 0.01–0.30; default 0.08).
- These settings are saved **per file** — switching files restores each file's individual settings.

### 3.7 Legend
- **Show Legend checkbox** — toggles the legend on/off.
- **Frame checkbox** — adds a border around the legend box.
- **Font size** — adjust legend text size.
- **Location** — choose legend anchor position, or drag the legend freely with the mouse.
- **Edit Labels** — opens a dialog to rename individual legend entries.

### 3.8 Reference Lines
Add horizontal or vertical guide lines to the plot:
1. Type a value in the **X =** or **Y =** entry box.
2. Choose a **line style** (dashed, dotted, solid, dash-dot) and **color**.
3. Click **Add X Line** or **Add Y Line**.
4. To edit or remove a line, select it in the listbox — its style/color loads into the dropdowns. Click **Remove** to delete it.

Reference lines in this tab are shared across all overlaid files.

### 3.9 Excel Export
Click **Export to Excel** to save the active file's data to an `.xlsx` file with two sheets:
- **Raw** — original data, one column group per selected cycle.
- **Corrected** — IR/RHE corrected data.

### 3.10 Editable Plot Title
Double-click anywhere in the title area above the plot to rename it.

---

## 4. Multi E.Chem Tab

Use this tab to view all loaded files **side by side** in a 2-column grid, each with its own independent plot.

### 4.1 Loading and Selecting Files
- Load files the same way as in the General tab.
- Click any subplot (or its toolbar) to make that file **active** — the left panel updates to show that file's settings.
- Alternatively, click the filename in the listbox.

### 4.2 Per-File Settings
Every control in the left panel (axis columns, units, range, cycles, correction, colors, gradient, legend, reference lines) applies **only to the active file**. Each file remembers its own settings independently.

### 4.3 Subplot Zoom
- **Double-click** any subplot to expand it to fill the entire right panel.
- A **← Back to Grid** button appears at the top. Click it to return to the 2-column grid.
- Figures remain live during zoom — no data is lost.

### 4.4 Editable Subplot Titles
Double-click the title strip above any subplot to rename it. The custom title is remembered even after replots.

### 4.5 Colors and Cycle Gradient
Same controls as the General tab, but settings are per-file. Switching files restores that file's color and gradient settings.

---

## 5. ECSA Calc Tab

Use this tab to extract the **electrochemical surface area (ECSA)** from cyclic voltammetry data using the double-layer capacitance (Cdl) method.

### 5.1 Overview of the Layout
The right panel has **two stacked plots**:
- **Upper plot (CV)** — the cyclic voltammetry curves for selected cycles.
- **Lower plot (Cdl)** — the linear fit of scan rate vs. Δj/2 used to extract Cdl and ECSA.

### 5.2 Loading Files and Selecting Cycles
Load files and select cycles the same way as in the General tab. Each cycle should correspond to a different scan rate.

### 5.3 Scan Rate Table
After selecting cycles, a scan rate input table appears:
- Enter the scan rate (mV/s) for each cycle.
- The CV legend updates as you type (no need to press Enter).

### 5.4 Setting ECSA Parameters
- **E_std (V)** — the potential at which Δj is measured (the standard potential in the non-Faradaic region). A red dashed vertical line marks this position on the CV plot.
- **Cs (mF/cm²)** — specific capacitance of the material (default: 0.040 mF/cm² for standard electrolyte). Used to convert Cdl to ECSA.

### 5.5 Extracting ECSA
1. Select the cycles, enter scan rates, and set E_std and Cs.
2. Click **Extract Cdl & ECSA**.
3. The Cdl plot shows a scatter of (scan rate, Δj/2) points with a linear fit. The legend shows:
   - The fit equation
   - Cdl (mF)
   - R² (goodness of fit)
   - ECSA (cm²)

### 5.6 ECSA Physics
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

### 5.7 Colors and Cycle Gradient
- The **Color combobox** changes the base color used for CV cycles.
- The **Gradient / Reverse / Step** controls work the same as in the General tab and are per-file.
- The Cdl scatter/fit plot always uses fixed colors (steelblue dots, tomato fit line).

### 5.8 Reference Lines
Separate reference line sections are provided for the CV plot and the Cdl plot. Each has its own listbox and add/remove controls.

### 5.9 Editable Plot Titles
Double-click the title strip on either the CV or Cdl plot to rename it.

---

## 6. Nyquist Plot Tab

Use this tab to visualize **electrochemical impedance spectroscopy (EIS)** data as a Nyquist diagram (Re(Z) vs. −Im(Z)).

### 6.1 Loading Files
Load `.txt` files that contain impedance columns (e.g. `Re(Z)/Ohm` and `-Im(Z)/Ohm`).

### 6.2 Axis Settings
- Select which columns to use for X and Y using the dropdowns.
- Adjust units with the unit comboboxes (Ω, kΩ, MΩ).

### 6.3 Display Options
- **Connect lines** — draws a line connecting the data points in frequency order.
- **Show markers** — shows point markers at each data point.

### 6.4 Multiple Files
All loaded files are overlaid on the same Nyquist plot. Each file is automatically assigned a **distinct color** and a **unique marker shape** so they can be told apart even without a legend.

### 6.5 Editable Plot Title
Double-click the title strip to rename the plot.

---

## 7. Common Controls (All Tabs)

### Mouse Interactions on the Plot
| Action | Effect |
|--------|--------|
| **Scroll wheel** | Zoom in/out around the cursor |
| **Left-drag** | Pan the plot |
| **Left-click** (on a data point) | Annotate that point with its coordinates |
| **Right-click** | Dismiss the annotation |
| **Double-click** (title strip) | Rename the plot title |

### Navigation Toolbar
Each plot has a toolbar below it:
- **Home** — reset the view to the auto-scaled limits from the last draw
- **Back / Forward** — navigate view history
- **Pan / Zoom** — standard matplotlib pan and zoom tools
- **Save** — save the current plot as an image file

### Zoom/Pan Preservation
Each file remembers its last zoom/pan state. Switching to another file and back restores the view exactly where you left it.

---

## 8. Tips and Shortcuts

- **Per-file independence** — every control in the left panel saves its value to the currently active file. Switch files freely; settings are never mixed up between files.
- **Gradient for tracking evolution** — turn on Gradient to see how your CV cycles evolve: lightest = earliest, darkest = latest (most evolved).
- **Cycle order matters** — cycles are plotted in the order they appear in the data file, so the gradient naturally reflects temporal evolution.
- **E_std position** — set E_std to the midpoint of the flat, featureless region of your CV for the most reliable Cdl extraction.
- **Cs value** — the default 0.040 mF/cm² is appropriate for Pt in 0.1 M HClO₄. Adjust for your material and electrolyte.
- **Multi E.Chem zoom** — use the zoom feature to inspect a single file in detail without losing the grid view; double-click anywhere on the subplot (not on the title) to zoom.
- **Sharing results** — use the toolbar's Save button to export any plot as PNG/PDF, or use Export to Excel (General tab) for the raw and corrected data.
- **Rebuilding the exe** — if you update the code and want a fresh exe, run `pyinstaller EchemGUI.spec` from the project folder (with PyInstaller installed).
