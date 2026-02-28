"""Quick diagnostic: opens a file dialog to pick a .mpr file and tests galvani.

Run as a regular Python script (not via pytest):
    python test_mpr.py
"""
import sys
import tkinter as tk
from tkinter import filedialog
import pandas as pd


def run_diagnostic():
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title="Select a .mpr file to test",
        filetypes=[("BioLogic MPR", "*.mpr"), ("All files", "*.*")]
    )
    if not path:
        print("No file selected.")
        sys.exit(0)

    print(f"Testing: {path}\n")

    # --- Step 1: import galvani ---
    try:
        from galvani import BioLogic
        print("[OK] galvani imported successfully")
    except ImportError as e:
        print(f"[FAIL] galvani import error: {e}")
        sys.exit(1)

    # --- Step 2: open the file ---
    try:
        mpr = BioLogic.MPRfile(path)
        print("[OK] MPRfile opened")
    except Exception as e:
        print(f"[FAIL] MPRfile() error: {e}")
        sys.exit(1)

    # --- Step 3: convert to DataFrame ---
    try:
        df = pd.DataFrame(mpr.data)
        print(f"[OK] DataFrame created: {len(df)} rows, {len(df.columns)} cols")
    except Exception as e:
        print(f"[FAIL] DataFrame conversion error: {e}")
        sys.exit(1)

    # --- Step 4: show raw column names ---
    print(f"\nRaw column names:")
    for c in df.columns:
        print(f"  {repr(c)}")

    # --- Step 5: apply cleanup and check desired set ---
    _MPR_DESIRED = frozenset({
        "time/s", "Ewe/V", "I/mA", "cycle number",
        "Re(Z)/Ohm", "-Im(Z)/Ohm", "freq/Hz", "Phase(Z)/deg",
    })

    cleaned = [c.strip().replace("<", "").replace(">", "") for c in df.columns]
    keep = [c for c in cleaned if c in _MPR_DESIRED]

    print(f"\nAfter cleanup, recognized columns ({len(keep)}):")
    for c in keep:
        print(f"  {c}")

    if not keep:
        print("\n[FAIL] No recognized columns — the file won't load.")
    else:
        print(f"\n[OK] File should load fine with {len(keep)} columns.")


if __name__ == "__main__":
    run_diagnostic()
