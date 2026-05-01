"""Session save/restore for EchemGUI – .echemsession (ZIP) format.

ZIP layout:
    manifest.json         version, timestamp, tab list
    preview.png           thumbnail from General E.Chem figure
    data/<hash>.csv       deduplicated raw DataFrames (SHA-256 of CSV bytes)
    general.json          EchemPanel state
    multi_echem.json      MultiEchemPanel state
    multi_echem2.json     MultiEchem2Panel state
    ecsa.json             ECSAPanel state
    nyquist.json          EISPanel state
"""

import io
import json
import hashlib
import zipfile
import datetime
from pathlib import Path

import pandas as pd

SESSION_VERSION = "1.0"
SESSION_EXT     = ".echemsession"

_AUTOSAVE_DIR  = Path.home() / ".echem_sessions"
AUTOSAVE_PATH  = _AUTOSAVE_DIR / "autosave.echemsession"

# Runtime-only keys stripped from file entries before serialisation
_FILE_RUNTIME = frozenset({
    "df", "df_raw",
    "fig", "ax", "ax_cv", "ax_cdl", "canvas", "canvas_cdl", "toolbar",
    "plot_frame",
    "legend",
    "label_var",
})

# Runtime-only keys stripped from group entries
_GROUP_RUNTIME = frozenset({
    "fig", "ax", "canvas", "toolbar", "plot_frame", "legend",
})


# ── DataFrame helpers ─────────────────────────────────────────────────

def _df_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def df_hash(df: pd.DataFrame) -> str:
    return hashlib.sha256(_df_csv_bytes(df)).hexdigest()[:20]


# ── Entry serialisation ───────────────────────────────────────────────

def serialise_file_entry(name: str, entry: dict, data_store: dict) -> dict:
    """Return a JSON-safe dict for one files-dict entry; stash df_raw in data_store."""
    rec: dict = {"name": name}
    df_raw = entry.get("df_raw")
    if df_raw is not None:
        h = df_hash(df_raw)
        data_store[h] = df_raw
        rec["data_hash"] = h
    for k, v in entry.items():
        if k in _FILE_RUNTIME:
            continue
        try:
            json.dumps(v)
            rec[k] = v
        except (TypeError, ValueError):
            pass
    return rec


def serialise_group_entry(name: str, gentry: dict) -> dict:
    """Return a JSON-safe dict for one groups-dict entry."""
    rec: dict = {"name": name}
    for k, v in gentry.items():
        if k in _GROUP_RUNTIME:
            continue
        try:
            json.dumps(v)
            rec[k] = v
        except (TypeError, ValueError):
            pass
    return rec


# ── Preview helper ────────────────────────────────────────────────────

def _capture_preview(panel) -> bytes | None:
    """Render the General E.Chem figure to a 72-dpi PNG; return bytes or None."""
    if panel is None:
        return None
    fig = getattr(panel, "fig", None)
    if fig is None:
        return None
    buf = io.BytesIO()
    try:
        fig.savefig(buf, format="png", dpi=72, bbox_inches="tight")
        return buf.getvalue()
    except Exception:
        return None


# ── Save ──────────────────────────────────────────────────────────────

def save_session(panels: dict, filepath: str) -> None:
    """Save all panel states to *filepath* (.echemsession ZIP).

    panels = {
        "general":     EchemPanel instance,
        "multi_echem": MultiEchemPanel instance,
        "multi_echem2":MultiEchem2Panel instance,
        "ecsa":        ECSAPanel instance,
        "nyquist":     EISPanel instance,
    }
    """
    data_store: dict[str, pd.DataFrame] = {}
    states: dict[str, dict] = {}

    for tab_id, panel in panels.items():
        fn = getattr(panel, "get_session_state", None)
        if fn is not None:
            try:
                states[tab_id] = fn(data_store)
            except Exception as exc:
                import traceback
                print(f"[session] error capturing {tab_id}: {exc}")
                traceback.print_exc()

    preview_bytes = _capture_preview(panels.get("general"))

    Path(filepath).parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(filepath, "w", zipfile.ZIP_DEFLATED) as zf:
        manifest = {
            "version":  SESSION_VERSION,
            "saved_at": datetime.datetime.now().isoformat(timespec="seconds"),
            "tabs":     list(states.keys()),
        }
        zf.writestr("manifest.json", json.dumps(manifest, indent=2))

        if preview_bytes:
            zf.writestr("preview.png", preview_bytes)

        for h, df in data_store.items():
            zf.writestr(f"data/{h}.csv", _df_csv_bytes(df).decode("utf-8"))

        for tab_id, state in states.items():
            zf.writestr(f"{tab_id}.json", json.dumps(state, indent=2, default=str))


# ── Load ──────────────────────────────────────────────────────────────

def load_session(panels: dict, filepath: str) -> None:
    """Restore all panel states from *filepath* (.echemsession ZIP).

    panels: same dict layout as save_session.
    """
    with zipfile.ZipFile(filepath, "r") as zf:
        names = set(zf.namelist())

        # Load all DataFrames into memory keyed by hash
        data_store: dict[str, pd.DataFrame] = {}
        for n in names:
            if n.startswith("data/") and n.endswith(".csv"):
                h = n[5:-4]
                try:
                    data_store[h] = pd.read_csv(io.BytesIO(zf.read(n)))
                except Exception:
                    pass

        # Restore each tab
        for tab_id, panel in panels.items():
            state_file = f"{tab_id}.json"
            if state_file not in names:
                continue
            try:
                state = json.loads(zf.read(state_file))
            except Exception:
                continue
            fn = getattr(panel, "restore_session_state", None)
            if fn is not None:
                try:
                    fn(state, data_store)
                except Exception as exc:
                    import traceback
                    print(f"[session] error restoring {tab_id}: {exc}")
                    traceback.print_exc()


# ── Autosave helpers ──────────────────────────────────────────────────

def autosave(panels: dict) -> None:
    """Save to the standard autosave location; silently swallows errors."""
    try:
        _AUTOSAVE_DIR.mkdir(parents=True, exist_ok=True)
        save_session(panels, str(AUTOSAVE_PATH))
    except Exception:
        pass


def autosave_exists() -> bool:
    return AUTOSAVE_PATH.is_file()


def autosave_info() -> str:
    """Return a human-readable summary of the autosave file (timestamp + tab list)."""
    if not autosave_exists():
        return ""
    try:
        with zipfile.ZipFile(str(AUTOSAVE_PATH), "r") as zf:
            manifest = json.loads(zf.read("manifest.json"))
        saved_at = manifest.get("saved_at", "unknown time")
        tabs = manifest.get("tabs", [])
        return f"Saved: {saved_at}\nTabs with data: {', '.join(tabs)}"
    except Exception:
        return ""
