# gf_sheet_utils.py
# tksheet quality-of-life helpers:
# - Plain-text paste (Ctrl+V and Ctrl+Shift+V) that respects tabs/newlines
# - Robust TSV/CSV clipboard parsing with csv.reader (quote-aware)
# - Optional "headers-only" paste (cap columns for Dialer)
# - Right-click menu: Copy / Add 100 Rows / Save Sheet / Delete Row / Undo
#   (PASTE REMOVED to avoid the row-1 jump for now)
# - Column resizing
# - Arrow/Tab keyboard navigation with wrap & auto-scroll
# - Column width persistence (load/apply/save on resize)

from __future__ import annotations

import json
import csv
import io
from pathlib import Path
from typing import Optional, List


# =========================
# Parsing helpers (TSV/CSV)
# =========================
def _parse_clipboard(clip: str) -> List[List[str]]:
    """
    Robustly parse clipboard text as TSV when tabs dominate; else CSV.
    Uses csv.reader for BOTH to keep quotes together and preserve empty trailing cells.
    """
    if not isinstance(clip, str):
        return []
    clip = clip.replace("\r\n", "\n").replace("\r", "\n")  # normalize newlines

    tab_count = clip.count("\t")
    comma_count = clip.count(",")

    try:
        if tab_count > 0 and tab_count >= comma_count:
            reader = csv.reader(io.StringIO(clip), delimiter="\t", quotechar='"')
        else:
            reader = csv.reader(io.StringIO(clip), delimiter=",", quotechar='"')
        return [row for row in reader]
    except Exception:
        # Fallback: split by tabs or lines
        if "\t" in clip:
            return [line.split("\t") for line in clip.split("\n")]
        return [[line] for line in clip.split("\n")]


# =========================
# Anchor helpers
# =========================
def _event_to_cell(sheet_obj, evt) -> Optional[tuple]:
    """
    Best-effort: map a mouse event to (row, col) across tksheet versions/widgets.
    """
    # 1) Newer tksheet helper (if available)
    fn = getattr(sheet_obj, "get_rc_popup_menu_grid_indexes", None)
    if callable(fn):
        try:
            r, c = fn(evt)
            if isinstance(r, int) and isinstance(c, int):
                return (r, c)
        except Exception:
            pass

    # 2) Fallback via MT identify (works for MT/RI/CH if present)
    mt = getattr(sheet_obj, "MT", None)
    if mt is not None:
        row_id = getattr(mt, "identify_row", None)
        col_id = getattr(mt, "identify_col", None)
        if callable(row_id) and callable(col_id):
            try:
                r = row_id(evt.y)
                c = col_id(evt.x)
                if isinstance(r, int) and isinstance(c, int):
                    return (r, c)
            except Exception:
                pass

    # 3) Last resort: use currently selected cell
    try:
        r, c = sheet_obj.get_currently_selected()
        if isinstance(r, int) and isinstance(c, int):
            return (r, c)
    except Exception:
        pass
    return None


def _selected_anchor(sheet_obj) -> tuple[int, int]:
    """
    Anchor for Ctrl+V: prefer the current cell, else first selected row, else (0,0).
    """
    r0 = c0 = None
    try:
        r0, c0 = sheet_obj.get_currently_selected()
    except Exception:
        r0 = c0 = None

    if not isinstance(r0, int) or r0 < 0:
        try:
            sel_rows = sheet_obj.get_selected_rows() or []
            if sel_rows:
                r0 = int(sel_rows[0])
        except Exception:
            r0 = None

    if not isinstance(r0, int) or r0 < 0:
        r0 = 0
    if not isinstance(c0, int) or c0 < 0:
        c0 = 0
    return r0, c0


# =========================
# Paste implementations
# =========================
def _get_clip_rows_from_root(tk_root) -> List[List[str]]:
    try:
        clip = tk_root.clipboard_get()
    except Exception:
        return []
    if not clip:
        return []
    return _parse_clipboard(clip)


def _do_plain_paste_at(sheet_obj, tk_root, r0: int, c0: int, headers_only_cols: Optional[int] = None):
    """
    Paste at explicit (row, col). Used by right-click in some builds; safe anchor.
    """
    rows = _get_clip_rows_from_root(tk_root)
    if not rows:
        return "break"

    if r0 is None or r0 < 0: r0 = 0
    if c0 is None or c0 < 0: c0 = 0

    # Grow rows if needed
    try:
        total_rows = sheet_obj.get_total_rows()
    except Exception:
        total_rows = 0
    need_rows = r0 + len(rows)
    if need_rows > total_rows:
        try:
            sheet_obj.insert_rows(total_rows, number_of_rows=(need_rows - total_rows))
        except Exception:
            try:
                sheet_obj.insert_rows(total_rows, amount=(need_rows - total_rows))
            except Exception:
                pass

    # Manual cell fill (reliable & respects headers_only_cols)
    for r_off, row in enumerate(rows):
        for c_off, val in enumerate(row):
            dest_c = c0 + c_off
            if headers_only_cols is not None and dest_c >= headers_only_cols:
                continue
            try:
                sheet_obj.set_cell_data(r0 + r_off, dest_c, val)
            except Exception:
                pass

    # Keep focus/selection on anchor
    try:
        if hasattr(sheet_obj, "select_cell"):
            sheet_obj.select_cell(r0, c0)
        else:
            sheet_obj.set_currently_selected(r0, c0)
        sheet_obj.see(r0, c0)
        sheet_obj.refresh()
    except Exception:
        pass
    return "break"


def _do_plain_paste(sheet_obj, tk_root, headers_only_cols: Optional[int] = None):
    """
    Ctrl+V behavior: paste at the current selection anchor (plain-text TSV/CSV).
    """
    r0, c0 = _selected_anchor(sheet_obj)
    return _do_plain_paste_at(sheet_obj, tk_root, r0, c0, headers_only_cols=headers_only_cols)


# =========================
# Bindings
# =========================
def bind_plaintext_paste(sheet_obj, tk_root, headers_only_cols: Optional[int] = None, *, save_callback=None):
    """
    Bind BOTH Ctrl+V and Ctrl+Shift+V to the plaintext paste routine across all subwidgets.
    Also auto-saves (if save_callback is provided) after a successful paste.
    """
    def _paste_plain_evt(_evt=None):
        res = _do_plain_paste(sheet_obj, tk_root, headers_only_cols=headers_only_cols)
        try:
            if callable(save_callback):
                save_callback()
        except Exception:
            pass
        return res  # must return "break" to stop default paste

    for w in filter(None, (sheet_obj, getattr(sheet_obj, "MT", None),
                           getattr(sheet_obj, "RI", None), getattr(sheet_obj, "CH", None),
                           getattr(sheet_obj, "Toplevel", None))):
        try:
            w.unbind("<Control-v>"); w.unbind("<Control-V>")
            w.unbind("<Control-Shift-v>"); w.unbind("<Control-Shift-V>")
        except Exception:
            pass
        try:
            w.bind("<Control-v>", _paste_plain_evt)
            w.bind("<Control-V>", _paste_plain_evt)
            w.bind("<Control-Shift-v>", _paste_plain_evt)
            w.bind("<Control-Shift-V>", _paste_plain_evt)
        except Exception:
            pass


def ensure_rc_menu_plain(
    sheet_obj,
    tk_root,
    headers_only_cols: Optional[int] = None,  # kept for API compat; unused now
    *,
    save_callback=None,
):
    """
    Custom right-click menu WITHOUT Paste (to avoid row-1 paste bug for now).
    Offers: Copy / Add 100 Rows / Save Sheet / Delete Row / Undo.
    """
    import tkinter as tk

    # Remove tksheet's default RC menu to avoid conflicts
    try:
        sheet_obj.disable_bindings(("right_click_popup_menu",))
    except Exception:
        pass

    def _popup(evt):
        # Visual feedback: move selection to clicked cell if we can resolve it
        try:
            cell = _event_to_cell(sheet_obj, evt)
            if cell:
                r_clicked, c_clicked = cell
                if hasattr(sheet_obj, "select_cell"):
                    sheet_obj.select_cell(r_clicked, c_clicked)
                else:
                    sheet_obj.set_currently_selected(r_clicked, c_clicked)
        except Exception:
            pass

        def _copy_from_menu():
            try:
                sheet_obj.copy()
            except Exception:
                pass

        def _add_100_rows():
            try:
                total = sheet_obj.get_total_rows()
                sheet_obj.insert_rows(total, number_of_rows=100)
                sheet_obj.refresh()
            except Exception:
                try:
                    total = sheet_obj.get_total_rows()
                    sheet_obj.insert_rows(total, amount=100)
                    sheet_obj.refresh()
                except Exception:
                    pass

        def _delete_row_under_cursor():
            rr = None
            try:
                rr, _ = sheet_obj.get_currently_selected()
            except Exception:
                rr = None
            if rr is None or rr < 0:
                return
            try:
                sheet_obj.delete_rows(rr, 1)
            except Exception:
                try:
                    sheet_obj.delete_rows(rr)
                except Exception:
                    try:
                        sheet_obj.del_rows(rr, 1)
                    except Exception:
                        pass
            try:
                sheet_obj.refresh()
            except Exception:
                pass

        def _undo_from_menu():
            try:
                sheet_obj.undo()
                sheet_obj.refresh()
            except Exception:
                pass

        m = tk.Menu(sheet_obj.MT, tearoff=0)
        # NOTE: Paste intentionally removed
        m.add_command(label="Copy", command=_copy_from_menu)
        m.add_separator()
        m.add_command(label="Add 100 Rows", command=_add_100_rows)
        if save_callback:
            m.add_command(label="Save Sheet", command=save_callback)
        m.add_command(label="Delete Row", command=_delete_row_under_cursor)
        m.add_command(label="Undo", command=_undo_from_menu)
        try:
            m.tk_popup(evt.x_root, evt.y_root)
        finally:
            try:
                m.grab_release()
            except Exception:
                pass

    # Bind popup on all relevant widgets so a right-click anywhere anchors correctly
    for w in filter(None, (sheet_obj, getattr(sheet_obj, "MT", None),
                           getattr(sheet_obj, "RI", None), getattr(sheet_obj, "CH", None),
                           getattr(sheet_obj, "top_left_corner", None))):
        try:
            w.bind("<Button-3>", _popup)
            w.bind("<Button-2>", _popup)  # middle button fallback
        except Exception:
            pass


# =========================
# Column resizing helpers
# =========================
def enable_column_resizing(sheet_obj):
    wanted = (
        "column_width_resize",
        "column_resize",
        "resize_columns",
        "drag_select",
        "column_drag_and_drop",
    )
    try:
        sheet_obj.enable_bindings(wanted)
        return
    except Exception:
        pass
    for fl in wanted:
        try:
            sheet_obj.enable_bindings((fl,))
        except Exception:
            pass


# =========================
# Keyboard navigation
# =========================
def enable_keyboard_nav(sheet_obj):
    try:
        total_rows = sheet_obj.get_total_rows()
    except Exception:
        total_rows = 0
    try:
        total_cols = sheet_obj.get_total_columns()
    except Exception:
        try:
            total_cols = len(getattr(sheet_obj, "headers", []))
        except Exception:
            total_cols = 0

    def _clamp(v, lo, hi):
        return max(lo, min(hi, v))

    def _current():
        try:
            r, c = sheet_obj.get_currently_selected()
        except Exception:
            r, c = 0, 0
        r = _clamp(r, 0, max(0, total_rows - 1))
        c = _clamp(c, 0, max(0, total_cols - 1))
        return r, c

    def _select(r, c):
        r = _clamp(r, 0, max(0, total_rows - 1))
        c = _clamp(c, 0, max(0, total_cols - 1))
        for fn in (getattr(sheet_obj, "select_cell", None),
                   getattr(sheet_obj, "set_currently_selected", None)):
            if callable(fn):
                try:
                    fn(r, c)
                    break
                except Exception:
                    pass
        try:
            sheet_obj.see(r, c); sheet_obj.refresh()
        except Exception:
            pass

    def _mv(dr, dc):
        r, c = _current()
        nr, nc = r + dr, c + dc
        if nc >= total_cols:
            nr += 1; nc = 0
        elif nc < 0:
            nr -= 1; nc = max(0, total_cols - 1)
        _select(nr, nc)
        return "break"

    for w in filter(None, (getattr(sheet_obj, "MT", None), sheet_obj)):
        try:
            w.bind("<Left>",  lambda e: _mv(0, -1))
            w.bind("<Right>", lambda e: _mv(0, 1))
            w.bind("<Up>",    lambda e: _mv(-1, 0))
            w.bind("<Down>",  lambda e: _mv(1, 0))
            w.bind("<Tab>",         lambda e: _mv(0, 1))
            w.bind("<ISO_Left_Tab>", lambda e: _mv(0, -1))
            w.bind("<Shift-Tab>",   lambda e: _mv(0, -1))
        except Exception:
            pass


# =========================
# Column width persistence
# =========================
def _safe_get_col_count(sheet_obj) -> int:
    try:
        return sheet_obj.get_total_columns()
    except Exception:
        try:
            return len(getattr(sheet_obj, "headers", []))
        except Exception:
            return 0


def _get_current_widths(sheet_obj, ncols: Optional[int] = None) -> List[int]:
    if ncols is None:
        ncols = _safe_get_col_count(sheet_obj)
    widths = []
    for c in range(ncols):
        w = None
        for getter in (
            getattr(sheet_obj, "column_width", None),
            getattr(getattr(sheet_obj, "MT", None), "column_width", None),
        ):
            if callable(getter):
                try:
                    w = getter(c)
                    if isinstance(w, (int, float)):
                        break
                except TypeError:
                    pass
                except Exception:
                    pass
        if not isinstance(w, (int, float)):
            w = 120
        widths.append(int(w))
    return widths


def apply_column_widths(sheet_obj, widths: List[int]):
    for c, w in enumerate(widths):
        try:
            sheet_obj.column_width(c, width=int(w))
        except Exception:
            try:
                sheet_obj.MT.column_width(c, width=int(w))
            except Exception:
                pass
    try:
        sheet_obj.refresh()
    except Exception:
        pass


def _load_prefs_json(p: Path) -> dict:
    try:
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def _save_prefs_json(p: Path, data: dict):
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def load_column_widths(pref_path: Path, key: str) -> List[int]:
    data = _load_prefs_json(pref_path)
    arr = data.get(key) or []
    if isinstance(arr, list):
        out = []
        for x in arr:
            try:
                out.append(int(float(x)))
            except Exception:
                pass
        return out
    return []


def save_column_widths(pref_path: Path, key: str, widths: List[int]):
    data = _load_prefs_json(pref_path)
    data[key] = [int(w) for w in widths]
    _save_prefs_json(pref_path, data)


def attach_column_width_persistence(sheet_obj, pref_path: Path, key: str, ncols: Optional[int] = None):
    last = {"w": []}

    def _maybe_save(_evt=None):
        try:
            current = _get_current_widths(sheet_obj, ncols)
        except Exception:
            current = []
        if current and current != last.get("w"):
            save_column_widths(pref_path, key, current)
            last["w"] = current

    try:
        last["w"] = _get_current_widths(sheet_obj, ncols)
    except Exception:
        last["w"] = []

    for w in filter(None, (getattr(sheet_obj, "MT", None), sheet_obj)):
        try:
            w.bind("<ButtonRelease-1>", _maybe_save)
        except Exception:
            pass


def restore_column_widths(sheet_obj, pref_path: Path, key: str, ncols: Optional[int] = None):
    try:
        widths = load_column_widths(pref_path, key)
    except Exception:
        widths = []
    if widths:
        try:
            apply_column_widths(sheet_obj, widths[: (ncols or len(widths))])
        except Exception:
            pass


def persist_widths_now(sheet_obj, pref_path: Path, key: str, ncols: Optional[int] = None):
    try:
        widths = _get_current_widths(sheet_obj, ncols)
        if widths:
            save_column_widths(pref_path, key, widths)
    except Exception:
        pass


# =========================
# One-call wiring utility
# =========================
def wire_sheet_defaults(
    sheet_obj,
    tk_root,
    headers_only_cols: Optional[int] = None,
    *,
    pref_path: Optional[Path] = None,
    persist_key: Optional[str] = None,
    ncols: Optional[int] = None,
    save_callback=None,
):
    """
    Set up a grid with:
      - Plaintext paste (Ctrl+V / Ctrl+Shift+V); optional headers-only cap
      - Right-click menu (NO Paste; Copy/Add 100 Rows/Save/Del Row/Undo)
      - Column resizing
      - Keyboard nav
      - (Optional) Column width persistence if pref_path + persist_key are provided
    """
    bind_plaintext_paste(
        sheet_obj,
        tk_root,
        headers_only_cols=headers_only_cols,
        save_callback=save_callback,  # autosave after Ctrl+V
    )
    ensure_rc_menu_plain(
        sheet_obj,
        tk_root,
        headers_only_cols=headers_only_cols,
        save_callback=save_callback,  # keep Save in menu
    )
    enable_column_resizing(sheet_obj)
    enable_keyboard_nav(sheet_obj)

    if pref_path is not None and persist_key:
        try:
            restore_column_widths(sheet_obj, pref_path, persist_key, ncols=ncols)
        except Exception:
            pass
        try:
            attach_column_width_persistence(sheet_obj, pref_path, persist_key, ncols=ncols)
        except Exception:
            pass
