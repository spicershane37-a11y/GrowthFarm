# gf_warm.py
# Warm Leads controller (Dialer-style selection), CSV IO, live note preview,
# cross-tab effects, and tksheet UX helpers (nav, col-width persistence; NO paste here).

from __future__ import annotations
import csv
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional

# Unified store & helpers
from gf_store import (
    APP_DIR,
    WARM_LEADS_PATH,
    WARM_V2_FIELDS,
    append_no_interest,
    update_customer_row_fields_by_company,
    append_order_row,
)

# Try analytics helpers (safe fallbacks if not present)
try:
    from gf_analytics import (
        refresh_analytics,
        increment_warm_generated,
        increment_new_customer,
        log_phone_call,
    )
except Exception:
    def refresh_analytics(_window=None):  # type: ignore
        pass
    def increment_warm_generated(_n: int = 1):  # type: ignore
        pass
    def increment_new_customer(_n: int = 1):  # type: ignore
        pass
    def log_phone_call(*_args, **_kwargs):  # type: ignore
        pass

# Optional customers live-reloader
try:
    from gf_customers import reload_customers_sheet  # type: ignore
except Exception:
    reload_customers_sheet = None  # fallback if not present

# tksheet
try:
    from tksheet import Sheet
    _TKSHEET_OK = True
except Exception as _e:
    Sheet = None
    _TKSHEET_OK = False
    print(f"[gf_warm] tksheet import failed: {_e}")

# ---- keep ONLY width persistence helpers from sheet utils ----
try:
    from gf_sheet_utils import (
        load_column_widths,
        apply_column_widths,
        attach_column_width_persistence,
    )
except Exception:
    def load_column_widths(*_a, **_k): return []
    def apply_column_widths(*_a, **_k): pass
    def attach_column_width_persistence(*_a, **_k): pass

_WARM_SHEET: Optional["Sheet"] = None
_WINDOW = None

# ---- watcher for auto-refresh ----
_LAST_WARM_MTIME: Optional[float] = None
_WATCH_INTERVAL_MS = 1500
_WATCH_STARTED = False

# ---- UI padding so you always have working room ----
MIN_WARM_ROWS = 120

# ---- prefs path for column width persistence ----
COLWIDTHS_JSON: Path = APP_DIR / "prefs" / "colwidths.json"


def _pad_matrix_rows(matrix: List[List[str]], min_rows: int = MIN_WARM_ROWS) -> List[List[str]]:
    need = max(0, min_rows - len(matrix))
    if need:
        matrix += [[""] * len(WARM_V2_FIELDS) for _ in range(need)]
    return matrix


def _prime_warm_mtime():
    global _LAST_WARM_MTIME
    try:
        _LAST_WARM_MTIME = WARM_LEADS_PATH.stat().st_mtime
    except Exception:
        _LAST_WARM_MTIME = None


def _warm_csv_changed() -> bool:
    global _LAST_WARM_MTIME
    try:
        mt = WARM_LEADS_PATH.stat().st_mtime
    except Exception:
        return False
    if _LAST_WARM_MTIME is None:
        _LAST_WARM_MTIME = mt
        return False
    if mt != _LAST_WARM_MTIME:
        _LAST_WARM_MTIME = mt
        return True
    return False


def _start_warm_file_watch(window):
    global _WATCH_STARTED
    if _WATCH_STARTED:
        return
    _WATCH_STARTED = True
    _prime_warm_mtime()

    def _tick():
        try:
            if _warm_csv_changed():
                reload_warm_sheet(window)
        finally:
            try:
                window.TKroot.after(_WATCH_INTERVAL_MS, _tick)
            except Exception:
                pass

    try:
        window.TKroot.after(_WATCH_INTERVAL_MS, _tick)
    except Exception:
        pass


# ---------------------------
# CSV helpers
# ---------------------------
def _ensure_file(path: Path, header: List[str]) -> None:
    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(header)


def _ensure_warm_file_once():
    _ensure_file(WARM_LEADS_PATH, WARM_V2_FIELDS)


def load_warm_leads_matrix_v2() -> List[List[str]]:
    _ensure_warm_file_once()
    with WARM_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        return []
    hdr = rows[0]
    idx = [hdr.index(h) if h in hdr else None for h in WARM_V2_FIELDS]
    out = []
    for r in rows[1:]:
        out.append([(r[ix] if ix is not None and ix < len(r) else "") for ix in idx])
    return out


def save_warm_leads_matrix_v2(matrix: List[List[str]]) -> None:
    _ensure_warm_file_once()
    tmp = WARM_LEADS_PATH.with_suffix(".csv.tmp")
    with tmp.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(WARM_V2_FIELDS)
        for row in matrix:
            r = (list(row) + [""] * len(WARM_V2_FIELDS))[:len(WARM_V2_FIELDS)]
            w.writerow(r)
    tmp.replace(WARM_LEADS_PATH)
    _refresh_sheet_from_file_if_mounted()


# ---------------------------
# Add from Dialer
# ---------------------------
def add_warm_lead_from_dialer(row_dict: Dict[str, str], call1_note: str, ts: Optional[str] = None) -> None:
    _ensure_warm_file_once()
    ts = ts or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    base = {k: row_dict.get(k, "") for k in WARM_V2_FIELDS}
    base["Company"] = base.get("Company") or row_dict.get("Company", "")
    base["Prospect Name"] = base.get("Prospect Name") or f"{row_dict.get('First Name','')} {row_dict.get('Last Name','')}".strip()
    base["Phone #"] = base.get("Phone #") or row_dict.get("Phone", "")
    base["Email"] = base.get("Email") or row_dict.get("Email", "")

    if not base.get("Location"):
        city = row_dict.get("City", "")
        state = row_dict.get("State", "")
        base["Location"] = f"{city}, {state}".strip(", ")

    base["Industry"] = base.get("Industry") or row_dict.get("Industry", "")
    base["Google Reviews"] = base.get("Google Reviews") or row_dict.get("Reviews", "")
    base["Timestamp"] = ts
    if "First Contact" in WARM_V2_FIELDS and not base.get("First Contact"):
        base["First Contact"] = ts
    if "Call 1" in WARM_V2_FIELDS:
        base["Call 1"] = (call1_note or "").strip()

    ordered = [base.get(h, "") for h in WARM_V2_FIELDS]
    with WARM_LEADS_PATH.open("a", encoding="utf-8", newline="") as f:
        csv.writer(f).writerow(ordered)

    if _WARM_SHEET is not None:
        try:
            data = _WARM_SHEET.get_sheet_data() or []
            data.append(ordered)
            _WARM_SHEET.set_sheet_data(_pad_matrix_rows(data))
            _WARM_SHEET.refresh()
        except Exception:
            pass

    try:
        increment_warm_generated(1)
        refresh_analytics(_WINDOW)
    except Exception:
        pass


# ---------------------------
# Controller (unchanged)
# ---------------------------
class WarmController:
    def __init__(self, window, sheet):
        self.window = window
        self.sheet = sheet
        self.state = {
            "row": None,
            "outcome": None,
            "note_col_by_row": {},
            "last_focus_row": None,
        }
        self._style_outcome_buttons(active=None)
        self._update_confirm_button()

    def _row_has_payload(self, r: Optional[int]) -> bool:
        if r is None or r < 0:
            return False
        try:
            row = self.sheet.get_row_data(r) or []
        except Exception:
            row = []
        limit = min(len(row), 8)
        for i in range(limit):
            if (row[i] or "").strip():
                return True
        return False

    def _coerce_index(self, obj):
        if obj is None:
            return None
        if isinstance(obj, set):
            if not obj:
                return None
            try:
                return int(sorted(obj)[0])
            except Exception:
                try:
                    return int(next(iter(obj)))
                except Exception:
                    return None
        if isinstance(obj, (list, tuple)):
            if not obj:
                return None
            try:
                return int(obj[0])
            except Exception:
                try:
                    inner = obj[0]
                    if isinstance(inner, (list, tuple)) and inner:
                        return int(inner[0])
                except Exception:
                    return None
                return None
        try:
            return int(obj)
        except Exception:
            return None

    def _row_selected(self) -> Optional[int]:
        sh = self.sheet
        if sh is None:
            return None
        try:
            rows = sh.get_selected_rows()
            r = self._coerce_index(rows)
            if r is not None and r >= 0 and self._row_has_payload(r):
                return r
        except Exception:
            pass
        try:
            sel = sh.get_currently_selected()
            if isinstance(sel, (list, tuple)):
                r = self._coerce_index(sel[0]) if len(sel) >= 1 else None
            else:
                r = self._coerce_index(sel)
            if r is not None and r >= 0 and self._row_has_payload(r):
                return r
        except Exception:
            pass
        return None

    def _set_working_row(self, r: Optional[int]) -> None:
        if r is None or not self._row_has_payload(r):
            self.state["row"] = None
            self.state["last_focus_row"] = None
            self._style_outcome_buttons(active=None)
            self._update_confirm_button()
            return

        current_c = 0
        try:
            cur_sel = self.sheet.get_currently_selected()
            if isinstance(cur_sel, (list, tuple)) and len(cur_sel) >= 2:
                current_c = int(cur_sel[1])
        except Exception:
            current_c = 0

        changed = False
        try:
            cur_r, cur_c = self.sheet.get_currently_selected()
            if cur_r != r or cur_c != current_c:
                changed = True
        except Exception:
            changed = True

        self.state["row"] = r
        if changed:
            try:
                self.sheet.set_currently_selected(r, current_c)
            except Exception:
                pass

        self.state["last_focus_row"] = r
        self._update_confirm_button()

    def _style_outcome_buttons(self, active: Optional[str]) -> None:
        spec = [
            ("green", "-WARM_SET_GREEN-", ("white", "#2E7D32")),
            ("gray",  "-WARM_SET_GRAY-",  ("black", "#DDDDDD")),
            ("red",   "-WARM_SET_RED-",   ("white", "#C62828")),
        ]
        for name, key, normal_colors in spec:
            try:
                self.window[key].update(button_color=normal_colors)
                try:
                    if active == name:
                        self.window[key].Widget.config(relief="sunken", bd=3)
                    else:
                        self.window[key].Widget.config(relief="raised", bd=1)
                except Exception:
                    pass
            except Exception:
                pass

    def _next_empty_call_col(self, row: int) -> Optional[int]:
        try:
            r = self.sheet.get_row_data(row) or []
        except Exception:
            return None
        try:
            c1 = WARM_V2_FIELDS.index("Call 1")
            cN = WARM_V2_FIELDS.index("Call 15")
        except ValueError:
            return None
        for c in range(c1, cN + 1):
            if c >= len(r) or not (r[c] or "").strip():
                return c
        return None

    def _current_note_text(self) -> str:
        try:
            return (self.window["-WARM_NOTE-"].get() or "").strip()
        except Exception:
            return ""

    def _confirm_enabled(self) -> bool:
        r = self.state["row"]
        if r is None:
            return False
        outcome = self.state.get("outcome")
        if not outcome:
            return False
        note_txt = self._current_note_text()
        if outcome in ("gray", "red"):
            if not note_txt:
                return False
            c = self.state["note_col_by_row"].get(r) or self._next_empty_call_col(r)
            return c is not None
        return True

    def _update_confirm_button(self) -> None:
        ok = self._confirm_enabled()
        try:
            self.window["-WARM_CONFIRM-"].update(
                disabled=not ok,
                button_color=("white", "#2E7D32" if ok else "#444444"),
            )
        except Exception:
            pass

    def tick(self):
        r = self._row_selected()
        if r is not None and r != self.state["row"]:
            self._set_working_row(r)
        elif r is None and self.state["row"] is not None and not self._row_has_payload(self.state["row"]):
            self._set_working_row(None)
        self._update_confirm_button()

    def handle_event(self, event, values) -> bool:
        if not (event and str(event).startswith("-WARM_")):
            return False

        self._set_working_row(self._row_selected())

        if event == "-WARM_SET_GREEN-":
            if self.state["row"] is None:
                self._hint("Pick a row first.")
                return True
            self.state["outcome"] = "green"
            self._style_outcome_buttons("green")
            self._update_confirm_button()
            return True

        if event == "-WARM_SET_GRAY-":
            if self.state["row"] is None:
                self._hint("Pick a row first.")
                return True
            self.state["outcome"] = "gray"
            self._style_outcome_buttons("gray")
            self._update_confirm_button()
            return True

        if event == "-WARM_SET_RED-":
            if self.state["row"] is None:
                self._hint("Pick a row first.")
                return True
            self.state["outcome"] = "red"
            self._style_outcome_buttons("red")
            self._update_confirm_button()
            return True

        if event == "-WARM_NOTE-":
            if self.state["row"] is None:
                return True
            r = self.state["row"]
            c = self.state["note_col_by_row"].get(r)
            if c is None:
                c = self._next_empty_call_col(r)
                self.state["note_col_by_row"][r] = c
            if c is not None:
                note = self._current_note_text()
                try:
                    xv = self.sheet.MT.xview(); yv = self.sheet.MT.yview()
                except Exception:
                    xv = yv = None
                try:
                    self.sheet.set_cell_data(r, c, note)
                    self.sheet.refresh()
                finally:
                    try:
                        if xv: self.sheet.MT.xview_moveto(xv[0])
                        if yv: self.sheet.MT.yview_moveto(yv[0])
                    except Exception:
                        pass
            self._update_confirm_button()
            return True

        if event == "-WARM_MARK_CUSTOMER-":
            if self.state["row"] is None:
                self._hint("Pick a row first.")
                return True
            self.state["outcome"] = "green"
            self._style_outcome_buttons("green")
            self._update_confirm_button()
            return True

        if event == "-WARM_ADD100-":
            try:
                try:
                    self.sheet.insert_rows(self.sheet.get_total_rows(), number_of_rows=100)
                except Exception:
                    self.sheet.insert_rows(self.sheet.get_total_rows(), amount=100)
                self.sheet.refresh()
                matrix = _matrix_from_sheet(self.sheet, len(WARM_V2_FIELDS))
                save_warm_leads_matrix_v2(matrix)
            except Exception:
                pass
            return True

        if event == "-WARM_RELOAD-":
            reload_warm_sheet(self.window)
            return True

        if event == "-WARM_EXPORT-":
            try:
                matrix = _matrix_from_sheet(self.sheet, len(WARM_V2_FIELDS))
                save_warm_leads_matrix_v2(matrix)
                import PySimpleGUI as sg
                sg.popup_ok("Warm leads saved.", keep_on_top=True)
            except Exception:
                pass
            return True

        if event == "-WARM_CONFIRM-":
            r = self.state["row"]
            if r is None:
                self._hint("Pick a row first.")
                return True

            outcome = (self.state.get("outcome") or "gray").lower()
            note_text = self._current_note_text()

            try:
                row_vals = self.sheet.get_row_data(r) or []
            except Exception:
                row_vals = []

            def get(name: str) -> str:
                try:
                    idx = WARM_V2_FIELDS.index(name)
                    return row_vals[idx] if idx < len(row_vals) else ""
                except ValueError:
                    return ""

            if note_text:
                c = self.state["note_col_by_row"].get(r) or self._next_empty_call_col(r)
                if c is not None:
                    try:
                        self.sheet.set_cell_data(r, c, note_text)
                        self.sheet.refresh()
                        row_vals = self.sheet.get_row_data(r) or row_vals
                    except Exception:
                        pass

            if outcome == "green":
                import PySimpleGUI as sg
                amt_s = sg.popup_get_text("Opening order amount (e.g. 199.00):",
                                          title="New Customer — Opening Order", keep_on_top=True)
                if not amt_s:
                    return True
                try:
                    amt = float(amt_s.replace("$", "").replace(",", "").strip())
                except Exception:
                    sg.popup_error("Please enter a valid number.", keep_on_top=True)
                    return True

                company = get("Company")
                cust_updates = {
                    "Company": company,
                    "Prospect Name": get("Prospect Name"),
                    "Phone #": get("Phone #"),
                    "Email": get("Email"),
                    "Industry": get("Industry"),
                    "City": (get("Location").split(",")[0].strip() if get("Location") else ""),
                    "State": (get("Location").split(",")[1].strip() if (get("Location") and "," in get("Location")) else ""),
                    "Reorder?": "Yes",
                }
                try:
                    update_customer_row_fields_by_company(company, cust_updates)
                except Exception:
                    pass

                try:
                    append_order_row(company, datetime.now().strftime("%Y-%m-%d"), f"{amt:.2f}")
                except Exception:
                    return True

                if (note_text or "").strip():
                    try:
                        log_phone_call(
                            source="Warm",
                            outcome="green",
                            note=note_text.strip(),
                            company=company,
                            email=get("Email"),
                            timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        )
                    except Exception:
                        pass

                try:
                    increment_new_customer(1)
                    refresh_analytics(self.window)
                except Exception:
                    pass

                try:
                    if reload_customers_sheet is not None:
                        reload_customers_sheet(self.window)
                except Exception:
                    pass

                self._delete_row_keep_focus(r)
                self._persist_grid()
                self._done_status("Converted to customer ✓")
                self._reset_after_confirm(r)
                return True

            if outcome == "red":
                row_dict = {
                    "Email": get("Email"),
                    "First Name": (get("Prospect Name").split(" ", 1)[0] if get("Prospect Name") else ""),
                    "Last Name": (get("Prospect Name").split(" ", 1)[1] if " " in get("Prospect Name") else ""),
                    "Company": get("Company"),
                    "Industry": get("Industry"),
                    "Phone": get("Phone #"),
                    "City": get("Location").split(",")[0].strip() if get("Location") else "",
                    "State": get("Location").split(",")[1].strip() if (get("Location") and "," in get("Location")) else "",
                    "Website": get("Website"),
                }
                try:
                    append_no_interest(row_dict, note_text, no_contact_flag=0, source="Warm")
                except Exception:
                    pass

                try:
                    log_phone_call(
                        source="Warm",
                        outcome="red",
                        note=(note_text or "").strip(),
                        company=row_dict.get("Company", ""),
                        email=row_dict.get("Email", ""),
                        timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    )
                except Exception:
                    pass

                self._delete_row_keep_focus(r)
                self._persist_grid()
                self._done_status("Marked No Interest ✓")
                self._reset_after_confirm(r)
                return True

            self._persist_grid()
            try:
                log_phone_call(
                    source="Warm",
                    outcome="gray",
                    note=(note_text or "").strip(),
                    company=get("Company"),
                    email=get("Email"),
                    timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                )
            except Exception:
                pass
            self._done_status("Saved note ✓")
            self._reset_after_confirm(r, keep_row=True)
            return True

        return True

    def _persist_grid(self):
        try:
            matrix = _matrix_from_sheet(self.sheet, len(WARM_V2_FIELDS))
            save_warm_leads_matrix_v2(matrix)
            try:
                refresh_analytics(self.window)
            except Exception:
                pass
        except Exception:
            pass

    def _delete_row_keep_focus(self, r: int):
        try:
            self.sheet.delete_rows(r, 1)
        except Exception:
            try:
                self.sheet.delete_rows(r)
            except Exception:
                try:
                    self.sheet.del_rows(r, 1)
                except Exception:
                    pass
        try:
            self.sheet.refresh()
        except Exception:
            pass
        try:
            total = self.sheet.get_total_rows()
        except Exception:
            total = 0
        if total <= 0:
            self.state["row"] = None
        else:
            new_idx = min(r, max(0, total - 1))
            try:
                self.sheet.set_currently_selected(new_idx, 0)
                self.sheet.see(new_idx, 0)
            except Exception:
                pass
            self.state["row"] = new_idx

    def _reset_after_confirm(self, r: int, keep_row: bool = False):
        try:
            self.window["-WARM_NOTE-"].update("")
        except Exception:
            pass
        self.state["outcome"] = None
        self.state["note_col_by_row"].pop(r, None)
        self._style_outcome_buttons(active=None)
        self._update_confirm_button()

    def _done_status(self, msg: str):
        for k in ("-WARM_STATUS-", "-WARM_STATUS_SIDE-"):
            try:
                self.window[k].update(msg)
            except Exception:
                pass

    def _hint(self, msg: str):
        try:
            self.window["-WARM_STATUS_SIDE-"].update(msg)
        except Exception:
            pass


def _matrix_from_sheet(sheet, headers_len: int) -> List[List[str]]:
    try:
        raw = sheet.get_sheet_data() or []
    except Exception:
        return []
    out = []
    for row in raw:
        r = (list(row) + [""] * headers_len)[:len(WARM_V2_FIELDS)]
        out.append(r)
    while out and not any((c or "").strip() for c in out[-1]):
        out.pop()
    return out


def _unbind_default_paste_and_rc(sheet):
    """Remove Ctrl+V and default RC popup so logic.py's global paste owns behavior."""
    for w in filter(None, (sheet, getattr(sheet, "MT", None),
                           getattr(sheet, "RI", None), getattr(sheet, "CH", None),
                           getattr(sheet, "Toplevel", None))):
        try:
            w.unbind("<Control-v>"); w.unbind("<Control-V>")
            w.unbind("<Control-Shift-v>"); w.unbind("<Control-Shift-V>")
            w.unbind("<Command-v>"); w.unbind("<Command-V>")
        except Exception:
            pass
    try:
        sheet.disable_bindings(("right_click_popup_menu",))
    except Exception:
        pass


# --- selection wiring + autosave (plain, no sheet_utils deps) ---
def _wire_tksheet_selection(sheet, controller) -> None:
    """
    Keep controller's working row synced with tksheet selection and
    persist the grid when a cell edit finishes.
    """
    if sheet is None or controller is None:
        return

    try:
        sheet.enable_bindings((
            "single_select", "row_select", "rc_select",
            "arrowkeys", "tab_key", "shift_tab_key",
            "drag_select", "copy", "cut", "delete", "undo",
            "edit_cell", "return_edit_cell", "select_all",
            "column_width_resize", "column_resize", "resize_columns",
            "column_drag_and_drop",
        ))
    except Exception:
        pass

    def _on_cell_select(ev: dict):
        r = ev.get("row")
        try:
            r = int(r) if r is not None else None
        except Exception:
            r = None
        if r is None or r < 0:
            try:
                r = controller._row_selected()
            except Exception:
                r = None
        try:
            controller._set_working_row(r)
        except Exception:
            pass

    def _on_end_edit_cell(_ev: dict):
        # Persist entire grid to CSV on edit completion
        try:
            matrix = _matrix_from_sheet(sheet, len(WARM_V2_FIELDS))
            save_warm_leads_matrix_v2(matrix)
        except Exception:
            pass
        # Keep controller state fresh
        try:
            controller.tick()
        except Exception:
            pass

    try:
        sheet.extra_bindings([
            ("cell_select",   _on_cell_select),
            ("rc_select",     _on_cell_select),
            ("end_edit_cell", _on_end_edit_cell),
        ])
    except Exception:
        pass
# --- end selection wiring ---


def mount_warm_grid(window, start_rows=100, col_width=130):
    """Create the Warm sheet, wire UX (NO paste), and start CSV auto-watch + selection tick."""
    global _WARM_SHEET, _WINDOW, _CTL
    _WINDOW = window
    if not _TKSHEET_OK:
        return None
    host = window["-WARM_HOST-"].Widget
    try:
        for child in host.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass
    except Exception:
        pass

    import PySimpleGUI as sg  # vendored
    holder = sg.tk.Frame(host, bg="#111111")
    holder.pack(side="top", fill="both", expand=True)

    try:
        matrix = load_warm_leads_matrix_v2()
    except Exception:
        matrix = []
    matrix = _pad_matrix_rows(matrix, min_rows=MIN_WARM_ROWS)

    sheet = Sheet(holder, data=matrix, headers=WARM_V2_FIELDS, show_x_scrollbar=True, show_y_scrollbar=True)
    sheet.enable_bindings((
        "single_select", "row_select", "rc_select",
        "arrowkeys", "tab_key", "shift_tab_key",
        "drag_select", "copy", "cut", "delete", "undo",
        "edit_cell", "return_edit_cell", "select_all",
        # intentionally no "right_click_popup_menu" (we disable it below)
        "column_width_resize", "column_resize", "resize_columns"
    ))
    try:
        sheet.set_options(
            expand_sheet_if_paste_too_big=True,
            data_change_detected=True,
            show_vertical_grid=True,
            show_horizontal_grid=True,
            row_selected_background=None,
            row_selected_foreground=None,
        )
    except Exception:
        pass

    sheet.pack(fill="both", expand=True)
    for c, name in enumerate(WARM_V2_FIELDS):
        w = col_width
        if name in ("Company", "Prospect Name"):
            w = 180
        if name in ("Email", "Location", "Industry"):
            w = 160
        if name.endswith("Call"):
            w = 160
        if name in ("Cost ($)", "Timestamp", "First Contact"):
            w = 120
        try:
            sheet.column_width(c, width=w)
        except Exception:
            pass

    # NO PASTE HERE: unbind any defaults and kill RC popup
    _unbind_default_paste_and_rc(sheet)

    # Width persistence only
    try:
        saved = load_column_widths(COLWIDTHS_JSON, "warm_leads")
        if saved:
            apply_column_widths(sheet, saved)
        attach_column_width_persistence(sheet, COLWIDTHS_JSON, "warm_leads", ncols=len(WARM_V2_FIELDS))
    except Exception:
        pass

    _WARM_SHEET = sheet
    _CTL = WarmController(window, sheet)
    _wire_tksheet_selection(sheet, _CTL)

    # Start CSV watcher and selection tick
    try:
        _start_warm_file_watch(window)
    except Exception:
        pass

    def _tick_wrap():
        try:
            if _CTL:
                _CTL.tick()
        finally:
            try:
                window.TKroot.after(250, _tick_wrap)
            except Exception:
                pass

    try:
        window.TKroot.after(250, _tick_wrap)
    except Exception:
        pass

    try:
        window["-WARM_CONFIRM-"].update(disabled=True, button_color=("white", "#444444"))
    except Exception:
        pass

    return sheet


def reload_warm_sheet(window):
    """Hard reload from file into the mounted sheet (keeps persistence in place)."""
    global _WARM_SHEET, _WINDOW, _CTL
    _WINDOW = window
    if _WARM_SHEET is None:
        return mount_warm_grid(window)
    try:
        try:
            xv = _WARM_SHEET.MT.xview(); yv = _WARM_SHEET.MT.yview()
        except Exception:
            xv = yv = None

        matrix = load_warm_leads_matrix_v2()
        matrix = _pad_matrix_rows(matrix, min_rows=MIN_WARM_ROWS)

        _WARM_SHEET.set_sheet_data(matrix)
        _WARM_SHEET.refresh()

        try:
            if xv: _WARM_SHEET.MT.xview_moveto(xv[0])
            if yv: _WARM_SHEET.MT.yview_moveto(yv[0])
        except Exception:
            pass
    except Exception:
        pass
    try:
        if _CTL:
            _CTL.tick()
    except Exception:
        pass
    return _WARM_SHEET


def _refresh_sheet_from_file_if_mounted():
    global _WARM_SHEET, _CTL
    if _WARM_SHEET is None:
        return
    try:
        try:
            xv = _WARM_SHEET.MT.xview(); yv = _WARM_SHEET.MT.yview()
        except Exception:
            xv = yv = None

        matrix = load_warm_leads_matrix_v2()
        matrix = _pad_matrix_rows(matrix, min_rows=MIN_WARM_ROWS)

        _WARM_SHEET.set_sheet_data(matrix)
        _WARM_SHEET.refresh()

        try:
            if xv: _WARM_SHEET.MT.xview_moveto(xv[0])
            if yv: _WARM_SHEET.MT.yview_moveto(yv[0])
        except Exception:
            pass
    except Exception:
        pass
    try:
        if _CTL:
            _CTL.tick()
    except Exception:
        pass


def warm_handle_event(event, values, window, _context) -> bool:
    """Route -WARM_* events via the controller."""
    if not (event and str(event).startswith("-WARM_")):
        return False
    if _CTL is None or _WARM_SHEET is None:
        return True
    try:
        handled = _CTL.handle_event(event, values)
    except Exception:
        handled = True
    try:
        _CTL.tick()
    except Exception:
        pass
    return handled

