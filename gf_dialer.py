# gf_dialer.py
from __future__ import annotations
import csv
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict
from gf_analytics import log_call, increment_warm_generated


# ===== Focused debug (disabled) =====
DEBUG = False
def _dbg(*_a):  # silenced
    pass

# Shared app store + schemas (single source of truth)
from gf_store import (
    APP_DIR,
    HEADER_FIELDS,
    WARM_LEADS_PATH,      # (compat)
    WARM_V2_FIELDS,       # only for warm append shape awareness
    NO_INTEREST_PATH,     # single path for no-interest
    DIALER_LEADS_PATH,    # grid storage
    DIALER_RESULTS_PATH,  # call log
)

# Warm module: live-append & UI update when green call is confirmed
from gf_warm import add_warm_lead_from_dialer

# --------------------------------
# Small file helpers (no duplicates)
# --------------------------------
EMOJI_GREEN = "🙂"
EMOJI_GRAY  = "😐"
EMOJI_RED   = "🙁"          # canonical going forward
EMOJI_RED_LEGACY = "☹️"     # tolerate on load

def _atomic_write_csv(path: Path, headers: List[str], rows: List[List[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        for row in rows:
            w.writerow((list(row) + [""] * len(headers))[:len(headers)])
    tmp.replace(path)

# --------------------------------
# Ensure core dialer files exist
# --------------------------------
def ensure_dialer_files() -> None:
    """Ensure the dialer call log exists with correct header."""
    if not DIALER_RESULTS_PATH.exists():
        with DIALER_RESULTS_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow([
                "Timestamp","Outcome",
                "Email","First Name","Last Name",
                "Company","Industry","Phone",
                "Address","City","State",
                "Reviews","Website","Note"
            ])

def ensure_dialer_leads_file() -> None:
    """Ensure the dialer grid CSV exists with expected headers."""
    if not DIALER_LEADS_PATH.exists():
        hdr = HEADER_FIELDS + [EMOJI_GREEN, EMOJI_GRAY, EMOJI_RED] + [f"Note{i}" for i in range(1, 9)]
        _atomic_write_csv(DIALER_LEADS_PATH, hdr, [])

def _ensure_no_interest_file_once() -> None:
    """Single path (from gf_store): ensure no_interest.csv has a header."""
    if not NO_INTEREST_PATH.exists():
        NO_INTEREST_PATH.parent.mkdir(parents=True, exist_ok=True)
        with NO_INTEREST_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow([
                "Timestamp","Email","First Name","Last Name","Company","Industry",
                "Phone","City","State","Website","Note","Source","NoContact"
            ])

# --------------------------------
# Dialer grid load/save (own CSV)
# --------------------------------
def load_dialer_leads_matrix() -> List[List[str]]:
    """
    Load rows; tolerate legacy ☹️ red header; fill missing dot cells with '○'.
    """
    ensure_dialer_leads_file()
    with DIALER_LEADS_PATH.open("r", encoding="utf-8", newline="") as f:
        rows = list(csv.reader(f))
    if not rows:
        return []
    hdr = rows[0]
    expected = HEADER_FIELDS + [EMOJI_GREEN, EMOJI_GRAY, EMOJI_RED] + [f"Note{i}" for i in range(1, 9)]
    head_ix = {h: (hdr.index(h) if h in hdr else None) for h in expected}
    if head_ix[EMOJI_RED] is None and EMOJI_RED_LEGACY in hdr:
        head_ix[EMOJI_RED] = hdr.index(EMOJI_RED_LEGACY)
    idx_order = [head_ix[h] for h in expected]

    out = []
    for r in rows[1:]:
        row = []
        for i, ix in enumerate(idx_order):
            if ix is None or ix >= len(r):  # missing column
                if len(HEADER_FIELDS) <= i < len(HEADER_FIELDS) + 3:
                    row.append("○")
                else:
                    row.append("")
            else:
                val = r[ix]
                if len(HEADER_FIELDS) <= i < len(HEADER_FIELDS) + 3:
                    row.append(val or "○")
                else:
                    row.append(val)
        out.append(row)
    return out

def save_dialer_leads_matrix(matrix: List[List[str]]) -> None:
    headers = HEADER_FIELDS + [EMOJI_GREEN, EMOJI_GRAY, EMOJI_RED] + [f"Note{i}" for i in range(1, 9)]
    _atomic_write_csv(DIALER_LEADS_PATH, headers, matrix)

# --------------------------------
# Call persistence + warm / no-interest
# --------------------------------
def dialer_save_call(row_dict: Dict[str,str], outcome: str, note: str) -> None:
    """
    Persist a single call to dialer_results.csv.
    Also appends to warm_leads.csv for green calls (with Call 1 filled),
    by delegating to gf_warm.add_warm_lead_from_dialer (which live-updates the grid).
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ensure_dialer_files()
    # Call log
    with DIALER_RESULTS_PATH.open("a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            ts, outcome,
            row_dict.get("Email",""),
            row_dict.get("First Name",""), row_dict.get("Last Name",""),
            row_dict.get("Company",""), row_dict.get("Industry",""),
            row_dict.get("Phone",""),
            row_dict.get("Address",""), row_dict.get("City",""), row_dict.get("State",""),
            row_dict.get("Reviews",""), row_dict.get("Website",""),
            note
        ])

    # Warm lead on green — delegate to warm module (writes CSV + live UI if mounted)
    if (outcome or "").lower() == "green":
        try:
            add_warm_lead_from_dialer(row_dict, note, ts=ts)
        except Exception:
            # Fallback: best-effort silent; call log is still persisted above
            pass

def add_no_interest(row_dict: Dict[str,str], note: str, no_contact_flag: int, source: str) -> None:
    """Append to no_interest.csv (single route using gf_store.NO_INTEREST_PATH)."""
    _ensure_no_interest_file_once()
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with NO_INTEREST_PATH.open("a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow([
            ts,
            row_dict.get("Email",""),
            row_dict.get("First Name",""), row_dict.get("Last Name",""),
            row_dict.get("Company",""), row_dict.get("Industry",""),
            row_dict.get("Phone",""),
            row_dict.get("City",""), row_dict.get("State",""),
            row_dict.get("Website",""),
            note, source, int(no_contact_flag or 0)
        ])

# =====================================================
# Dialer Controller — owns the Dialer tab UI behavior
# =====================================================
class DialerController:
    """Owns all Dialer tab behavior, coloring, and persistence."""

    # dot colors (single cell background preview only)
    DOT_BG = {"green": "#2E7D32", "gray": "#BDBDBD", "red": "#C62828"}  # preview color
    DOT_FG = {"green": "#FFFFFF", "gray": "#000000", "red": "#FFFFFF"}

    # full-row persisted colors (APPLIED ONLY ON CONFIRM)
    ROW_BG = {"green": "#1f3d2a", "gray": "#d9d9d9", "red": "#3d1f1f"}   # gray lighter
    ROW_FG = {"green": "#ffffff", "gray": "#000000", "red": "#ffffff"}

    SOFT_BLUE = "#CCE5FF"  # tksheet shows selection blue; we don't paint blue ourselves

    def __init__(self, window, sheet, header_fields=None):
        self.window = window
        self.sheet = sheet
        self.header_fields = header_fields or HEADER_FIELDS
        self.cols = self._cols_info()
        self.state = {
            "row": None,
            "outcome": None,            # current pending disposition (preview only)
            "note_col_by_row": {},      # row -> reserved "NoteX" slot (for live preview)
            "last_focus_row": None,
            "gray_rows": set(),         # rows persisted as gray (confirmed)
            "row_preview_outcome": {},  # row -> preview intent (no row tint)
        }
        # initialize outcome button visuals as "none selected"
        self._style_outcome_buttons(active=None)
        # paint any previously gray rows
        self.repaint_all_rows()

    # ---------- layout helpers ----------
    def _cols_info(self):
        first_dot = len(self.header_fields)
        last_dot = first_dot + 2
        first_note = len(self.header_fields) + 3
        last_note = first_note + 7
        return {"first_dot": first_dot, "last_dot": last_dot, "first_note": first_note, "last_note": last_note}

    # ---------- payload test ----------
    def _row_has_payload(self, r: Optional[int]) -> bool:
        if r is None or r < 0:
            return False
        try:
            row = self.sheet.get_row_data(r) or []
        except Exception:
            row = []
        limit = min(len(row), len(self.header_fields))
        for i in range(limit):
            if (row[i] or "").strip():
                return True
        return False

    # --- selection coercion ---
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

    # ----- scroll helper: vertical-only see (keeps horizontal xview) -----
    def _see_row_vert_only(self, r: int) -> None:
        try:
            xv = self.sheet.MT.xview()
        except Exception:
            xv = None
        try:
            self.sheet.see(r, 0)
        except Exception:
            pass
        try:
            if xv:
                self.sheet.MT.xview_moveto(xv[0])
        except Exception:
            pass

    # ----- base paint helper (white or persisted gray) -----
    def _apply_base_row_paint(self, r: int) -> None:
        """Apply the correct non-preview paint for a row."""
        try:
            middle_dot = (self.sheet.get_cell_data(r, self.cols["first_dot"] + 1) or "").strip()
        except Exception:
            middle_dot = ""
        is_gray = (middle_dot == "●")
        try:
            if is_gray:
                self.sheet.highlight_rows(rows=[r], bg=self.ROW_BG["gray"], fg=self.ROW_FG["gray"])
                self.state["gray_rows"].add(r)
            else:
                self.sheet.highlight_rows(rows=[r], bg=None, fg=None)
                self.state["gray_rows"].discard(r)
        except Exception:
            pass

    # ----- outcome button visuals -----
    def _style_outcome_buttons(self, active: Optional[str]) -> None:
        """
        Give the selected outcome button a 'sunken' look and reset others to 'raised'.
        active ∈ {"green","gray","red", None}
        """
        spec = [
            ("green", "-DIAL_SET_GREEN-", ("white", "#2E7D32")),
            ("gray",  "-DIAL_SET_GRAY-",  ("black", "#DDDDDD")),
            ("red",   "-DIAL_SET_RED-",   ("white", "#C62828")),
        ]
        for name, key, normal_colors in spec:
            try:
                # keep original colors; just toggle relief/border for "pressed" feel
                self.window[key].update(button_color=normal_colors)
                try:
                    if active == name:
                        self.window[key].Widget.config(relief="sunken", bd=3)
                    else:
                        self.window[key].Widget.config(relief="raised", bd=1)
                except Exception:
                    # Some themes/widgets may not expose .Widget; color still updates
                    pass
            except Exception:
                pass

    # ----- preview revert (only clears backgrounds; we never changed symbols) -----
    def _revert_preview_on_row(self, r: Optional[int]) -> None:
        if r is None:
            return
        base = self.cols["first_dot"]
        # Clear background highlight on the three dot cells
        try:
            for i in range(3):
                c = base + i
                try:
                    self.sheet.highlight_cells(row=r, column=c, bg=None, fg=None)
                except Exception:
                    pass
        except Exception:
            pass
        # Clear temp preview note
        c = self.state["note_col_by_row"].get(r)
        if c is not None:
            try:
                self.sheet.set_cell_data(r, c, "")
            except Exception:
                pass
            self.state["note_col_by_row"].pop(r, None)
        # Clear intent flag & outcome
        self.state["row_preview_outcome"].pop(r, None)
        if self.state.get("row") == r:
            self.state["outcome"] = None
        # Restore base row paint
        self._apply_base_row_paint(r)
        try:
            self.sheet.refresh()
        except Exception:
            pass
        # Reset the button visuals (none selected)
        self._style_outcome_buttons(active=None)

    def _set_working_row(self, r: Optional[int]) -> None:
        prev = self.state.get("last_focus_row")
        if prev is not None and (r != prev):
            self._revert_preview_on_row(prev)

        if r is None or not self._row_has_payload(r):
            self.state["row"] = None
            self.state["last_focus_row"] = None
            # also clear button press state
            self._style_outcome_buttons(active=None)
            return

        self.state["row"] = r
        try:
            self.sheet.set_currently_selected(r, 0)
        except Exception:
            pass
        self._see_row_vert_only(r)
        self.state["last_focus_row"] = r  # blue selection is handled by tksheet

    def repaint_all_rows(self) -> None:
        try:
            total = self.sheet.get_total_rows()
        except Exception:
            total = 0

        self.state["gray_rows"].clear()
        for r in range(total):
            # clear lingering dot bg highlights
            for i in range(3):
                try:
                    self.sheet.highlight_cells(row=r, column=self.cols["first_dot"] + i, bg=None, fg=None)
                except Exception:
                    pass
            self._apply_base_row_paint(r)

        # No preview row tint at all (by design)
        try:
            self.sheet.refresh()
        except Exception:
            pass

    # ---------- preview helpers ----------
    def _preview_dot_only(self, row: int, outcome: str) -> None:
        """Highlight ONLY the dot cell background; do not change symbols or row tint."""
        base = self.cols["first_dot"]
        idx = {"green": 0, "gray": 1, "red": 2}[outcome]
        try:
            # clear any previous dot highlights
            for i in range(3):
                try:
                    self.sheet.highlight_cells(row=row, column=base + i, bg=None, fg=None)
                except Exception:
                    pass
            # apply background preview to the chosen dot
            c = base + idx
            self.sheet.highlight_cells(row=row, column=c, bg=self.DOT_BG[outcome], fg=self.DOT_FG[outcome])
            self.sheet.refresh()
        except Exception:
            pass
        self.state["row_preview_outcome"][row] = outcome
        # reflect selection in the buttons
        self._style_outcome_buttons(active=outcome)

    def _next_empty_note_col(self, row: int) -> Optional[int]:
        try:
            r = self.sheet.get_row_data(row) or []
        except Exception:
            return None
        for c in range(self.cols["first_note"], self.cols["last_note"] + 1):
            if c >= len(r) or not (r[c] or "").strip():
                return c
        return None

    def _move_to_next_row(self, current_row: int) -> int:
        try:
            total = self.sheet.get_total_rows()
        except Exception:
            total = 0
        nxt = current_row + 1 if total == 0 else min(current_row + 1, max(0, total - 1))
        try:
            self.sheet.set_currently_selected(nxt, 0)
        except Exception:
            pass
        self._see_row_vert_only(nxt)
        self._set_working_row(nxt if self._row_has_payload(nxt) else None)
        return nxt

    def _save_grid_csv(self) -> None:
        try:
            data = self.sheet.get_sheet_data() or []
        except Exception:
            data = []
        expected = len(self.header_fields) + 3 + 8
        matrix = []
        for row in data:
            r = (list(row) + [""] * expected)[:expected]
            # ensure unfilled dots are hollow; we never set them during preview
            for i in range(len(self.header_fields), len(self.header_fields) + 3):
                if not (r[i] or "").strip():
                    r[i] = "○"
            matrix.append(r)
        save_dialer_leads_matrix(matrix)
        self.repaint_all_rows()

    def _current_note_text(self) -> str:
        try:
            v = (self.window["-DIAL_NOTE-"].get() or "").strip()
            return v
        except Exception:
            return ""

    def _confirm_enabled(self) -> bool:
        r = self.state["row"]
        if r is None:
            return False
        sticky = self.state["note_col_by_row"].get(r)
        have_slot = (sticky is not None) or (self._next_empty_note_col(r) is not None)
        ok = (self.state["outcome"] in ("green", "gray", "red")) and bool(self._current_note_text()) and have_slot
        return ok

    def _update_confirm_button(self) -> None:
        ok = self._confirm_enabled()
        try:
            self.window["-DIAL_CONFIRM-"].update(
                disabled=not ok,
                button_color=("white", "#2E7D32" if ok else "#444444"),
            )
        except Exception:
            pass

    # ---------- public: selection follower ----------
    def tick(self):
        r = self._row_selected()
        if r is not None and r != self.state["row"]:
            self._set_working_row(r)
        elif r is None and self.state["row"] is not None and not self._row_has_payload(self.state["row"]):
            self._set_working_row(None)
        self._update_confirm_button()

    # ---------- public: event router ----------
    def handle_event(self, event, values) -> bool:
        if not (event and str(event).startswith("-DIAL_")):
            return False

        # Re-sync working row BEFORE handling
        r_live = None
        try:
            r_live = self._row_selected()
        except Exception:
            pass
        self._set_working_row(r_live)  # will clear if None/empty

        # Disposition buttons (PREVIEW: dot background only)
        if event == "-DIAL_SET_GREEN-":
            if self.state["row"] is None:
                self.window["-DIAL_MSG-"].update("Pick a row first.")
                return True
            self.state["outcome"] = "green"
            self._preview_dot_only(self.state["row"], "green")
            self._update_confirm_button()
            return True

        if event == "-DIAL_SET_GRAY-":
            if self.state["row"] is None:
                self.window["-DIAL_MSG-"].update("Pick a row first.")
                return True
            self.state["outcome"] = "gray"
            self._preview_dot_only(self.state["row"], "gray")
            self._update_confirm_button()
            return True

        if event == "-DIAL_SET_RED-":
            if self.state["row"] is None:
                self.window["-DIAL_MSG-"].update("Pick a row first.")
                return True
            self.state["outcome"] = "red"
            self._preview_dot_only(self.state["row"], "red")
            self._update_confirm_button()
            return True

        # Live note preview (writes into a reserved Note cell, but that's transient)
        if event == "-DIAL_NOTE-":
            if self.state["row"] is None:
                return True
            r = self.state["row"]
            c = self.state["note_col_by_row"].get(r)
            if c is None:
                c = self._next_empty_note_col(r)
                self.state["note_col_by_row"][r] = c
            if c is not None:
                try:
                    xv = self.sheet.MT.xview(); yv = self.sheet.MT.yview()
                except Exception:
                    xv = yv = None
                try:
                    self.sheet.set_cell_data(r, c, self._current_note_text())
                    self.sheet.refresh()
                finally:
                    try:
                        if xv: self.sheet.MT.xview_moveto(xv[0])
                        if yv: self.sheet.MT.yview_moveto(yv[0])
                    except Exception:
                        pass
            self._update_confirm_button()
            return True

        # Confirm/save
        if event == "-DIAL_CONFIRM-":
            if self.state["row"] is None:
                self.window["-DIAL_MSG-"].update("Pick a row first.")
                return True

            r = self.state["row"]
            note_text = self._current_note_text()
            if not note_text:
                self.window["-DIAL_MSG-"].update("Type a note.")
                return True

            outcome = (self.state["outcome"] or "gray").lower()

            # FINALIZE DOT SYMBOLS (set ● on the chosen dot), and clear any preview bg
            base = self.cols["first_dot"]
            try:
                for i in range(3):
                    c = base + i
                    try:
                        self.sheet.highlight_cells(row=r, column=c, bg=None, fg=None)
                    except Exception:
                        pass
                    self.sheet.set_cell_data(r, c, "○")
                c_idx = {"green": 0, "gray": 1, "red": 2}[outcome]
                self.sheet.set_cell_data(r, base + c_idx, "●")
            except Exception:
                pass

            # Persist gray row tint; green/red rows will be deleted so no need to tint
            if outcome == "gray":
                try:
                    self.sheet.highlight_rows(rows=[r], bg=self.ROW_BG["gray"], fg=self.ROW_FG["gray"])
                except Exception:
                    pass
                self.state["gray_rows"].add(r)

            # Finalize the preview note into the row
            c = self.state["note_col_by_row"].get(r)
            if c is None:
                c = self._next_empty_note_col(r)
            if c is not None:
                try:
                    xv = self.sheet.MT.xview(); yv = self.sheet.MT.yview()
                except Exception:
                    xv = yv = None
                try:
                    self.sheet.set_cell_data(r, c, note_text)
                    self.sheet.refresh()
                finally:
                    try:
                        if xv: self.sheet.MT.xview_moveto(xv[0])
                        if yv: self.sheet.MT.yview_moveto(yv[0])
                    except Exception:
                        pass

            # Build base row dict from the sheet row (HEADER_FIELDS only)
            try:
                row_vals = self.sheet.get_row_data(r) or []
            except Exception:
                row_vals = []
            base_row = {h: (row_vals[i] if i < len(row_vals) else "") for i, h in enumerate(self.header_fields)}

            # Persist call + side effects
            try:
                dialer_save_call(base_row, outcome, note_text)

                # --- NEW: analytics call log (CSV + daily counters) ---
                try:
                    log_call(
                        source="dialer",
                        outcome=outcome,                    # "green" | "gray" | "red"
                        note=note_text,
                        company=base_row.get("Company", ""),
                        prospect=base_row.get("Prospect Name", ""),
                        email=base_row.get("Email", ""),
                        phone=base_row.get("Phone #", ""),
                    )
                except Exception as _e:
                    print("Call logging failed:", _e)

                # If GREEN, we consider that a warm lead was generated → bump pipeline counter
                if outcome == "green":
                    try:
                        increment_warm_generated(1)
                    except Exception as _e:
                        print("Warm counter bump failed:", _e)

                # Existing “no interest” logic
                if outcome == "red":
                    add_no_interest(base_row, note_text, no_contact_flag=0, source="Dialer")
                elif outcome == "gray":
                    # if 8 notes filled, flag as No Contact
                    filled = 0
                    row_vals2 = self.sheet.get_row_data(r) or []
                    for k in range(self.cols["first_note"], self.cols["last_note"] + 1):
                        if k < len(row_vals2) and (row_vals2[k] or "").strip():
                            filled += 1
                    if filled >= 8:
                        add_no_interest(base_row, "No Contact after 8 calls. " + note_text, no_contact_flag=1, source="Dialer")

            except Exception as e:
                self.window["-DIAL_MSG-"].update(f"Save error: {e}")
                return True

            self.window["-DIAL_MSG-"].update("Saved ✓")
            try:
                self.window["-DIAL_NOTE-"].update("")
            except Exception:
                pass
            self.state["outcome"] = None
            self.state["note_col_by_row"].pop(r, None)
            self.state["row_preview_outcome"].pop(r, None)

            # Reset button visuals after a confirm
            self._style_outcome_buttons(active=None)

            # Remove row on green/red, keep (persist gray color) on gray
            if outcome in ("green", "red"):
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
                self._save_grid_csv()
                self.repaint_all_rows()
                try:
                    total = self.sheet.get_total_rows()
                except Exception:
                    total = 0
                if total <= 0:
                    self.state["row"] = None
                else:
                    new_idx = min(r, max(0, total - 1))
                    self._set_working_row(new_idx if self._row_has_payload(new_idx) else None)
            else:
                # Persist grid and advance to next row
                self._save_grid_csv()
                self._move_to_next_row(r)

            self._update_confirm_button()
            return True

        if event == "-DIAL_ADD100-":
            try:
                add = [[""] * len(self.header_fields) + ["○", "○", "○"] + (([""] * 8)) for _ in range(100)]
                try:
                    cur = self.sheet.get_sheet_data() or []
                except Exception:
                    cur = []
                self.sheet.set_sheet_data((cur or []) + add)
                self.sheet.refresh()
                self._save_grid_csv()
                self.repaint_all_rows()
            except Exception as e:
                self.window["-DIAL_MSG-"].update(f"Add rows error: {e}")
            return True

        return True

# --- tksheet selection wiring ---
def _wire_tksheet_selection(sheet, controller: "DialerController") -> None:
    if sheet is None or controller is None:
        return
    try:
        sheet.enable_bindings((
            "single_select",
            "row_select",
            "arrowkeys",
            "tab_key", "shift_tab_key",
            "drag_select",
            "copy", "cut", "paste",
            "delete", "undo",
            "edit_cell", "return_edit_cell",
            "select_all",
            "rc_select",
            "right_click_popup_menu",
            "column_width_resize", "column_resize", "resize_columns",
            "column_drag_and_drop",
        ))
    except Exception:
        pass

    def _on_cell_select(ev: dict):
        r = ev.get("row")
        if isinstance(r, int) and r >= 0 and controller._row_has_payload(r):
            controller._set_working_row(r)
        else:
            controller._set_working_row(None)

    def _on_data_change(_ev=None):
        controller.repaint_all_rows()

    try:
        sheet.extra_bindings([
            ("cell_select", _on_cell_select),
            ("data_change", _on_data_change),
        ])
        return
    except Exception:
        pass

    # Fallback: raw canvas click
    try:
        def _on_click(event):
            try:
                r = sheet.identify_row(event.y)
            except Exception:
                r = None
            if isinstance(r, int) and r >= 0 and controller._row_has_payload(r):
                controller._set_working_row(r)
            else:
                controller._set_working_row(None)
            controller.repaint_all_rows()
        sheet.MT.bind("<Button-1>", _on_click, add="+")
    except Exception:
        pass

# ---------------
# Simple factory
# ---------------
def attach_dialer(window, dial_sheet) -> DialerController:
    ensure_dialer_files()
    ensure_dialer_leads_file()
    ctrl = DialerController(window, dial_sheet, HEADER_FIELDS)
    _wire_tksheet_selection(dial_sheet, ctrl)
    try:
        ctrl.repaint_all_rows()
    except Exception:
        pass
    return ctrl
