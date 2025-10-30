# gf_ui_logic.py
# Mounts all grids (Leads, Dialer, Warm, Customers) and drives the event loop.

from __future__ import annotations

import sys, os, csv
from pathlib import Path
from datetime import datetime

from gf_map import open_customer_map

from gf_store import (
    APP_DIR,
    ensure_app_files,
    # Email Leads
    EMAIL_LEADS_PATH,
    HEADER_FIELDS,
    load_email_leads_matrix,
    save_email_leads_matrix,
    # Warm v2
    WARM_LEADS_PATH,
    WARM_V2_FIELDS,
    load_warm_leads_matrix_v2,
    save_warm_leads_matrix_v2,
    # Customers
    CUSTOMERS_PATH,
    CUSTOMER_FIELDS,
    load_customers_matrix,
    save_customers_matrix,
    append_order_row,
    # Results (for “Emails Sent” and campaign resp% / results UI)
    RESULTS_PATH,
    load_results_rows_sorted,
)

# Analytics (right-side panels + pipeline counters)
from gf_analytics import init_analytics
from gf_analytics import increment_warm_generated, increment_new_customer  # noqa - imported elsewhere

# Campaigns
from gf_campaigns import (
    list_campaign_keys,
    load_campaign_by_key,
    save_campaign_by_key,
    delete_campaign_by_key,
    summarize_campaign_for_table,
    normalize_campaign_steps,
    normalize_campaign_settings,
)

# Dialer (controller owns its own coloring/preview logic)
from gf_dialer import (
    attach_dialer,
    load_dialer_leads_matrix,
    save_dialer_leads_matrix,
    EMOJI_GREEN, EMOJI_GRAY, EMOJI_RED,
)

# Warm module owns its own grid + events
from gf_warm import (
    mount_warm_grid,
    reload_warm_sheet,
    warm_handle_event,
)

# ---- Column width persistence helpers
from gf_sheet_utils import (
    load_column_widths,
    apply_column_widths,
    attach_column_width_persistence,
    # (do NOT import bind_plaintext_paste anymore; we implement it locally to fix anchor)
)

# --- updater (safe import) ---
try:
    from gf_updater import update_ui_flow  # performs check/download/install + UI popups
except Exception as _upd_err:
    def update_ui_flow(window=None):
        # Fallback so the app doesn't crash if gf_updater.py is missing
        try:
            import PySimpleGUI as _sg
            _sg.popup_error(
                "Updater module not found. Please make sure gf_updater.py is in the app folder.",
                title="Update",
                keep_on_top=True,
            )
        except Exception:
            pass

# --- Force vendored PySimpleGUI 4.60.x ---
from pathlib import Path as _P
_VENDOR_PSG = _P(__file__).parent / "vendor_psg"
if str(_VENDOR_PSG) not in sys.path:
    sys.path.insert(0, str(_VENDOR_PSG))
import PySimpleGUI as sg  # noqa
# -----------------------------------------

# tksheet import (friendly message if missing)
try:
    from tksheet import Sheet
    _TKSHEET_OK = True
except Exception as _e:
    Sheet = None
    _TKSHEET_OK = False
    print(f"[gf_ui_logic] tksheet import failed: {_e}")

# ==============================
# Small helpers
# ==============================
_SOFT_BLUE = "#CCE5FF"

# Where we persist column widths
PREFS_DIR = APP_DIR / "prefs"
PREFS_DIR.mkdir(parents=True, exist_ok=True)
COLWIDTHS_PATH = PREFS_DIR / "column_widths.json"

_LEADS_PREF_KEY = "leads_colwidths"
_DIALER_PREF_KEY = "dialer_colwidths"
_CUSTOMERS_PREF_KEY = "customers_colwidths"

def _to_float_money(s):
    try:
        return float((s or "").replace("$", "").replace(",", "").strip() or "0")
    except Exception:
        return 0.0

def _parse_any_date(s):
    if not s:
        return None
    s = str(s).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S","%Y-%m-%d","%m/%d/%Y %I:%M:%S %p","%m/%d/%Y %I:%M %p","%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None

def _clear_children(tk_parent):
    try:
        for child in tk_parent.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass
    except Exception:
        pass

# ==============================
# Plain-text paste with correct selection anchor (local)
# ==============================

def _selected_anchor(sheet_obj):
    """
    Prefer the actual currently selected cell; else first selected row (col 0);
    else (0,0). Never silently coerce to row 0 if a valid selection exists.
    """
    # Try current cell first
    try:
        sel = sheet_obj.get_currently_selected()
        if isinstance(sel, (list, tuple)) and len(sel) >= 2:
            r, c = sel[0], sel[1]
            if isinstance(r, int) and r >= 0 and isinstance(c, int) and c >= 0:
                return r, c
        if isinstance(sel, tuple) and len(sel) == 2 and all(isinstance(x, int) for x in sel):
            r, c = sel
            if r >= 0 and c >= 0:
                return r, c
    except Exception:
        pass
    # Fallback: first selected row -> column 0
    try:
        rows = sheet_obj.get_selected_rows() or []
        if rows:
            r0 = int(rows[0])
            if r0 >= 0:
                return r0, 0
    except Exception:
        pass
    # Final fallback
    return 0, 0

def _parse_clipboard_text(tk_root):
    """Return rows from clipboard; prefer raw TSV (no CSV parsing), else CSV-aware fallback."""
    try:
        clip = tk_root.clipboard_get()
    except Exception:
        return []
    if not clip:
        return []

    # Normalize newlines
    clip = clip.replace("\r\n", "\n").replace("\r", "\n")

    # If it looks like tab-separated (Google Sheets / Excel), split only on tabs.
    # This preserves literal quotes/parentheses and keeps empty columns.
    if "\t" in clip:
        lines = clip.split("\n")
        # drop a single trailing empty line if present (Sheets often appends it)
        if lines and lines[-1] == "":
            lines = lines[:-1]
        return [line.split("\t") for line in lines]

    # Otherwise, fall back to CSV parsing (comma-delimited) with quote handling.
    rows = []
    for line in clip.split("\n"):
        if line == "":
            continue
        try:
            for parsed in csv.reader([line]):
                rows.append(parsed)
        except Exception:
            rows.append([line])
    return rows

def _unbind_default_paste(sheet):
    """Remove any default paste bindings on all subwidgets to avoid conflicts."""
    targets = (sheet, getattr(sheet, "MT", None), getattr(sheet, "RI", None),
               getattr(sheet, "CH", None), getattr(sheet, "Toplevel", None))
    for w in filter(None, targets):
        for seq in ("<Control-v>", "<Control-V>",
                    "<Control-Shift-v>", "<Control-Shift-V>",
                    "<Command-v>", "<Command-V>",
                    "<<Paste>>", "<Shift-Insert>", "<Control-Insert>"):
            try:
                w.unbind(seq)
            except Exception:
                pass

def _manual_plain_paste(sheet, tk_root, headers_only_cols=None):
    """Always paste as plain text at selected cell; expand rows; optional cap on columns."""
    rows = _parse_clipboard_text(tk_root)
    if not rows:
        return "break"

    r0, c0 = _selected_anchor(sheet)

    # Ensure enough rows exist
    try:
        total_rows = sheet.get_total_rows()
    except Exception:
        total_rows = 0
    need_rows = r0 + len(rows)
    if need_rows > total_rows:
        try:
            sheet.insert_rows(total_rows, number_of_rows=(need_rows - total_rows))
        except Exception:
            try:
                sheet.insert_rows(total_rows, amount=(need_rows - total_rows))
            except Exception:
                pass

    # Write cells
    for r_off, row in enumerate(rows):
        for c_off, val in enumerate(row):
            dest_c = c0 + c_off
            if headers_only_cols is not None and dest_c >= headers_only_cols:
                continue
            try:
                sheet.set_cell_data(r0 + r_off, dest_c, val)
            except Exception:
                pass

    # Restore selection & refresh
    try:
        if hasattr(sheet, "select_cell"):
            sheet.select_cell(r0, c0)
        else:
            sheet.set_currently_selected(r0, c0)
        sheet.see(r0, c0)
        sheet.refresh()
    except Exception:
        pass
    return "break"

def _bind_plaintext_paste(sheet, tkroot, *, headers_only_cols=None, save_callback=None):
    """Bind Ctrl/Cmd+V to our plain-text paste at the correct anchor, then autosave."""
    if sheet is None:
        return
    _unbind_default_paste(sheet)  # kill any defaults that might paste at (0,0)

    def _on_paste(_evt=None):
        res = _manual_plain_paste(sheet, tkroot, headers_only_cols=headers_only_cols)
        try:
            if callable(save_callback):
                save_callback()
        except Exception:
            pass
        return res

    # Bind on all subwidgets to be safe
    targets = (sheet, getattr(sheet, "MT", None), getattr(sheet, "RI", None),
               getattr(sheet, "CH", None), getattr(sheet, "Toplevel", None))
    for w in filter(None, targets):
        for seq in ("<Control-v>", "<Control-V>",
                    "<Control-Shift-v>", "<Control-Shift-V>",
                    "<Command-v>", "<Command-V>",
                    "<<Paste>>", "<Shift-Insert>", "<Control-Insert>"):
            try:
                w.bind(seq, _on_paste)
            except Exception:
                pass

# ==============================
# Analytics refresh shim
# ==============================

def _trigger_analytics_refresh(window):
    """Signal the UI loop to refresh the right-hand analytics panels."""
    try:
        window.write_event_value("-ANALYTICS_REFRESH-", True)
    except Exception:
        pass

# ==============================
# Persistence helpers
# ==============================

def _matrix_from_sheet(sheet, headers_len):
    try:
        raw = sheet.get_sheet_data() or []
    except Exception:
        return []
    out = []
    for row in raw:
        r = (list(row) + [""] * headers_len)[:headers_len]
        out.append(r)
    while out and not any((c or "").strip() for c in out[-1]):
        out.pop()
    return out

def _save_leads(sheet):
    try:
        data = _matrix_from_sheet(sheet, len(HEADER_FIELDS))
        save_email_leads_matrix(data)
    except Exception:
        pass

def _save_customers(sheet):
    try:
        data = _matrix_from_sheet(sheet, len(CUSTOMER_FIELDS))
        save_customers_matrix(data)
    except Exception:
        pass

def _save_warm(sheet):
    try:
        data = _matrix_from_sheet(sheet, len(WARM_V2_FIELDS))
        save_warm_leads_matrix_v2(data)
    except Exception:
        pass

def _save_all(context):
    """Persist all grids best-effort."""
    try:
        if context.get("sheet"):
            _save_leads(context["sheet"])
    except Exception:
        pass
    try:
        if context.get("customer_sheet"):
            _save_customers(context["customer_sheet"])
    except Exception:
        pass
    try:
        if context.get("warm_sheet"):
            _save_warm(context["warm_sheet"])
    except Exception:
        pass
    try:
        if context.get("dial_sheet"):
            data = _matrix_from_sheet(context["dial_sheet"], len(HEADER_FIELDS) + 3 + 8)
            save_dialer_leads_matrix(data)
    except Exception:
        pass

def _autosave_on_edit(sheet, save_fn):
    """Bind end_edit_cell to persist after manual edits (not just paste)."""
    if not sheet:
        return
    def _on_end_edit(_ev=None):
        try:
            save_fn(sheet)
        except Exception:
            pass
    try:
        sheet.extra_bindings([("end_edit_cell", _on_end_edit)])
    except Exception:
        try:
            sheet.bind("<Return>", lambda _e: _on_end_edit())
        except Exception:
            pass

def _disable_rc_menu(sheet):
    """Remove right-click popup entirely (per your request to remove Paste)."""
    try:
        sheet.disable_bindings(("right_click_popup_menu",))
    except Exception:
        pass

# ==============================
# Grid mounting
# ==============================

def _mount_leads(window, start_rows=200, col_width=140, context=None):
    if not _TKSHEET_OK:
        return None
    host = window["-LEADS_HOST-"].Widget
    _clear_children(host)
    holder = sg.tk.Frame(host, bg="#111111")
    holder.pack(side="top", fill="both", expand=True)
    try:
        rows = load_email_leads_matrix()
    except Exception:
        rows = []
    if len(rows) < start_rows:
        rows += [[""] * len(HEADER_FIELDS) for _ in range(start_rows - len(rows))]
    sheet = Sheet(holder, data=rows, headers=HEADER_FIELDS, show_x_scrollbar=True, show_y_scrollbar=True)
    sheet.enable_bindings((
        "single_select","row_select","arrowkeys","tab_key","shift_tab_key",
        "drag_select","copy","cut","delete","undo",
        "edit_cell","return_edit_cell","select_all",
        "rc_select",
        "column_width_resize","column_resize","resize_columns"
    ))
    try:
        sheet.set_options(
            expand_sheet_if_paste_too_big=True,
            data_change_detected=True,
            show_vertical_grid=True,
            show_horizontal_grid=True,
        )
    except Exception:
        pass
    sheet.pack(fill="both", expand=True)

    for c in range(len(HEADER_FIELDS)):
        try:
            sheet.column_width(c, width=col_width)
        except Exception:
            pass

    saved = load_column_widths(COLWIDTHS_PATH, _LEADS_PREF_KEY)
    if saved:
        apply_column_widths(sheet, saved)
    attach_column_width_persistence(sheet, COLWIDTHS_PATH, _LEADS_PREF_KEY, ncols=len(HEADER_FIELDS))

    # 🔒 Remove right-click menu (no Paste there)
    _disable_rc_menu(sheet)

    # ✅ Proven plain-text paste anchored to selected cell, autosave after paste + refresh analytics
    _bind_plaintext_paste(
        sheet, window.TKroot,
        headers_only_cols=None,
        save_callback=lambda: (_save_leads(sheet), _trigger_analytics_refresh(window))
    )

    # ✅ Autosave on manual edits too (and refresh)
    _autosave_on_edit(sheet, lambda _s: (_save_leads(sheet), _trigger_analytics_refresh(window)))

    return sheet


def _mount_dialer(window, start_rows=100, col_width=120, context=None):
    if not _TKSHEET_OK:
        return None
    host = window["-DIAL_HOST-"].Widget
    _clear_children(host)
    holder = sg.tk.Frame(host, bg="#111111")
    holder.pack(side="top", fill="both", expand=True)
    try:
        matrix = load_dialer_leads_matrix()
    except Exception:
        matrix = []
    if not matrix:
        try:
            base = load_email_leads_matrix()
        except Exception:
            base = []
        if not base:
            base = [[""] * len(HEADER_FIELDS) for _ in range(50)]
        matrix = [row + ["○","○","○"] + ([""]*8) for row in base]
    if len(matrix) < start_rows:
        padrow = [""]*len(HEADER_FIELDS) + ["○","○","○"] + ([""]*8)
        matrix += [padrow[:] for _ in range(start_rows - len(matrix))]
    headers = HEADER_FIELDS + [EMOJI_GREEN, EMOJI_GRAY, EMOJI_RED] + [f"Note{i}" for i in range(1,9)]
    sheet = Sheet(holder, data=matrix, headers=headers, show_x_scrollbar=True, show_y_scrollbar=True)
    sheet.enable_bindings((
        "single_select","row_select","arrowkeys","tab_key","shift_tab_key",
        "drag_select","copy","cut","delete","undo",
        "edit_cell","return_edit_cell","select_all",
        "rc_select",
        "column_width_resize","column_resize","resize_columns"
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

    first_dot = len(HEADER_FIELDS)
    first_note = first_dot + 3
    last_note = first_note + 7
    for c in range(len(headers)):
        w = col_width
        if first_dot <= c < first_note:
            w = 42
        if first_note <= c <= last_note:
            w = 140
        try:
            sheet.column_width(c, width=w)
        except Exception:
            pass

    saved = load_column_widths(COLWIDTHS_PATH, _DIALER_PREF_KEY)
    if saved:
        apply_column_widths(sheet, saved)
    attach_column_width_persistence(sheet, COLWIDTHS_PATH, _DIALER_PREF_KEY, ncols=len(headers))

    _disable_rc_menu(sheet)

    # ✅ Paste only into lead columns (don’t overwrite dots/notes), autosave + refresh analytics
    _bind_plaintext_paste(
        sheet, window.TKroot,
        headers_only_cols=len(HEADER_FIELDS),
        save_callback=lambda: (save_dialer_leads_matrix(_matrix_from_sheet(sheet, len(headers))),
                               _trigger_analytics_refresh(window))
    )

    _autosave_on_edit(
        sheet,
        lambda _s: (save_dialer_leads_matrix(_matrix_from_sheet(sheet, len(headers))),
                    _trigger_analytics_refresh(window))
    )

    return sheet


def _mount_customers(window, start_rows=50, col_width=130, context=None):
    if not _TKSHEET_OK:
        return None
    host = window["-CUST_HOST-"].Widget
    _clear_children(host)
    holder = sg.tk.Frame(host, bg="#111111")
    holder.pack(side="top", fill="both", expand=True)
    try:
        matrix = load_customers_matrix()
    except Exception:
        matrix = []
    if len(matrix) < start_rows:
        matrix += [[""] * len(CUSTOMER_FIELDS) for _ in range(start_rows - len(matrix))]
    sheet = Sheet(holder, data=matrix, headers=CUSTOMER_FIELDS, show_x_scrollbar=True, show_y_scrollbar=True)
    sheet.enable_bindings((
        "single_select","row_select","arrowkeys","tab_key","shift_tab_key",
        "drag_select","copy","cut","delete","undo",
        "edit_cell","return_edit_cell","select_all",
        "rc_select",
        "column_width_resize","column_resize","resize_columns"
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

    for c, name in enumerate(CUSTOMER_FIELDS):
        w = col_width
        if name in ("Company","Prospect Name"):
            w = 180
        if name in ("Email","Industry","Address","City"):
            w = 160
        if name in ("State","ZIP","Lat","Lon"):
            w = 80
        if name in ("CLTV","Sales/Day","Days"):
            w = 100
        try:
            sheet.column_width(c, width=w)
        except Exception:
            pass

    saved = load_column_widths(COLWIDTHS_PATH, _CUSTOMERS_PREF_KEY)
    if saved:
        apply_column_widths(sheet, saved)
    attach_column_width_persistence(sheet, COLWIDTHS_PATH, _CUSTOMERS_PREF_KEY, ncols=len(CUSTOMER_FIELDS))

    _disable_rc_menu(sheet)

    # ✅ Plain-text paste + autosave at correct anchor + refresh analytics
    _bind_plaintext_paste(
        sheet, window.TKroot,
        headers_only_cols=None,
        save_callback=lambda: (_save_customers(sheet), _trigger_analytics_refresh(window))
    )
    _autosave_on_edit(sheet, lambda _s: (_save_customers(sheet), _trigger_analytics_refresh(window)))

    return sheet


# ---- Campaigns table refresh ----

def _refresh_campaign_table(window):
    try:
        keys = list_campaign_keys() or ["default"]
        table_rows = []
        for k in keys:
            base = summarize_campaign_for_table(k)
            # add resp% (best-effort)
            try:
                rows = load_results_rows_sorted()
                sent = replied = 0
                subs = {
                    (s.get("subject", "") or "").strip()
                    for s in normalize_campaign_steps(load_campaign_by_key(k)[0])
                    if (s.get("subject", "") or "").strip()
                }
                for r in rows:
                    if (r.get("Subject", "") or "").strip() in subs:
                        if r.get("DateSent"):
                            sent += 1
                        if r.get("DateReplied"):
                            replied += 1
                resp = "0.0%" if sent == 0 else f"{(replied / sent) * 100:.1f}%"
            except Exception:
                resp = ""
            table_rows.append(base + [resp])
        try:
            window["-CAMP_TABLE-"].update(values=table_rows)
        except Exception:
            pass
        try:
            window["-CAMP_KEY-"].update(values=keys)
        except Exception:
            pass
        try:
            window["-CAMP_STATUS-"].update("Campaigns loaded ✓")
        except Exception:
            pass
    except Exception as e:
        try:
            window["-CAMP_STATUS-"].update(f"Campaign refresh error: {e}")
        except Exception:
            pass

# ==============================
# Public entry points
# ==============================

def mount_grids(window, _context):
    # Create & mount all grids; start analytics watcher.
    ensure_app_files()

    context = dict(_context or {})

    leads_sheet = _mount_leads(window, context=context)
    dial_sheet  = _mount_dialer(window, context=context)
    warm_sheet  = mount_warm_grid(window)   # lives in gf_warm.py
    cust_sheet  = _mount_customers(window, context=context)

    # Also wire Ctrl+V for Warm sheet (gf_warm removed paste on purpose)
    if isinstance(warm_sheet, Sheet):
        _disable_rc_menu(warm_sheet)
        _bind_plaintext_paste(
            warm_sheet, window.TKroot,
            headers_only_cols=None,
            save_callback=lambda: (_save_warm(warm_sheet), _trigger_analytics_refresh(window))
        )
        _autosave_on_edit(warm_sheet, lambda _s: (_save_warm(warm_sheet), _trigger_analytics_refresh(window)))

    # Start analytics (updates Daily/Monthly + right-side panels)
    init_analytics(window)

    # --- populate Campaigns table immediately on startup ---
    try:
        _refresh_campaign_table(window)
    except Exception:
        pass

    # Attach dialer controller (owns -DIAL_* events)
    dialer_ctl = None
    try:
        if dial_sheet is not None:
            dialer_ctl = attach_dialer(window, dial_sheet)
    except Exception as _e:
        print("[dialer] attach failed:", _e)

    context.update({
        "sheet": leads_sheet,
        "dial_sheet": dial_sheet,
        "warm_sheet": warm_sheet,
        "customer_sheet": cust_sheet,
        "dialer_ctl": dialer_ctl,
        "_dial_last_row": None,
        "_warm_last_row": None,
        "_cust_last_row": None,
        "_active_sheet_name": "leads",
        "_active_sheet": leads_sheet,
    })
    return context


def run_event_loop(window, context):
    sheet        = context.get("sheet")
    dial_sheet   = context.get("dial_sheet")
    warm_sheet   = context.get("warm_sheet")
    cust_sheet   = context.get("customer_sheet")
    dialer_ctl   = context.get("dialer_ctl")

    def _grid_highlight_row(sheet, new_row, last_row_holder: dict, last_key: str, color=_SOFT_BLUE):
        if sheet is None or new_row is None:
            return
        last = last_row_holder.get(last_key)
        try:
            if last is not None and last != new_row:
                sheet.highlight_rows(rows=[last], bg=None, fg=None)
        except Exception:
            pass
        try:
            sheet.highlight_rows(rows=[new_row], bg=color, fg="#000000")
            sheet.see(new_row, 0)
            last_row_holder[last_key] = new_row
        except Exception:
            pass
        try:
            sheet.refresh()
        except Exception:
            pass

    def _row_selected(sheet_obj):
        if sheet_obj is None:
            return None
        try:
            sel = sheet_obj.get_selected_rows() or []
            if sel:
                return sel[0]
        except Exception:
            pass
        try:
            r, _ = sheet_obj.get_currently_selected()
            if isinstance(r, int) and r >= 0:
                return r
        except Exception:
            pass
        return None

    while True:
        event, values = window.read(timeout=250)
        if event in (sg.WINDOW_CLOSE_ATTEMPTED_EVENT, sg.WIN_CLOSED):
            # SAVE-ON-EXIT (bulletproof persistence)
            _save_all(context)
            break

        # Global analytics refresh hook
        if event == "-ANALYTICS_REFRESH-":
            try:
                init_analytics(window)
            except Exception:
                pass

        # ---- Update button ----
        if event == "-UPDATE-":
            try:
                window["-UPDATE-"].update(disabled=True)
            except Exception:
                pass
            try:
                update_ui_flow(window)
            finally:
                try:
                    window["-UPDATE-"].update(disabled=False)
                except Exception:
                    pass
            continue

        # Warm events are handled in gf_warm and may fully handle the tick
        if warm_handle_event(event, values, window, context):
            continue

        # Route dialer events to controller (it manages its own coloring & state)
        if dialer_ctl and event and str(event).startswith("-DIAL_"):
            try:
                r_clicked = _row_selected(dial_sheet)
            except Exception:
                r_clicked = None
            target_row = r_clicked if r_clicked is not None else context.get("_dial_last_row")
            if target_row is not None:
                try:
                    dial_sheet.set_currently_selected(target_row, 0)
                    dial_sheet.see(target_row, 0)
                except Exception:
                    pass
                context["_dial_last_row"] = target_row
                try:
                    if hasattr(dialer_ctl, "_set_working_row"):
                        dialer_ctl._set_working_row(target_row)  # type: ignore[attr-defined]
                    elif hasattr(dialer_ctl, "state") and isinstance(dialer_ctl.state, dict):
                        dialer_ctl.state["row"] = target_row
                except Exception:
                    try:
                        if hasattr(dialer_ctl, "state") and isinstance(dialer_ctl.state, dict):
                            dialer_ctl.state["row"] = target_row
                    except Exception:
                        pass
            try:
                handled = dialer_ctl.handle_event(event, values)
            except Exception as _e:
                print("[dialer] handle_event error:", _e)
                handled = True
            if handled:
                continue
            try:
                dialer_ctl.tick()
            except Exception:
                pass

        # Leads tab
        if event == "-SAVECSV-":
            try:
                _save_leads(sheet)
                _trigger_analytics_refresh(window)
                window["-STATUS-"].update("Saved CSV")
            except Exception as e:
                window["-STATUS-"].update(f"Save error: {e}")

        elif event == "-ADDROWS-":
            if sheet is None:
                sg.popup_error("Leads sheet not initialized.")
            else:
                try:
                    sheet.insert_rows(sheet.get_total_rows(), number_of_rows=1000)
                    sheet.refresh()
                    _save_leads(sheet)
                    _trigger_analytics_refresh(window)
                except Exception:
                    try:
                        sheet.insert_rows(sheet.get_total_rows(), amount=1000)
                        sheet.refresh()
                        _save_leads(sheet)
                        _trigger_analytics_refresh(window)
                    except Exception as e:
                        sg.popup_error(f"Could not add rows: {e}")

        elif event == "-DELROWS-":
            if sheet is None:
                sg.popup_error("Leads sheet not initialized.")
            else:
                try:
                    sels = sheet.get_selected_rows() or []
                    if sels:
                        for r in sorted(sels, reverse=True):
                            try:
                                sheet.delete_rows(r, 1)
                            except Exception:
                                try:
                                    sheet.delete_rows(r)
                                except Exception:
                                    try:
                                        sheet.del_rows(r, 1)
                                    except Exception:
                                        pass
                        sheet.refresh()
                        _save_leads(sheet)
                        _trigger_analytics_refresh(window)
                except Exception as e:
                    sg.popup_error(f"Could not delete rows: {e}")

        elif event == "-OPENFOLDER-":
            try:
                os.startfile(str(APP_DIR))
            except Exception as e:
                window["-STATUS-"].update(f"Open folder error: {e}")

        elif event == "-LEADS_RELOAD-":
            new_sheet = _mount_leads(window, context=context)
            context["sheet"] = new_sheet
            sheet = new_sheet
            _trigger_analytics_refresh(window)

        # Customers tab
        elif event == "-CUST_ADD50-":
            if cust_sheet is not None:
                try:
                    cust_sheet.insert_rows(cust_sheet.get_total_rows(), number_of_rows=50)
                    cust_sheet.refresh()
                    _save_customers(cust_sheet)
                    _trigger_analytics_refresh(window)
                except Exception:
                    pass

        elif event == "-CUST_RELOAD-" or event == "-CUSTOMERS_RELOAD-":
            new_sheet = _mount_customers(window, context=context)
            context["customer_sheet"] = new_sheet
            cust_sheet = new_sheet
            _trigger_analytics_refresh(window)

        elif event == "-CUST_EXPORT-":
            try:
                _save_customers(cust_sheet)
                _trigger_analytics_refresh(window)
                sg.popup_ok("Customers saved.", keep_on_top=True)
            except Exception:
                pass

        elif event == "-CUST_ADD_ORDER-":
            def _row_selected_local(sheet_obj):
                return _row_selected(sheet_obj)
            try:
                r_clicked = _row_selected_local(cust_sheet)
            except Exception:
                r_clicked = None
            target_row = r_clicked if r_clicked is not None else context.get("_cust_last_row")
            if target_row is not None:
                try:
                    cust_sheet.set_currently_selected(target_row, 0)
                    cust_sheet.see(target_row, 0)
                except Exception:
                    pass
                _grid_highlight_row(cust_sheet, target_row, context, "_cust_last_row", color=_SOFT_BLUE)

            company = sg.popup_get_text("Company name for the order:", title="Add Order")
            if not company:
                continue
            date_s = sg.popup_get_text("Order date (YYYY-MM-DD or MM/DD/YYYY):", title="Add Order") or ""
            amount_s = sg.popup_get_text("Amount (e.g. 199.00):", title="Add Order") or ""
            try:
                append_order_row(company, date_s, amount_s)
                _save_customers(cust_sheet)
                new_sheet = _mount_customers(window, context=context)
                context["customer_sheet"] = new_sheet
                cust_sheet = new_sheet
                _trigger_analytics_refresh(window)
                sg.popup_ok("Order added and customers updated.", keep_on_top=True)
            except Exception as e:
                sg.popup_error(f"Add order failed: {e}", keep_on_top=True)

        # Warm reload hook
        elif event == "-WARM_RELOAD-":
            try:
                reload_warm_sheet(window)
            except Exception:
                warm_sheet = mount_warm_grid(window)
                context["warm_sheet"] = warm_sheet
            _trigger_analytics_refresh(window)

        # Map tab
        elif event == "-OPEN_MAP-":
            try:
                open_customer_map(window)
            except Exception:
                pass

        # Campaigns UI ...
        elif event in ("-CAMP_ADD_NEW-", "-CAMP_NEW-"):
            import re
            key = sg.popup_get_text(
                "Name your new campaign (e.g., 'butcher shop', 'farm market'):",
                title="Add New Campaign",
                keep_on_top=True,
            )
            if not key:
                continue
            key = re.sub(r"\s+", " ", key.strip())
            if not key:
                continue
            try:
                keys = list_campaign_keys()
            except Exception:
                keys = []
            if key not in keys:
                keys.append(key)
            window["-CAMP_KEY-"].update(values=keys, value=key)
            steps = [
                {"subject": "", "body": "", "delay_days": 0},
                {"subject": "", "body": "", "delay_days": 0},
                {"subject": "", "body": "", "delay_days": 0},
            ]
            for i, st in enumerate(steps, start=1):
                window[f"-CAMP_SUBJ_{i}-"].update(st["subject"])
                window[f"-CAMP_BODY_{i}-"].update(st["body"])
                window[f"-CAMP_DELAY_{i}-"].update(str(st["delay_days"]))
            window["-CAMP_SEND_TO_DIALER-"].update(True)
            window["-CAMP_EMPTY_WRAP-"].update(visible=False)
            window["-CAMP_EDITOR_WRAP-"].update(visible=True)
            try:
                window["-CAMP_STATUS-"].update("New campaign ready. Fill in fields and click Save.")
            except Exception:
                pass

        elif event in ("-CAMP_LOAD-", "-CAMP_KEY-"):
            key = (values.get("-CAMP_KEY-") or "").strip()
            if not key:
                continue
            try:
                steps, settings = load_campaign_by_key(key)
            except Exception:
                steps, settings = [], {}
            steps = normalize_campaign_steps(steps or [])
            for i, st in enumerate(steps + [{"subject": "", "body": "", "delay_days": 0}] * 3, start=1):
                if i > 3:
                    break
                window[f"-CAMP_SUBJ_{i}-"].update(st.get("subject", ""))
                window[f"-CAMP_BODY_{i}-"].update(st.get("body", ""))
                window[f"-CAMP_DELAY_{i}-"].update(str(st.get("delay_days", 0)))
            settings = normalize_campaign_settings(settings or {})
            window["-CAMP_SEND_TO_DIALER-"].update(bool(settings.get("send_to_dialer_after") in ("1", True, "true", "yes", "on")))
            window["-CAMP_EMPTY_WRAP-"].update(visible=False)
            window["-CAMP_EDITOR_WRAP-"].update(visible=True)
            try:
                window["-CAMP_STATUS-"].update(f"Loaded '{key}'.")
            except Exception:
                pass

        elif event == "-CAMP_SAVE-":
            try:
                key = (values.get("-CAMP_KEY-") or "").strip()
                if not key:
                    sg.popup_error("Provide a campaign name first (use New).", keep_on_top=True)
                    continue

                steps = []
                for i in (1, 2, 3):
                    subj = values.get(f"-CAMP_SUBJ_{i}-", "")
                    body = values.get(f"-CAMP_BODY_{i}-", "")
                    delay_raw = values.get(f"-CAMP_DELAY_{i}-", "0")
                    try:
                        delay_i = int(str(delay_raw or "0").strip() or "0")
                    except Exception:
                        delay_i = 0
                    steps.append({"subject": subj or "", "body": body or "", "delay_days": max(0, delay_i)})

                steps = normalize_campaign_steps(steps)
                send_to_dialer = bool(values.get("-CAMP_SEND_TO_DIALER-", False))
                settings = normalize_campaign_settings({"send_to_dialer_after": "1" if send_to_dialer else "0"})

                try:
                    save_campaign_by_key(key, steps, settings)
                    window["-CAMP_STATUS-"].update("Saved ✓")
                except Exception as e:
                    window["-CAMP_STATUS-"].update(f"Save error: {e}")

                _refresh_campaign_table(window)
                window["-CAMP_EMPTY_WRAP-"].update(visible=False)
                window["-CAMP_EDITOR_WRAP-"].update(visible=True)
            except Exception as e:
                try:
                    window["-CAMP_STATUS-"].update(f"Save error: {e}")
                except Exception:
                    pass

        elif event == "-CAMP_DELETE-":
            key = (values.get("-CAMP_KEY-") or "").strip()
            if not key:
                continue
            yn = sg.popup_yes_no(f"Delete campaign '{key}'? This cannot be undone.", keep_on_top=True)
            if yn == "Yes":
                try:
                    delete_campaign_by_key(key)
                    window["-CAMP_STATUS-"].update("Deleted ✓")
                except Exception as e:
                    window["-CAMP_STATUS-"].update(f"Delete error: {e}")
                _refresh_campaign_table(window)

        elif event == "-CAMP_REFRESH_LIST-":
            _refresh_campaign_table(window)

    window.close()

