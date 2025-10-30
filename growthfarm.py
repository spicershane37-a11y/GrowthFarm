# ============================================================
# growthfarm.py — Main App Entry Point (CSV-only license + first-run prompt)
# ============================================================

import sys
from pathlib import Path

from gf_store import ensure_app_files
from gf_ui_layout import build_window
from gf_ui_logic import mount_grids, run_event_loop
from gf_license import (
    get_banner_text,
    accounts_csv_exists,
    create_or_replace_account,
    purge_legacy_json,
)

# ---- Optional: allow vendored PySimpleGUI path like gf_ui_layout does ----
import sys as _sys
from pathlib import Path as _Path
_VENDOR_PSG = _Path(__file__).parent / "vendor_psg"
if str(_VENDOR_PSG) not in _sys.path:
    _sys.path.insert(0, str(_VENDOR_PSG))
import PySimpleGUI as sg  # noqa
# --------------------------------------------------------------------------

APP_VERSION = "2025-10-14"


def _first_run_prompt() -> bool:
    """
    Shows a tiny modal dialog to collect email/password/user/company.
    Returns True if the CSV was created, False if user canceled.
    """
    layout = [
        [sg.Text("Welcome to GrowthFarm — set up your account", text_color="#9EE493")],
        [sg.Text("Email", size=(12, 1)), sg.Input(key="-EMAIL-", size=(36, 1))],
        [sg.Text("Password", size=(12, 1)), sg.Input(key="-PASS-", size=(36, 1), password_char="*")],
        [sg.Text("User (your name)", size=(12, 1)), sg.Input(key="-USER-", size=(36, 1))],
        [sg.Text("Company", size=(12, 1)), sg.Input(key="-COMP-", size=(36, 1))],
        [sg.Push(),
         sg.Button("Save", key="-SAVE-", button_color=("white", "#2E7D32")),
         sg.Button("Cancel", key="-CANCEL-")]
    ]
    win = sg.Window("GrowthFarm — First Run", layout, modal=True, keep_on_top=True)
    created = False
    try:
        while True:
            ev, values = win.read()
            if ev in (sg.WINDOW_CLOSED, "-CANCEL-"):
                break
            if ev == "-SAVE-":
                email = (values.get("-EMAIL-") or "").strip()
                pw = (values.get("-PASS-") or "").strip()
                user = (values.get("-USER-") or "").strip()
                comp = (values.get("-COMP-") or "").strip()
                if not (email and pw and user and comp):
                    sg.popup_ok("Please fill all four fields.", keep_on_top=True)
                    continue
                create_or_replace_account(email, pw, user, comp)
                created = True
                break
    finally:
        try:
            win.close()
        except Exception:
            pass
    return created


def main():
    # --- Initialize storage (creates CSV shells, folders, etc.) ---
    ensure_app_files()

    # --- Kill legacy JSON if it exists ---
    purge_legacy_json()

    # --- Ensure we have accounts.csv; if not, prompt once to create it ---
    if not accounts_csv_exists():
        made = _first_run_prompt()
        if not made:
            # User canceled setup; exit gracefully.
            return 0

    # --- Build banner text from accounts.csv (e.g., "Shane's GF Test Account Growth Farm") ---
    display_name = get_banner_text()

    # --- Build UI window + shared context (handles, sheets, etc.) ---
    window, context = build_window(APP_VERSION, user_display_name=display_name)

    # --- Mount grids (tksheet) BEFORE entering the loop ---
    context = mount_grids(window, context)

    # --- Event loop driver (core app logic) ---
    try:
        run_event_loop(window, context)
    finally:
        try:
            window.close()
        except Exception:
            pass

    return 0


if __name__ == "__main__":
    sys.exit(main())

