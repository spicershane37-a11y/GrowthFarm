# gf_ui_layout.py
# Builds the GrowthFarm main window (layout only — no grid mounting here)

# --- Force vendored PySimpleGUI 4.60.x ---
import sys
from pathlib import Path

_VENDOR_PSG = Path(__file__).parent / "vendor_psg"
if str(_VENDOR_PSG) not in sys.path:
    sys.path.insert(0, str(_VENDOR_PSG))
import PySimpleGUI as sg  # noqa
# -----------------------------------------


def _set_theme():
    set_theme = getattr(sg, "theme", getattr(sg, "ChangeLookAndFeel", None))
    ok = False
    if callable(set_theme):
        try:
            set_theme("DarkGrey13")
            ok = True
        except Exception:
            ok = False
    if not ok:
        fallback = getattr(sg, "SetOptions", None)
        if callable(fallback):
            try:
                fallback(background_color="#1B1B1B", text_color="#FFFFFF")
            except Exception:
                pass


# Safe global options for PSG 4.60.x (no 5.x-only args)
try:
    sg.set_options(
        keep_on_top=False,
        resizable=True,
        margins=(0, 0),
    )
except Exception:
    pass


def _step_row(i: int):
    body_h = 6 if i == 1 else 5
    return [
        [sg.Text(f"Step {i}", text_color="#CCCCCC"),
         sg.Push(),
         sg.Text("Resp:", text_color="#9EE493"),
         sg.Text("—", key=f"-CAMP_RESP_{i}-", text_color="#FFFFFF")],
        [sg.Column([[sg.Text("Subject", text_color="#9EE493")],
                    [sg.Input(key=f"-CAMP_SUBJ_{i}-", size=(48, 1), enable_events=True)]], pad=(0, 0)),
         sg.Text("   "),
         sg.Column([[sg.Text("Body", text_color="#9EE493")],
                    [sg.Multiline(key=f"-CAMP_BODY_{i}-", size=(90, body_h),
                                  font=("Consolas", 10),
                                  text_color="#EEE", background_color="#111",
                                  enable_events=True, no_scrollbar=False)]],
                   pad=(0, 0), expand_x=True)],
        [sg.Text("Delay (days) after previous send:", text_color="#CCCCCC"),
         sg.Input("0", key=f"-CAMP_DELAY_{i}-", size=(6, 1), enable_events=True)],
        [sg.HorizontalSeparator(color="#333333")]
    ]


def _editor_toolbar(key_suffix: str = ""):
    """Top/bottom editor toolbar. Pass key_suffix='' (top) or '_BOTTOM' (bottom)."""
    return [
        sg.Text("Campaign niche / industry:", text_color="#9EE493"),
        sg.Combo(values=["default"], default_value="default",
                 key=f"-CAMP_KEY{key_suffix}-", size=(28, 1), enable_events=True),
        sg.Button("New",  key=f"-CAMP_NEW{key_suffix}-"),
        sg.Button("Load", key=f"-CAMP_LOAD{key_suffix}-"),
        sg.Button("Save Campaign", key=f"-CAMP_SAVE{key_suffix}-", button_color=("white", "#2E7D32")),
        sg.Button("Delete This Campaign", key=f"-CAMP_DELETE{key_suffix}-", button_color=("white", "#8B0000")),
        sg.Push(),
        sg.Text("", key=f"-CAMP_STATUS{key_suffix}-", text_color="#A0FFA0")
    ]


def build_window(app_version: str, user_display_name: str = ""):
    """
    Return (window, context).
    Context contains references to 'host' *Columns* where the grids are mounted later.
    Mount onto context['leads_host'].Widget (and same for others).
    """
    _set_theme()

    # ---------------- Toolbar ----------------
    top_bar = [
        sg.Text(f"GrowthFarm v{app_version}", text_color="#9EE493"),
        sg.Push(),
        sg.Button("Update", key="-UPDATE-", button_color=("white", "#444444")),
    ]

    # ---------------- Banner (green, big, single line) ----------------
    banner_text_value = (user_display_name or "GrowthFarm") + " Growth Farm"
    banner_row = [
        sg.Text(
            banner_text_value,
            key="-BANNER-",
            text_color="#9EE493",
            font=("Segoe UI", 20, "bold"),
            justification="center",
            expand_x=True,
            pad=(0, 12),
        )
    ]

    # ================== SCOREBOARDS ==================
    da_header = [[sg.Text("DAILY ACTIVITY TRACKER",
                          text_color="#9EE493",
                          font=("Segoe UI", 16, "bold"),
                          justification="center",
                          expand_x=True)]]
    da_left = [
        [sg.Text("CALLS:",          text_color="#CCCCCC"), sg.Text("0",     key="-DA_CALLS-",  text_color="#A0FFA0")],
        [sg.Text("EMAILS:",         text_color="#CCCCCC"), sg.Text("0",     key="-DA_EMAILS-", text_color="#A0FFA0")],
        [sg.Text("NEW WARM LEADS:", text_color="#CCCCCC"), sg.Text("0",     key="-DA_WARMS-",  text_color="#A0FFA0")],
    ]
    da_right = [
        [sg.Text("NEW ACCOUNTS:",   text_color="#CCCCCC"), sg.Text("0",     key="-DA_NEWCUS-", text_color="#A0FFA0")],
        [sg.Text("DAILY SALES:",    text_color="#CCCCCC"), sg.Text("$0.00", key="-DA_SALES-",  text_color="#A0FFA0")],
    ]
    daily_scoreboard = sg.Frame(
        "",
        da_header + [[sg.Column(da_left, pad=(6, 6)), sg.Text("   "), sg.Column(da_right, pad=(6, 6))]],
        relief=sg.RELIEF_GROOVE, border_width=2,
        background_color="#1B1B1B", title_color="#9EE493",
        expand_x=False, expand_y=False
    )

    mo_header = [[sg.Text("MONTHLY RESULTS",
                          text_color="#9EE493",
                          font=("Segoe UI", 16, "bold"),
                          justification="center",
                          expand_x=True)]]
    mo_col = [
        [sg.Text("NEW WARM LEADS:", text_color="#CCCCCC"), sg.Text("0",     key="-MO_WARMS-",  text_color="#A0FFA0")],
        [sg.Text("NEW CUSTOMERS:",  text_color="#CCCCCC"), sg.Text("0",     key="-MO_NEWCUS-", text_color="#A0FFA0")],
        [sg.Text("TOTAL SALES:",    text_color="#CCCCCC"), sg.Text("$0.00", key="-MO_SALES-",  text_color="#A0FFA0")],
    ]
    monthly_scoreboard = sg.Frame(
        "",
        mo_header + [[sg.Column(mo_col, pad=(6, 6))]],
        relief=sg.RELIEF_GROOVE, border_width=2,
        background_color="#1B1B1B", title_color="#9EE493",
        expand_x=False, expand_y=False
    )

    scoreboards_row = [
        sg.Column(
            [[daily_scoreboard, sg.Text("  "), monthly_scoreboard]],
            pad=((500, 0), (50, 6)),
            background_color="#202020",
            expand_x=False, expand_y=False
        )
    ]
    # ================== /SCOREBOARDS ==================

    # ---------- Email Leads tab ----------
    leads_mount = sg.Column(
        [[sg.Text("Loading grid…", key="-LOADING-", text_color="#9EE493")]],
        key="-LEADS_MOUNT-",
        pad=(0, 0),
        expand_x=True,
        expand_y=True,
        background_color=None,
        scrollable=False,
    )
    leads_host = sg.Frame(
        "EMAIL LEADS (paste directly from Google Sheets / Excel)",
        [[leads_mount]],
        relief=sg.RELIEF_GROOVE, border_width=2, background_color="#1B1B1B",
        title_color="#9EE493", expand_x=True, expand_y=True, key="-LEADS_HOST-",
    )
    leads_buttons_row1 = [
        sg.Button("Open Folder", key="-OPENFOLDER-"),
        sg.Button("Add 1,000 Rows", key="-ADDROWS-"),
        sg.Button("Delete Selected Rows", key="-DELROWS-"),
        sg.Button("Save Now", key="-SAVECSV-"),
        sg.Text("Status:", text_color="#A0A0A0"),
        sg.Text("Idle", key="-STATUS-", text_color="#FFFFFF"),
    ]
    leads_buttons_row2 = [
        sg.Button("Fire Emails", key="-FIRE-", size=(22, 2), disabled=True, button_color=("white", "#700000")),
        sg.Text(" (disabled: add valid NEW leads)", key="-FIRE_HINT-", text_color="#BBBBBB")
    ]
    leads_tab = [
        [leads_host],
        [sg.Column([leads_buttons_row1], pad=(0, 0))],
        [sg.Column([leads_buttons_row2], pad=(0, 0))],
    ]

    # ---------- Campaigns tab ----------
    empty_state = [
        [sg.Text("No campaigns available yet.", text_color="#CCCCCC", key="-CAMP_EMPTY_MSG-")],
        [sg.Button("➕  Add New Campaign", key="-CAMP_ADD_NEW-", button_color=("white", "#2E7D32"))]
    ]

    # Editor header (top)
    editor_header_top = [_editor_toolbar("")]

    # Three steps
    step_rows = _step_row(1) + _step_row(2) + _step_row(3)

    # Editor footer (bottom duplicate toolbar + toggle + saved table)
    editor_toolbar_bottom = [_editor_toolbar("_BOTTOM")]
    editor_footer_controls = [
        [sg.Checkbox("Send to Dialer automatically if they complete the campaign without replying",
                     key="-CAMP_SEND_TO_DIALER-", default=True, text_color="#EEEEEE")]
    ]

    saved_table_section = [
        [sg.HorizontalSeparator(color="#4CAF50")],
        [sg.Text("Saved Campaigns", text_color="#9EE493")],
        [sg.Table(values=[],
                  headings=["Campaign", "Enabled Steps", "Delays (days)", "To Dialer", "Auto Sync", "Hourly Runner", "Resp %"],
                  auto_size_columns=False, col_widths=[24, 14, 18, 10, 10, 14, 14], justification="left", num_rows=8,
                  key="-CAMP_TABLE-", enable_events=True, alternating_row_color="#2a2a2a",
                  text_color="#EEE", background_color="#111",
                  header_text_color="#FFF", header_background_color="#333")],
        [sg.Button("Refresh List", key="-CAMP_REFRESH_LIST-")]
    ]

    # Compose scrollable editor column: top header → empty/editor → steps → bottom toolbar → saved campaigns
    campaigns_editor_column = sg.Column(
        [[sg.Text("Email Campaigns let you schedule up to 3 follow-ups.", text_color="#CCCCCC")],
         [sg.Column(empty_state, key="-CAMP_EMPTY_WRAP-", visible=True, expand_x=True)],
         [sg.Column(editor_header_top + step_rows + editor_footer_controls,
                    key="-CAMP_EDITOR_WRAP-", visible=False, expand_x=True)],
         [sg.Column(editor_toolbar_bottom, key="-CAMP_EDITOR_BOTTOM-", visible=False, expand_x=True)],
         [sg.Column(saved_table_section, key="-CAMP_SAVED_WRAP-", visible=True, expand_x=True)]],
        size=(980, 620),
        scrollable=True,
        vertical_scroll_only=True,
        expand_x=True,
        expand_y=False,              # keep False so scrollbar works consistently
        pad=(0, 0),
        background_color=None,
        key="-CAMP_SCROLL-"          # << needed so logic can force a scrollregion refresh
    )

    campaigns_tab = [[campaigns_editor_column]]

    # ---------- Email Results tab ----------
    results_tab = [
        [sg.Text("Sync replies from Outlook; tag Green (good), Gray (neutral), Red (negative).", text_color="#CCCCCC")],
        [sg.Text("Lookback days:", text_color="#CCCCCC"), sg.Input("60", key="-LOOKBACK-", size=(6, 1)),
         sg.Button("Sync from Outlook", key="-SYNC-"),
         sg.Checkbox("Auto Sync (hourly)", key="-AUTO_SYNC-", default=False, text_color="#EEEEEE"),
         sg.Text("", key="-RS_STATUS-", text_color="#A0FFA0")],
        [sg.Table(values=[], headings=["Ref", "Email", "Company", "Industry", "DateSent", "DateReplied", "Status", "Subject"],
                  auto_size_columns=False, col_widths=[10, 26, 26, 14, 18, 18, 8, 40], justification="left", num_rows=15,
                  key="-RSTABLE-", enable_events=True, alternating_row_color="#2a2a2a",
                  text_color="#EEE", background_color="#111", header_text_color="#FFF", header_background_color="#333")],
        [sg.Button("Mark Green", key="-MARK_GREEN-", button_color=("white", "#2E7D32")),
         sg.Button("Mark Gray",  key="-MARK_GRAY-",  button_color=("black", "#DDDDDD")),
         sg.Button("Mark Red",   key="-MARK_RED-",   button_color=("white", "#C62828")),
         sg.Text("   Warm Leads:", text_color="#A0A0A0"), sg.Text("0", key="-WARM-", text_color="#9EE493"),
         sg.Text("   Replies:", text_color="#A0A0A0"), sg.Text("0 / 0", key="-REPLRATE-", text_color="#FFFFFF")]
    ]

    # ---------- Dialer tab ----------
    dialer_mount = sg.Column(
        [[sg.Text("Loading dialer grid…", key="-DIAL_LOAD-", text_color="#9EE493")]],
        key="-DIAL_MOUNT-",
        pad=(0, 0),
        expand_x=True,
        expand_y=True,
        background_color=None,
        scrollable=False,
    )
    dialer_host = sg.Frame(
        "DIALER GRID",
        [[dialer_mount]],
        relief=sg.RELIEF_GROOVE, border_width=2, background_color="#1B1B1B",
        title_color="#9EE493", expand_x=True, expand_y=True, key="-DIAL_HOST-",
    )
    dialer_controls_right = [
        [sg.Text("Outcome:", text_color="#CCCCCC")],
        [sg.Button("🟢 Green", key="-DIAL_SET_GREEN-", button_color=("white", "#2E7D32"), size=(14, 1))],
        [sg.Button("⚪ Gray",  key="-DIAL_SET_GRAY-",  button_color=("black", "#DDDDDD"), size=(14, 1))],
        [sg.Button("🔴 Red",   key="-DIAL_SET_RED-",   button_color=("white", "#C62828"), size=(14, 1))],
        [sg.Text("Note:", text_color="#CCCCCC", pad=((0, 0), (10, 0)))],
        [sg.Multiline(key="-DIAL_NOTE-", size=(28, 6), font=("Consolas", 10), background_color="#111", text_color="#EEE")],
        [sg.Button("Confirm Call", key="-DIAL_CONFIRM-", size=(16, 2), disabled=True, button_color=("white", "#444444"))],
        [sg.Text("", key="-DIAL_MSG-", text_color="#A0FFA0", size=(28, 2))],
        [sg.Button("Add 100 Rows", key="-DIAL_ADD100-")],
    ]
    dialer_tab = [
        [sg.Column([[dialer_host]], expand_x=True, expand_y=True),
         sg.Column(dialer_controls_right, vertical_alignment="top", pad=((10, 0), (0, 0)))]
    ]

    # ---------- Warm Leads tab ----------
    warm_mount = sg.Column(
        [[sg.Text("Loading warm grid…", key="-WARM_LOAD-", text_color="#9EE493")]],
        key="-WARM_MOUNT-",
        pad=(0, 0),
        expand_x=True,
        expand_y=True,
        background_color=None,
        scrollable=False,
    )
    warm_host = sg.Frame(
        "WARM LEADS GRID",
        [[warm_mount]],
        relief=sg.RELIEF_GROOVE, border_width=2, background_color="#1B1B1B",
        title_color="#9EE493", expand_x=True, expand_y=True, key="-WARM_HOST-",
    )
    warm_controls_right = [
        [sg.Text("Outcome:", text_color="#CCCCCC")],
        [sg.Button("🟢 Green", key="-WARM_SET_GREEN-", button_color=("white", "#2E7D32"), size=(14, 1))],
        [sg.Button("⚪ Gray",  key="-WARM_SET_GRAY-",  button_color=("black", "#DDDDDD"), size=(14, 1))],
        [sg.Button("🔴 Red",   key="-WARM_SET_RED-",   button_color=("white", "#C62828"), size=(14, 1))],
        [sg.Text("Note:", text_color="#CCCCCC", pad=((0, 0), (10, 0)))],
        [sg.Multiline(key="-WARM_NOTE-", size=(28, 6), font=("Consolas", 10),
                      background_color="#111", text_color="#EEE", enable_events=True)],
        [sg.Button("Confirm", key="-WARM_CONFIRM-", size=(16, 2), disabled=True, button_color=("white", "#444444"))],
        [sg.Text("", key="-WARM_STATUS_SIDE-", text_color="#A0FFA0", size=(28, 2))],
        [sg.Button("Export Warm Leads CSV", key="-WARM_EXPORT-")],
        [sg.Button("Reload Warm", key="-WARM_RELOAD-")],
        [sg.Button("Add 100 Rows", key="-WARM_ADD100-")],
        [sg.Button("→ Confirm New Customer", key="-WARM_MARK_CUSTOMER-", button_color=("white", "#2E7D32"))],
        [sg.Text("", key="-WARM_STATUS-", text_color="#A0FFA0")],
    ]
    warm_tab = [
        [sg.Column([[warm_host]], expand_x=True, expand_y=True),
         sg.Column(warm_controls_right, vertical_alignment="top", pad=((10, 0), (0, 0)))]
    ]

    # ---------- Customers tab ----------
    cust_mount = sg.Column(
        [[sg.Text("Loading customers grid…", key="-CUST_LOAD-", text_color="#9EE493")]],
        key="-CUST_MOUNT-",
        pad=(0, 0),
        expand_x=True,
        expand_y=True,
        background_color=None,
        scrollable=False,
    )
    customers_host = sg.Frame(
        "CUSTOMERS GRID",
        [[cust_mount]],
        relief=sg.RELIEF_GROOVE, border_width=2, background_color="#1B1B1B",
        title_color="#9EE493", expand_x=True, expand_y=True, key="-CUST_HOST-",
    )
    customers_buttons_under = [
        sg.Button("Export Customers CSV", key="-CUST_EXPORT-"),
        sg.Button("Reload Customers", key="-CUST_RELOAD-"),
        sg.Button("Add 50 Rows", key="-CUST_ADD50-"),
        sg.Button("Add Order", key="-CUST_ADD_ORDER-", button_color=("white", "#2E7D32")),
        sg.Text("", key="-CUST_STATUS-", text_color="#A0FFA0")
    ]
    an_customer = [
        [sg.Text("CUSTOMER ANALYTICS", text_color="#9EE493")],
        [sg.Text("Total Sales"),  sg.Text("0.00", key="-AN_TOTALSALES-", text_color="#A0FFA0")],
        [sg.Text("CAC"),          sg.Text("0.00", key="-AN_CAC-",        text_color="#A0FFA0")],
        [sg.Text("LTV"),          sg.Text("0.00", key="-AN_LTV-",        text_color="#A0FFA0")],
        [sg.Text("LTV : CAC"),    sg.Text("1 : 0", key="-AN_CACLTV-",    text_color="#A0FFA0")],
        [sg.Text("Reorder Rate"), sg.Text("0%",   key="-AN_REORDER-",    text_color="#A0FFA0")],
    ]
    an_pipeline = [
        [sg.HorizontalSeparator(color="#4CAF50")],
        [sg.Text("PIPELINE ANALYTICS", text_color="#9EE493")],
        [sg.Text("Warm Leads Generated"),  sg.Text("0",  key="-AN_WARMS-",     text_color="#A0FFA0")],
        [sg.Text("New Accounts Acquired"), sg.Text("0",  key="-AN_NEWCUS-",    text_color="#A0FFA0")],
        [sg.Text("Close Rate"),            sg.Text("0%", key="-AN_CLOSERATE-", text_color="#A0FFA0")],
    ]
    analytics_panel = sg.Frame(
        "",
        [[sg.Column(an_customer, pad=(6, 6), expand_x=True, expand_y=False)],
         [sg.Column(an_pipeline, pad=(6, 0), expand_x=True, expand_y=False)]],
        relief=sg.RELIEF_GROOVE, border_width=2,
        background_color="#1B1B1B", title_color="#9EE493",
        expand_x=False, expand_y=False
    )
    customers_tab = [
        [sg.Column([[customers_host],
                    [sg.Column([customers_buttons_under], pad=(0, 0))]],
                   expand_x=True, expand_y=True),
         sg.Column([[analytics_panel]],
                   vertical_alignment="top",
                   pad=((10, 0), (0, 0)),
                   size=(320, 340))]
    ]

    # ---------- Map tab ----------
    map_tab = [
        [sg.Text("Customer Map", text_color="#9EE493", font=("Segoe UI", 14, "bold"))],
        [sg.Text("Opens a live Leaflet map with pins for each geocoded customer (Company, CLTV, Sales/Day).",
                 text_color="#CCCCCC")],
        [sg.Button("🗺️ Open Customer Map", key="-OPEN_MAP-", button_color=("white", "#2D6CDF"), size=(24, 2)),
         sg.Text("", key="-MAP_STATUS-", text_color="#A0FFA0")]
    ]

    # ---------- Compose layout ----------
    layout = [
        top_bar,
        banner_row,
        scoreboards_row,
        [sg.TabGroup([[sg.Tab("Email Leads",     leads_tab,     expand_x=True, expand_y=True),
                       sg.Tab("Email Campaigns", campaigns_tab, expand_x=True, expand_y=True),
                       sg.Tab("Email Results",   results_tab,   expand_x=True, expand_y=True),
                       sg.Tab("Dialer",          dialer_tab,    expand_x=True, expand_y=True),
                       sg.Tab("Warm Leads",      warm_tab,      expand_x=True, expand_y=True),
                       sg.Tab("Customers",       customers_tab, expand_x=True, expand_y=True),
                       sg.Tab("Map",             map_tab,       expand_x=True, expand_y=True)]],
                     key="-TABGROUP-",
                     expand_x=True, expand_y=True)]
    ]

    window = sg.Window(
        f"GrowthFarm — {app_version}",
        layout,
        finalize=True,
        resizable=True,
        grab_anywhere=True,
        background_color="#202020",
        size=(1200, 800),
        return_keyboard_events=True,
        enable_close_attempted_event=True,
    )

    # Ensure OS titlebar is present and usable
    try:
        window.TKroot.overrideredirect(False)
        try:
            window.TKroot.attributes("-topmost", False)
        except Exception:
            pass
    except Exception:
        pass

    # Start on-screen
    try:
        window.move(50, 50)
    except Exception:
        pass

    # Nudge key containers to expand fully
    for k in ("-TABGROUP-", "-LEADS_HOST-", "-DIAL_HOST-", "-WARM_HOST-", "-CUST_HOST-",
              "-LEADS_MOUNT-", "-DIAL_MOUNT-", "-WARM_MOUNT-", "-CUST_MOUNT-"):
        try:
            window[k].expand(expand_x=True, expand_y=True)
        except Exception:
            pass

    # Context: references the *inner mount columns* so logic can mount tksheet later
    context = {
        "leads_host": window["-LEADS_MOUNT-"],       # mount here
        "dialer_host": window["-DIAL_MOUNT-"],       # mount here
        "warm_host": window["-WARM_MOUNT-"],         # mount here
        "customers_host": window["-CUST_MOUNT-"],    # mount here
    }

    return window, context
