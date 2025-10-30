# gf_campaigns.py
from __future__ import annotations
import json, csv, os
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Tuple

# Store paths we already have
from gf_store import (
    APP_DIR,
    RESULTS_PATH,
    HEADER_FIELDS,
    load_results_rows_sorted,
)

# Helpers from your toolkit
from gf_helpers import (
    apply_placeholders,    # fills {First Name}, {Company}, etc.
    blocks_to_html,        # text->HTML blocks (preserves newlines)
    pick_store,            # select Outlook store
    require_pywin32,       # True if pywin32 available/usable
    upsert_result,         # cache into results.csv
    _lead_row_from_email_company,
    _parse_any_datetime,
)

# ---------- Constants ----------
GROWTHFARM_SUBFOLDER = "GrowthFarm"   # Draft subfolder name under Outlook Drafts

# ---------- Paths ----------
CAMPAIGNS_DIR = APP_DIR / "campaigns"
CAMPAIGNS_DIR.mkdir(parents=True, exist_ok=True)

# Small prefs file (remembers last chosen campaign)
_PREFS_PATH = CAMPAIGNS_DIR / "_prefs.json"

def _atomic_write_text(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        f.write(text)
    tmp.replace(path)

# ---------- Normalization ----------
def normalize_campaign_steps(steps: List[Dict]) -> List[Dict]:
    """Ensure exactly 3 steps with keys: subject, body, delay_days (strings)."""
    out = []
    for i in range(3):
        base = {"subject": "", "body": "", "delay_days": "0"}
        try:
            s = steps[i] if i < len(steps) else {}
        except Exception:
            s = {}
        subj = (s.get("subject") or "").strip()
        body = (s.get("body") or "").strip()
        try:
            dd = str(s.get("delay_days", "0")).strip()
        except Exception:
            dd = "0"
        base.update({"subject": subj, "body": body, "delay_days": dd or "0"})
        out.append(base)
    return out

def normalize_campaign_settings(settings: Dict) -> Dict:
    """Currently just 'send_to_dialer_after' -> '1' or '0' strings."""
    st = dict(settings or {})
    val = str(st.get("send_to_dialer_after", "1")).strip().lower()
    st["send_to_dialer_after"] = "1" if val in ("1", "true", "yes", "on") else "0"
    return st

# ---------- Persistence (per-campaign JSON) ----------
def _campaign_path_for_key(key: str) -> Path:
    safe = (key or "default").strip()
    for ch in r'\/:*?"<>|':
        safe = safe.replace(ch, "_")
    return CAMPAIGNS_DIR / f"{safe}.json"

def list_campaign_keys() -> List[str]:
    keys = [p.stem for p in CAMPAIGNS_DIR.glob("*.json")]
    if "default" not in keys:
        keys.insert(0, "default")
    return sorted(set(keys), key=lambda k: (k != "default", k.lower()))

def load_campaign_by_key(key: str) -> Tuple[List[Dict], Dict]:
    p = _campaign_path_for_key(key)
    if not p.exists():
        return normalize_campaign_steps([]), normalize_campaign_settings({})
    try:
        with p.open("r", encoding="utf-8") as f:
            data = json.load(f)
        steps = normalize_campaign_steps(data.get("steps", []))
        settings = normalize_campaign_settings(data.get("settings", {}))
        return steps, settings
    except Exception:
        return normalize_campaign_steps([]), normalize_campaign_settings({})

def save_campaign_by_key(key: str, steps: List[Dict], settings: Dict):
    key = (key or "default").strip()
    steps = normalize_campaign_steps(steps or [])
    settings = normalize_campaign_settings(settings or {})
    payload = {"key": key, "steps": steps, "settings": settings, "saved_at": datetime.now().isoformat()}
    path = _campaign_path_for_key(key)
    _atomic_write_text(path, json.dumps(payload, ensure_ascii=False, indent=2))

def delete_campaign_by_key(key: str):
    p = _campaign_path_for_key(key)
    if p.exists():
        try:
            p.unlink()
        except Exception:
            pass

# ---------- UI helpers ----------
def summarize_campaign_for_table(key: str) -> List[str]:
    """Return [key, enabled_steps, delays_csv, to_dialer, auto_sync, hourly] (Resp% appended by UI)."""
    steps, settings = load_campaign_by_key(key)
    enabled = sum(1 for s in steps if (s.get("subject") or s.get("body")))
    delays = ",".join(str(s.get("delay_days", "0") or "0") for s in steps)
    to_dialer = "Yes" if settings.get("send_to_dialer_after", "1") == "1" else "No"
    auto_sync = "—"
    hourly = "—"
    return [key, str(enabled), delays, to_dialer, auto_sync, hourly]

# ---------- Response-rate helper used by UI (optional) ----------
def _response_rate_by_subjects(subjects_set) -> str:
    if not subjects_set:
        return ""
    try:
        rows = load_results_rows_sorted()
    except Exception:
        return ""
    sent = replied = 0
    for r in rows:
        subj = (r.get("Subject", "") or "").strip()
        if subj in subjects_set:
            if r.get("DateSent"):
                sent += 1
            if r.get("DateReplied"):
                replied += 1
    return "0.0%" if sent == 0 else f"{(replied / sent) * 100:.1f}%"

# ---------- Campaign enrollment CSV (simple queue) ----------
ENROLL_PATH = APP_DIR / "campaigns_enrollments.csv"

def _ensure_enroll_file():
    if not ENROLL_PATH.exists():
        with ENROLL_PATH.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["Ref", "Email", "Company", "CampaignKey", "Stage", "DivertToDialer"])

def campaigns_enroll(ref_short: str, email: str, company: str,
                     campaign_key: str = "default",
                     divert_to_dialer: bool = True):
    _ensure_enroll_file()
    rows = []
    exists = False
    ref_l = (ref_short or "").strip().lower()
    with ENROLL_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            if (r.get("Ref","") or "").strip().lower() == ref_l:
                exists = True
            rows.append(r)
    if not exists:
        rows.append({
            "Ref": ref_short or "",
            "Email": email or "",
            "Company": company or "",
            "CampaignKey": campaign_key or "default",
            "Stage": "0",
            "DivertToDialer": "1" if divert_to_dialer else "0",
        })
    with ENROLL_PATH.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["Ref","Email","Company","CampaignKey","Stage","DivertToDialer"])
        w.writeheader()
        for r in rows:
            w.writerow(r)

def campaigns_is_enrolled(ref_short: str) -> bool:
    _ensure_enroll_file()
    ref_l = (ref_short or "").strip().lower()
    with ENROLL_PATH.open("r", encoding="utf-8", newline="") as f:
        rdr = csv.DictReader(f)
        for r in rdr:
            if (r.get("Ref","") or "").strip().lower() == ref_l:
                return True
    return False

def campaigns_enroll_from_results_row(res_row: dict, campaign_key="default", divert_to_dialer=True):
    ref = res_row.get("Ref","") or ""
    email = res_row.get("Email","") or ""
    company = res_row.get("Company","") or ""
    if not ref:
        return
    campaigns_enroll(ref, email, company, campaign_key, divert_to_dialer)

def campaigns_bulk_enroll_from_status(status="gray", campaign_key="default", divert_to_dialer=True, max_rows=2000):
    try:
        rows = load_results_rows_sorted()
    except Exception:
        return 0
    count = 0
    st_l = (status or "").strip().lower()
    for r in rows:
        if count >= max_rows:
            break
        if (r.get("Status","") or "").strip().lower() == st_l and not (r.get("DateReplied") or "").strip():
            campaigns_enroll_from_results_row(r, campaign_key, divert_to_dialer)
            count += 1
    return count

# ---------- Outlook draft helpers ----------
def _ensure_outlook_folder_drafts_sub(session, name: str):
    """Returns/creates a subfolder under Drafts."""
    store = pick_store(session)
    drafts_root = store.GetDefaultFolder(16)  # olFolderDrafts
    for i in range(1, drafts_root.Folders.Count + 1):
        f = drafts_root.Folders.Item(i)
        if (f.Name or "").lower() == (name or "").strip().lower():
            return f
    return drafts_root.Folders.Add(name or GROWTHFARM_SUBFOLDER)

def _draft_one_outlook(ref_short: str, email: str, subj_text: str, body_text: str):
    """Create a single Outlook draft under the GrowthFarm subfolder."""
    import win32com.client as win32
    outlook = win32.Dispatch("Outlook.Application")
    session = outlook.GetNamespace("MAPI")
    target_folder = _ensure_outlook_folder_drafts_sub(session, GROWTHFARM_SUBFOLDER)
    body_html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"></head>
    <body style="margin:0;padding:0;">
      <div style="font-family:Segoe UI, Arial, sans-serif; font-size:14px; line-height:1.5; color:#111;">
        {blocks_to_html(body_text)}
        <!-- ref:{ref_short} -->
      </div>
    </body></html>"""
    msg = target_folder.Items.Add("IPM.Note")
    msg.To = email or ""
    msg.Subject = f"{subj_text} [ref:{ref_short}]"
    msg.BodyFormat = 2
    msg.HTMLBody = body_html
    msg.Save()
    # Update results.csv cache for visibility in the grid
    try:
        upsert_result(ref_short, email or "", "", "", subj_text)
    except Exception:
        pass
    return True

# ---------- Placeholders & row dict ----------
def _rowdict_for_placeholders(results_row: dict):
    """Prefer the original lead row (for First Name, etc.)."""
    lead = _lead_row_from_email_company(results_row.get("Email",""), results_row.get("Company",""))
    if lead:
        return lead
    d = {h: "" for h in HEADER_FIELDS}
    d["Email"] = results_row.get("Email","")
    d["Company"] = results_row.get("Company","")
    d["Industry"] = results_row.get("Industry","")
    return d

# ---------- Delay / stage helpers anchored to DateSent ----------
def _days_since(dt) -> int:
    if not dt:
        return 0
    try:
        return max(0, (datetime.now() - dt).days)
    except Exception:
        return 0

def _get_step_delays_for_key(campaign_key: str) -> Tuple[int, int]:
    """
    Returns (delay_e2_days, delay_e3_days) from the campaign definition.
    """
    try:
        steps, _settings = load_campaign_by_key(campaign_key or "default")
        steps = normalize_campaign_steps(steps)
        d2 = int(str(steps[1].get("delay_days", 0)).strip() or "0")  # step 2 delay
        d3 = int(str(steps[2].get("delay_days", 0)).strip() or "0")  # step 3 delay
        return (max(0, d2), max(0, d3))
    except Exception:
        return (3, 7)  # safe fallback

def _is_due_for_next(results_row: dict, next_stage: int, campaign_key: str) -> bool:
    """
    - next_stage 2: days_since(DateSent) >= delay(step2)
    - next_stage 3: days_since(DateSent) >= delay(step2) + delay(step3)
    """
    sent_dt = _parse_any_datetime(results_row.get("DateSent",""))
    if not sent_dt:
        return False
    d2, d3 = _get_step_delays_for_key(campaign_key)
    elapsed = _days_since(sent_dt)
    if next_stage == 2:
        return elapsed >= d2
    if next_stage == 3:
        return elapsed >= (d2 + d3)
    return False

def _get_subject_body_for_stage(campaign_key: str, stage_num: int) -> Tuple[str, str]:
    """Pull subject/body for the given stage (1..3). Provide safe defaults if missing."""
    steps, _settings = load_campaign_by_key(campaign_key or "default")
    steps = normalize_campaign_steps(steps)
    idx = max(1, min(3, stage_num)) - 1
    subj = (steps[idx].get("subject") or "").strip()
    body = (steps[idx].get("body") or "").strip()
    if not subj:
        subj = ["Quick hello for {Company}",
                "Following up for {Company}",
                "Worth a quick chat about {Company}?"][idx]
    if not body:
        body_defaults = [
            "Hi {First Name},\n\nWanted to share something relevant to {Company}.\n\nCheers,\nMe",
            "Hi {First Name},\n\nCircling back in case my note missed you.\n\nBest,\nMe",
            "Hi {First Name},\n\nLast follow-up from me—open to a quick call?\n\nThanks,\nMe",
        ]
        body = body_defaults[idx]
    return subj, body

# ---------- Draft next stage (public) ----------
def draft_next_stage_from_config(ref: str, email: str, company: str,
                                 campaign_key: str, next_stage: int) -> bool:
    """
    Create an Outlook draft for stage 1..3 under Drafts/ GrowthFarm.
    Preconditions:
      - pywin32 is available
      - there is a results.csv row for ref (not replied)
      - for stages > 1 the delays are satisfied based on DateSent
    """
    try:
        if not require_pywin32():
            return False

        # find results row by Ref
        rows = load_results_rows_sorted()
        r = None
        ref_l = (ref or "").strip().lower()
        for rr in rows:
            if (rr.get("Ref","") or "").strip().lower() == ref_l:
                r = rr
                break
        if not r:
            return False

        # already replied? skip
        if _parse_any_datetime(r.get("DateReplied","")):
            return False

        # target email
        target_email = (email or "").strip() or (r.get("Email","") or "").strip()
        if not target_email:
            return False

        # delays for 2/3
        if next_stage in (2, 3):
            if not _is_due_for_next(r, next_stage, campaign_key):
                return False

        # render subject/body with placeholders
        subj_tpl, body_tpl = _get_subject_body_for_stage(campaign_key, next_stage)
        rowd = _rowdict_for_placeholders(r)
        subj_text = apply_placeholders(subj_tpl, rowd)
        body_text = apply_placeholders(body_tpl, rowd)

        _draft_one_outlook(ref, target_email, subj_text, body_text)
        return True
    except Exception:
        return False

# ============================================================
# NEW: Outlook SEND + results logging (updates analytics)
# ============================================================

# Cached Outlook COM app instance
_OUTLOOK_APP = None

def _get_outlook_app():
    """Start or reuse the Outlook COM Application."""
    global _OUTLOOK_APP
    if _OUTLOOK_APP is not None:
        return _OUTLOOK_APP
    if not require_pywin32():
        raise RuntimeError("pywin32 not installed. Run: pip install pywin32")
    import win32com.client as win32
    _OUTLOOK_APP = win32.Dispatch("Outlook.Application")
    return _OUTLOOK_APP

def send_email_via_outlook(
    to_email: str,
    subject: str,
    body_html: str | None = None,
    body_text: str | None = None,
    attachments: list[str] | None = None,
) -> bool:
    """
    Create and send an email through Outlook desktop.
    Returns True if .Send() succeeded, False otherwise.
    """
    try:
        app = _get_outlook_app()
        # 0 = olMailItem
        mail = app.CreateItem(0)
        mail.To = to_email or ""
        mail.Subject = subject or ""
        if body_html:
            mail.BodyFormat = 2  # olFormatHTML
            mail.HTMLBody = body_html
        else:
            mail.Body = body_text or ""

        if attachments:
            try:
                for path in attachments:
                    if path:
                        mail.Attachments.Add(Path(path).resolve().as_posix())
            except Exception:
                pass

        mail.Send()
        return True
    except Exception as e:
        print(f"[campaigns] Outlook send failed: {e}")
        return False

def _ensure_results_csv_with_header(header: List[str]):
    """If results.csv doesn't exist, create it with the given header."""
    RESULTS_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not RESULTS_PATH.exists():
        with RESULTS_PATH.open("w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerow(header)

def _flexible_write_results_row(rowdict: Dict[str, str]):
    """
    Write a row to results.csv respecting existing headers if present.
    Ensures analytics fields exist: DateSent, To, Subject (also writes Email, Company, Ref, Campaign, Stage, Status when available).
    """
    default_header = ["Ref","Email","Company","Campaign","Stage","Subject","DateSent","DateReplied","Status","To"]
    RESULTS_PATH.parent.mkdir(parents=True, exist_ok=True)

    # Detect existing header (if any)
    header = None
    if RESULTS_PATH.exists():
        try:
            with RESULTS_PATH.open("r", encoding="utf-8", newline="") as f:
                rdr = csv.reader(f)
                header = next(rdr, None)
        except Exception:
            header = None

    if not header or not isinstance(header, list) or len(header) == 0:
        header = default_header
        _ensure_results_csv_with_header(header)

    # Build row in that header's order, fill blanks for missing keys
    out = []
    for col in header:
        out.append(rowdict.get(col, ""))
    with RESULTS_PATH.open("a", encoding="utf-8", newline="") as f:
        csv.writer(f).writerow(out)

def log_email_sent(*, ref: str = "", to_email: str = "", subject: str = "", campaign: str = "", company: str = "", stage: int | None = None, status: str = "sent"):
    """
    Append a 'sent' record so analytics updates Emails Sent today.
    Also plays nice with your existing results.csv (keeps whatever header it already has).
    """
    try:
        ts = datetime.now().isoformat(timespec="seconds")
        row = {
            "Ref": ref or "",
            "Email": to_email or "",
            "To": to_email or "",
            "Company": company or "",
            "Campaign": campaign or "",
            "Stage": str(stage or "") if stage else "",
            "Subject": subject or "",
            "DateSent": ts,
            "DateReplied": "",
            "Status": status or "sent",
        }
        _flexible_write_results_row(row)
    except Exception:
        # never crash UI for logging problems
        pass

def _render_subject_body_for(ref: str, email: str, company: str, campaign_key: str, stage_num: int) -> Tuple[str, str, Dict]:
    """
    Helper: fetch results row (by Ref), render placeholders, return (subject, body_text, results_row_dict).
    """
    rows = load_results_rows_sorted()
    r = None
    ref_l = (ref or "").strip().lower()
    for rr in rows:
        if (rr.get("Ref","") or "").strip().lower() == ref_l:
            r = rr
            break
    if not r:
        # synthesize minimal row if missing (still allow send)
        r = {"Ref": ref or "", "Email": email or "", "Company": company or ""}

    subj_tpl, body_tpl = _get_subject_body_for_stage(campaign_key, stage_num)
    rowd = _rowdict_for_placeholders(r)
    subj_text = apply_placeholders(subj_tpl, rowd)
    body_text = apply_placeholders(body_tpl, rowd)
    return subj_text, body_text, r

def send_stage_now(ref: str, email: str, company: str, campaign_key: str, stage_num: int, attachments: list[str] | None = None) -> bool:
    """
    Send stage 1..3 **now** via Outlook and log the send to results.csv.
    Respects delays for stages 2 and 3.
    Returns True if sent, False otherwise.
    """
    try:
        if not require_pywin32():
            return False

        # Find current results row (if any)
        rows = load_results_rows_sorted()
        r = None
        ref_l = (ref or "").strip().lower()
        for rr in rows:
            if (rr.get("Ref","") or "").strip().lower() == ref_l:
                r = rr
                break

        # If results row exists and it's a follow-up, enforce delay
        if r and stage_num in (2, 3) and not _is_due_for_next(r, stage_num, campaign_key):
            return False

        subj_text, body_text, r0 = _render_subject_body_for(ref, email, company, campaign_key, stage_num)

        # Build HTML (same style as drafts)
        body_html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"></head>
        <body style="margin:0;padding:0;">
          <div style="font-family:Segoe UI, Arial, sans-serif; font-size:14px; line-height:1.5; color:#111;">
            {blocks_to_html(body_text)}
            <!-- ref:{ref} -->
          </div>
        </body></html>"""

        target_email = (email or "").strip() or (r0.get("Email","") or "").strip()
        if not target_email:
            return False

        ok = send_email_via_outlook(
            to_email=target_email,
            subject=subj_text,
            body_html=body_html,
            body_text=None,
            attachments=attachments,
        )
        if ok:
            # Ensure it shows in your results UI AND analytics tile
            try:
                upsert_result(ref, target_email, company or "", "", subj_text)
            except Exception:
                pass
            log_email_sent(ref=ref, to_email=target_email, subject=subj_text, campaign=campaign_key, company=company, stage=stage_num, status="sent")
        return ok
    except Exception as e:
        print(f"[campaigns] send_stage_now error: {e}")
        return False

# ============================================================
# NEW: Campaign chooser popup for 'Fire Emails' flow
# ============================================================

def _campaign_subjects(key: str) -> List[str]:
    """Return normalized subject strings for a campaign (empty strings removed)."""
    steps, _ = load_campaign_by_key(key)
    steps = normalize_campaign_steps(steps)
    return [(s.get("subject") or "").strip() for s in steps if (s.get("subject") or "").strip()]

def _campaign_stats(key: str) -> Tuple[int, int, float]:
    """
    Compute (sent, replies, resp_pct) for a campaign by matching Subject to any
    of the campaign's step subjects.
    """
    subjects = set(_campaign_subjects(key))
    if not subjects:
        return (0, 0, 0.0)
    try:
        rows = load_results_rows_sorted()
    except Exception:
        rows = []
    sent = replied = 0
    for r in rows:
        subj = (r.get("Subject","") or "").strip()
        if subj in subjects:
            if (r.get("DateSent") or "").strip():
                sent += 1
            if (r.get("DateReplied") or "").strip():
                replied += 1
    pct = 0.0 if sent == 0 else (replied / sent) * 100.0
    return (sent, replied, pct)

def _load_last_selected() -> str:
    try:
        with _PREFS_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
        v = str(data.get("last_campaign", "")).strip()
        return v
    except Exception:
        return ""

def _save_last_selected(key: str):
    try:
        payload = {"last_campaign": (key or "").strip()}
        _atomic_write_text(_PREFS_PATH, json.dumps(payload, ensure_ascii=False, indent=2))
    except Exception:
        pass

def select_campaign_for_send(window=None) -> str | None:
    """
    Modal popup: lets user pick a campaign before sending.
    Returns the chosen campaign key, or None if cancelled.
    """
    # Lazy import PSG to keep this module safe for non-UI usage
    try:
        import PySimpleGUI as sg
    except Exception:
        # Fallback: headless environments
        keys = list_campaign_keys() or ["default"]
        chosen = _load_last_selected() or (keys[0] if keys else "default")
        return chosen

    keys = list_campaign_keys() or ["default"]

    # Build table data with stats
    table_rows = []
    for k in keys:
        sent, replied, pct = _campaign_stats(k)
        table_rows.append([k, sent, replied, f"{pct:.1f}%"])

    headings = ["Campaign", "Emails Sent", "Responses", "Resp %"]
    last = _load_last_selected()
    default_key = last if last in keys else (keys[0] if keys else "default")

    layout = [
        [sg.Text("Select a campaign for this batch:", text_color="#9EE493")],
        [sg.Table(values=table_rows,
                  headings=headings,
                  key="-CSEL_TABLE-",
                  num_rows=min(10, max(5, len(table_rows))),
                  auto_size_columns=False,
                  col_widths=[26, 12, 12, 8],
                  justification="left",
                  enable_events=True,
                  alternating_row_color="#2a2a2a",
                  text_color="#EEE", background_color="#111",
                  header_text_color="#FFF", header_background_color="#333")],
        [sg.Text("Campaign:", text_color="#CCCCCC"),
         sg.Combo(values=keys, default_value=default_key, key="-CSEL_KEY-", size=(32, 1), enable_events=True),
         sg.Push(),
         sg.Button("Refresh Stats", key="-CSEL_REFRESH-")],
        [sg.HorizontalSeparator(color="#333333")],
        [sg.Push(),
         sg.Button("Use Campaign", key="-CSEL_USE-", button_color=("white", "#2E7D32")),
         sg.Button("Cancel", key="-CSEL_CANCEL-", button_color=("white", "#555555"))]
    ]

    # Make it modal; center relative to main window if provided
    kwargs = dict(modal=True, finalize=True, keep_on_top=True)
    if window is not None:
        try:
            x = window.current_location()[0] + 60
            y = window.current_location()[1] + 60
            kwargs["location"] = (x, y)
        except Exception:
            pass

    win = sg.Window("Choose Campaign", layout, **kwargs)

    chosen = None
    try:
        while True:
            ev, vals = win.read()
            if ev in (sg.WIN_CLOSED, "-CSEL_CANCEL-"):
                chosen = None
                break

            if ev == "-CSEL_TABLE-":
                try:
                    selected = vals.get("-CSEL_TABLE-", [])
                    if selected:
                        idx = int(selected[0])
                        key_from_row = table_rows[idx][0]
                        win["-CSEL_KEY-"].update(value=key_from_row)
                except Exception:
                    pass

            elif ev == "-CSEL_REFRESH-":
                # recompute stats
                table_rows[:] = []
                for k in keys:
                    sent, replied, pct = _campaign_stats(k)
                    table_rows.append([k, sent, replied, f"{pct:.1f}%"])
                try:
                    win["-CSEL_TABLE-"].update(values=table_rows)
                except Exception:
                    pass

            elif ev == "-CSEL_USE-":
                k = (vals.get("-CSEL_KEY-") or "").strip()
                if not k:
                    sg.popup_error("Please select a campaign.", keep_on_top=True)
                    continue
                chosen = k
                _save_last_selected(k)
                break
    finally:
        try:
            win.close()
        except Exception:
            pass

    return chosen

