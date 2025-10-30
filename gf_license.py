# gf_license.py
from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, Tuple

from gf_store import get_app_dir

ACCOUNTS_FILENAME = "accounts.csv"

# We keep the original 4 headers (lowercase) for full backward compatibility.
REQUIRED_HEADERS = ["email", "password", "user", "company"]


# ---------- Paths ----------
def accounts_csv_path() -> Path:
    return Path(get_app_dir()) / ACCOUNTS_FILENAME


def accounts_json_path() -> Path:
    # Legacy file we want to remove if present
    return Path(get_app_dir()) / "accounts.json"


# ---------- Legacy cleanup ----------
def purge_legacy_json() -> None:
    """Silently delete old accounts.json if it exists."""
    try:
        p = accounts_json_path()
        if p.exists():
            # Python 3.8+ supports missing_ok; fall back if needed.
            try:
                p.unlink(missing_ok=True)  # type: ignore[arg-type]
            except TypeError:
                p.unlink()
    except Exception:
        pass


# ---------- Basic checks ----------
def accounts_csv_exists() -> bool:
    p = accounts_csv_path()
    if not p.exists():
        return False
    try:
        with p.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.reader(f)
            header = next(rdr, None)
            return bool(header)
    except Exception:
        return False


# ---------- IO helpers ----------
def _normalize_row_keys(row: Dict[str, str]) -> Dict[str, str]:
    """
    Return a dict with multiple key aliases so we can read files that have
    'Email' or 'email' or 'EMAIL', etc. Missing keys -> "".
    """
    raw = dict(row or {})
    # Build a lookup with several variants for each incoming column
    lut: Dict[str, str] = {}
    for k, v in raw.items():
        if not k:
            continue
        v2 = (v or "").strip()
        k_trim = k.strip()
        lut[k_trim] = v2
        lut[k_trim.lower()] = v2
        lut[k_trim.replace(" ", "_").lower()] = v2

    out: Dict[str, str] = {}
    for k in REQUIRED_HEADERS:
        # prefer exact lowercase, then Title Case, etc.
        out[k] = (
            lut.get(k, "")
            or lut.get(k.capitalize(), "")
            or lut.get(k.upper(), "")
            or lut.get(k.replace("_", " "), "")
        )
    return out


def _read_first_row(path: Path) -> Dict[str, str]:
    """
    Read the very first data row of accounts.csv and normalize the keys.
    Returns {} if file is missing or empty.
    """
    if not path.exists():
        return {}
    try:
        with path.open("r", encoding="utf-8", newline="") as f:
            rdr = csv.DictReader(f)
            first = next(iter(rdr), None)
            if not first:
                return {}
            return _normalize_row_keys(first)
    except Exception:
        return {}


# ---------- Public API ----------
def load_active_account() -> Dict[str, str]:
    """
    Load the single active account (first row). Removes legacy json first.
    Returns normalized dict with keys: email, password, user, company.
    Missing values come back as "".
    """
    purge_legacy_json()
    return _read_first_row(accounts_csv_path())


def _possessive(name: str) -> str:
    n = (name or "").strip()
    if not n:
        return ""
    return f"{n}'" if n[-1].lower() == "s" else f"{n}'s"


def get_banner_text(default: str = "GrowthFarm") -> str:
    """
    Return a clean display string WITHOUT the 'Growth Farm' suffix.
    The UI adds ' Growth Farm' itself to avoid doubling.
    """
    acct = load_active_account()
    if not acct:
        return default
    user = (acct.get("user", "") or "").strip()
    company = (acct.get("company", "") or "").strip()
    if user and company:
        return f"{_possessive(user)} {company}"
    if user:
        return _possessive(user)
    if company:
        return company
    return default


def create_or_replace_account(email: str, password: str, user: str, company: str) -> None:
    """
    Overwrite accounts.csv with exactly one row (CSV header + one record).
    Kept intentionally simple/compatible with your current first-run flow.
    """
    path = accounts_csv_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(REQUIRED_HEADERS)
        w.writerow([email or "", password or "", user or "", company or ""])


# ---------- Optional convenience helpers (safe no-ops if file is missing) ----------
def verify_login(email: str, password: str) -> bool:
    """
    Compare given credentials with stored ones (case-insensitive email).
    Returns True on match, False otherwise.
    """
    acct = load_active_account()
    if not acct:
        return False
    stored_email = (acct.get("email") or "").strip().lower()
    stored_pw = (acct.get("password") or "").strip()
    return (email or "").strip().lower() == stored_email and (password or "").strip() == stored_pw


def update_account(**fields: str) -> bool:
    """
    Update one or more fields in the single account row. Accepted keys:
    'email', 'password', 'user', 'company'. Returns True on success.
    """
    path = accounts_csv_path()
    acct = load_active_account()
    if not acct:
        return False
    for k in REQUIRED_HEADERS:
        if k in fields:
            acct[k] = (fields.get(k) or "").strip()
    try:
        with path.open("w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(REQUIRED_HEADERS)
            w.writerow([acct.get(k, "") for k in REQUIRED_HEADERS])
        return True
    except Exception:
        return False


def reset_account() -> None:
    """
    Remove the accounts.csv so app shows first-run prompt next launch.
    """
    try:
        p = accounts_csv_path()
        if p.exists():
            p.unlink()
    except Exception:
        pass


# ---------- Minimal validation you can call before saving ----------
def validate_inputs(email: str, password: str, user: str, company: str) -> Tuple[bool, str]:
    """
    Lightweight checks used by the first-run UI before calling create_or_replace_account().
    """
    if not all([(email or "").strip(), (password or "").strip(), (user or "").strip(), (company or "").strip()]):
        return False, "Please fill all four fields."
    if "@" not in email or "." not in email.split("@")[-1]:
        return False, "Email doesn’t look valid."
    if len(password.strip()) < 3:
        return False, "Password is too short."
    return True, ""
