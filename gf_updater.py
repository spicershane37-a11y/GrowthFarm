# gf_updater.py
from __future__ import annotations

import json
import re
import sys
import webbrowser
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

# ---- Configure this to your repo ----
REPO_OWNER = "spicershane37-a11y"
REPO_NAME  = "GrowthFarm"

GITHUB_API_LATEST = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest"
GITHUB_RELEASES   = f"https://github.com/{REPO_OWNER}/{REPO_NAME}/releases"
GITHUB_LATEST_TAG = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/tags"

# We use PySimpleGUI only for popups; imported lazily inside functions


def _http_get_json(url: str) -> dict | None:
    try:
        req = Request(url, headers={"User-Agent": "GrowthFarm-Updater"})
        with urlopen(req, timeout=10) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except (URLError, HTTPError, TimeoutError, ValueError):
        return None


def _pick_installer_asset(assets: list[dict]) -> str | None:
    """
    Try to find a Windows installer asset (.exe or .msi).
    Prefer files that contain 'GrowthFarm' and end correctly.
    """
    if not assets:
        return None
    # Score candidates: extension + name match
    ranked = []
    for a in assets:
        name = (a.get("name") or "").lower()
        dl = a.get("browser_download_url") or ""
        score = 0
        if name.endswith(".exe") or name.endswith(".msi"):
            score += 10
        if "growthfarm" in name:
            score += 5
        if score > 0:
            ranked.append((score, dl))
    if ranked:
        ranked.sort(reverse=True)
        return ranked[0][1]
    # fallback: first asset URL if nothing matched
    return assets[0].get("browser_download_url") or None


_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def _parse_version(v: str) -> tuple:
    """
    Accepts either YYYY-MM-DD or semver-ish x.y.z (or x.y).
    Returns a tuple that compares properly with Python tuple comparison.
    """
    s = (v or "").strip().lstrip("vV")
    if _DATE_RE.match(s):
        y, m, d = s.split("-")
        return ("date", int(y), int(m), int(d))
    nums = [int(n) if n.isdigit() else 0 for n in re.split(r"[^\d]+", s) if n != ""]
    while len(nums) < 3:
        nums.append(0)
    return ("semver", nums[0], nums[1], nums[2])


def is_newer(latest: str, current: str) -> bool:
    try:
        return _parse_version(latest) > _parse_version(current)
    except Exception:
        # If parsing fails, assume not newer
        return False


def get_latest_release_info() -> tuple[str | None, str | None, str | None]:
    """
    Returns (latest_version, notes, download_url) or (None, None, None) on failure.
    """
    data = _http_get_json(GITHUB_API_LATEST)
    if data and isinstance(data, dict):
        tag = (data.get("tag_name") or "").strip()
        notes = (data.get("body") or "").strip()
        assets = data.get("assets") or []
        dl = _pick_installer_asset(assets)
        return (tag or None, notes or None, dl or None)

    # Fallback: if releases API blocked, try tags to at least get a version
    tags = _http_get_json(GITHUB_LATEST_TAG)
    if isinstance(tags, list) and tags:
        tag = (tags[0].get("name") or "").strip()
        return (tag or None, None, None)

    return (None, None, None)


def check_and_prompt(window, current_version: str):
    """
    Check GitHub for a newer version. If found, prompt user to open the download.
    """
    import PySimpleGUI as sg

    try:
        window["-STATUS-"].update("Checking for updates...")
    except Exception:
        pass

    latest, notes, dl = get_latest_release_info()
    if not latest:
        sg.popup_ok(
            "Could not check for updates right now.\n\n"
            "You can always visit the releases page:\n" + GITHUB_RELEASES,
            keep_on_top=True,
            title="GrowthFarm — Update"
        )
        try:
            window["-STATUS-"].update("Update check failed")
        except Exception:
            pass
        return

    if is_newer(latest, current_version):
        msg = f"A new version is available.\n\nCurrent: {current_version}\nLatest:  {latest}\n"
        if notes:
            msg += f"\nRelease notes:\n{notes[:1000]}"  # keep popup sane
        msg += "\n\nDownload and install now?"

        yn = sg.popup_yes_no(msg, keep_on_top=True, title="GrowthFarm — Update Available")
        if yn == "Yes":
            url = dl or GITHUB_RELEASES
            try:
                webbrowser.open(url)
            except Exception:
                sg.popup_ok("Open this page to download:\n" + url, keep_on_top=True)
        try:
            window["-STATUS-"].update(f"Opened download for {latest}")
        except Exception:
            pass
    else:
        sg.popup_ok(f"You're up to date.\n\nCurrent version: {current_version}", keep_on_top=True)
        try:
            window["-STATUS-"].update("Up to date")
        except Exception:
            pass
