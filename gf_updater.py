# gf_updater.py
# Safe, GitHub-backed updater for GrowthFarm (Windows / Python build)
from __future__ import annotations

import os, sys, json, time, zipfile, shutil
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

# -------- Config / constants --------
APP_ROOT = Path(__file__).resolve().parent
APP_EXE  = APP_ROOT / "growthfarm.py"       # if running as script
IS_FROZEN = getattr(sys, "frozen", False)   # if you later bundle with PyInstaller
RUNTIME_PATH = Path(sys.executable).resolve() if IS_FROZEN else None

DEFAULT_REPO = "YOUR_GITHUB_USERNAME/growthfarm"  # <--- set via app.ini; this is a fallback
DEFAULT_ASSET_HINT = "growthfarm-windows.zip"      # release asset to download

INI_PATH = (APP_ROOT / "app.ini")
def _read_ini_section(section: str) -> dict:
    out, key = {}, None
    if not INI_PATH.exists(): return out
    try:
        for line in INI_PATH.read_text(encoding="utf-8").splitlines():
            s = line.strip()
            if not s or s.startswith(";") or s.startswith("#"): continue
            if s.startswith("[") and s.endswith("]"):
                key = s[1:-1].strip().lower()
                continue
            if key == section.lower() and "=" in s:
                k, v = s.split("=", 1)
                out[k.strip().lower()] = v.strip()
    except Exception:
        pass
    return out

def _get_repo_and_asset_hint():
    cfg = _read_ini_section("updates")
    repo = cfg.get("repo", DEFAULT_REPO)
    hint = cfg.get("asset_hint", DEFAULT_ASSET_HINT)
    return repo, hint

def _local_version() -> str:
    # Read growthfarm.APP_VERSION without importing the whole app twice if possible
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("growthfarm_local", APP_ROOT / "growthfarm.py")
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)  # type: ignore
        v = getattr(mod, "APP_VERSION", "")
        return str(v or "").strip()
    except Exception:
        return ""

def _parse_ver(v: str):
    # Accept YYYY-MM-DD or semver-ish; fall back to string compare
    s = (v or "").strip().lower().lstrip("v")
    # Try YYYY-MM-DD
    parts = s.replace(".", "-").split("-")
    try:
        nums = tuple(int(p) for p in parts if p.isdigit())
        return nums if nums else (0,)
    except Exception:
        return (0,)

def _api_get(url: str) -> dict | None:
    try:
        req = Request(url, headers={"User-Agent": "GrowthFarm-Updater"})
        with urlopen(req, timeout=20) as r:
            return json.loads(r.read().decode("utf-8"))
    except (URLError, HTTPError, TimeoutError, ValueError):
        return None

def _download(url: str, dest: Path) -> bool:
    try:
        req = Request(url, headers={"User-Agent": "GrowthFarm-Updater"})
        with urlopen(req, timeout=60) as r, dest.open("wb") as f:
            shutil.copyfileobj(r, f)
        return True
    except Exception as e:
        print("[updater] download failed:", e)
        return False

def _find_asset(assets: list[dict], hint: str) -> dict | None:
    hint_l = (hint or "").lower()
    for a in assets or []:
        name = (a.get("name") or "").lower()
        if hint_l in name:
            return a
    # fallback: first asset
    return (assets or [None])[0]

def _latest_release(repo: str) -> tuple[str, dict | None]:
    api = f"https://api.github.com/repos/{repo}/releases/latest"
    data = _api_get(api)
    if not data: return "", None
    tag = str(data.get("tag_name", "") or data.get("name", "")).strip()
    return tag, data

def _stage_extract(zip_path: Path) -> Path | None:
    stage = APP_ROOT / "_update_stage"
    if stage.exists():
        shutil.rmtree(stage, ignore_errors=True)
    stage.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(stage)
        return stage
    except Exception as e:
        print("[updater] extract failed:", e)
        return None

def _copy_tree(src: Path, dst: Path, exclude: set[str]):
    for root, dirs, files in os.walk(src):
        r = Path(root)
        rel = r.relative_to(src) if r != src else Path("")
        # Skip excluded top-level entries
        if rel.parts and rel.parts[0] in exclude:
            continue
        # Ensure dir
        (dst / rel).mkdir(parents=True, exist_ok=True)
        for fn in files:
            p = r / fn
            # never overwrite user data folders; we only replace app code
            if rel.parts and rel.parts[0] in exclude:
                continue
            shutil.copy2(p, dst / rel / fn)

def _build_swap_script(stage_dir: Path, relaunch_cmd: list[str]) -> Path:
    """
    Creates a .bat that waits for the current process to exit, copies files from stage,
    and relaunches GrowthFarm.
    """
    bat = APP_ROOT / "_swap_and_restart.bat"
    app_dir = APP_ROOT
    # Keep user data SAFE: it lives under %APPDATA%\GrowthFarm, so we exclude that entirely.
    # We only replace the app folder contents.
    script = rf"""@echo off
setlocal
REM Wait for the parent (this Python) to exit
ping 127.0.0.1 -n 2 >nul
REM Copy staged files over app directory
xcopy "{stage_dir}" "{app_dir}" /E /I /Y
REM Clean stage
rmdir /S /Q "{stage_dir}"
REM Relaunch
start "" {" ".join('"%s"' % c for c in relaunch_cmd)}
endlocal
"""
    bat.write_text(script, encoding="utf-8")
    return bat

def _relaunch_cmd() -> list[str]:
    # If frozen one day: return [Path to exe]
    if IS_FROZEN and RUNTIME_PATH:
        return [str(RUNTIME_PATH)]
    # Script mode
    return [sys.executable, str(APP_EXE)]

def check_for_update() -> tuple[bool, str, str]:
    """Returns (update_available, latest_version, local_version)."""
    local = _local_version()
    repo, hint = _get_repo_and_asset_hint()
    tag, _data = _latest_release(repo)
    if not tag:
        return (False, "", local)
    return (_parse_ver(tag) > _parse_ver(local), tag, local)

def run_update(window=None) -> bool:
    """
    Download newest release asset, stage it, swap files via a .bat, and exit.
    Returns True if the swap script is launched (i.e., update is proceeding).
    """
    repo, hint = _get_repo_and_asset_hint()
    tag, data = _latest_release(repo)
    if not tag or not data:
        if window: 
            try: window["-STATUS-"].update("No release info available.")
            except Exception: pass
        return False

    assets = data.get("assets", []) or []
    asset = _find_asset(assets, hint)
    if not asset:
        if window:
            try: window["-STATUS-"].update("No suitable asset found.")
            except Exception: pass
        return False

    url = asset.get("browser_download_url")
    if not url:
        return False

    dl_dir = APP_ROOT / "tmp_dl"
    dl_dir.mkdir(parents=True, exist_ok=True)
    zip_path = dl_dir / f"update_{tag}.zip"
    if window:
        try: window["-STATUS-"].update("Downloading update…")
        except Exception: pass
    ok = _download(url, zip_path)
    if not ok:
        if window:
            try: window["-STATUS-"].update("Download failed.")
            except Exception: pass
        return False

    if window:
        try: window["-STATUS-"].update("Extracting…")
        except Exception: pass
    stage = _stage_extract(zip_path)
    if not stage:
        if window:
            try: window["-STATUS-"].update("Extract failed.")
            except Exception: pass
        return False

    # Heuristic: if the zip contains a top-level folder, use it as the real root
    children = list(stage.iterdir())
    src = children[0] if len(children) == 1 and children[0].is_dir() else stage

    # Build swap script and launch it; exclude NOTHING here because user data is NOT in app dir
    bat = _build_swap_script(src, _relaunch_cmd())

    # Launch the swapper and exit current app
    try:
        os.startfile(str(bat))  # Windows-only
    except Exception:
        # Fallback
        import subprocess
        subprocess.Popen(["cmd", "/c", str(bat)], close_fds=True)
    return True

# -------- UI helper --------
def update_ui_flow(window=None):
    """High-level UI flow used by the -UPDATE- button."""
    try:
        if window:
            try: window["-STATUS-"].update("Checking for updates…")
            except Exception: pass
        has, latest, local = check_for_update()
        if not has:
            if window:
                try:
                    import PySimpleGUI as sg
                    sg.popup_ok(f"You are up to date.\n\nLocal: {local or '(unknown)'}\nLatest: {latest or '(unknown)'}",
                                keep_on_top=True)
                    window["-STATUS-"].update("Up to date")
                except Exception: pass
            return

        # Confirm
        if window:
            try:
                import PySimpleGUI as sg
                yn = sg.popup_yes_no(f"Update available: {local} → {latest}\n\nDownload and restart now?",
                                     keep_on_top=True)
                if yn != "Yes":
                    window["-STATUS-"].update("Update cancelled.")
                    return
            except Exception:
                pass

        ok = run_update(window)
        if ok and window:
            try:
                import PySimpleGUI as sg
                sg.popup_ok("Updating… the app will restart automatically.", keep_on_top=True)
            except Exception:
                pass
            # Ask outer loop to save and exit
            try: window.write_event_value("-PLEASE_EXIT_FOR_UPDATE-", True)
            except Exception: pass
        elif window:
            try: window["-STATUS-"].update("Update failed.")
            except Exception: pass
    except Exception as e:
        print("[updater] flow error:", e)
