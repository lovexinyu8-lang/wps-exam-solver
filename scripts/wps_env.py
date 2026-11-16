"""Multi-version WPS Office COM environment discovery, isolation, and safe file locking."""

from __future__ import annotations

import gc
import json
import os
import shutil
import time
import zipfile
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator

try:
    import winreg
except ImportError:  # pragma: no cover - Windows-only module
    winreg = None  # type: ignore

try:
    import win32api
    import win32com.client.dynamic as dyn
except ImportError:  # pragma: no cover
    win32api = None  # type: ignore
    dyn = None  # type: ignore

PROGID_CANDIDATES: dict[str, tuple[str, ...]] = {
    "word": ("Kwps.Application", "KWPS.Application", "KWPS.Application.9", "WPS.Application"),
    "excel": ("Ket.Application", "KET.Application", "KET.Application.9", "ET.Application"),
    "ppt": ("Kwpp.Application", "KWPP.Application", "KWPP.Application.9", "WPP.Application"),
}

MS_PROGIDS: dict[str, str] = {
    "word": "Word.Application",
    "excel": "Excel.Application",
    "ppt": "PowerPoint.Application",
}

EXT_KIND_MAP: dict[str, str] = {
    ".docx": "word",
    ".doc": "word",
    ".wps": "word",
    ".wpt": "word",
    ".rtf": "word",
    ".xlsx": "excel",
    ".xls": "excel",
    ".et": "excel",
    ".ett": "excel",
    ".csv": "excel",
    ".pptx": "ppt",
    ".ppt": "ppt",
    ".dps": "ppt",
    ".dpt": "ppt",
}


def infer_kind(file_path: str) -> str:
    """Infer Office application kind ('word', 'excel', 'ppt') from file extension."""
    ext = Path(file_path).suffix.lower()
    if ext not in EXT_KIND_MAP:
        raise ValueError(f"Unsupported file extension: {ext!r} ({file_path})")
    return EXT_KIND_MAP[ext]


def _read_reg_default(subkey: str, wow_flag: int = 0) -> str | None:
    if winreg is None:
        return None
    access = winreg.KEY_READ | wow_flag
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, subkey, 0, access) as key:
            val, _ = winreg.QueryValueEx(key, "")
            return str(val) if val is not None else None
    except OSError:
        return None


def _extract_exe_path(command_str: str | None) -> str | None:
    if not command_str:
        return None
    s = command_str.strip()
    if s.startswith('"'):
        end = s.find('"', 1)
        if end > 1:
            return s[1:end]
    low = s.lower()
    exe_idx = low.find(".exe")
    if exe_idx != -1:
        return s[: exe_idx + 4].strip()
    return s.split()[0]


def _get_file_version(exe_path: str | None) -> str | None:
    if not exe_path or not os.path.exists(exe_path) or win32api is None:
        return None
    try:
        info = win32api.GetFileVersionInfo(exe_path, "\\")
        ms = info["FileVersionMS"]
        ls = info["FileVersionLS"]
        return f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
    except Exception:
        return None


def probe_progid(progid: str) -> dict[str, Any] | None:
    """Inspect Windows Registry registration for a single COM ProgID."""
    if winreg is None:
        return None
    clsid = _read_reg_default(f"{progid}\\CLSID")
    if not clsid:
        return None
    curver = _read_reg_default(f"{progid}\\CurVer")
    server64 = _read_reg_default(f"CLSID\\{clsid}\\LocalServer32", winreg.KEY_WOW64_64KEY)
    server32 = _read_reg_default(f"CLSID\\{clsid}\\LocalServer32", winreg.KEY_WOW64_32KEY)
    server = server64 or server32
    exe_path = _extract_exe_path(server)
    return {
        "progid": progid,
        "clsid": clsid,
        "curver": curver,
        "server": server,
        "exe_path": exe_path,
        "exe_exists": bool(exe_path and os.path.exists(exe_path)),
        "version": _get_file_version(exe_path),
    }


def detect_wps_environment() -> dict[str, Any]:
    """Detect installed WPS / MS Office COM ProgIDs, versions, and potential conflicts."""
    if os.name != "nt":
        return {"platform": os.name, "supported": False, "error": "WPS COM automation requires Windows."}

    wps_reg: dict[str, list[dict[str, Any]]] = {}
    for kind, candidates in PROGID_CANDIDATES.items():
        found: list[dict[str, Any]] = []
        for pid in candidates:
            info = probe_progid(pid)
            if info:
                found.append(info)
        wps_reg[kind] = found

    ms_reg: dict[str, dict[str, Any] | None] = {
        kind: probe_progid(pid) for kind, pid in MS_PROGIDS.items()
    }

    local_app_data = os.environ.get("LOCALAPPDATA", "")
    gen_py_dir = os.path.join(local_app_data, "Temp", "gen_py") if local_app_data else ""
    has_gen_py = bool(gen_py_dir and os.path.isdir(gen_py_dir))
    has_msoffice = any(v is not None for v in ms_reg.values())

    warnings: list[str] = []
    for kind, items in wps_reg.items():
        if not items:
            warnings.append(f"No WPS COM ProgID registered for {kind}.")
        elif not any(item["exe_exists"] for item in items):
            warnings.append(f"WPS COM ProgID for {kind} points to a missing executable.")
    if has_msoffice:
        warnings.append(
            "MS Office is also installed: always use win32com.client.dynamic.Dispatch "
            "to avoid gen_py early-binding collisions, and watch OLE embedded chart associations."
        )
    if has_gen_py:
        warnings.append(f"gen_py cache detected at {gen_py_dir}; dynamic.Dispatch is enforced.")

    return {
        "platform": "nt",
        "supported": True,
        "wps": wps_reg,
        "msoffice": ms_reg,
        "has_msoffice": has_msoffice,
        "has_gen_py": has_gen_py,
        "warnings": warnings,
    }


def resolve_progid(kind: str, preferred_progid: str | None = None) -> str:
    """Resolve the first registered WPS COM ProgID for the given document kind."""
    candidates = ([preferred_progid] if preferred_progid else []) + list(PROGID_CANDIDATES[kind])
    for pid in candidates:
        info = probe_progid(pid)
        if info and info["exe_exists"]:
            return pid
    for pid in candidates:
        if probe_progid(pid):
            return pid
    raise RuntimeError(
        f"No registered WPS COM ProgID found for {kind!r}. Checked: {', '.join(candidates)}"
    )


def inspect_container(file_path: str) -> dict[str, Any]:
    """Inspect file existence, lock files, write lock status, and binary/OOXML container format."""
    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"Target file not found: {abs_path}")

    p = Path(abs_path)
    lock_candidates = [
        p.with_name(f"~${p.name}"),
        p.with_name(f".~lock.{p.name}#"),
    ]
    active_locks = [str(lp) for lp in lock_candidates if lp.exists()]

    with open(abs_path, "rb") as fp:
        header = fp.read(8)

    if zipfile.is_zipfile(abs_path):
        container = "ooxml"
    elif header.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        container = "ole2"
    else:
        container = "unknown"

    return {
        "path": abs_path,
        "kind": infer_kind(abs_path),
        "size": p.stat().st_size,
        "container": container,
        "lock_files": active_locks,
    }


def check_write_lock(file_path: str) -> None:
    """Perform 1ms write-lock preflight before COM Open() to prevent modal lock popups."""
    abs_path = os.path.abspath(file_path)
    p = Path(abs_path)
    lock_file = p.with_name(f"~${p.name}")
    try:
        with open(abs_path, "r+b"):
            pass
    except PermissionError as exc:
        lock_hint = f" (lock file present: {lock_file})" if lock_file.exists() else ""
        raise PermissionError(
            f"Target file is locked by another process{lock_hint}: {abs_path}. "
            "Close it in WPS/Office before running COM automation."
        ) from exc


def safe_backup(file_path: str, backup_path: str | None = None) -> str:
    """Create a timestamped or explicit backup copy of the target file before editing."""
    abs_path = os.path.abspath(file_path)
    if backup_path is None:
        p = Path(abs_path)
        backup_path = str(p.with_name(f".{p.stem}.{int(time.time())}.bak{p.suffix}"))
    shutil.copy2(abs_path, backup_path)
    return backup_path


def restore_backup(backup_path: str, target_path: str) -> None:
    """Restore target_path from backup_path safely."""
    if os.path.exists(backup_path):
        shutil.copy2(backup_path, target_path)


def _collection_count(app: Any, kind: str) -> int:
    try:
        if kind == "word":
            return int(app.Documents.Count)
        if kind == "excel":
            return int(app.Workbooks.Count)
        if kind == "ppt":
            return int(app.Presentations.Count)
    except Exception:
        return 0
    return 0


def _set_quiet(app: Any, kind: str, visible: bool | None, alerts: bool) -> dict[str, Any]:
    prev: dict[str, Any] = {}
    try:
        prev["Visible"] = app.Visible
        if visible is not None and kind in ("word", "excel"):
            app.Visible = visible
    except Exception:
        pass
    try:
        prev["DisplayAlerts"] = app.DisplayAlerts
        if kind == "ppt":
            app.DisplayAlerts = 1 if not alerts else 2  # ppAlertsNone=1
        else:
            app.DisplayAlerts = alerts
    except Exception:
        pass
    return prev


@contextmanager
def wps_session(
    kind: str,
    file_path: str,
    *,
    read_only: bool = False,
    with_window: bool = False,
    backup: bool = False,
    keep_backup: bool = False,
    preferred_progid: str | None = None,
) -> Iterator[tuple[Any, Any, dict[str, Any]]]:
    """Open a WPS document via dynamic COM Dispatch with lock check, backup/rollback, and window isolation.

    Yields:
        (app, doc_obj, meta) where meta contains progid, backup_path, container, and pre_count.
    """
    if dyn is None:
        raise RuntimeError("pywin32 (win32com.client.dynamic) is required on Windows.")

    info = inspect_container(file_path)
    abs_path = info["path"]
    if not read_only:
        check_write_lock(abs_path)

    backup_path = safe_backup(abs_path) if (backup and not read_only) else None
    progid = resolve_progid(kind, preferred_progid=preferred_progid)
    app = dyn.Dispatch(progid)
    pre_count = _collection_count(app, kind)

    # Preserve foreground user visibility if WPS already has open documents.
    target_visible: bool | None = None if pre_count > 0 else False
    prev_state = _set_quiet(app, kind, visible=target_visible, alerts=False)
    doc_obj = None
    succeeded = False

    try:
        if kind == "word":
            doc_obj = app.Documents.Open(abs_path, ReadOnly=read_only)
        elif kind == "excel":
            doc_obj = app.Workbooks.Open(abs_path, ReadOnly=read_only)
        elif kind == "ppt":
            doc_obj = app.Presentations.Open(
                abs_path,
                ReadOnly=-1 if read_only else 0,
                Untitled=0,
                WithWindow=-1 if with_window else 0,
            )
        else:
            raise ValueError(f"Unknown kind: {kind!r}")

        meta = {
            "progid": progid,
            "path": abs_path,
            "container": info["container"],
            "backup_path": backup_path,
            "pre_count": pre_count,
        }
        yield app, doc_obj, meta
        succeeded = True
    finally:
        if doc_obj is not None:
            try:
                if kind in ("word", "excel"):
                    doc_obj.Close(SaveChanges=False)
                else:
                    doc_obj.Close()
            except Exception:
                pass
            doc_obj = None

        if not succeeded and backup_path and not read_only:
            try:
                restore_backup(backup_path, abs_path)
            except Exception:
                pass

        if backup_path and not keep_backup and succeeded:
            try:
                os.remove(backup_path)
            except Exception:
                pass

        post_count = _collection_count(app, kind)
        # Never quit an existing WPS process that already had open user/exam documents.
        if pre_count == 0 and post_count == 0:
            try:
                app.Quit()
            except Exception:
                pass
        else:
            for k, v in prev_state.items():
                try:
                    setattr(app, k, v)
                except Exception:
                    pass
        app = None
        gc.collect()


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Probe WPS COM environment and target file status.")
    parser.add_argument("--probe", action="store_true", help="Print WPS/MS Office COM registry probe.")
    parser.add_argument("file", nargs="?", help="Optional file path to inspect container & lock status.")
    args = parser.parse_args()

    out: dict[str, Any] = {}
    if args.probe or not args.file:
        out["environment"] = detect_wps_environment()
    if args.file:
        info = inspect_container(args.file)
        try:
            check_write_lock(args.file)
            info["writable"] = True
        except PermissionError as exc:
            info["writable"] = False
            info["lock_error"] = str(exc)
        out["file"] = info
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
