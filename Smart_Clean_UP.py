"""
================================================================================
SMART CLEAN UP PROJECT — REAL-TIME LOCAL HOST SERVER & MULTI-USER WEB APP
================================================================================
Run locally on the Admin Desktop:
    pip install fastapi uvicorn openpyxl python-multipart
    python smart_cleanup_server.py

Features:
- Zero dummy/sample data (automatically ignores old sample workspace paths).
- Native Windows File & Folder Browse dialogs right from the Admin Dashboard.
- Direct Excel / Native / Text folder import & live real-time synchronization.
- Token-based User Workspace with non-destructive "Updated <Column>" columns
  and "As Per Doc" one-click verification.
================================================================================
"""

import json
import os
import shutil
import socket
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
import openpyxl
from pydantic import BaseModel
import uvicorn


# ==============================================================================
# 1. CONFIGURATION, THREAD LOCK & REAL-TIME CONFIG STORAGE (NO DUMMY DATA)
# ==============================================================================
APP_HOST = "0.0.0.0"
APP_PORT = 8000

ADMIN_USER_ID = os.environ.get("SMART_CLEANUP_ADMIN_ID", "Jaydev Tejwani")
ADMIN_PASSWORD = os.environ.get("SMART_CLEANUP_ADMIN_PASS", "892371")

BASE_DIR = Path(__file__).resolve().parent / "smart_cleanup_workspace"
IMPORTED_DIR = BASE_DIR / "imported_sources"
CONFIG_FILE = BASE_DIR / "host_config.json"

EXCEL_WRITE_LOCK = threading.RLock()


def get_lan_ip() -> str:
    """Detect the Admin machine's local network IPv4 address for sharing."""
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.settimeout(1.0)
        sock.connect(("8.8.8.8", 80))
        ip = sock.getsockname()[0]
        sock.close()
        return ip
    except Exception:
        return "127.0.0.1"


def _is_legacy_sample_path(path_str: str) -> bool:
    """Detect if a saved path belongs to the old auto-generated dummy workspace."""
    if not path_str:
        return False
    norm = path_str.replace("\\", "/").lower()
    legacy_markers = (
        "smart_cleanup_workspace/database/assignment_master.xlsx",
        "smart_cleanup_workspace/native_files",
        "smart_cleanup_workspace/text_files",
    )
    return any(marker in norm for marker in legacy_markers)


def load_host_config() -> Dict[str, str]:
    """Load saved real-time folder/workbook paths and purge any legacy dummy paths."""
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    empty_cfg = {
        "excel_path": "",
        "sheet_name": "",
        "native_folder": "",
        "text_folder": "",
    }
    if CONFIG_FILE.exists():
        try:
            loaded = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                cleaned = dict(empty_cfg)
                had_legacy = False
                for k in empty_cfg:
                    val = str(loaded.get(k) or "").strip()
                    if _is_legacy_sample_path(val):
                        cleaned[k] = ""
                        had_legacy = True
                    else:
                        cleaned[k] = val
                if had_legacy:
                    CONFIG_FILE.write_text(json.dumps(cleaned, indent=2), encoding="utf-8")
                return cleaned
        except Exception:
            pass

    CONFIG_FILE.write_text(json.dumps(empty_cfg, indent=2), encoding="utf-8")
    return empty_cfg


HOST_CONFIG: Dict[str, str] = load_host_config()


def save_host_config() -> None:
    """Persist Admin source folder/workbook paths safely."""
    with EXCEL_WRITE_LOCK:
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_FILE.write_text(json.dumps(HOST_CONFIG, indent=2), encoding="utf-8")


# ==============================================================================
# 2. NATIVE OS FILE / FOLDER PICKER HELPER (RUNS ON ADMIN HOST DESKTOP)
# ==============================================================================
def open_native_os_dialog(mode: str) -> str:
    """
    Open a native Windows/macOS/Linux file or folder selection dialog on the Host PC.
    mode: 'excel' | 'native_folder' | 'text_folder'
    """
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as exc:
        raise RuntimeError(f"Native GUI picker unavailable on this OS: {exc}")

    selected_path = ""
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    root.update()

    try:
        if mode == "excel":
            selected_path = filedialog.askopenfilename(
                parent=root,
                title="Select Real-Time Master Excel Workbook (.xlsx)",
                filetypes=[
                    ("Excel Workbooks", "*.xlsx *.xlsm"),
                    ("All Files", "*.*"),
                ],
            )
        elif mode == "native_folder":
            selected_path = filedialog.askdirectory(
                parent=root,
                title="Select Native Files Folder (PDF, JPG, PNG, DOCX, XLSX, TIF...)",
                mustexist=True,
            )
        elif mode == "text_folder":
            selected_path = filedialog.askdirectory(
                parent=root,
                title="Select Text / OCR Files Folder (.txt)",
                mustexist=True,
            )
    finally:
        try:
            root.destroy()
        except Exception:
            pass

    return str(selected_path or "").strip()


# ==============================================================================
# 3. REAL-TIME EXCEL WORKBOOK ENGINE & DOCUMENT FINDER HELPERS
# ==============================================================================
def normalize_header_key(header: str) -> str:
    return "".join(ch.lower() for ch in str(header or "") if ch.isalnum())


def resolve_sheet(wb: openpyxl.Workbook, preferred_sheet: str = ""):
    if preferred_sheet and preferred_sheet in wb.sheetnames:
        return wb[preferred_sheet]
    return wb.active


def _match_file_in_folder(doc_id: str, folder_str: str, prefer_ext: Optional[str] = None) -> Optional[Path]:
    """
    Find a real file matching doc_id in folder_str (or immediate subfolders).
    Matches exact stem (e.g., DOC-101.pdf) first, then prefix matches (e.g., DOC-101_scan.pdf).
    """
    if not doc_id or not folder_str:
        return None
    folder = Path(folder_str)
    if not folder.exists() or not folder.is_dir():
        return None

    clean_id = doc_id.strip()
    target_lower = clean_id.lower()

    if prefer_ext:
        direct = folder / f"{clean_id}.{prefer_ext.lstrip('.')}"
        if direct.exists() and direct.is_file():
            return direct

    exact_matches: List[Path] = []
    prefix_matches: List[Path] = []

    try:
        entries = list(folder.iterdir())
    except Exception:
        return None

    for entry in entries:
        if entry.is_file():
            stem_lower = entry.stem.strip().lower()
            if stem_lower == target_lower:
                exact_matches.append(entry)
            elif (
                stem_lower.startswith(target_lower + "_")
                or stem_lower.startswith(target_lower + "-")
                or stem_lower.startswith(target_lower + " ")
            ):
                prefix_matches.append(entry)

    if exact_matches:
        if prefer_ext:
            for m in exact_matches:
                if m.suffix.lstrip(".").lower() == prefer_ext.lower():
                    return m
        return exact_matches[0]

    if prefix_matches:
        return sorted(prefix_matches, key=lambda p: len(p.name))[0]

    for sub in entries:
        if sub.is_dir():
            try:
                for entry in sub.iterdir():
                    if entry.is_file() and entry.stem.strip().lower() == target_lower:
                        return entry
            except Exception:
                continue

    return None


def find_native_file_for_doc(doc_id: str, native_folder_str: str) -> Optional[Path]:
    return _match_file_in_folder(doc_id, native_folder_str, prefer_ext=None)


def find_text_file_for_doc(doc_id: str, text_folder_str: str) -> Optional[Path]:
    return _match_file_in_folder(doc_id, text_folder_str, prefer_ext="txt")


def read_excel_records(token_filter: Optional[str] = None) -> Dict[str, Any]:
    """Read live rows and headers directly from the connected Excel file."""
    excel_path_str = HOST_CONFIG.get("excel_path", "").strip()
    if not excel_path_str:
        raise FileNotFoundError("No Excel workbook connected yet. Admin must select an .xlsx file first.")

    excel_path = Path(excel_path_str)
    if not excel_path.exists() or not excel_path.is_file():
        raise FileNotFoundError(f"Connected Excel file not found on disk: {excel_path_str}")

    with EXCEL_WRITE_LOCK:
        wb = openpyxl.load_workbook(str(excel_path), data_only=True)
        ws = resolve_sheet(wb, HOST_CONFIG.get("sheet_name", ""))
        sheet_names = list(wb.sheetnames)
        active_sheet_name = ws.title

        raw_headers: List[str] = []
        max_col = ws.max_column or 0
        for col_idx in range(1, max_col + 1):
            val = ws.cell(row=1, column=col_idx).value
            raw_headers.append(str(val).strip() if val is not None and str(val).strip() else f"Column {col_idx}")

        doc_id_col = None
        token_col = None
        for idx, h in enumerate(raw_headers, start=1):
            norm = normalize_header_key(h)
            if doc_id_col is None and norm in (
                "documentid", "docid", "id", "documentno", "docno", "documentname", "filename", "bates", "batesid"
            ):
                doc_id_col = idx
            if token_col is None and norm in (
                "tokenid", "token", "assigntoken", "tokenno", "userid", "batchid", "queueid"
            ):
                token_col = idx

        if doc_id_col is None and raw_headers:
            doc_id_col = 1

        records: List[Dict[str, Any]] = []
        native_folder = HOST_CONFIG.get("native_folder", "")
        text_folder = HOST_CONFIG.get("text_folder", "")

        for row_idx in range(2, (ws.max_row or 1) + 1):
            row_vals = [ws.cell(row=row_idx, column=c).value for c in range(1, max_col + 1)]
            if all(v is None or str(v).strip() == "" for v in row_vals):
                continue

            doc_id_val = ""
            if doc_id_col and doc_id_col <= len(row_vals) and row_vals[doc_id_col - 1] is not None:
                doc_id_val = str(row_vals[doc_id_col - 1]).strip()

            token_val = ""
            if token_col and token_col <= len(row_vals) and row_vals[token_col - 1] is not None:
                token_val = str(row_vals[token_col - 1]).strip()

            if token_filter is not None and str(token_filter).strip() != "":
                if token_val.lower() != str(token_filter).strip().lower():
                    continue

            cells: Dict[str, str] = {}
            for col_idx, header_name in enumerate(raw_headers, start=1):
                v = row_vals[col_idx - 1]
                if isinstance(v, datetime):
                    cells[header_name] = (
                        v.strftime("%Y-%m-%d %H:%M:%S")
                        if (v.hour or v.minute or v.second)
                        else v.strftime("%Y-%m-%d")
                    )
                else:
                    cells[header_name] = "" if v is None else str(v)

            native_path = find_native_file_for_doc(doc_id_val, native_folder)
            text_path = find_text_file_for_doc(doc_id_val, text_folder)

            records.append({
                "row_index": row_idx,
                "doc_id": doc_id_val,
                "token_id": token_val,
                "cells": cells,
                "native_file": native_path.name if native_path else "",
                "native_ext": native_path.suffix.lstrip(".").lower() if native_path and native_path.suffix else "",
                "has_native": bool(native_path),
                "text_file": text_path.name if text_path else "",
                "has_text": bool(text_path),
            })

        wb.close()

    return {
        "headers": raw_headers,
        "rows": records,
        "sheet_names": sheet_names,
        "active_sheet": active_sheet_name,
        "has_token_column": token_col is not None,
    }


def _find_or_create_column(ws, header_name: str) -> int:
    """Find an existing header column case-insensitively or create a new one at the end."""
    target_norm = normalize_header_key(header_name)
    for col_idx in range(1, (ws.max_column or 0) + 1):
        existing = ws.cell(row=1, column=col_idx).value
        if existing is not None and normalize_header_key(str(existing)) == target_norm:
            return col_idx
    new_col = (ws.max_column or 0) + 1
    ws.cell(row=1, column=new_col, value=header_name)
    return new_col


def update_excel_row_in_place(
    row_index: int,
    updates: Dict[str, str],
    as_per_doc: bool,
    user_name: str,
    token_id: str,
) -> Dict[str, Any]:
    """Non-destructively update a row in the connected Excel file in real time."""
    excel_path_str = HOST_CONFIG.get("excel_path", "").strip()
    excel_path = Path(excel_path_str)
    if not excel_path_str or not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path_str}")

    with EXCEL_WRITE_LOCK:
        wb = openpyxl.load_workbook(str(excel_path))
        ws = resolve_sheet(wb, HOST_CONFIG.get("sheet_name", ""))

        if row_index < 2 or row_index > ws.max_row:
            wb.close()
            raise ValueError(f"Invalid row index: {row_index}")

        token_col = None
        for idx in range(1, ws.max_column + 1):
            if normalize_header_key(str(ws.cell(row=1, column=idx).value or "")) in (
                "tokenid", "token", "assigntoken", "tokenno", "userid", "batchid", "queueid"
            ):
                token_col = idx
                break
        if token_col:
            row_token = str(ws.cell(row=row_index, column=token_col).value or "").strip()
            if row_token and token_id and row_token.lower() != token_id.strip().lower():
                wb.close()
                raise PermissionError(f"Row {row_index} belongs to Token ID '{row_token}', not '{token_id}'.")

        remark_col = _find_or_create_column(ws, "Remark / Action")
        updated_by_col = _find_or_create_column(ws, "Updated By")
        user_name_col = _find_or_create_column(ws, "User Name")
        timestamp_col = _find_or_create_column(ws, "Last Updated")

        if as_per_doc:
            ws.cell(row=row_index, column=remark_col, value="As Per Doc")
            for field_label in updates.keys():
                clean_label = field_label.strip()
                if not clean_label:
                    continue
                target_header = clean_label if clean_label.lower().startswith("updated ") else f"Updated {clean_label}"
                upd_col = _find_or_create_column(ws, target_header)
                ws.cell(row=row_index, column=upd_col, value="As Per Doc")
        else:
            wrote_any = False
            for field_label, new_val in updates.items():
                clean_label = field_label.strip()
                val_str = "" if new_val is None else str(new_val).strip()
                if not clean_label or val_str == "":
                    continue
                target_header = clean_label if clean_label.lower().startswith("updated ") else f"Updated {clean_label}"
                upd_col = _find_or_create_column(ws, target_header)
                ws.cell(row=row_index, column=upd_col, value=val_str)
                wrote_any = True
            if wrote_any:
                ws.cell(row=row_index, column=remark_col, value="Updated")
            else:
                wb.close()
                raise ValueError("Please enter at least one updated field value or click 'As Per Doc'.")

        actor_str = f"{user_name} (Token {token_id})" if token_id else user_name
        ws.cell(row=row_index, column=user_name_col, value=user_name)
        ws.cell(row=row_index, column=updated_by_col, value=actor_str)
        ws.cell(row=row_index, column=timestamp_col, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        try:
            wb.save(str(excel_path))
        except PermissionError:
            wb.close()
            raise PermissionError(
                "Excel file is currently locked/open in Microsoft Excel. "
                "Please close the Excel window on the Admin PC and retry saving."
            )
        wb.close()

    return {"status": "ok", "row_index": row_index}


def admin_update_row_in_excel(
    row_index: int,
    token_id: Optional[str] = None,
    cells: Optional[Dict[str, str]] = None,
) -> Dict[str, Any]:
    """Allow Admin to update Token ID or cell values directly in the connected Excel file."""
    excel_path_str = HOST_CONFIG.get("excel_path", "").strip()
    excel_path = Path(excel_path_str)
    if not excel_path_str or not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path_str}")

    with EXCEL_WRITE_LOCK:
        wb = openpyxl.load_workbook(str(excel_path))
        ws = resolve_sheet(wb, HOST_CONFIG.get("sheet_name", ""))
        if row_index < 2 or row_index > ws.max_row:
            wb.close()
            raise ValueError(f"Invalid row index: {row_index}")

        header_to_col: Dict[str, int] = {}
        token_col = None
        for c in range(1, ws.max_column + 1):
            h_val = ws.cell(row=1, column=c).value
            if h_val is not None:
                h_str = str(h_val).strip()
                header_to_col[h_str] = c
                if normalize_header_key(h_str) in (
                    "tokenid", "token", "assigntoken", "tokenno", "userid", "batchid", "queueid"
                ):
                    token_col = c

        if token_id is not None:
            if token_col is None:
                token_col = _find_or_create_column(ws, "Token ID")
            ws.cell(row=row_index, column=token_col, value=str(token_id).strip())

        if cells:
            for h_name, val in cells.items():
                if h_name in header_to_col:
                    ws.cell(row=row_index, column=header_to_col[h_name], value=val)

        try:
            wb.save(str(excel_path))
        except PermissionError:
            wb.close()
            raise PermissionError(
                "Excel file is currently open/locked in Microsoft Excel. Close it and try again."
            )
        wb.close()

    return {"status": "ok", "row_index": row_index}


def admin_append_row_to_excel(row_values: Dict[str, str]) -> Dict[str, Any]:
    """Append a new row to the connected Excel workbook."""
    excel_path_str = HOST_CONFIG.get("excel_path", "").strip()
    excel_path = Path(excel_path_str)
    if not excel_path_str or not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path_str}")

    with EXCEL_WRITE_LOCK:
        wb = openpyxl.load_workbook(str(excel_path))
        ws = resolve_sheet(wb, HOST_CONFIG.get("sheet_name", ""))
        next_row = (ws.max_row or 1) + 1

        for header_name, val in row_values.items():
            if header_name.strip():
                col_idx = _find_or_create_column(ws, header_name.strip())
                ws.cell(row=next_row, column=col_idx, value=val)

        try:
            wb.save(str(excel_path))
        except PermissionError:
            wb.close()
            raise PermissionError(
                "Excel file is currently open/locked in Microsoft Excel. Close it and try again."
            )
        wb.close()

    return {"status": "ok", "row_index": next_row}


# ==============================================================================
# 4. REQUEST MODELS & FASTAPI APPLICATION
# ==============================================================================
class AdminLoginRequest(BaseModel):
    user_id: str
    password: str


class UserLoginRequest(BaseModel):
    name: str
    token_id: str


class SourceConfigRequest(BaseModel):
    excel_path: str
    sheet_name: Optional[str] = ""
    native_folder: str
    text_folder: str


class InspectExcelRequest(BaseModel):
    excel_path: str


class BrowseRequest(BaseModel):
    mode: str  # 'excel' | 'native_folder' | 'text_folder'


class UserRowSaveRequest(BaseModel):
    row_index: int
    user_name: str
    token_id: str
    updates: Dict[str, str] = {}
    as_per_doc: bool = False


class AdminRowEditRequest(BaseModel):
    row_index: int
    token_id: Optional[str] = None
    cells: Optional[Dict[str, str]] = None


class AdminAddRowRequest(BaseModel):
    values: Dict[str, str]


app = FastAPI(title="Smart Clean UP Real-Time Server", version="3.1.0")


@app.get("/", response_class=HTMLResponse)
def serve_frontend() -> HTMLResponse:
    return HTMLResponse(content=FRONTEND_HTML)


@app.get("/api/host-info")
def api_host_info() -> Dict[str, Any]:
    lan_ip = get_lan_ip()
    excel_str = HOST_CONFIG.get("excel_path", "").strip()
    native_str = HOST_CONFIG.get("native_folder", "").strip()
    text_str = HOST_CONFIG.get("text_folder", "").strip()

    excel_path = Path(excel_str) if excel_str else None
    native_path = Path(native_str) if native_str else None
    text_path = Path(text_str) if text_str else None

    excel_ok = bool(excel_path and excel_path.exists() and excel_path.is_file())
    native_ok = bool(native_path and native_path.exists() and native_path.is_dir())
    text_ok = bool(text_path and text_path.exists() and text_path.is_dir())

    native_count = 0
    if native_ok and native_path:
        try:
            native_count = sum(1 for p in native_path.iterdir() if p.is_file())
        except Exception:
            native_count = 0

    text_count = 0
    if text_ok and text_path:
        try:
            text_count = sum(1 for p in text_path.iterdir() if p.is_file())
        except Exception:
            text_count = 0

    return {
        "lan_ip": lan_ip,
        "port": APP_PORT,
        "share_url": f"http://{lan_ip}:{APP_PORT}",
        "local_url": f"http://localhost:{APP_PORT}",
        "config": HOST_CONFIG,
        "status": {
            "excel_exists": excel_ok,
            "excel_name": excel_path.name if excel_ok and excel_path else "",
            "native_exists": native_ok,
            "text_exists": text_ok,
            "native_file_count": native_count,
            "text_file_count": text_count,
            "ready_for_users": excel_ok and native_ok and text_ok,
        },
    }


@app.post("/api/admin/login")
def api_admin_login(req: AdminLoginRequest) -> Dict[str, Any]:
    if req.user_id.strip() == ADMIN_USER_ID and req.password == ADMIN_PASSWORD:
        return {"ok": True, "role": "admin", "user_id": req.user_id.strip()}
    raise HTTPException(status_code=401, detail="Invalid Admin User ID or Password.")


@app.post("/api/admin/browse")
def api_admin_browse(req: BrowseRequest) -> Dict[str, Any]:
    """Trigger native OS file/folder dialog on the Admin Host PC and auto-save if valid."""
    if req.mode not in ("excel", "native_folder", "text_folder"):
        raise HTTPException(status_code=400, detail="Invalid browse mode.")

    try:
        chosen = open_native_os_dialog(req.mode)
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc))

    if not chosen:
        return {"ok": False, "cancelled": True, "path": ""}

    p = Path(chosen)
    sheet_names: List[str] = []
    active_sheet = ""

    if req.mode == "excel":
        if not p.exists() or not p.is_file():
            raise HTTPException(status_code=400, detail="Selected Excel file does not exist.")
        try:
            with EXCEL_WRITE_LOCK:
                wb = openpyxl.load_workbook(str(p), read_only=True)
                sheet_names = list(wb.sheetnames)
                active_sheet = wb.active.title if wb.active else (sheet_names[0] if sheet_names else "")
                wb.close()
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Could not read Excel workbook: {exc}")

        HOST_CONFIG["excel_path"] = str(p)
        HOST_CONFIG["sheet_name"] = active_sheet
        save_host_config()

    elif req.mode == "native_folder":
        if not p.exists() or not p.is_dir():
            raise HTTPException(status_code=400, detail="Selected Native folder does not exist.")
        HOST_CONFIG["native_folder"] = str(p)
        save_host_config()

    elif req.mode == "text_folder":
        if not p.exists() or not p.is_dir():
            raise HTTPException(status_code=400, detail="Selected Text folder does not exist.")
        HOST_CONFIG["text_folder"] = str(p)
        save_host_config()

    return {
        "ok": True,
        "cancelled": False,
        "mode": req.mode,
        "path": str(p),
        "sheet_names": sheet_names,
        "active_sheet": active_sheet,
        "config": HOST_CONFIG,
    }


@app.post("/api/admin/import-upload")
async def api_admin_import_upload(
    target: str = Form(...),  # 'excel' | 'native' | 'text'
    files: List[UploadFile] = File(...),
) -> Dict[str, Any]:
    """Allow importing Excel file or Native/Text files directly via browser upload."""
    if not files:
        raise HTTPException(status_code=400, detail="No files selected for import.")

    IMPORTED_DIR.mkdir(parents=True, exist_ok=True)

    if target == "excel":
        f = files[0]
        fname = Path(f.filename or "Master_Workbook.xlsx").name
        if not fname.lower().endswith((".xlsx", ".xlsm")):
            raise HTTPException(status_code=400, detail="Please upload an .xlsx or .xlsm Excel file.")
        dest_dir = IMPORTED_DIR / "excel"
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest_path = dest_dir / fname
        content = await f.read()
        dest_path.write_bytes(content)

        with EXCEL_WRITE_LOCK:
            wb = openpyxl.load_workbook(str(dest_path), read_only=True)
            sheet_names = list(wb.sheetnames)
            active_sheet = wb.active.title if wb.active else (sheet_names[0] if sheet_names else "")
            wb.close()

        HOST_CONFIG["excel_path"] = str(dest_path)
        HOST_CONFIG["sheet_name"] = active_sheet
        save_host_config()
        return {
            "ok": True,
            "path": str(dest_path),
            "sheet_names": sheet_names,
            "active_sheet": active_sheet,
            "count": 1,
        }

    elif target in ("native", "text"):
        sub_name = "native_imported" if target == "native" else "text_imported"
        dest_dir = IMPORTED_DIR / sub_name
        dest_dir.mkdir(parents=True, exist_ok=True)
        saved_count = 0
        for f in files:
            if not f.filename:
                continue
            clean_name = Path(f.filename).name
            if not clean_name:
                continue
            data = await f.read()
            (dest_dir / clean_name).write_bytes(data)
            saved_count += 1

        cfg_key = "native_folder" if target == "native" else "text_folder"
        HOST_CONFIG[cfg_key] = str(dest_dir)
        save_host_config()
        return {"ok": True, "path": str(dest_dir), "count": saved_count}

    raise HTTPException(status_code=400, detail="Invalid import target.")


@app.post("/api/admin/disconnect")
def api_admin_disconnect() -> Dict[str, Any]:
    """Reset connected sources so Admin can attach a brand-new dataset."""
    HOST_CONFIG["excel_path"] = ""
    HOST_CONFIG["sheet_name"] = ""
    HOST_CONFIG["native_folder"] = ""
    HOST_CONFIG["text_folder"] = ""
    save_host_config()
    return {"ok": True, "config": HOST_CONFIG}


@app.post("/api/admin/inspect-excel")
def api_admin_inspect_excel(req: InspectExcelRequest) -> Dict[str, Any]:
    raw_path = req.excel_path.strip().strip('"').strip("'")
    if not raw_path:
        raise HTTPException(status_code=400, detail="Please enter or browse a valid .xlsx file path.")
    p = Path(raw_path)
    if not p.exists() or not p.is_file():
        raise HTTPException(status_code=400, detail=f"File not found: {p}")
    if p.suffix.lower() not in (".xlsx", ".xlsm"):
        raise HTTPException(status_code=400, detail="Please select a valid .xlsx or .xlsm Excel workbook.")

    try:
        with EXCEL_WRITE_LOCK:
            wb = openpyxl.load_workbook(str(p), read_only=True)
            sheets = list(wb.sheetnames)
            active = wb.active.title if wb.active else (sheets[0] if sheets else "")
            wb.close()
        return {"ok": True, "sheet_names": sheets, "active_sheet": active}
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not read Excel workbook: {exc}")


@app.post("/api/user/login")
def api_user_login(req: UserLoginRequest) -> Dict[str, Any]:
    name = req.name.strip()
    token_id = req.token_id.strip()
    if not name or not token_id:
        raise HTTPException(status_code=400, detail="Both Name and Token ID are required.")

    try:
        data = read_excel_records(token_filter=token_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed reading Excel file: {exc}")

    if not data.get("has_token_column"):
        raise HTTPException(
            status_code=400,
            detail="The connected Excel sheet does not have a 'Token ID' column yet. Ask Admin to assign Token IDs.",
        )

    if len(data["rows"]) == 0:
        raise HTTPException(
            status_code=404,
            detail=f"No rows assigned to Token ID '{token_id}' in the connected Excel file.",
        )

    return {
        "ok": True,
        "role": "user",
        "name": name,
        "token_id": token_id,
        "assigned_count": len(data["rows"]),
    }


@app.post("/api/admin/config")
def api_admin_save_config(req: SourceConfigRequest) -> Dict[str, Any]:
    excel_clean = req.excel_path.strip().strip('"').strip("'")
    native_clean = req.native_folder.strip().strip('"').strip("'")
    text_clean = req.text_folder.strip().strip('"').strip("'")

    excel_p = Path(excel_clean)
    native_p = Path(native_clean)
    text_p = Path(text_clean)

    if not excel_clean or not excel_p.exists() or not excel_p.is_file():
        raise HTTPException(status_code=400, detail=f"Excel file does not exist: {excel_clean}")
    if excel_p.suffix.lower() not in (".xlsx", ".xlsm"):
        raise HTTPException(status_code=400, detail="Excel file must be .xlsx or .xlsm format.")
    if not native_clean or not native_p.exists() or not native_p.is_dir():
        raise HTTPException(status_code=400, detail=f"Native folder does not exist: {native_clean}")
    if not text_clean or not text_p.exists() or not text_p.is_dir():
        raise HTTPException(status_code=400, detail=f"Text folder does not exist: {text_clean}")

    sheet_clean = (req.sheet_name or "").strip()
    try:
        with EXCEL_WRITE_LOCK:
            wb = openpyxl.load_workbook(str(excel_p), read_only=True)
            if sheet_clean and sheet_clean not in wb.sheetnames:
                available = ", ".join(wb.sheetnames)
                wb.close()
                raise HTTPException(
                    status_code=400,
                    detail=f"Worksheet '{sheet_clean}' not found in workbook. Available sheets: {available}",
                )
            wb.close()
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Unable to open Excel file: {exc}")

    HOST_CONFIG["excel_path"] = str(excel_p)
    HOST_CONFIG["sheet_name"] = sheet_clean
    HOST_CONFIG["native_folder"] = str(native_p)
    HOST_CONFIG["text_folder"] = str(text_p)
    save_host_config()

    return {"ok": True, "config": HOST_CONFIG}


@app.get("/api/records")
def api_get_records(token_id: Optional[str] = Query(default=None)) -> Dict[str, Any]:
    try:
        data = read_excel_records(token_filter=token_id)
        return {"ok": True, **data}
    except FileNotFoundError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Error reading Excel: {exc}")


@app.post("/api/user/save-row")
def api_user_save_row(req: UserRowSaveRequest) -> Dict[str, Any]:
    try:
        res = update_excel_row_in_place(
            row_index=req.row_index,
            updates=req.updates,
            as_per_doc=req.as_per_doc,
            user_name=req.user_name,
            token_id=req.token_id,
        )
        return {"ok": True, **res}
    except PermissionError as exc:
        raise HTTPException(status_code=423, detail=str(exc))
    except (FileNotFoundError, ValueError) as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed saving to Excel: {exc}")


@app.post("/api/admin/update-row")
def api_admin_update_row(req: AdminRowEditRequest) -> Dict[str, Any]:
    try:
        res = admin_update_row_in_excel(
            row_index=req.row_index,
            token_id=req.token_id,
            cells=req.cells,
        )
        return {"ok": True, **res}
    except PermissionError as exc:
        raise HTTPException(status_code=423, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/admin/add-row")
def api_admin_add_row(req: AdminAddRowRequest) -> Dict[str, Any]:
    try:
        res = admin_append_row_to_excel(req.values)
        return {"ok": True, **res}
    except PermissionError as exc:
        raise HTTPException(status_code=423, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.get("/api/admin/download-excel")
def api_admin_download_excel() -> Response:
    excel_path_str = HOST_CONFIG.get("excel_path", "").strip()
    excel_path = Path(excel_path_str)
    if not excel_path_str or not excel_path.exists():
        raise HTTPException(status_code=404, detail="Connected Excel file not found.")

    with EXCEL_WRITE_LOCK:
        content = excel_path.read_bytes()

    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{excel_path.name}"'},
    )


@app.get("/api/doc/text/{doc_id:path}")
def api_get_doc_text(doc_id: str) -> Dict[str, Any]:
    text_folder = HOST_CONFIG.get("text_folder", "").strip()
    if not text_folder:
        return {
            "ok": False,
            "doc_id": doc_id,
            "filename": f"{doc_id}.txt",
            "path": "",
            "content": "[Text Folder is not connected yet. Ask Admin to select the Text Folder.]",
        }

    text_file = find_text_file_for_doc(doc_id, text_folder)
    if not text_file:
        return {
            "ok": False,
            "doc_id": doc_id,
            "filename": f"{doc_id}.txt",
            "path": str(Path(text_folder) / f"{doc_id}.txt"),
            "content": f"[No text file found matching Document ID '{doc_id}' in {text_folder}]",
        }

    try:
        raw_bytes = text_file.read_bytes()
        content = raw_bytes.decode("utf-8", errors="replace")
    except Exception as exc:
        content = f"[Error reading text file: {exc}]"

    return {
        "ok": True,
        "doc_id": doc_id,
        "filename": text_file.name,
        "path": str(text_file),
        "content": content,
    }


@app.get("/api/doc/native/{doc_id:path}")
def api_get_doc_native(
    doc_id: str,
    download: bool = Query(default=False),
    inline: bool = Query(default=False),
):
    native_folder = HOST_CONFIG.get("native_folder", "").strip()
    if not native_folder:
        raise HTTPException(status_code=400, detail="Native Folder is not connected yet.")

    native_file = find_native_file_for_doc(doc_id, native_folder)
    if not native_file:
        raise HTTPException(
            status_code=404,
            detail=f"No native document found matching '{doc_id}' in {native_folder}.",
        )

    ext = native_file.suffix.lstrip(".").lower()
    mime_map = {
        "pdf": "application/pdf",
        "txt": "text/plain; charset=utf-8",
        "csv": "text/csv; charset=utf-8",
        "json": "application/json; charset=utf-8",
        "xml": "application/xml; charset=utf-8",
        "html": "text/html; charset=utf-8",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "png": "image/png",
        "gif": "image/gif",
        "webp": "image/webp",
        "tif": "image/tiff",
        "tiff": "image/tiff",
        "bmp": "image/bmp",
        "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "xls": "application/vnd.ms-excel",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "doc": "application/msword",
    }
    media_type = mime_map.get(ext, "application/octet-stream")

    if download:
        return FileResponse(
            path=str(native_file),
            filename=native_file.name,
            media_type=media_type,
        )

    if inline:
        content = native_file.read_bytes()
        return Response(
            content=content,
            media_type=media_type,
            headers={"Content-Disposition": f'inline; filename="{native_file.name}"'},
        )

    text_like_exts = {"txt", "csv", "json", "xml", "md", "log", "html"}
    if ext in text_like_exts:
        text_preview = native_file.read_bytes().decode("utf-8", errors="replace")
        return JSONResponse({
            "ok": True,
            "doc_id": doc_id,
            "filename": native_file.name,
            "ext": ext,
            "path": str(native_file),
            "size_bytes": native_file.stat().st_size,
            "is_text_preview": True,
            "is_image": False,
            "is_pdf": False,
            "preview_text": text_preview,
            "download_url": f"/api/doc/native/{doc_id}?download=1",
        })

    is_image = ext in {"jpg", "jpeg", "png", "gif", "webp", "bmp"}
    is_pdf = ext == "pdf"
    return JSONResponse({
        "ok": True,
        "doc_id": doc_id,
        "filename": native_file.name,
        "ext": ext,
        "path": str(native_file),
        "size_bytes": native_file.stat().st_size,
        "is_text_preview": False,
        "is_image": is_image,
        "is_pdf": is_pdf,
        "stream_url": f"/api/doc/native/{doc_id}?inline=1",
        "download_url": f"/api/doc/native/{doc_id}?download=1",
    })


# ==============================================================================
# 5. COMPLETE REAL-TIME WEB FRONTEND (ADMIN VIEW + USER SPLIT WORKSPACE)
# ==============================================================================
FRONTEND_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Smart Clean UP — Real-Time Excel & Document Cleanup</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
  <style>
    body { font-family: 'Inter', sans-serif; background-color: #f8fafc; color: #0f172a; }
    .mono { font-family: 'JetBrains Mono', monospace; }
    mark.hl { background-color: #fde047; color: #0f172a; padding: 1px 3px; border-radius: 3px; font-weight: 600; }
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: #f1f5f9; }
    ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
    ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
  </style>
</head>
<body class="min-h-screen flex flex-col">

  <!-- TOP HEADER BAR -->
  <header class="bg-slate-900 text-white border-b border-slate-800 px-6 py-3 flex flex-wrap items-center justify-between gap-3 sticky top-0 z-30 shadow-sm">
    <div class="flex items-center gap-3">
      <div class="w-9 h-9 rounded-lg bg-indigo-600 flex items-center justify-center font-bold text-lg tracking-tight shadow">
        SC
      </div>
      <div>
        <div class="flex items-center gap-2">
          <h1 class="font-bold text-base tracking-tight">Smart Clean UP</h1>
          <span class="text-[11px] font-semibold uppercase tracking-wider px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
            Real-Time Live Sync
          </span>
        </div>
        <p class="text-xs text-slate-400">Local LAN Multi-User Excel &amp; Document Cleanup Server</p>
      </div>
    </div>

    <div class="flex items-center gap-3">
      <div id="connectionBadges" class="hidden md:flex items-center gap-2 text-xs"></div>
      <div id="sessionPill" class="hidden items-center gap-2 bg-slate-800 border border-slate-700 px-3 py-1.5 rounded-lg text-xs">
        <span id="sessionLabel" class="font-medium text-slate-200"></span>
        <button onclick="logout()" class="ml-2 text-rose-400 hover:text-rose-300 font-semibold">Logout</button>
      </div>
    </div>
  </header>

  <!-- TOAST NOTIFICATION -->
  <div id="toastBox" class="fixed bottom-5 right-5 z-50 hidden max-w-md rounded-xl px-4 py-3 text-sm font-medium shadow-xl border"></div>

  <!-- ====================================================================== -->
  <!-- VIEW 1: DUAL LOGIN PORTAL (ADMIN & USER)                               -->
  <!-- ====================================================================== -->
  <main id="loginView" class="flex-1 flex items-center justify-center p-6">
    <div class="w-full max-w-4xl">
      <!-- Live Server Status Banner -->
      <div class="mb-6 bg-white border border-slate-200 rounded-2xl p-4 shadow-sm flex flex-wrap items-center justify-between gap-4">
        <div>
          <div class="text-xs font-semibold uppercase tracking-wider text-slate-400">Host Server LAN URL (Share With Team)</div>
          <div class="flex items-center gap-2 mt-1">
            <code id="loginShareUrl" class="mono text-sm font-semibold text-indigo-600 bg-indigo-50 px-2.5 py-1 rounded-md">Loading...</code>
            <button onclick="copyShareUrl()" class="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium px-2.5 py-1 rounded-md transition">
              Copy URL
            </button>
          </div>
        </div>
        <div id="loginStatusSummary" class="flex flex-wrap gap-2 text-xs"></div>
      </div>

      <div class="grid md:grid-cols-2 gap-6">
        <!-- USER LOGIN CARD -->
        <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex flex-col justify-between">
          <div>
            <div class="inline-flex items-center gap-1.5 text-xs font-semibold text-indigo-700 bg-indigo-50 px-2.5 py-1 rounded-full mb-3">
              Team Member / Operator
            </div>
            <h2 class="text-xl font-bold text-slate-900">User Workspace Login</h2>
            <p class="text-sm text-slate-500 mt-1">
              Enter your name and assigned <strong>Token ID</strong> to load your real-time queue and documents.
            </p>

            <form id="userLoginForm" onsubmit="handleUserLogin(event)" class="mt-5 space-y-4">
              <div>
                <label class="block text-xs font-semibold text-slate-700 uppercase tracking-wide mb-1">Your Name</label>
                <input type="text" id="userNameInput" required placeholder="e.g., Rahul Sharma"
                  class="w-full rounded-xl border border-slate-300 px-3.5 py-2.5 text-sm focus:border-indigo-500 focus:ring-2 focus:ring-indigo-200 outline-none" />
              </div>
              <div>
                <label class="block text-xs font-semibold text-slate-700 uppercase tracking-wide mb-1">Assigned Token ID</label>
                <input type="text" id="userTokenInput" required placeholder="e.g., 5 or TOKEN-01"
                  class="w-full rounded-xl border border-slate-300 px-3.5 py-2.5 text-sm focus:border-indigo-500 focus:ring-2 focus:ring-indigo-200 outline-none" />
              </div>
              <button type="submit"
                class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-2.5 px-4 rounded-xl text-sm transition shadow-sm">
                Open Assigned Queue
              </button>
            </form>
          </div>
          <p class="text-xs text-slate-400 mt-4">
            Only rows matching your Token ID in the Admin's connected Excel sheet will appear.
          </p>
        </div>

        <!-- ADMIN LOGIN CARD -->
        <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm flex flex-col justify-between">
          <div>
            <div class="inline-flex items-center gap-1.5 text-xs font-semibold text-amber-800 bg-amber-50 border border-amber-200 px-2.5 py-1 rounded-full mb-3">
              Host Administrator
            </div>
            <h2 class="text-xl font-bold text-slate-900">Admin Host Control</h2>
            <p class="text-sm text-slate-500 mt-1">
              Log in to browse &amp; connect your real Excel workbook, Native file folder, and Text folder.
            </p>

            <form id="adminLoginForm" onsubmit="handleAdminLogin(event)" class="mt-5 space-y-4">
              <div>
                <label class="block text-xs font-semibold text-slate-700 uppercase tracking-wide mb-1">Admin User ID</label>
                <input type="text" id="adminIdInput" required placeholder="Enter Admin User ID"
                  class="w-full rounded-xl border border-slate-300 px-3.5 py-2.5 text-sm focus:border-slate-800 focus:ring-2 focus:ring-slate-200 outline-none" />
              </div>
              <div>
                <label class="block text-xs font-semibold text-slate-700 uppercase tracking-wide mb-1">Admin Password</label>
                <input type="password" id="adminPassInput" required placeholder="••••••••"
                  class="w-full rounded-xl border border-slate-300 px-3.5 py-2.5 text-sm focus:border-slate-800 focus:ring-2 focus:ring-slate-200 outline-none" />
              </div>
              <button type="submit"
                class="w-full bg-slate-900 hover:bg-slate-800 text-white font-semibold py-2.5 px-4 rounded-xl text-sm transition shadow-sm">
                Sign In as Admin
              </button>
            </form>
          </div>
          <p class="text-xs text-slate-400 mt-4">
            After login, click Browse or Import to load your real Excel sheet and document folders.
          </p>
        </div>
      </div>
    </div>
  </main>

  <!-- ====================================================================== -->
  <!-- VIEW 2: ADMIN HOST DASHBOARD                                           -->
  <!-- ====================================================================== -->
  <section id="adminView" class="hidden flex-1 p-6 max-w-[1600px] w-full mx-auto space-y-6">

    <!-- Hidden Browser File Inputs for Direct Import -->
    <input type="file" id="uploadExcelInput" accept=".xlsx,.xlsm" class="hidden" onchange="handleDirectUpload('excel', this.files)" />
    <input type="file" id="uploadNativeInput" webkitdirectory directory multiple class="hidden" onchange="handleDirectUpload('native', this.files)" />
    <input type="file" id="uploadTextInput" webkitdirectory directory multiple class="hidden" onchange="handleDirectUpload('text', this.files)" />

    <!-- Source Connection Card -->
    <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
      <div class="flex flex-wrap items-center justify-between gap-4 border-b border-slate-100 pb-4 mb-5">
        <div>
          <h2 class="text-lg font-bold text-slate-900">1. Connect or Import Real-Time Data Sources</h2>
          <p class="text-xs text-slate-500 mt-0.5">
            Click <strong>Browse</strong> to pick directly from your PC Explorer (live path), or <strong>Import Folder/File</strong> to upload into the workspace.
          </p>
        </div>
        <div class="flex flex-wrap items-center gap-2">
          <span class="text-xs text-slate-500">Team LAN Link:</span>
          <code id="adminShareUrl" class="mono text-xs font-semibold bg-indigo-50 text-indigo-700 px-2.5 py-1 rounded"></code>
          <button onclick="copyShareUrl()" class="text-xs bg-slate-100 hover:bg-slate-200 px-2.5 py-1 rounded font-medium">Copy</button>
          <button onclick="disconnectSources()" class="text-xs bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 px-2.5 py-1 rounded font-semibold">
            Clear / Reset Paths
          </button>
        </div>
      </div>

      <form onsubmit="saveAdminConfig(event)" class="grid md:grid-cols-2 lg:grid-cols-4 gap-4">
        <!-- EXCEL FILE -->
        <div class="lg:col-span-2">
          <div class="flex items-center justify-between mb-1">
            <label class="text-xs font-semibold text-slate-700">1. Excel Workbook (.xlsx)</label>
            <span id="statusTagExcel" class="text-[11px] font-semibold text-amber-600">Not Connected</span>
          </div>
          <div class="flex gap-1.5">
            <input type="text" id="cfgExcelPath" required placeholder="Click Browse to select your .xlsx file..."
              class="mono flex-1 rounded-xl border border-slate-300 px-3 py-2 text-xs focus:border-indigo-500 outline-none" />
            <button type="button" onclick="browseSource('excel')"
              class="bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold px-3 py-2 rounded-xl whitespace-nowrap">
              Browse .xlsx
            </button>
            <button type="button" onclick="document.getElementById('uploadExcelInput').click()"
              class="bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-semibold px-3 py-2 rounded-xl whitespace-nowrap">
              Import
            </button>
          </div>
        </div>

        <!-- WORKSHEET TAB -->
        <div>
          <label class="block text-xs font-semibold text-slate-700 mb-1">Worksheet / Tab Name</label>
          <div class="flex gap-1.5">
            <input type="text" id="cfgSheetName" list="sheetNamesList" placeholder="Active sheet"
              class="flex-1 rounded-xl border border-slate-300 px-3 py-2 text-xs focus:border-indigo-500 outline-none" />
            <button type="button" onclick="inspectExcelSheets()"
              class="bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-semibold px-2.5 py-2 rounded-xl whitespace-nowrap">
              Tabs
            </button>
          </div>
          <datalist id="sheetNamesList"></datalist>
        </div>

        <!-- SAVE BUTTON -->
        <div class="flex items-end">
          <button type="submit"
            class="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-2 px-4 rounded-xl text-xs transition shadow-sm">
            &#10003; Apply &amp; Load Live Data
          </button>
        </div>

        <!-- NATIVE FILES FOLDER -->
        <div class="lg:col-span-2">
          <div class="flex items-center justify-between mb-1">
            <label class="text-xs font-semibold text-slate-700">2. Native Files Folder (PDF, JPG, PNG, DOCX, XLSX, TIF...)</label>
            <span id="statusTagNative" class="text-[11px] font-semibold text-amber-600">Not Connected</span>
          </div>
          <div class="flex gap-1.5">
            <input type="text" id="cfgNativeFolder" required placeholder="Click Browse to select Native files folder..."
              class="mono flex-1 rounded-xl border border-slate-300 px-3 py-2 text-xs focus:border-indigo-500 outline-none" />
            <button type="button" onclick="browseSource('native_folder')"
              class="bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold px-3 py-2 rounded-xl whitespace-nowrap">
              Browse Folder
            </button>
            <button type="button" onclick="document.getElementById('uploadNativeInput').click()"
              class="bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-semibold px-3 py-2 rounded-xl whitespace-nowrap">
              Import Folder
            </button>
          </div>
        </div>

        <!-- TEXT FILES FOLDER -->
        <div class="lg:col-span-2">
          <div class="flex items-center justify-between mb-1">
            <label class="text-xs font-semibold text-slate-700">3. Text Files Folder (.txt / OCR Output)</label>
            <span id="statusTagText" class="text-[11px] font-semibold text-amber-600">Not Connected</span>
          </div>
          <div class="flex gap-1.5">
            <input type="text" id="cfgTextFolder" required placeholder="Click Browse to select Text files folder..."
              class="mono flex-1 rounded-xl border border-slate-300 px-3 py-2 text-xs focus:border-indigo-500 outline-none" />
            <button type="button" onclick="browseSource('text_folder')"
              class="bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold px-3 py-2 rounded-xl whitespace-nowrap">
              Browse Folder
            </button>
            <button type="button" onclick="document.getElementById('uploadTextInput').click()"
              class="bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-semibold px-3 py-2 rounded-xl whitespace-nowrap">
              Import Folder
            </button>
          </div>
        </div>
      </form>
    </div>

    <!-- Live Master Excel Table -->
    <div class="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm">
      <div class="flex flex-wrap items-center justify-between gap-4 mb-4">
        <div>
          <h2 class="text-lg font-bold text-slate-900">2. Live Connected Excel Workbook &amp; Token Assignment</h2>
          <p class="text-xs text-slate-500" id="adminTableMeta">Select your Excel workbook above to load real rows.</p>
        </div>
        <div class="flex flex-wrap items-center gap-2">
          <input type="text" id="adminSearchInput" oninput="renderAdminTable()" placeholder="Filter rows..."
            class="rounded-xl border border-slate-300 px-3 py-1.5 text-xs focus:border-indigo-500 outline-none" />
          <button onclick="openAddRowModal()"
            class="bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold px-3.5 py-2 rounded-xl">
            + Add Row
          </button>
          <button onclick="loadAdminRecords()"
            class="bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-semibold px-3.5 py-2 rounded-xl">
            Refresh Live Excel
          </button>
          <a href="/api/admin/download-excel"
            class="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-3.5 py-2 rounded-xl">
            Download Updated .xlsx
          </a>
        </div>
      </div>

      <div class="overflow-x-auto border border-slate-200 rounded-xl max-h-[62vh]">
        <table class="w-full text-left border-collapse text-xs">
          <thead id="adminTableHead" class="bg-slate-100 text-slate-700 sticky top-0 z-10"></thead>
          <tbody id="adminTableBody" class="divide-y divide-slate-200">
            <tr>
              <td class="py-10 text-center text-slate-400">No Excel file connected yet. Click <strong>Browse .xlsx</strong> above to load your real data.</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </section>

  <!-- ADD ROW MODAL (ADMIN) -->
  <div id="addRowModal" class="hidden fixed inset-0 bg-slate-900/50 z-50 flex items-center justify-center p-4">
    <div class="bg-white rounded-2xl max-w-lg w-full p-6 shadow-2xl border border-slate-200">
      <div class="flex items-center justify-between border-b border-slate-100 pb-3 mb-4">
        <h3 class="font-bold text-base text-slate-900">Append New Row to Live Excel</h3>
        <button onclick="closeAddRowModal()" class="text-slate-400 hover:text-slate-700 text-lg font-bold">&times;</button>
      </div>
      <form id="addRowForm" onsubmit="submitNewRow(event)" class="space-y-3 max-h-[65vh] overflow-y-auto pr-1"></form>
      <div class="flex justify-end gap-2 mt-5 pt-3 border-t border-slate-100">
        <button type="button" onclick="closeAddRowModal()" class="px-4 py-2 rounded-xl text-xs font-semibold bg-slate-100 text-slate-700">Cancel</button>
        <button type="submit" form="addRowForm" class="px-4 py-2 rounded-xl text-xs font-semibold bg-indigo-600 text-white">Append to Excel</button>
      </div>
    </div>
  </div>

  <!-- ====================================================================== -->
  <!-- VIEW 3: USER SPLIT-SCREEN CLEANUP WORKSPACE                           -->
  <!-- ====================================================================== -->
  <section id="userView" class="hidden flex-1 grid grid-cols-1 lg:grid-cols-12 gap-4 p-4 max-w-[1800px] w-full mx-auto">

    <!-- LEFT PANE (5 COLS): ASSIGNED ROWS + NON-DESTRUCTIVE EDITOR -->
    <div class="lg:col-span-5 flex flex-col gap-4">
      <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm flex flex-col">
        <div class="flex items-center justify-between gap-2 mb-3">
          <div>
            <h2 class="font-bold text-sm text-slate-900">Your Assigned Records</h2>
            <p class="text-xs text-slate-500" id="userQueueStats">Loading assigned rows...</p>
          </div>
          <div class="flex items-center gap-2">
            <input type="text" id="userRowSearch" oninput="renderUserQueue()" placeholder="Search ID / name..."
              class="rounded-lg border border-slate-300 px-2.5 py-1 text-xs w-36 focus:border-indigo-500 outline-none" />
            <button onclick="loadUserRecords()" class="text-xs bg-slate-100 hover:bg-slate-200 px-2.5 py-1 rounded-lg font-medium">
              Sync
            </button>
          </div>
        </div>

        <div class="overflow-auto border border-slate-200 rounded-xl max-h-60">
          <table class="w-full text-left border-collapse text-xs">
            <thead class="bg-slate-100 text-slate-700 sticky top-0">
              <tr>
                <th class="py-2 px-3 font-semibold">Row</th>
                <th class="py-2 px-3 font-semibold">Document ID</th>
                <th class="py-2 px-3 font-semibold">Files</th>
                <th class="py-2 px-3 font-semibold">Status</th>
              </tr>
            </thead>
            <tbody id="userQueueBody" class="divide-y divide-slate-200"></tbody>
          </table>
        </div>
      </div>

      <!-- Non-Destructive Row Editor Card -->
      <div class="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm flex-1 flex flex-col">
        <div class="flex items-center justify-between border-b border-slate-100 pb-3 mb-3">
          <div>
            <span class="text-[11px] font-semibold uppercase tracking-wider text-indigo-600">Non-Destructive Update</span>
            <h3 class="font-bold text-base text-slate-900" id="editorDocTitle">Select a row above</h3>
          </div>
          <div class="flex items-center gap-1.5">
            <button onclick="selectPrevRow()" class="px-2.5 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs font-semibold">&larr; Prev</button>
            <button onclick="selectNextRow()" class="px-2.5 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs font-semibold">Next &rarr;</button>
          </div>
        </div>

        <div id="editorFieldsContainer" class="space-y-3 overflow-y-auto flex-1 pr-1 max-h-[46vh]">
          <p class="text-xs text-slate-400">Click any record in your queue to inspect original Excel values and enter updates.</p>
        </div>

        <div class="pt-3 mt-3 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2">
          <button type="button" onclick="saveCurrentRow(true)"
            class="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-4 py-2.5 rounded-xl transition shadow-sm">
            &#10003; Mark "As Per Doc"
          </button>
          <button type="button" onclick="saveCurrentRow(false)"
            class="bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold px-5 py-2.5 rounded-xl transition shadow-sm">
            Save Updates to Live Excel
          </button>
        </div>
      </div>
    </div>

    <!-- RIGHT PANE (7 COLS): LIVE TEXT & NATIVE DOCUMENT VIEWER + SEARCH -->
    <div class="lg:col-span-7 bg-white border border-slate-200 rounded-2xl p-4 shadow-sm flex flex-col min-h-[78vh]">
      <div class="flex flex-wrap items-center justify-between gap-3 border-b border-slate-100 pb-3 mb-3">
        <div class="flex items-center gap-2">
          <button id="tabTextBtn" onclick="switchDocTab('text')"
            class="px-3.5 py-1.5 rounded-xl text-xs font-semibold bg-indigo-600 text-white transition">
            Text File (.txt)
          </button>
          <button id="tabNativeBtn" onclick="switchDocTab('native')"
            class="px-3.5 py-1.5 rounded-xl text-xs font-semibold bg-slate-100 text-slate-700 hover:bg-slate-200 transition">
            Native File
          </button>
          <span id="docFileBadge" class="mono text-xs text-slate-500 ml-1"></span>
        </div>

        <div class="flex items-center gap-2">
          <div class="relative">
            <input type="text" id="docKeywordSearch" oninput="applyDocSearch()" placeholder="Search inside document..."
              class="rounded-xl border border-slate-300 pl-3 pr-16 py-1.5 text-xs w-56 focus:border-indigo-500 outline-none" />
            <span id="searchMatchCount" class="absolute right-2.5 top-1.5 text-[11px] font-semibold text-slate-400"></span>
          </div>
          <a id="downloadNativeBtn" href="#" target="_blank"
            class="hidden bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold px-3 py-1.5 rounded-xl">
            Download Native
          </a>
        </div>
      </div>

      <!-- TEXT DOCUMENT PANE -->
      <div id="textViewerPane" class="flex-1 flex flex-col">
        <pre id="textViewerContent"
          class="mono text-xs leading-relaxed bg-slate-50 border border-slate-200 rounded-xl p-4 flex-1 overflow-auto whitespace-pre-wrap text-slate-800 min-h-[64vh]">Select a record to view its linked Text file.</pre>
      </div>

      <!-- NATIVE DOCUMENT PANE -->
      <div id="nativeViewerPane" class="hidden flex-1 flex flex-col">
        <div id="nativeViewerBody" class="flex-1 border border-slate-200 rounded-xl bg-slate-50 overflow-auto flex items-center justify-center min-h-[64vh]">
          <p class="text-xs text-slate-400">Select a record to preview or download its Native document.</p>
        </div>
      </div>
    </div>
  </section>

  <!-- ====================================================================== -->
  <!-- JAVASCRIPT APPLICATION CONTROLLER                                      -->
  <!-- ====================================================================== -->
  <script>
    const STATE = {
      role: null,
      adminId: '',
      userName: '',
      tokenId: '',
      hostInfo: null,
      adminData: { headers: [], rows: [], sheet_names: [], active_sheet: '' },
      userData: { headers: [], rows: [], active_sheet: '' },
      selectedUserIdx: 0,
      activeDocTab: 'text',
      currentTextRaw: '',
      currentNativeMeta: null,
    };

    function showToast(msg, type = 'info') {
      const box = document.getElementById('toastBox');
      box.textContent = msg;
      box.className = 'fixed bottom-5 right-5 z-50 max-w-md rounded-xl px-4 py-3 text-sm font-medium shadow-xl border ';
      if (type === 'error') {
        box.className += 'bg-rose-50 text-rose-900 border-rose-300';
      } else if (type === 'success') {
        box.className += 'bg-emerald-50 text-emerald-900 border-emerald-300';
      } else {
        box.className += 'bg-slate-900 text-white border-slate-700';
      }
      box.classList.remove('hidden');
      clearTimeout(window._toastTimer);
      window._toastTimer = setTimeout(() => box.classList.add('hidden'), 4200);
    }

    function escapeHtml(str) {
      return String(str ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
    }

    async function fetchHostInfo() {
      try {
        const res = await fetch('/api/host-info');
        const info = await res.json();
        STATE.hostInfo = info;

        document.getElementById('loginShareUrl').textContent = info.share_url;
        document.getElementById('adminShareUrl').textContent = info.share_url;

        document.getElementById('cfgExcelPath').value = info.config.excel_path || '';
        document.getElementById('cfgSheetName').value = info.config.sheet_name || '';
        document.getElementById('cfgNativeFolder').value = info.config.native_folder || '';
        document.getElementById('cfgTextFolder').value = info.config.text_folder || '';

        const st = info.status;

        // Update inline status tags in Admin panel
        const tagExcel = document.getElementById('statusTagExcel');
        const tagNative = document.getElementById('statusTagNative');
        const tagText = document.getElementById('statusTagText');
        if (tagExcel) {
          tagExcel.textContent = st.excel_exists ? `Connected (${st.excel_name})` : 'Not Connected';
          tagExcel.className = `text-[11px] font-semibold ${st.excel_exists ? 'text-emerald-600' : 'text-amber-600'}`;
        }
        if (tagNative) {
          tagNative.textContent = st.native_exists ? `Connected (${st.native_file_count} files)` : 'Not Connected';
          tagNative.className = `text-[11px] font-semibold ${st.native_exists ? 'text-emerald-600' : 'text-amber-600'}`;
        }
        if (tagText) {
          tagText.textContent = st.text_exists ? `Connected (${st.text_file_count} files)` : 'Not Connected';
          tagText.className = `text-[11px] font-semibold ${st.text_exists ? 'text-emerald-600' : 'text-amber-600'}`;
        }

        const badgeHtml = `
          <span class="px-2.5 py-1 rounded-full font-semibold ${st.excel_exists ? 'bg-emerald-50 text-emerald-700 border border-emerald-200' : 'bg-amber-50 text-amber-700 border border-amber-200'}">
            Excel: ${st.excel_exists ? escapeHtml(st.excel_name) : 'Not Connected'}
          </span>
          <span class="px-2.5 py-1 rounded-full font-semibold ${st.native_exists ? 'bg-emerald-50 text-emerald-700 border border-emerald-200' : 'bg-amber-50 text-amber-700 border border-amber-200'}">
            Native: ${st.native_exists ? st.native_file_count + ' files' : 'Not Connected'}
          </span>
          <span class="px-2.5 py-1 rounded-full font-semibold ${st.text_exists ? 'bg-emerald-50 text-emerald-700 border border-emerald-200' : 'bg-amber-50 text-amber-700 border border-amber-200'}">
            Text: ${st.text_exists ? st.text_file_count + ' files' : 'Not Connected'}
          </span>
        `;
        document.getElementById('loginStatusSummary').innerHTML = badgeHtml;
        const hdrBadges = document.getElementById('connectionBadges');
        hdrBadges.innerHTML = badgeHtml;
        hdrBadges.classList.remove('hidden');
      } catch (e) {
        console.error(e);
      }
    }

    function copyShareUrl() {
      if (!STATE.hostInfo) return;
      navigator.clipboard.writeText(STATE.hostInfo.share_url);
      showToast('Copied LAN URL: ' + STATE.hostInfo.share_url, 'success');
    }

    function switchView(viewName) {
      document.getElementById('loginView').classList.add('hidden');
      document.getElementById('adminView').classList.add('hidden');
      document.getElementById('userView').classList.add('hidden');
      document.getElementById(viewName).classList.remove('hidden');

      const pill = document.getElementById('sessionPill');
      if (viewName === 'loginView') {
        pill.classList.add('hidden');
        pill.classList.remove('flex');
      } else {
        pill.classList.remove('hidden');
        pill.classList.add('flex');
      }
    }

    function logout() {
      STATE.role = null;
      switchView('loginView');
      fetchHostInfo();
    }

    // ------------------------------------------------------------------------
    // ADMIN FLOW (BROWSE, IMPORT, LOAD LIVE EXCEL)
    // ------------------------------------------------------------------------
    async function handleAdminLogin(e) {
      e.preventDefault();
      const user_id = document.getElementById('adminIdInput').value;
      const password = document.getElementById('adminPassInput').value;
      const res = await fetch('/api/admin/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ user_id, password })
      });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || 'Admin login failed', 'error');
        return;
      }
      STATE.role = 'admin';
      STATE.adminId = data.user_id;
      document.getElementById('sessionLabel').textContent = `Admin: ${data.user_id}`;
      switchView('adminView');
      await fetchHostInfo();

      if (STATE.hostInfo?.status?.excel_exists) {
        await loadAdminRecords();
        showToast('Logged in — Live Excel & Folders loaded', 'success');
      } else {
        showToast('Logged in! Click "Browse" or "Import" to connect your real Excel & document folders.', 'info');
      }
    }

    async function browseSource(mode) {
      showToast('Opening file/folder picker on Host PC...', 'info');
      try {
        const res = await fetch('/api/admin/browse', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ mode })
        });
        const data = await res.json();
        if (!res.ok) {
          showToast(data.detail || 'Browse failed', 'error');
          return;
        }
        if (data.cancelled) {
          return;
        }

        if (mode === 'excel') {
          document.getElementById('cfgExcelPath').value = data.path;
          document.getElementById('cfgSheetName').value = data.active_sheet || '';
          const dl = document.getElementById('sheetNamesList');
          dl.innerHTML = (data.sheet_names || []).map(s => `<option value="${escapeHtml(s)}">`).join('');
          showToast('Excel workbook connected & loaded!', 'success');
        } else if (mode === 'native_folder') {
          document.getElementById('cfgNativeFolder').value = data.path;
          showToast('Native folder connected!', 'success');
        } else if (mode === 'text_folder') {
          document.getElementById('cfgTextFolder').value = data.path;
          showToast('Text folder connected!', 'success');
        }

        await fetchHostInfo();
        if (STATE.hostInfo?.status?.excel_exists) {
          await loadAdminRecords();
        }
      } catch (err) {
        showToast('Error opening native picker: ' + err, 'error');
      }
    }

    async function handleDirectUpload(target, fileList) {
      if (!fileList || fileList.length === 0) return;
      showToast(`Importing ${fileList.length} file(s)...`, 'info');

      const fd = new FormData();
      fd.append('target', target);
      for (const f of fileList) {
        fd.append('files', f);
      }

      try {
        const res = await fetch('/api/admin/import-upload', {
          method: 'POST',
          body: fd
        });
        const data = await res.json();
        if (!res.ok) {
          showToast(data.detail || 'Import failed', 'error');
          return;
        }
        showToast(`Imported ${data.count} file(s) successfully!`, 'success');
        await fetchHostInfo();
        if (STATE.hostInfo?.status?.excel_exists) {
          await loadAdminRecords();
        }
      } catch (err) {
        showToast('Upload error: ' + err, 'error');
      }
    }

    async function disconnectSources() {
      await fetch('/api/admin/disconnect', { method: 'POST' });
      STATE.adminData = { headers: [], rows: [], sheet_names: [], active_sheet: '' };
      document.getElementById('adminTableHead').innerHTML = '';
      document.getElementById('adminTableBody').innerHTML =
        `<tr><td class="py-10 text-center text-slate-400">Sources cleared. Select your real Excel file and folders above.</td></tr>`;
      document.getElementById('adminTableMeta').textContent = 'No Excel workbook connected.';
      await fetchHostInfo();
      showToast('All paths cleared. Ready to connect your real dataset.', 'info');
    }

    async function inspectExcelSheets() {
      const excel_path = document.getElementById('cfgExcelPath').value;
      const res = await fetch('/api/admin/inspect-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ excel_path })
      });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || 'Unable to inspect Excel file', 'error');
        return;
      }
      const dl = document.getElementById('sheetNamesList');
      dl.innerHTML = data.sheet_names.map(s => `<option value="${escapeHtml(s)}">`).join('');
      if (!document.getElementById('cfgSheetName').value && data.active_sheet) {
        document.getElementById('cfgSheetName').value = data.active_sheet;
      }
      showToast(`Found ${data.sheet_names.length} worksheet(s): ${data.sheet_names.join(', ')}`, 'success');
    }

    async function saveAdminConfig(e) {
      e.preventDefault();
      const payload = {
        excel_path: document.getElementById('cfgExcelPath').value,
        sheet_name: document.getElementById('cfgSheetName').value,
        native_folder: document.getElementById('cfgNativeFolder').value,
        text_folder: document.getElementById('cfgTextFolder').value,
      };
      const res = await fetch('/api/admin/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || 'Failed saving configuration', 'error');
        return;
      }
      showToast('Live Excel & Document folders connected!', 'success');
      await fetchHostInfo();
      await loadAdminRecords();
    }

    async function loadAdminRecords() {
      const res = await fetch('/api/records');
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || 'Could not read Excel rows', 'error');
        return;
      }
      STATE.adminData = data;
      document.getElementById('adminTableMeta').textContent =
        `Sheet: "${data.active_sheet}" • Total Rows: ${data.rows.length} • Live synced with disk`;
      renderAdminTable();
    }

    function renderAdminTable() {
      const { headers, rows } = STATE.adminData;
      const query = (document.getElementById('adminSearchInput').value || '').toLowerCase();

      const headEl = document.getElementById('adminTableHead');
      headEl.innerHTML = `<tr>
        <th class="py-2.5 px-3 border-b border-slate-200">Row</th>
        <th class="py-2.5 px-3 border-b border-slate-200">Assign Token ID</th>
        <th class="py-2.5 px-3 border-b border-slate-200">Matched Docs</th>
        ${headers.map(h => `<th class="py-2.5 px-3 border-b border-slate-200 whitespace-nowrap">${escapeHtml(h)}</th>`).join('')}
      </tr>`;

      const filtered = rows.filter(r => {
        if (!query) return true;
        return JSON.stringify(r).toLowerCase().includes(query);
      });

      const bodyEl = document.getElementById('adminTableBody');
      if (filtered.length === 0) {
        bodyEl.innerHTML = `<tr><td colspan="${headers.length + 3}" class="py-8 text-center text-slate-400">No rows found in connected Excel sheet.</td></tr>`;
        return;
      }

      bodyEl.innerHTML = filtered.map(r => {
        const docBadges = `
          <span class="px-1.5 py-0.5 rounded text-[10px] font-semibold ${r.has_text ? 'bg-emerald-100 text-emerald-800' : 'bg-rose-100 text-rose-700'}">TXT</span>
          <span class="px-1.5 py-0.5 rounded text-[10px] font-semibold ${r.has_native ? 'bg-indigo-100 text-indigo-800' : 'bg-rose-100 text-rose-700'}">${r.native_ext ? r.native_ext.toUpperCase() : 'NAT'}</span>
        `;
        return `<tr class="hover:bg-slate-50">
          <td class="py-2 px-3 font-mono text-slate-500">${r.row_index}</td>
          <td class="py-2 px-3">
            <div class="flex items-center gap-1">
              <input type="text" value="${escapeHtml(r.token_id)}" id="tok_${r.row_index}"
                class="w-20 rounded border border-slate-300 px-2 py-1 text-xs focus:border-indigo-500 outline-none" />
              <button onclick="updateRowToken(${r.row_index})"
                class="bg-slate-800 hover:bg-slate-900 text-white px-2 py-1 rounded text-[11px] font-semibold">Set</button>
            </div>
          </td>
          <td class="py-2 px-3 whitespace-nowrap">${docBadges}</td>
          ${headers.map(h => `<td class="py-2 px-3 whitespace-nowrap max-w-xs truncate" title="${escapeHtml(r.cells[h])}">${escapeHtml(r.cells[h])}</td>`).join('')}
        </tr>`;
      }).join('');
    }

    async function updateRowToken(rowIndex) {
      const tokenVal = document.getElementById(`tok_${rowIndex}`).value;
      const res = await fetch('/api/admin/update-row', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ row_index: rowIndex, token_id: tokenVal })
      });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || 'Failed updating Token ID', 'error');
        return;
      }
      showToast(`Row ${rowIndex} Token ID updated to "${tokenVal}" in live Excel`, 'success');
      await loadAdminRecords();
    }

    function openAddRowModal() {
      const headers = STATE.adminData.headers.length
        ? STATE.adminData.headers
        : ['Document ID', 'Token ID'];
      const form = document.getElementById('addRowForm');
      form.innerHTML = headers.map(h => `
        <div>
          <label class="block text-xs font-semibold text-slate-700 mb-1">${escapeHtml(h)}</label>
          <input type="text" name="${escapeHtml(h)}"
            class="w-full rounded-xl border border-slate-300 px-3 py-2 text-xs focus:border-indigo-500 outline-none" />
        </div>
      `).join('');
      document.getElementById('addRowModal').classList.remove('hidden');
    }

    function closeAddRowModal() {
      document.getElementById('addRowModal').classList.add('hidden');
    }

    async function submitNewRow(e) {
      e.preventDefault();
      const fd = new FormData(e.target);
      const values = {};
      for (const [k, v] of fd.entries()) {
        values[k] = v;
      }
      const res = await fetch('/api/admin/add-row', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ values })
      });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || 'Failed appending row', 'error');
        return;
      }
      closeAddRowModal();
      showToast(`Appended new row #${data.row_index} to live Excel`, 'success');
      await loadAdminRecords();
    }

    // ------------------------------------------------------------------------
    // USER FLOW
    // ------------------------------------------------------------------------
    async function handleUserLogin(e) {
      e.preventDefault();
      const name = document.getElementById('userNameInput').value;
      const token_id = document.getElementById('userTokenInput').value;
      const res = await fetch('/api/user/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, token_id })
      });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || 'User login failed', 'error');
        return;
      }
      STATE.role = 'user';
      STATE.userName = data.name;
      STATE.tokenId = data.token_id;
      STATE.selectedUserIdx = 0;

      document.getElementById('sessionLabel').textContent = `${data.name} • Token: ${data.token_id}`;
      switchView('userView');
      await loadUserRecords();
      showToast(`Loaded ${data.assigned_count} assigned record(s)`, 'success');
    }

    async function loadUserRecords() {
      const res = await fetch(`/api/records?token_id=${encodeURIComponent(STATE.tokenId)}`);
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || 'Error loading assigned records', 'error');
        return;
      }
      STATE.userData = data;
      if (STATE.selectedUserIdx >= data.rows.length) {
        STATE.selectedUserIdx = 0;
      }
      renderUserQueue();
      if (data.rows.length > 0) {
        selectUserRow(STATE.selectedUserIdx);
      }
    }

    function renderUserQueue() {
      const rows = STATE.userData.rows || [];
      const q = (document.getElementById('userRowSearch').value || '').toLowerCase();

      const completedCount = rows.filter(r => {
        const rem = (r.cells['Remark / Action'] || '').trim();
        return rem.length > 0;
      }).length;
      document.getElementById('userQueueStats').textContent =
        `Token "${STATE.tokenId}" • ${completedCount} of ${rows.length} completed`;

      const tbody = document.getElementById('userQueueBody');
      tbody.innerHTML = rows.map((r, idx) => {
        if (q && !JSON.stringify(r).toLowerCase().includes(q)) return '';
        const active = idx === STATE.selectedUserIdx;
        const remark = (r.cells['Remark / Action'] || '').trim();
        let statusBadge = `<span class="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-amber-100 text-amber-800">Pending</span>`;
        if (remark === 'As Per Doc') {
          statusBadge = `<span class="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-emerald-100 text-emerald-800">As Per Doc</span>`;
        } else if (remark) {
          statusBadge = `<span class="px-2 py-0.5 rounded-full text-[10px] font-semibold bg-indigo-100 text-indigo-800">${escapeHtml(remark)}</span>`;
        }

        return `<tr onclick="selectUserRow(${idx})"
          class="cursor-pointer transition ${active ? 'bg-indigo-50/90 font-semibold' : 'hover:bg-slate-50'}">
          <td class="py-2 px-3 font-mono text-slate-500">#${r.row_index}</td>
          <td class="py-2 px-3 text-slate-900">${escapeHtml(r.doc_id || '(No ID)')}</td>
          <td class="py-2 px-3 whitespace-nowrap">
            <span class="text-[10px] font-bold ${r.has_text ? 'text-emerald-600' : 'text-slate-300'}">TXT</span> /
            <span class="text-[10px] font-bold ${r.has_native ? 'text-indigo-600' : 'text-slate-300'}">${r.native_ext ? r.native_ext.toUpperCase() : 'NAT'}</span>
          </td>
          <td class="py-2 px-3">${statusBadge}</td>
        </tr>`;
      }).join('');
    }

    function selectPrevRow() {
      if (STATE.selectedUserIdx > 0) selectUserRow(STATE.selectedUserIdx - 1);
    }

    function selectNextRow() {
      if (STATE.selectedUserIdx < STATE.userData.rows.length - 1) selectUserRow(STATE.selectedUserIdx + 1);
    }

    async function selectUserRow(idx) {
      STATE.selectedUserIdx = idx;
      renderUserQueue();
      const row = STATE.userData.rows[idx];
      if (!row) return;

      document.getElementById('editorDocTitle').textContent = `Row #${row.row_index} — ${row.doc_id}`;

      const systemCols = new Set([
        'tokenid', 'token', 'assigntoken', 'username', 'updatedby', 'lastupdated', 'remarkaction'
      ]);
      const baseHeaders = STATE.userData.headers.filter(h => {
        const norm = h.toLowerCase().replace(/[^a-z0-9]/g, '');
        return !h.toLowerCase().startsWith('updated ') && !systemCols.has(norm);
      });

      const container = document.getElementById('editorFieldsContainer');
      container.innerHTML = baseHeaders.map(h => {
        const origVal = row.cells[h] ?? '';
        const updatedHeader = `Updated ${h}`;
        const existingUpdatedVal = row.cells[updatedHeader] ?? '';
        const isDocId = ['documentid', 'docid', 'id'].includes(h.toLowerCase().replace(/[^a-z0-9]/g, ''));
        return `
          <div class="p-3 rounded-xl border border-slate-200 bg-slate-50/60">
            <div class="flex items-center justify-between text-xs mb-1">
              <span class="font-semibold text-slate-800">${escapeHtml(h)}</span>
              <span class="mono text-[11px] text-slate-500">Original: <strong>${escapeHtml(origVal || '—')}</strong></span>
            </div>
            ${isDocId ? `<div class="text-[11px] text-slate-400">Document Key (Locked)</div>` : `
              <input type="text" data-field="${escapeHtml(h)}" value="${escapeHtml(existingUpdatedVal)}"
                placeholder="Enter updated ${escapeHtml(h)} (writes to '${escapeHtml(updatedHeader)}')"
                class="user-edit-input w-full rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs focus:border-indigo-500 outline-none" />
            `}
          </div>
        `;
      }).join('');

      await loadDocumentPreviews(row.doc_id);
    }

    async function saveCurrentRow(asPerDoc) {
      const row = STATE.userData.rows[STATE.selectedUserIdx];
      if (!row) return;

      const inputs = document.querySelectorAll('.user-edit-input');
      const updates = {};
      inputs.forEach(inp => {
        const field = inp.getAttribute('data-field');
        updates[field] = inp.value;
      });

      const res = await fetch('/api/user/save-row', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          row_index: row.row_index,
          user_name: STATE.userName,
          token_id: STATE.tokenId,
          updates,
          as_per_doc: asPerDoc
        })
      });
      const data = await res.json();
      if (!res.ok) {
        showToast(data.detail || 'Save failed', 'error');
        return;
      }
      showToast(
        asPerDoc
          ? `Row #${row.row_index} marked "As Per Doc" in live Excel!`
          : `Row #${row.row_index} updated non-destructively in live Excel!`,
        'success'
      );
      await loadUserRecords();
    }

    // ------------------------------------------------------------------------
    // DOCUMENT VIEWER & LIVE KEYWORD SEARCH
    // ------------------------------------------------------------------------
    function switchDocTab(tab) {
      STATE.activeDocTab = tab;
      const textBtn = document.getElementById('tabTextBtn');
      const nativeBtn = document.getElementById('tabNativeBtn');
      const textPane = document.getElementById('textViewerPane');
      const nativePane = document.getElementById('nativeViewerPane');

      if (tab === 'text') {
        textBtn.className = 'px-3.5 py-1.5 rounded-xl text-xs font-semibold bg-indigo-600 text-white transition';
        nativeBtn.className = 'px-3.5 py-1.5 rounded-xl text-xs font-semibold bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
        textPane.classList.remove('hidden');
        nativePane.classList.add('hidden');
      } else {
        nativeBtn.className = 'px-3.5 py-1.5 rounded-xl text-xs font-semibold bg-indigo-600 text-white transition';
        textBtn.className = 'px-3.5 py-1.5 rounded-xl text-xs font-semibold bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
        nativePane.classList.remove('hidden');
        textPane.classList.add('hidden');
      }
      updateDocBadgeAndDownload();
      applyDocSearch();
    }

    function updateDocBadgeAndDownload() {
      const badge = document.getElementById('docFileBadge');
      const dlBtn = document.getElementById('downloadNativeBtn');
      if (STATE.currentNativeMeta && STATE.currentNativeMeta.ok) {
        dlBtn.href = STATE.currentNativeMeta.download_url;
        dlBtn.classList.remove('hidden');
      } else {
        dlBtn.classList.add('hidden');
      }

      if (STATE.activeDocTab === 'native' && STATE.currentNativeMeta?.ok) {
        badge.textContent = STATE.currentNativeMeta.filename;
      }
    }

    async function loadDocumentPreviews(docId) {
      if (!docId) return;

      try {
        const tRes = await fetch(`/api/doc/text/${encodeURIComponent(docId)}`);
        const tData = await tRes.json();
        STATE.currentTextRaw = tData.content || '';
        if (STATE.activeDocTab === 'text') {
          document.getElementById('docFileBadge').textContent = tData.filename || '';
        }
      } catch (e) {
        STATE.currentTextRaw = '[Error fetching text document]';
      }

      const nativeBody = document.getElementById('nativeViewerBody');
      try {
        const nRes = await fetch(`/api/doc/native/${encodeURIComponent(docId)}`);
        if (!nRes.ok) {
          const err = await nRes.json();
          STATE.currentNativeMeta = null;
          nativeBody.innerHTML = `<div class="p-6 text-center text-xs text-rose-600">${escapeHtml(err.detail || 'Native file not found')}</div>`;
        } else {
          const nData = await nRes.json();
          STATE.currentNativeMeta = nData;
          if (nData.is_text_preview) {
            nativeBody.innerHTML = `<pre id="nativeTextPre" class="mono text-xs p-4 w-full h-full overflow-auto whitespace-pre-wrap text-slate-800">${escapeHtml(nData.preview_text)}</pre>`;
          } else if (nData.is_image) {
            nativeBody.innerHTML = `<img src="${nData.stream_url}" alt="${escapeHtml(nData.filename)}" class="max-h-[68vh] max-w-full object-contain rounded-lg shadow" />`;
          } else if (nData.is_pdf) {
            nativeBody.innerHTML = `<iframe src="${nData.stream_url}" class="w-full h-[68vh] rounded-xl border-0"></iframe>`;
          } else {
            nativeBody.innerHTML = `
              <div class="text-center p-8 space-y-3">
                <div class="text-sm font-bold text-slate-800">${escapeHtml(nData.filename)}</div>
                <div class="text-xs text-slate-500">Format: .${escapeHtml(nData.ext)} (${(nData.size_bytes / 1024).toFixed(1)} KB)</div>
                <a href="${nData.download_url}" class="inline-block bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold px-4 py-2 rounded-xl">
                  Download Native File to Desktop
                </a>
              </div>`;
          }
        }
      } catch (e) {
        STATE.currentNativeMeta = null;
        nativeBody.innerHTML = `<div class="p-6 text-xs text-rose-600">Error loading native file.</div>`;
      }

      updateDocBadgeAndDownload();
      applyDocSearch();
    }

    function highlightText(rawText, query) {
      const escaped = escapeHtml(rawText);
      if (!query) return { html: escaped, count: 0 };
      const safeRegex = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const regex = new RegExp(`(${safeRegex})`, 'gi');
      let count = 0;
      const html = escaped.replace(regex, (m) => {
        count++;
        return `<mark class="hl">${m}</mark>`;
      });
      return { html, count };
    }

    function applyDocSearch() {
      const query = (document.getElementById('docKeywordSearch').value || '').trim();
      const countEl = document.getElementById('searchMatchCount');

      if (STATE.activeDocTab === 'text') {
        const res = highlightText(STATE.currentTextRaw, query);
        document.getElementById('textViewerContent').innerHTML = res.html;
        countEl.textContent = query ? `${res.count} hit${res.count === 1 ? '' : 's'}` : '';
      } else if (STATE.currentNativeMeta?.is_text_preview) {
        const pre = document.getElementById('nativeTextPre');
        if (pre) {
          const res = highlightText(STATE.currentNativeMeta.preview_text || '', query);
          pre.innerHTML = res.html;
          countEl.textContent = query ? `${res.count} hit${res.count === 1 ? '' : 's'}` : '';
        }
      } else {
        countEl.textContent = '';
      }
    }

    // Boot
    fetchHostInfo();
  </script>
</body>
</html>
"""


if __name__ == "__main__":
    lan = get_lan_ip()
    print("=" * 72)
    print("SMART CLEAN UP — REAL-TIME LOCAL HOST SERVER")
    print("=" * 72)
    print(f"Local Admin URL : http://localhost:{APP_PORT}")
    print(f"Team LAN URL    : http://{lan}:{APP_PORT}")
    print(f"Admin User ID   : {ADMIN_USER_ID}")
    print("=" * 72)
    uvicorn.run(app, host=APP_HOST, port=APP_PORT)