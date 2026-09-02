# -*- coding: utf-8 -*-
"""Standalone **Excel Formula Tracer** - Dataiku Standard webapp Python pane.

Open any .xlsx / .xlsm workbook - by managed-folder id + filename, or by uploading it -
and walk its formulas node by node: which IF branch is live, what each sub-expression
evaluates to, what a VLOOKUP matched, the precedents back to the raw inputs, and a
workbook-wide lint.

Needs ``webapp_core`` and ``excel_deep_trace`` on the project's Python libraries
(both are in this bundle's ``python-lib/``).  No config, no auth, read-only.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import threading
import uuid
from datetime import datetime

from flask import request, jsonify

try:  # pragma: no cover
    import dataiku
    _HAS_DATAIKU = True
except Exception:  # noqa: BLE001
    dataiku = None
    _HAS_DATAIKU = False

import excel_deep_trace
from excel_deep_trace import ExcelDeepTraceEngine

_ENGINE_BUILD = getattr(excel_deep_trace, "ENGINE_BUILD", "unknown")
_ENGINE_OUTDATED = ("The excel_deep_trace.py on the project's Python libraries is out of "
                    "date for this webapp. Redeploy the whole python-lib/ folder from this "
                    "bundle and Restart the backend.")

try:  # pragma: no cover - Dataiku provides `app`
    app  # type: ignore  # noqa: F821
except NameError:  # pragma: no cover
    from flask import Flask
    app = Flask(__name__)

_SESSIONS: dict = {}
_LOCK = threading.RLock()
_MAX_UPLOAD_BYTES = 120 * 1024 * 1024
_ALLOWED = (".xlsx", ".xlsm")


def _err(msg, code=400, details=None):
    return jsonify({"success": False, "error": str(msg), "details": details}), code


def _session(sid):
    with _LOCK:
        s = _SESSIONS.get(sid)
    if not s:
        raise KeyError("Trace session not found - the backend may have restarted. "
                       "Open the workbook again.")
    return s


def _engine(sid, need=()):
    eng = _session(sid)["engine"]
    missing = [m for m in need if not hasattr(eng, m)]
    if missing:
        raise RuntimeError("%s (missing: %s)" % (_ENGINE_OUTDATED, ", ".join(missing)))
    return eng


def _open_bytes(content, filename):
    directory = tempfile.mkdtemp(prefix="excel_tracer_")
    safe = os.path.basename(filename).replace("/", "_").replace("\\", "_") or "workbook.xlsx"
    path = os.path.join(directory, safe)
    with open(path, "wb") as fh:
        fh.write(content if isinstance(content, bytes) else bytes(content))
    try:
        engine = ExcelDeepTraceEngine(path, max_expanded_range=20000)
    except TypeError:  # older engine signature
        engine = ExcelDeepTraceEngine(path)
    summary = engine.scan()
    sid = uuid.uuid4().hex
    with _LOCK:
        if len(_SESSIONS) > 24:  # keep the store small
            old = min(_SESSIONS, key=lambda k: _SESSIONS[k]["opened"])
            shutil.rmtree(_SESSIONS[old]["directory"], ignore_errors=True)
            _SESSIONS.pop(old, None)
        _SESSIONS[sid] = {"engine": engine, "directory": directory, "filename": safe,
                          "opened": datetime.utcnow().isoformat()}
    return sid, summary


# --------------------------------------------------------------------------- routes
@app.route("/api/health", methods=["GET"])
def api_health():
    return jsonify({"success": True, "dataiku": _HAS_DATAIKU,
                    "engine_build": _ENGINE_BUILD,
                    "engine_ok": hasattr(ExcelDeepTraceEngine, "formula_transformation"),
                    "time": datetime.utcnow().isoformat() + "Z"})


@app.route("/api/folders", methods=["GET"])
def api_folders():
    if not _HAS_DATAIKU:
        return jsonify({"success": True, "folders": []})
    try:
        client = dataiku.api_client()
        project = client.get_project(dataiku.default_project_key())
        return jsonify({"success": True, "folders": [
            {"id": m["id"], "name": m.get("name", m["id"])}
            for m in project.list_managed_folders()]})
    except Exception as exc:  # noqa: BLE001
        return jsonify({"success": True, "folders": [], "note": str(exc)})


@app.route("/api/folder-files", methods=["GET"])
def api_folder_files():
    folder_id = request.args.get("folder_id")
    if not folder_id:
        return _err("folder_id is required")
    if not _HAS_DATAIKU:
        return _err("not running inside Dataiku - upload a file instead")
    try:
        folder = dataiku.Folder(folder_id)
        paths = [p for p in folder.list_paths_in_partition()
                 if p.lower().endswith(_ALLOWED)]
        return jsonify({"success": True, "files": sorted(paths)})
    except Exception as exc:  # noqa: BLE001
        return _err("could not list folder %s: %s" % (folder_id, exc), 400)


@app.route("/api/open", methods=["POST"])
def api_open():
    body = request.get_json(force=True, silent=True) or {}
    folder_id, path = body.get("folder_id"), body.get("path")
    if not folder_id or not path:
        return _err("folder_id and path are required")
    if not _HAS_DATAIKU:
        return _err("not running inside Dataiku - upload a file instead")
    if not str(path).lower().endswith(_ALLOWED):
        return _err("only .xlsx / .xlsm workbooks are supported")
    try:
        folder = dataiku.Folder(folder_id)
        norm = path if str(path).startswith("/") else "/" + str(path)
        with folder.get_download_stream(norm) as stream:
            content = stream.read()
    except Exception as exc:  # noqa: BLE001
        return _err("could not read '%s' from folder %s: %s" % (path, folder_id, exc), 404)
    sid, summary = _open_bytes(content, os.path.basename(path))
    return jsonify({"success": True, "session_id": sid,
                    "filename": os.path.basename(path), "summary": summary})


@app.route("/api/upload", methods=["POST"])
def api_upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return _err("no file uploaded")
    if os.path.splitext(f.filename)[1].lower() not in _ALLOWED:
        return _err("only .xlsx / .xlsm workbooks are supported")
    content = f.read()
    if len(content) > _MAX_UPLOAD_BYTES:
        return _err("file exceeds the upload limit", 413)
    sid, summary = _open_bytes(content, f.filename)
    return jsonify({"success": True, "session_id": sid,
                    "filename": os.path.basename(f.filename), "summary": summary})


@app.route("/api/columns", methods=["POST"])
def api_columns():
    body = request.get_json(force=True, silent=True) or {}
    try:
        return jsonify({"success": True,
                        **_session(body["session_id"])["engine"].sheet_columns(body["sheet"])})
    except KeyError as exc:
        return _err(str(exc), 404)
    except Exception as exc:  # noqa: BLE001
        return _err(str(exc), 400)


@app.route("/api/journey", methods=["POST"])
def api_journey():
    body = request.get_json(force=True, silent=True) or {}
    try:
        eng = _session(body["session_id"])["engine"]
        return jsonify({"success": True,
                        "journey": eng.row_journey(body["sheet"], int(body["row"]))})
    except KeyError as exc:
        return _err(str(exc), 404)
    except Exception as exc:  # noqa: BLE001
        return _err(str(exc), 400)


@app.route("/api/transformation", methods=["POST"])
def api_transformation():
    body = request.get_json(force=True, silent=True) or {}
    try:
        eng = _engine(body["session_id"], need=["formula_transformation"])
        return jsonify({"success": True,
                        "transformation": eng.formula_transformation(body["cell"])})
    except KeyError as exc:
        return _err(str(exc), 404)
    except Exception as exc:  # noqa: BLE001
        return _err(str(exc), 400)


@app.route("/api/explain-plain", methods=["POST"])
def api_explain_plain():
    body = request.get_json(force=True, silent=True) or {}
    try:
        eng = _engine(body["session_id"], need=["explain_plain"])
        return jsonify({"success": True, **eng.explain_plain(body["cell"])})
    except KeyError as exc:
        return _err(str(exc), 404)
    except Exception as exc:  # noqa: BLE001
        return _err(str(exc), 400)


@app.route("/api/tree", methods=["POST"])
def api_tree():
    body = request.get_json(force=True, silent=True) or {}
    try:
        eng = _session(body["session_id"])["engine"]
        return jsonify({"success": True,
                        "tree": eng.expression_tree(body["cell"], int(body.get("max_depth", 8)))})
    except KeyError as exc:
        return _err(str(exc), 404)
    except Exception as exc:  # noqa: BLE001
        return _err(str(exc), 400)


@app.route("/api/explain", methods=["POST"])
def api_explain():
    body = request.get_json(force=True, silent=True) or {}
    try:
        eng = _session(body["session_id"])["engine"]
        key = eng._normalize_key(body["cell"])
        rec = eng.records.get(key)
        depth = int(body.get("max_depth", 8))
        return jsonify({
            "success": True, "cell": key, "value": eng.value(key),
            "formula": rec.formula if rec else None,
            "formula_metadata": {
                "functions": rec.functions if rec else [],
                "formula_family": rec.family if rec else None,
                "dependency_count": rec.dependency_count if rec else 0,
            },
            "branch_analysis": eng.branch_analysis(key),
            "lookup_analysis": eng.lookup_analysis(key),
            "precedent_graph": eng.trace(key, "precedents", depth, 300),
            "dependent_graph": eng.trace(key, "dependents", depth, 300),
        })
    except KeyError as exc:
        return _err(str(exc), 404)
    except Exception as exc:  # noqa: BLE001
        return _err(str(exc), 400)


@app.route("/api/mismatch", methods=["POST"])
def api_mismatch():
    body = request.get_json(force=True, silent=True) or {}
    try:
        eng = _session(body["session_id"])["engine"]
        return jsonify({"success": True, **eng.mismatch_rows(
            body["sheet"], body["left"], body["right"], float(body.get("tolerance", 0.01)))})
    except KeyError as exc:
        return _err(str(exc), 404)
    except Exception as exc:  # noqa: BLE001
        return _err(str(exc), 400)


@app.route("/api/findings", methods=["GET"])
def api_findings():
    try:
        eng = _session(request.args.get("session_id"))["engine"]
        return jsonify({"success": True, "findings": [
            {"severity": f.severity, "code": f.code, "cell": f.cell, "message": f.message}
            for f in eng.findings]})
    except KeyError as exc:
        return _err(str(exc), 404)


@app.route("/api/formulas", methods=["GET"])
def api_formulas():
    try:
        eng = _session(request.args.get("session_id"))["engine"]
        sheet = request.args.get("sheet")
        func = (request.args.get("function") or "").upper()
        out = []
        for rec in eng.records.values():
            if sheet and rec.sheet != sheet:
                continue
            if func and func not in rec.functions:
                continue
            out.append({"key": rec.key, "sheet": rec.sheet, "cell": rec.cell,
                        "formula": rec.formula, "value": rec.cached_value,
                        "family": rec.family, "functions": rec.functions,
                        "is_conditional": rec.is_conditional, "is_lookup": rec.is_lookup})
        return jsonify({"success": True, "total": len(out), "formulas": out[:2000]})
    except KeyError as exc:
        return _err(str(exc), 404)
