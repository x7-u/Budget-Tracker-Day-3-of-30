"""Day 3. Budget vs Actual Tracker, local Flask server.

Pure-local: bound to 127.0.0.1:1003 by default. Day-N port convention,
each day binds to port 1000 + N (Day 3 = 1003) so multiple days can run
side by side.

Routes:
  GET  /                         renders index.html, sets CSRF cookie
  POST /api/analyse              workbook upload, returns variance JSON
  GET  /api/status               environment + sample availability
  GET  /api/download/<filename>  serve a file from outputs/
  POST /api/shutdown             debug-only clean stop
  GET  /favicon.ico              static SVG icon

Same hardening as Days 1 and 2: 5 MB upload cap, secure_filename and
safe_join, single-flight semaphore on /api/analyse, CSRF double-submit
cookie, generic 500 to client and full traceback in logs/server.log.
"""
from __future__ import annotations

import argparse
import logging
import logging.handlers
import os
import secrets
import sys
import threading
import time
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from comparison import compare
from comparison import to_dict as compare_to_dict
from comparison import write_csv as compare_write_csv
from comparison import write_workbook as compare_write_workbook
from cost_log import CostLog
from csv_writer import write_csv
from excel_writer import write_workbook
from flask import Flask, abort, jsonify, make_response, render_template, request, send_file
from history_store import HistoryStore
from pipeline import analyse, to_dict
from power_bi import write_power_bi_assets
from pptx_writer import is_available as pptx_available
from pptx_writer import write_pptx
from run_cache import RunCache
from werkzeug.utils import safe_join, secure_filename

from shared.config import ANTHROPIC_API_KEY

HERE = Path(__file__).resolve().parent
SAMPLE_DIR = HERE / "sample_data"
OUTPUTS = HERE / "outputs"
UPLOADS = HERE / "uploads"
LOGS = HERE / "logs"
POWER_BI_DIR = HERE / "power_bi"

SAMPLES: dict[str, tuple[str, str]] = {
    "general":   ("sample_general.xlsx",
                  "Sample Co Ltd. 6 cost centres x 5 categories, single month, all cost lines."),
    "marketing": ("sample_marketing.xlsx",
                  "Marketing function deep-dive. 4 sub-teams x 6 categories, "
                  "single month, per-row tolerance overrides."),
    "quarterly": ("sample_quarterly.xlsx",
                  "Quarterly view. 3 months, 4 cost centres x 4 categories, "
                  "mixed cost and revenue lines."),
}

MAX_UPLOAD_BYTES = 5 * 1024 * 1024
ALLOWED_EXTS = {".xlsx", ".csv"}
CSRF_COOKIE_NAME = "csrf_token"
CSRF_HEADER_NAME = "X-CSRF-Token"

app = Flask(
    __name__,
    template_folder=str(HERE / "templates"),
    static_folder=str(HERE / "static"),
)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES

_analyse_lock = threading.Lock()
_cost_log = CostLog(OUTPUTS / "runs.jsonl")
_history = HistoryStore(OUTPUTS / "history.jsonl")
_run_cache = RunCache(OUTPUTS / "runs")


# ---- Logging ---------------------------------------------------------

LOGS.mkdir(parents=True, exist_ok=True)
_handler = logging.handlers.RotatingFileHandler(
    LOGS / "server.log", maxBytes=512_000, backupCount=3, encoding="utf-8",
)
_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
logging.basicConfig(level=logging.INFO, handlers=[_handler, logging.StreamHandler()])
log = logging.getLogger("day03.server")


# ---- Helpers ---------------------------------------------------------

def _env_key_ok() -> bool:
    return bool(ANTHROPIC_API_KEY) and not ANTHROPIC_API_KEY.startswith("sk-ant-placeholder")


def _ensure_csrf_cookie(resp):
    if not request.cookies.get(CSRF_COOKIE_NAME):
        resp.set_cookie(
            CSRF_COOKIE_NAME, secrets.token_urlsafe(24),
            samesite="Strict", httponly=False, secure=False, max_age=24 * 3600,
        )
    return resp


def _csrf_check() -> bool:
    cookie = request.cookies.get(CSRF_COOKIE_NAME, "")
    header = request.headers.get(CSRF_HEADER_NAME, "")
    return bool(cookie) and secrets.compare_digest(cookie, header)


def _samples_for_template():
    out = []
    for sid, (fname, label) in SAMPLES.items():
        if (SAMPLE_DIR / fname).exists():
            out.append({"id": sid, "filename": fname, "label": label})
    return out


# ---- Routes ----------------------------------------------------------

@app.route("/")
def index():
    resp = make_response(render_template(
        "index.html",
        env_key_ok=_env_key_ok(),
        samples=_samples_for_template(),
        max_upload_mb=MAX_UPLOAD_BYTES // (1024 * 1024),
    ))
    return _ensure_csrf_cookie(resp)


@app.route("/api/status")
def status():
    return jsonify(
        env_key_ok=_env_key_ok(),
        samples=_samples_for_template(),
        max_upload_mb=MAX_UPLOAD_BYTES // (1024 * 1024),
        cost_log=_cost_log_dict(),
    )


@app.route("/api/cost-log")
def cost_log_route():
    return jsonify(
        summary=_cost_log_dict(),
        entries=_cost_log.entries(limit=200),
    )


@app.route("/api/cost-log/clear", methods=["POST"])
def cost_log_clear():
    if not _csrf_check():
        return jsonify(error="CSRF token missing or invalid. Refresh the page."), 403
    n = _cost_log.clear()
    log.info("cost log cleared: %d entries removed", n)
    return jsonify(cleared=n, summary=_cost_log_dict())


@app.route("/api/history/clear", methods=["POST"])
def history_clear():
    if not _csrf_check():
        return jsonify(error="CSRF token missing or invalid. Refresh the page."), 403
    n = _history.clear()
    log.info("history store cleared: %d records removed", n)
    return jsonify(cleared=n, stats=_history.stats())


@app.route("/api/runs")
def runs_list():
    """Return cost-log entries augmented with a 'cached' flag so the UI
    can show whether re-opening a past run is possible."""
    entries = _cost_log.entries(limit=200)
    out = []
    for e in entries:
        eid = e.get("id")
        cached = bool(eid) and (_run_cache.root / f"{eid}.json").is_file()
        out.append({**e, "cached": cached})
    return jsonify(entries=out, summary=_cost_log_dict())


@app.route("/api/runs/<run_id>")
def run_get(run_id: str):
    if not run_id or len(run_id) > 64 or not run_id.replace("-", "").isalnum():
        return jsonify(error="Invalid run id."), 400
    payload = _run_cache.get(run_id)
    if payload is None:
        return jsonify(error="Run not cached. Older runs are evicted to disk."), 404
    return jsonify(payload)


@app.route("/api/runs/<run_id>", methods=["DELETE"])
def run_delete(run_id: str):
    if not _csrf_check():
        return jsonify(error="CSRF token missing or invalid. Refresh the page."), 403
    if not run_id or len(run_id) > 64 or not run_id.replace("-", "").isalnum():
        return jsonify(error="Invalid run id."), 400
    if _run_cache.remove(run_id):
        return jsonify(removed=run_id)
    return jsonify(error="Not cached."), 404


def _cost_log_dict() -> dict:
    s = _cost_log.summary()
    return {
        "runs": s.runs,
        "cost_usd_total": s.cost_usd_total,
        "rows_total": s.rows_total,
        "last_run_at": s.last_run_at,
        "cost_usd_30d": s.cost_usd_30d,
        "runs_30d": s.runs_30d,
    }


@app.route("/api/analyse", methods=["POST"])
def api_analyse():
    if not _csrf_check():
        return jsonify(error="CSRF token missing or invalid. Refresh the page."), 403

    if not _analyse_lock.acquire(blocking=False):
        return jsonify(error="Another analysis is already in flight. Wait for it to finish."), 429

    started = time.time()
    try:
        skip_ai = request.form.get("skip_ai") == "true"
        api_key_override = (request.form.get("api_key") or "").strip() or None
        model_choice = (request.form.get("model") or "").strip() or None
        use_samples = request.form.get("use_samples") == "true"
        sample_id = (request.form.get("sample_id") or "").strip()

        if use_samples:
            if sample_id not in SAMPLES:
                return jsonify(error=f"Unknown sample id: '{sample_id}'."), 400
            fname, _ = SAMPLES[sample_id]
            sample_path = SAMPLE_DIR / fname
            if not sample_path.exists():
                return jsonify(error=f"Sample file missing on disk: {fname}."), 500
            file_bytes = sample_path.read_bytes()
            display_name = fname
        else:
            upload = request.files.get("file")
            if upload is None or not upload.filename:
                return jsonify(error="No file uploaded. Pick an .xlsx workbook."), 400
            safe_name = secure_filename(upload.filename) or "upload.xlsx"
            ext = Path(safe_name).suffix.lower()
            if ext not in ALLOWED_EXTS:
                return jsonify(error=f"Unsupported file type: {ext} (only .xlsx is supported)."), 400
            file_bytes = upload.read()
            UPLOADS.mkdir(parents=True, exist_ok=True)
            (UPLOADS / f"{uuid.uuid4().hex[:8]}_{safe_name}").write_bytes(file_bytes)
            display_name = safe_name

        try:
            result = analyse(
                file_bytes=file_bytes,
                source_filename=display_name,
                model=model_choice,
                api_key=api_key_override,
                skip_ai=skip_ai,
                history=_history,
            )
        except ValueError as e:
            log.warning("analyse validation error: %s", e)
            return jsonify(error=str(e)), 400

        # Always write the outputs; the client gets filenames to download.
        xlsx_path = write_workbook(result, OUTPUTS)
        csv_path = write_csv(result, OUTPUTS / "budget_variance.csv")
        pptx_path = write_pptx(result, OUTPUTS) if pptx_available() else None

        elapsed_ms = int((time.time() - started) * 1000)
        cost = result.commentary.cost_usd
        log_entry = _cost_log.append(
            company=result.metadata.company,
            period_label=result.metadata.period_label,
            rows=len(result.rows),
            cost_usd=cost,
            model=result.commentary.model,
            skipped=result.commentary.skipped,
            elapsed_ms=elapsed_ms,
            source_filename=display_name,
            total_variance=result.headline.total_variance,
            total_variance_pct=result.headline.total_variance_pct,
            rag_red=result.headline.rag_counts.get("red", 0),
        )
        log.info(
            "analyse OK rows=%d ai=%s ms=%d cost_usd=%.4f",
            len(result.rows), not skip_ai, elapsed_ms, cost,
        )

        body = to_dict(result)
        anomaly_count = sum(1 for r in result.rows if r.z_score is not None and abs(r.z_score) >= 2.0)
        body.update(
            xlsx_filename=xlsx_path.name,
            csv_filename=csv_path.name,
            pptx_filename=(pptx_path.name if pptx_path is not None else None),
            elapsed_ms=elapsed_ms,
            total_cost_usd=round(cost, 6),
            cost_log=_cost_log_dict(),
            anomaly_count=anomaly_count,
            history_stats=_history.stats(),
            run_id=log_entry["id"],
        )
        # Cache the full payload so the user can re-open this run later.
        try:
            _run_cache.save(log_entry["id"], body)
        except Exception:
            log.exception("failed to cache run %s", log_entry["id"])
        return jsonify(body)
    except Exception:
        log.exception("analyse unexpected error")
        return jsonify(
            error="Server error during analysis. See logs/server.log for details."
        ), 500
    finally:
        _analyse_lock.release()


@app.route("/api/compare", methods=["POST"])
def api_compare():
    """Compare two workbooks at the row level. AI is skipped on both runs."""
    if not _csrf_check():
        return jsonify(error="CSRF token missing or invalid. Refresh the page."), 403

    if not _analyse_lock.acquire(blocking=False):
        return jsonify(error="Another analysis is already in flight. Wait for it to finish."), 429

    try:
        files = [request.files.get("file_a"), request.files.get("file_b")]
        if any(f is None or not f.filename for f in files):
            return jsonify(error="Comparison needs two .xlsx files: file_a and file_b."), 400

        results = []
        for f in files:
            safe_name = secure_filename(f.filename) or "upload.xlsx"
            ext = Path(safe_name).suffix.lower()
            if ext not in ALLOWED_EXTS:
                return jsonify(error=f"Unsupported file type: {ext} (only .xlsx)."), 400
            data = f.read()
            try:
                results.append(analyse(file_bytes=data, source_filename=safe_name, skip_ai=True))
            except ValueError as e:
                return jsonify(error=f"{safe_name}: {e}"), 400

        diff = compare(results[0], results[1])
        # Write the diff to disk so the user can download Excel + CSV.
        compare_xlsx = compare_write_workbook(diff, OUTPUTS)
        compare_csv = compare_write_csv(diff, OUTPUTS / "compare.csv")
        body = compare_to_dict(diff)
        body["xlsx_filename"] = compare_xlsx.name
        body["csv_filename"] = compare_csv.name
        return jsonify(body)
    except Exception:
        log.exception("compare unexpected error")
        return jsonify(error="Server error during comparison. See logs/server.log for details."), 500
    finally:
        _analyse_lock.release()


@app.errorhandler(413)
def _too_large(_e):
    return jsonify(error=f"Upload exceeds {MAX_UPLOAD_BYTES // 1024 // 1024} MB limit."), 413


@app.route("/api/download/<path:filename>")
def download(filename: str):
    safe = secure_filename(filename) or ""
    if not safe:
        abort(400)
    full = safe_join(str(OUTPUTS), safe)
    if not full or not Path(full).is_file():
        return jsonify(error=f"Not found: {safe}"), 404
    return send_file(full, as_attachment=True, download_name=safe)


@app.route("/api/shutdown", methods=["POST"])
def shutdown():
    if not (app.debug or os.getenv("DAY03_ALLOW_SHUTDOWN") == "1"):
        return jsonify(error="Shutdown not enabled. Run with --debug or DAY03_ALLOW_SHUTDOWN=1."), 403
    if not _csrf_check():
        return jsonify(error="CSRF token missing."), 403
    threading.Thread(target=lambda: (time.sleep(0.2), os._exit(0)), daemon=True).start()
    return jsonify(stopped=True)


@app.route("/favicon.ico")
def favicon():
    p = HERE / "static" / "favicon.svg"
    if p.exists():
        return send_file(p)
    return ("", 204)


@app.route("/power_bi/<path:filename>")
def power_bi_file(filename: str):
    """Serve the Power BI M script and README. Lazily regenerate the script
    on first request so the embedded outputs path stays accurate after the
    user moves the project."""
    safe = secure_filename(filename) or ""
    if safe not in {"power_query.m", "README.md"}:
        return jsonify(error="Unknown Power BI asset."), 404
    path = POWER_BI_DIR / safe
    if not path.is_file() or safe == "power_query.m":
        try:
            write_power_bi_assets(POWER_BI_DIR)
        except Exception:
            log.exception("failed to regenerate Power BI assets")
    if not path.is_file():
        return jsonify(error="Power BI asset missing on disk."), 500
    return send_file(path, as_attachment=True, download_name=safe)


# ---- CLI -------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=int(os.getenv("DAY03_PORT", "1003")))
    parser.add_argument(
        "--host", default="127.0.0.1",
        help="Loopback by default. Keep it that way unless you know why.",
    )
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()
    print()
    print("  Day 3. Budget vs Actual Tracker")
    print(f"  Local URL:  http://{args.host}:{args.port}/")
    print("  Press Ctrl+C to stop.")
    print()
    app.run(host=args.host, port=args.port, debug=args.debug, use_reloader=args.debug)


if __name__ == "__main__":
    main()
