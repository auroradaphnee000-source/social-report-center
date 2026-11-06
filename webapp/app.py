from __future__ import annotations

import json
import logging
import os
import queue
import re
import sqlite3
import threading
import traceback
import uuid
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pydantic import BaseModel

from social_report import generate_report
from webapp.meta import MetaError, MetaGraphClient, MetaRepository, MetaSettings
from webapp.popo import PopoDeliveryError, PopoSettings, deliver_job
from webapp.popo_auth import (
    PopoAuthError,
    begin_popo_login,
    fabric_cli,
    popo_status,
    resume_popo_login,
    start_auth_refresher,
)


CN_TZ = timezone(timedelta(hours=8))
LOGGER = logging.getLogger("social_report.meta")
ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = Path(os.getenv("APP_DATA_DIR", ROOT / "runtime")).resolve()
JOBS_DIR = DATA_DIR / "jobs"
DB_PATH = DATA_DIR / "jobs.sqlite3"
STATIC_DIR = Path(__file__).resolve().parent / "static"
MAX_FILE_BYTES = int(os.getenv("MAX_FILE_BYTES", str(25 * 1024 * 1024)))
ALLOWED_IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg"}
ALLOWED_SHEET_SUFFIXES = {".xls", ".xlsx"}
REPORT_TYPE_LABELS = {
    "monthly": "月报",
    "quarterly": "季度报",
    "annual": "年报",
    "custom": "自定义报告",
}

DATA_DIR.mkdir(parents=True, exist_ok=True)
JOBS_DIR.mkdir(parents=True, exist_ok=True)
_wake = queue.Queue()
_popo_wake = queue.Queue()
_meta_sync_wake = queue.Queue()
_stop_workers = threading.Event()
_worker_threads: list[threading.Thread] = []
META_SETTINGS = MetaSettings.from_env()
META_REPOSITORY = MetaRepository(DB_PATH, META_SETTINGS)
META_CLIENT = MetaGraphClient(META_SETTINGS)
POPO_SETTINGS = PopoSettings.from_env()


def _now_cn() -> str:
    return datetime.now(CN_TZ).replace(tzinfo=None).isoformat(timespec="seconds")


@contextmanager
def _db():
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def _init_db():
    with _db() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS jobs (
                id TEXT PRIMARY KEY,
                brand TEXT NOT NULL,
                report_type TEXT NOT NULL DEFAULT 'monthly',
                report_start TEXT NOT NULL,
                report_end TEXT NOT NULL,
                status TEXT NOT NULL,
                error TEXT,
                created_at TEXT NOT NULL,
                started_at TEXT,
                finished_at TEXT,
                created_by TEXT,
                popo_status TEXT,
                popo_url TEXT,
                popo_doc_id TEXT,
                popo_error TEXT,
                popo_requested_at TEXT,
                popo_finished_at TEXT,
                popo_attempts INTEGER NOT NULL DEFAULT 0
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS meta_sync_jobs (
                id TEXT PRIMARY KEY,
                page_id TEXT NOT NULL,
                instagram_id TEXT,
                report_start TEXT NOT NULL,
                report_end TEXT NOT NULL,
                status TEXT NOT NULL,
                progress INTEGER NOT NULL DEFAULT 0,
                message TEXT,
                error TEXT,
                error_code INTEGER,
                reconnect INTEGER NOT NULL DEFAULT 0,
                sync_id TEXT,
                created_at TEXT NOT NULL,
                started_at TEXT,
                finished_at TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_meta_sync_jobs_status_created
            ON meta_sync_jobs(status, created_at)
            """
        )
        columns = {
            row["name"] for row in conn.execute("PRAGMA table_info(jobs)").fetchall()
        }
        if "report_type" not in columns:
            conn.execute(
                "ALTER TABLE jobs ADD COLUMN report_type TEXT NOT NULL DEFAULT 'monthly'"
            )
        popo_columns = {
            "popo_status": "TEXT",
            "popo_url": "TEXT",
            "popo_doc_id": "TEXT",
            "popo_error": "TEXT",
            "popo_requested_at": "TEXT",
            "popo_finished_at": "TEXT",
            "popo_attempts": "INTEGER NOT NULL DEFAULT 0",
        }
        for name, declaration in popo_columns.items():
            if name not in columns:
                conn.execute(f"ALTER TABLE jobs ADD COLUMN {name} {declaration}")
        conn.execute(
            "UPDATE jobs SET status='queued', started_at=NULL "
            "WHERE status='running'"
        )
        conn.execute(
            "UPDATE jobs SET popo_status='pending' "
            "WHERE popo_status='processing'"
        )
        conn.execute(
            """
            UPDATE meta_sync_jobs
            SET status='queued', progress=5,
                message='服务已恢复，等待继续同步', started_at=NULL
            WHERE status='running'
            """
        )


def _set_job(job_id: str, **values):
    if not values:
        return
    fields = ", ".join(f"{key}=?" for key in values)
    with _db() as conn:
        conn.execute(
            f"UPDATE jobs SET {fields} WHERE id=?",
            [*values.values(), job_id],
        )


def _set_meta_sync_job(sync_job_id: str, **values):
    if not values:
        return
    fields = ", ".join(f"{key}=?" for key in values)
    with _db() as conn:
        conn.execute(
            f"UPDATE meta_sync_jobs SET {fields} WHERE id=?",
            [*values.values(), sync_job_id],
        )


def _meta_sync_job_dict(row):
    if not row:
        return None
    result = dict(row)
    result["reconnect"] = bool(result.get("reconnect"))
    sync_id = result.get("sync_id")
    if result.get("status") == "succeeded" and sync_id:
        sync = META_REPOSITORY.get_sync(sync_id)
        if sync:
            payload = dict(sync["result"])
            payload["sync_id"] = sync_id
            result["result"] = payload
    return result


def _get_meta_sync_job(sync_job_id: str):
    with _db() as conn:
        row = conn.execute(
            "SELECT * FROM meta_sync_jobs WHERE id=?", (sync_job_id,)
        ).fetchone()
    if not row:
        raise HTTPException(404, "Meta 同步任务不存在")
    return _meta_sync_job_dict(row)


def _run_meta_sync_job(row):
    sync_job_id = row["id"]
    _set_meta_sync_job(
        sync_job_id,
        status="running",
        progress=10,
        message="正在连接 Meta 并读取账号数据",
        error=None,
        error_code=None,
        reconnect=0,
        started_at=_now_cn(),
        finished_at=None,
    )

    def update_progress(progress: int, message: str):
        _set_meta_sync_job(
            sync_job_id,
            progress=max(10, min(int(progress), 95)),
            message=_clean_text(message, 200),
        )

    try:
        asset = META_REPOSITORY.get_asset(row["page_id"])
        result = META_CLIENT.sync(
            asset,
            instagram_id=row["instagram_id"] or None,
            report_start=row["report_start"],
            report_end=row["report_end"],
            progress_callback=update_progress,
        )
        update_progress(96, "正在保存同步结果")
        sync_id = META_REPOSITORY.save_sync(
            page_id=asset["page_id"],
            instagram_id=row["instagram_id"],
            report_start=row["report_start"],
            report_end=row["report_end"],
            result=result,
        )
        META_REPOSITORY.touch_connection()
        _set_meta_sync_job(
            sync_job_id,
            status="succeeded",
            progress=100,
            message="Meta 数据同步完成",
            sync_id=sync_id,
            error=None,
            error_code=None,
            reconnect=0,
            finished_at=_now_cn(),
        )
    except MetaError as exc:
        LOGGER.warning(
            "Meta background sync %s failed (code=%s, reconnect=%s): %s",
            sync_job_id,
            exc.code,
            exc.reconnect,
            exc,
        )
        _set_meta_sync_job(
            sync_job_id,
            status="failed",
            message="Meta 数据同步失败",
            error=_clean_text(str(exc), 800),
            error_code=exc.code,
            reconnect=1 if exc.reconnect else 0,
            finished_at=_now_cn(),
        )
    except Exception as exc:
        message = f"{type(exc).__name__}: {exc}"
        LOGGER.exception("Meta background sync %s crashed", sync_job_id)
        _set_meta_sync_job(
            sync_job_id,
            status="failed",
            message="服务器处理 Meta 数据时失败",
            error=_clean_text(message, 800),
            error_code=None,
            reconnect=0,
            finished_at=_now_cn(),
        )


def _meta_sync_worker():
    while not _stop_workers.is_set():
        with _db() as conn:
            row = conn.execute(
                """
                SELECT * FROM meta_sync_jobs
                WHERE status='queued'
                ORDER BY created_at LIMIT 1
                """
            ).fetchone()
        if row:
            _run_meta_sync_job(row)
            continue
        try:
            _meta_sync_wake.get(timeout=5)
        except queue.Empty:
            pass


def _artifact_list(job_id: str):
    output_dir = JOBS_DIR / job_id / "outputs"
    if not output_dir.exists():
        return []
    artifacts = []
    for path in sorted(output_dir.iterdir()):
        if not path.is_file() or path.suffix.lower() not in {".docx", ".json"}:
            continue
        artifacts.append(
            {
                "name": path.name,
                "size": path.stat().st_size,
                "url": f"/api/jobs/{job_id}/files/{path.name}",
                "kind": "report" if path.suffix.lower() == ".docx" else "analysis",
            }
        )
    return artifacts


def _job_dict(row):
    if not row:
        return None
    result = dict(row)
    result["report_type"] = result.get("report_type") or "monthly"
    result["report_type_label"] = REPORT_TYPE_LABELS.get(
        result["report_type"], REPORT_TYPE_LABELS["custom"]
    )
    result["artifacts"] = _artifact_list(result["id"])
    return result


def _get_job(job_id: str):
    with _db() as conn:
        row = conn.execute("SELECT * FROM jobs WHERE id=?", (job_id,)).fetchone()
    if not row:
        raise HTTPException(404, "任务不存在")
    return _job_dict(row)


def _run_job(row):
    job_id = row["id"]
    job_dir = JOBS_DIR / job_id
    log_path = job_dir / "run.log"
    request_path = job_dir / "request.json"
    _set_job(job_id, status="running", started_at=_now_cn(), error=None)
    with log_path.open("a", encoding="utf-8") as log:
        try:
            request = json.loads(request_path.read_text(encoding="utf-8"))
            log.write(f"[{_now_cn()}] 开始生成：{request['brand']}\n")
            log.flush()
            output_path, analysis = generate_report(job_dir, request)
            linkedin = analysis["linkedin"]
            linkedin_summary = (
                f"LinkedIn 原生展示 {linkedin['content']['impressions_organic']}"
                if linkedin.get("mode") == "company"
                else (
                    "LinkedIn 个人账号 Top 10 样本 "
                    f"{linkedin['posts']['count']} 条"
                    if linkedin.get("mode") == "personal"
                    else "本期无 LinkedIn 数据"
                )
            )
            log.write(
                f"[{_now_cn()}] 生成完成：{output_path.name}，"
                f"{linkedin_summary}\n"
            )
            _set_job(
                job_id,
                status="succeeded",
                finished_at=_now_cn(),
                error=None,
                popo_status="pending",
                popo_error=None,
                popo_requested_at=_now_cn(),
            )
            _popo_wake.put(job_id)
        except Exception as exc:
            message = f"{type(exc).__name__}: {exc}"
            log.write(f"[{_now_cn()}] 生成失败：{message}\n")
            log.write(traceback.format_exc())
            _set_job(
                job_id,
                status="failed",
                finished_at=_now_cn(),
                error=message[:800],
            )


def _worker():
    while not _stop_workers.is_set():
        with _db() as conn:
            row = conn.execute(
                "SELECT * FROM jobs WHERE status='queued' ORDER BY created_at LIMIT 1"
            ).fetchone()
        if row:
            _run_job(row)
            continue
        try:
            _wake.get(timeout=5)
        except queue.Empty:
            pass


def _run_popo_delivery(row):
    job_id = row["id"]
    job_dir = JOBS_DIR / job_id
    log_path = job_dir / "run.log"
    _set_job(
        job_id,
        popo_status="processing",
        popo_error=None,
        popo_attempts=int(row["popo_attempts"] or 0) + 1,
    )
    with log_path.open("a", encoding="utf-8") as log:
        try:
            log.write(f"[{_now_cn()}] 开始交付 POPO 原生报告\n")
            log.flush()
            result = deliver_job(job_dir, log)
            _set_job(
                job_id,
                popo_status="completed",
                popo_url=result["url"][:500],
                popo_doc_id=result["doc_id"][:120],
                popo_error=None,
                popo_finished_at=_now_cn(),
            )
            log.write(
                f"[{_now_cn()}] POPO 交付完成：{result['url']}"
                f"（{'新建' if result['created'] else '同名复用'}）\n"
            )
        except Exception as exc:
            message = f"{type(exc).__name__}: {exc}"
            log.write(f"[{_now_cn()}] POPO 交付失败：{message}\n")
            log.write(traceback.format_exc())
            _set_job(
                job_id,
                popo_status="failed",
                popo_error=message[:800],
                popo_finished_at=_now_cn(),
            )


def _popo_worker():
    while not _stop_workers.is_set():
        with _db() as conn:
            row = conn.execute(
                "SELECT * FROM jobs "
                "WHERE status='succeeded' AND popo_status='pending' "
                "ORDER BY popo_requested_at, created_at LIMIT 1"
            ).fetchone()
        if row:
            _run_popo_delivery(row)
            continue
        try:
            _popo_wake.get(timeout=5)
        except queue.Empty:
            pass


def _validate_date(value: str, label: str):
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except ValueError as exc:
        raise HTTPException(400, f"{label}必须是 YYYY-MM-DD") from exc


def _clean_text(value: str, max_len: int = 200) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())[:max_len]


def _optional_nonnegative_int(value: str, label: str) -> int | None:
    text = (value or "").strip().replace(",", "")
    if not text:
        return None
    try:
        number = int(text)
    except ValueError as exc:
        raise HTTPException(400, f"{label}必须是非负整数") from exc
    if number < 0:
        raise HTTPException(400, f"{label}不能为负数")
    return number


def _meta_status_payload() -> dict:
    base = {
        "configured": META_SETTINGS.configured,
        "graph_version": META_SETTINGS.graph_version,
        "redirect_uri": META_SETTINGS.redirect_uri,
    }
    if not META_SETTINGS.configured:
        return {
            **base,
            "connected": False,
            "reason": "server_not_configured",
            "message": "服务器尚未配置 Meta 开发者应用，当前可使用手动补充入口。",
        }
    connection = META_REPOSITORY.get_connection()
    if not connection:
        return {
            **base,
            "connected": False,
            "reason": "not_connected",
            "message": "尚未连接 Meta 账号。",
        }
    expires_at = connection.get("token_expires_at")
    if expires_at and int(expires_at) <= int(datetime.now(timezone.utc).timestamp()):
        return {
            **base,
            "connected": False,
            "reason": "authorization_expired",
            "message": "Meta 授权已过期，请重新连接。",
            "user_name": connection.get("user_name"),
        }
    return {
        **base,
        "connected": True,
        "reason": "connected",
        "message": "Meta 已连接，可以自动读取概览数据。",
        "user_name": connection.get("user_name"),
        "connected_at": connection.get("connected_at"),
        "last_verified_at": connection.get("last_verified_at"),
        "assets": META_REPOSITORY.public_assets(),
    }


def _raise_meta_http(exc: MetaError):
    status = 401 if exc.reconnect else 502
    raise HTTPException(
        status,
        {
            "message": str(exc),
            "reconnect": exc.reconnect,
            "code": exc.code,
        },
    ) from exc


async def _save_upload(
    upload: UploadFile, target_dir: Path, role: str, allowed_suffixes: set[str]
) -> str:
    original = Path(upload.filename or "").name
    suffix = Path(original).suffix.lower()
    if suffix not in allowed_suffixes:
        raise HTTPException(
            400, f"{role} 文件类型不正确，请上传：{', '.join(sorted(allowed_suffixes))}"
        )
    safe_stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(original).stem)[:60] or role
    filename = f"{role}_{safe_stem}{suffix}"
    target = target_dir / filename
    size = 0
    try:
        with target.open("wb") as stream:
            while True:
                chunk = await upload.read(1024 * 1024)
                if not chunk:
                    break
                size += len(chunk)
                if size > MAX_FILE_BYTES:
                    raise HTTPException(
                        413, f"{role} 超过 {MAX_FILE_BYTES // 1024 // 1024} MB 限制"
                    )
                stream.write(chunk)
    finally:
        await upload.close()
    if size == 0:
        target.unlink(missing_ok=True)
        raise HTTPException(400, f"{role} 是空文件")
    magic = target.read_bytes()[:8]
    if allowed_suffixes == ALLOWED_IMAGE_SUFFIXES:
        valid = magic.startswith(b"\x89PNG") or magic[:3] == b"\xff\xd8\xff"
    else:
        valid = magic.startswith(b"\xD0\xCF\x11\xE0") or magic.startswith(b"PK")
    if not valid:
        target.unlink(missing_ok=True)
        raise HTTPException(400, f"{role} 的实际文件格式与扩展名不一致")
    return filename


app = FastAPI(title="Social Report Center", version="1.5.6")


@app.on_event("startup")
def startup():
    _init_db()
    META_REPOSITORY.init_db()
    _stop_workers.clear()
    report_worker = threading.Thread(
        target=_worker, name="social-report-worker", daemon=True
    )
    popo_worker = threading.Thread(
        target=_popo_worker,
        name="social-report-popo-worker",
        daemon=True,
    )
    meta_sync_worker = threading.Thread(
        target=_meta_sync_worker,
        name="social-report-meta-sync-worker",
        daemon=True,
    )
    _worker_threads[:] = [report_worker, popo_worker, meta_sync_worker]
    for worker in _worker_threads:
        worker.start()
    start_auth_refresher()


@app.on_event("shutdown")
def shutdown():
    _stop_workers.set()
    _wake.put(None)
    _popo_wake.put(None)
    _meta_sync_wake.put(None)
    for worker in _worker_threads:
        worker.join(timeout=6)
    _worker_threads.clear()


@app.get("/api/health")
def health():
    with _db() as conn:
        queued = conn.execute(
            "SELECT COUNT(*) FROM jobs WHERE status IN ('queued','running')"
        ).fetchone()[0]
        meta_sync_queued = conn.execute(
            """
            SELECT COUNT(*) FROM meta_sync_jobs
            WHERE status IN ('queued','running')
            """
        ).fetchone()[0]
    return {
        "ok": True,
        "service": "social-report",
        "version": app.version,
        "queued": queued,
        "meta_sync_queued": meta_sync_queued,
        "max_file_mb": MAX_FILE_BYTES // 1024 // 1024,
        "meta_configured": META_SETTINGS.configured,
        "meta_connected": bool(
            META_SETTINGS.configured and META_REPOSITORY.get_connection()
        ),
        "popo_configured": POPO_SETTINGS.configured,
        "popo_available": bool(fabric_cli()),
    }


@app.get("/api/templates/linkedin-personal-top10.xlsx")
def linkedin_personal_template():
    book = Workbook()
    sheet = book.active
    sheet.title = "Top10帖子"
    headers = [
        "帖子发布日期",
        "帖子标题或内容摘要",
        "帖子链接",
        "浏览量/展示量",
        "点赞数",
        "评论数",
        "转发数",
        "点击数",
    ]
    sheet.append(headers)
    sheet.append(
        [
            "2026-07-01",
            "示例：产品应用或案例内容",
            "https://www.linkedin.com/posts/...",
            100,
            5,
            1,
            0,
            "",
        ]
    )
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="123A63")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for column, width in {
        "A": 16,
        "B": 42,
        "C": 42,
        "D": 18,
        "E": 12,
        "F": 12,
        "G": 12,
        "H": 12,
    }.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"
    notes = book.create_sheet("填写说明")
    notes.append(["填写规则"])
    notes.append(["1. 每行一条帖子，最多填写 10 条，系统会按浏览量/展示量重新排序。"])
    notes.append(["2. 发布日期、内容摘要、浏览量/展示量为必填；其余字段可留空。"])
    notes.append(["3. 仅填写自然流量数据，不要混入赞助或推广数据。"])
    notes.append(["4. 同一报告周期沿用相同筛选标准，便于后续对比。"])
    notes.column_dimensions["A"].width = 90
    notes["A1"].font = Font(bold=True, color="FFFFFF")
    notes["A1"].fill = PatternFill("solid", fgColor="123A63")
    stream = BytesIO()
    book.save(stream)
    stream.seek(0)
    return StreamingResponse(
        stream,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": (
                'attachment; filename="LinkedIn_personal_Top10_template.xlsx"'
            )
        },
    )


class MetaSyncBody(BaseModel):
    page_id: str
    instagram_id: str | None = None
    report_start: str
    report_end: str


@app.get("/api/meta/status")
def meta_status():
    return _meta_status_payload()


@app.get("/api/meta/connect")
def meta_connect():
    if not META_SETTINGS.configured:
        raise HTTPException(503, "服务器尚未配置 Meta 开发者应用")
    state = META_REPOSITORY.create_state()
    return RedirectResponse(META_CLIENT.authorization_url(state), status_code=302)


@app.get("/api/meta/callback")
def meta_callback(
    state: str = "",
    code: str = "",
    error: str = "",
    error_reason: str = "",
):
    if not META_SETTINGS.configured:
        return RedirectResponse("/?meta=error&reason=server_not_configured", 302)
    if not state or not META_REPOSITORY.consume_state(state):
        return RedirectResponse("/?meta=error&reason=invalid_state", 302)
    if error or not code:
        reason = _clean_text(error_reason or error or "authorization_cancelled", 60)
        return RedirectResponse(
            "/?meta=error&reason=" + re.sub(r"[^a-zA-Z0-9_-]+", "_", reason),
            302,
        )
    callback_stage = "exchange_code"
    try:
        token_result = META_CLIENT.exchange_code(code)
        token = token_result["access_token"]
        callback_stage = "get_profile"
        try:
            profile = META_CLIENT.get_profile(token)
        except MetaError:
            # Business Login tokens can grant Page/Instagram access while the
            # personal profile lookup is unavailable. The report workflow only
            # needs the managed assets, so keep profile metadata best-effort.
            profile = {"id": "", "name": "Meta account"}
        callback_stage = "save_connection"
        META_REPOSITORY.save_connection(
            user_id=str(profile.get("id") or ""),
            user_name=str(profile.get("name") or "Meta 用户"),
            token=token,
            expires_in=token_result.get("expires_in"),
            scopes=META_SETTINGS.scopes,
        )
        callback_stage = "list_pages"
        pages = META_CLIENT.list_pages(token)
        callback_stage = "save_assets"
        META_REPOSITORY.save_assets(pages)
        META_REPOSITORY.touch_connection()
    except MetaError as exc:
        LOGGER.warning(
            "Meta OAuth callback failed at %s (code=%s, subcode=%s, reconnect=%s): %s",
            callback_stage,
            exc.code,
            exc.subcode,
            exc.reconnect,
            exc,
        )
        message = str(exc)
        if exc.reconnect:
            reason = "reconnect"
        elif callback_stage == "exchange_code" and (
            "超时" in message or "网络" in message
        ):
            reason = "meta_network"
        elif callback_stage == "list_pages":
            reason = "asset_read_failed"
        else:
            reason = "api_error"
        return RedirectResponse(f"/?meta=error&reason={reason}", 302)
    return RedirectResponse("/?meta=connected", 302)


@app.get("/api/meta/assets")
def meta_assets(refresh: bool = False):
    status = _meta_status_payload()
    if not status.get("connected"):
        raise HTTPException(401, status.get("message") or "Meta 尚未连接")
    try:
        assets = META_REPOSITORY.public_assets()
        if refresh or not assets:
            connection = META_REPOSITORY.get_connection(include_token=True)
            pages = META_CLIENT.list_pages(connection["token"])
            META_REPOSITORY.save_assets(pages)
            META_REPOSITORY.touch_connection()
            assets = META_REPOSITORY.public_assets()
        return {"items": assets}
    except MetaError as exc:
        _raise_meta_http(exc)


@app.post("/api/meta/sync", status_code=202)
def meta_sync(body: MetaSyncBody):
    status = _meta_status_payload()
    if not status.get("connected"):
        raise HTTPException(401, status.get("message") or "Meta 尚未连接")
    start = _validate_date(body.report_start, "报告开始日期")
    end = _validate_date(body.report_end, "报告结束日期")
    if end < start:
        raise HTTPException(400, "报告结束日期不能早于开始日期")
    try:
        page_id = _clean_text(body.page_id, 80)
        instagram_id = _clean_text(body.instagram_id or "", 80) or None
        asset = META_REPOSITORY.get_asset(page_id)
        if instagram_id and instagram_id != asset.get("instagram_id"):
            raise HTTPException(
                400, "Instagram 账号与所选 Facebook 页面不匹配"
            )
        with _db() as conn:
            existing = conn.execute(
                """
                SELECT * FROM meta_sync_jobs
                WHERE page_id=? AND COALESCE(instagram_id, '')=COALESCE(?, '')
                  AND report_start=? AND report_end=?
                  AND status IN ('queued','running')
                ORDER BY created_at DESC LIMIT 1
                """,
                (
                    page_id,
                    instagram_id,
                    body.report_start,
                    body.report_end,
                ),
            ).fetchone()
            if existing:
                return _meta_sync_job_dict(existing)
            sync_job_id = uuid.uuid4().hex[:12]
            conn.execute(
                """
                INSERT INTO meta_sync_jobs(
                    id,page_id,instagram_id,report_start,report_end,status,
                    progress,message,created_at
                ) VALUES(?,?,?,?,?,'queued',5,?,?)
                """,
                (
                    sync_job_id,
                    page_id,
                    instagram_id,
                    body.report_start,
                    body.report_end,
                    "同步任务已进入后台队列",
                    _now_cn(),
                ),
            )
        _meta_sync_wake.put(sync_job_id)
        return _get_meta_sync_job(sync_job_id)
    except HTTPException:
        raise
    except MetaError as exc:
        _raise_meta_http(exc)


@app.get("/api/meta/sync/{sync_job_id}")
def meta_sync_status(sync_job_id: str):
    job = _get_meta_sync_job(_clean_text(sync_job_id, 40))
    if job["status"] == "failed":
        job["detail"] = {
            "message": job.get("error") or "Meta 数据同步失败",
            "reconnect": job.get("reconnect", False),
            "code": job.get("error_code"),
        }
    return job


@app.post("/api/meta/disconnect")
def meta_disconnect():
    META_REPOSITORY.disconnect()
    return {"ok": True}


@app.get("/api/integrations/popo/status")
def integration_popo_status():
    status = popo_status()
    return {
        **status,
        "configured": POPO_SETTINGS.configured,
        "folder_url": POPO_SETTINGS.folder_url,
    }


@app.post("/api/integrations/popo/login")
def integration_popo_login():
    if not POPO_SETTINGS.configured:
        raise HTTPException(503, "POPO 团队空间或社媒报告文件夹尚未配置")
    try:
        return begin_popo_login()
    except PopoAuthError as exc:
        raise HTTPException(503, str(exc)) from exc


@app.post("/api/integrations/popo/resume")
def integration_popo_resume():
    try:
        return resume_popo_login()
    except PopoAuthError as exc:
        detail = {
            "message": str(exc),
            "code": exc.code,
            "retryable": exc.retryable,
        }
        raise HTTPException(409 if exc.retryable else 503, detail) from exc


@app.post("/api/jobs")
async def create_job(
    brand: str = Form(...),
    account_label: str = Form(""),
    report_type: str = Form("monthly"),
    report_start: str = Form(...),
    report_end: str = Form(...),
    created_by: str = Form(...),
    meta_sync_id: str = Form(""),
    manual_override: str = Form(""),
    facebook_posts: str = Form(""),
    facebook_views: str = Form(""),
    facebook_reach: str = Form(""),
    facebook_viewers: str = Form(""),
    facebook_interactions: str = Form(""),
    facebook_likes: str = Form(""),
    instagram_posts: str = Form(""),
    instagram_views: str = Form(""),
    instagram_reach: str = Form(""),
    instagram_interactions: str = Form(""),
    instagram_likes: str = Form(""),
    facebook_highlight: str = Form(""),
    instagram_highlight: str = Form(""),
    linkedin_mode: str = Form("company"),
    linkedin_period_impressions: str = Form(""),
    linkedin_profile_views: str = Form(""),
    linkedin_current_followers: str = Form(""),
    linkedin_highlight: str = Form(""),
    facebook_screenshot: UploadFile = File(...),
    instagram_screenshot: UploadFile = File(...),
    linkedin_content: UploadFile | None = File(None),
    linkedin_followers: UploadFile | None = File(None),
    linkedin_visitors: UploadFile | None = File(None),
    linkedin_personal_top10: UploadFile | None = File(None),
    linkedin_personal_screenshots: list[UploadFile] | None = File(None),
):
    brand = _clean_text(brand, 80)
    created_by = _clean_text(created_by, 40)
    if not brand:
        raise HTTPException(400, "客户/品牌名称不能为空")
    if not created_by:
        raise HTTPException(400, "“你是谁”不能为空")
    report_type = _clean_text(report_type, 20).lower()
    if report_type not in REPORT_TYPE_LABELS:
        raise HTTPException(400, "报告类型仅支持月报、季度报、年报或自定义报告")
    start = _validate_date(report_start, "报告开始日期")
    end = _validate_date(report_end, "报告结束日期")
    if end < start:
        raise HTTPException(400, "报告结束日期不能早于开始日期")
    linkedin_mode = _clean_text(linkedin_mode, 20).lower()
    if linkedin_mode not in {"company", "personal", "none"}:
        raise HTTPException(400, "LinkedIn 数据来源仅支持公共主页、个人账号或本期无数据")

    def has_upload(upload: UploadFile | None) -> bool:
        return bool(upload and Path(upload.filename or "").name)

    personal_screenshots = [
        upload
        for upload in (linkedin_personal_screenshots or [])
        if has_upload(upload)
    ]
    if linkedin_mode == "company" and not all(
        has_upload(upload)
        for upload in (linkedin_content, linkedin_followers, linkedin_visitors)
    ):
        raise HTTPException(
            400, "LinkedIn 公共主页模式需要上传内容、关注者和访客三份导出"
        )
    if (
        linkedin_mode == "personal"
        and not has_upload(linkedin_personal_top10)
        and not personal_screenshots
    ):
        raise HTTPException(
            400, "LinkedIn 个人账号模式至少需要上传 Top 10 汇总表或一张截图"
        )
    linkedin_manual = {
        "period_impressions": _optional_nonnegative_int(
            linkedin_period_impressions, "LinkedIn 报告期总浏览量"
        ),
        "profile_views": _optional_nonnegative_int(
            linkedin_profile_views, "LinkedIn 主页访问量"
        ),
        "current_followers": _optional_nonnegative_int(
            linkedin_current_followers, "LinkedIn 当前关注者"
        ),
        "highlight": _clean_text(linkedin_highlight, 500),
    }

    submitted_facebook = {
        "posts": _optional_nonnegative_int(facebook_posts, "Facebook 发布内容数"),
        "views": _optional_nonnegative_int(facebook_views, "Facebook 浏览量"),
        "reach": _optional_nonnegative_int(facebook_reach, "Facebook 覆盖人数"),
        "viewers": _optional_nonnegative_int(facebook_viewers, "Facebook 浏览人数"),
        "interactions": _optional_nonnegative_int(
            facebook_interactions, "Facebook 互动次数"
        ),
        "likes": _optional_nonnegative_int(facebook_likes, "Facebook 点赞数"),
    }
    submitted_instagram = {
        "posts": _optional_nonnegative_int(instagram_posts, "Instagram 发布内容数"),
        "views": _optional_nonnegative_int(instagram_views, "Instagram 浏览量"),
        "reach": _optional_nonnegative_int(instagram_reach, "Instagram 覆盖人数"),
        "interactions": _optional_nonnegative_int(
            instagram_interactions, "Instagram 互动次数"
        ),
        "likes": _optional_nonnegative_int(instagram_likes, "Instagram 点赞数"),
    }
    clean_meta_sync_id = _clean_text(meta_sync_id, 40)
    sync = META_REPOSITORY.get_sync(clean_meta_sync_id)
    stale_sync_snapshot = bool(clean_meta_sync_id and not sync)
    if sync and (
        sync["report_start"] != report_start or sync["report_end"] != report_end
    ):
        raise HTTPException(400, "报告日期已变化，请重新同步 Meta 数据")

    used_manual = False
    if sync:
        synced = sync["result"]
        facebook_data = dict(synced["facebook"])
        instagram_data = dict(synced["instagram"])
        override_all = manual_override.strip().lower() in {"1", "true", "yes", "on"}
        for target, submitted in (
            (facebook_data, submitted_facebook),
            (instagram_data, submitted_instagram),
        ):
            for key, value in submitted.items():
                if value is None:
                    continue
                if override_all or target.get(key) is None:
                    if target.get(key) != value:
                        used_manual = True
                    target[key] = value
        meta_source = "mixed" if used_manual else "api"
        meta_info = {
            "source": meta_source,
            "sync_id": sync["id"],
            "page": synced.get("page"),
            "instagram_account": synced.get("instagram_account"),
            "missing_metrics": [
                item
                for item in synced.get("missing_metrics", [])
                if (
                    item.split(".", 1)[0] == "facebook"
                    and facebook_data.get(item.split(".", 1)[1]) is None
                )
                or (
                    item.split(".", 1)[0] == "instagram"
                    and instagram_data.get(item.split(".", 1)[1]) is None
                )
            ],
            "warnings": synced.get("warnings", []),
            "synced_at": synced.get("synced_at"),
            "graph_version": synced.get("graph_version"),
        }
        if not account_label:
            account_label = (synced.get("page") or {}).get("name") or ""
        facebook_highlight = (
            _clean_text(facebook_highlight, 300)
            or synced.get("facebook_highlight", "")
        )
        instagram_highlight = (
            _clean_text(instagram_highlight, 300)
            or synced.get("instagram_highlight", "")
        )
    else:
        facebook_data = submitted_facebook
        instagram_data = submitted_instagram
        meta_info = {
            "source": "cached_api" if stale_sync_snapshot else "manual",
            "sync_id": clean_meta_sync_id if stale_sync_snapshot else None,
            "page": ({"name": account_label} if stale_sync_snapshot and account_label else None),
            "instagram_account": None,
            "missing_metrics": [
                f"{platform}.{key}"
                for platform, values in (
                    ("facebook", facebook_data),
                    ("instagram", instagram_data),
                )
                for key, value in values.items()
                if value is None
            ],
            "warnings": [
                (
                    "原 Meta 同步记录已失效；本任务使用页面保留的最近一次 API 同步快照。"
                    if stale_sync_snapshot
                    else "本任务未使用 Meta API 自动同步数据。"
                )
            ],
            "synced_at": None,
            "graph_version": None,
        }
        facebook_highlight = _clean_text(facebook_highlight, 300)
        instagram_highlight = _clean_text(instagram_highlight, 300)

    job_id = uuid.uuid4().hex[:12]
    job_dir = JOBS_DIR / job_id
    input_dir = job_dir / "inputs"
    input_dir.mkdir(parents=True, exist_ok=True)
    try:
        files: dict[str, object] = {
            "facebook_screenshot": await _save_upload(
                facebook_screenshot,
                input_dir,
                "facebook",
                ALLOWED_IMAGE_SUFFIXES,
            ),
            "instagram_screenshot": await _save_upload(
                instagram_screenshot,
                input_dir,
                "instagram",
                ALLOWED_IMAGE_SUFFIXES,
            ),
        }
        if linkedin_mode == "company":
            files["linkedin_content"] = await _save_upload(
                linkedin_content,
                input_dir,
                "linkedin_content",
                ALLOWED_SHEET_SUFFIXES,
            )
            files["linkedin_followers"] = await _save_upload(
                linkedin_followers,
                input_dir,
                "linkedin_followers",
                ALLOWED_SHEET_SUFFIXES,
            )
            files["linkedin_visitors"] = await _save_upload(
                linkedin_visitors,
                input_dir,
                "linkedin_visitors",
                ALLOWED_SHEET_SUFFIXES,
            )
        elif linkedin_mode == "personal":
            if has_upload(linkedin_personal_top10):
                files["linkedin_personal_top10"] = await _save_upload(
                    linkedin_personal_top10,
                    input_dir,
                    "linkedin_personal_top10",
                    ALLOWED_SHEET_SUFFIXES,
                )
            files["linkedin_personal_screenshots"] = [
                await _save_upload(
                    upload,
                    input_dir,
                    f"linkedin_personal_{index}",
                    ALLOWED_IMAGE_SUFFIXES,
                )
                for index, upload in enumerate(personal_screenshots, start=1)
            ]
    except Exception:
        for path in input_dir.glob("*"):
            path.unlink(missing_ok=True)
        raise

    request = {
        "brand": brand,
        "account_label": _clean_text(account_label, 100),
        "report_type": report_type,
        "report_start": report_start,
        "report_end": report_end,
        "created_by": created_by,
        "facebook": facebook_data,
        "instagram": instagram_data,
        "facebook_highlight": facebook_highlight,
        "instagram_highlight": instagram_highlight,
        "linkedin_mode": linkedin_mode,
        "linkedin_manual": linkedin_manual,
        "meta": meta_info,
        "files": files,
    }
    (job_dir / "request.json").write_text(
        json.dumps(request, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    with _db() as conn:
        conn.execute(
            "INSERT INTO jobs(id,brand,report_type,report_start,report_end,status,created_at,created_by) "
            "VALUES(?,?,?,?,?,?,?,?)",
            (
                job_id,
                brand,
                report_type,
                report_start,
                report_end,
                "queued",
                _now_cn(),
                created_by,
            ),
        )
    _wake.put(job_id)
    return _get_job(job_id)


@app.get("/api/jobs")
def list_jobs(limit: int = 30):
    limit = min(max(limit, 1), 100)
    with _db() as conn:
        rows = conn.execute(
            "SELECT * FROM jobs ORDER BY created_at DESC LIMIT ?", (limit,)
        ).fetchall()
    return {"items": [_job_dict(row) for row in rows]}


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str):
    return _get_job(job_id)


@app.post("/api/jobs/{job_id}/retry")
def retry_job(job_id: str):
    job = _get_job(job_id)
    if job["status"] in {"queued", "running"}:
        return job
    if job["status"] != "failed":
        raise HTTPException(409, "只有生成失败的任务才能重试")
    job_dir = JOBS_DIR / job_id
    request_path = job_dir / "request.json"
    input_dir = job_dir / "inputs"
    if not request_path.is_file() or not input_dir.is_dir():
        raise HTTPException(409, "原任务文件已不存在，请重新上传后生成")
    _set_job(
        job_id,
        status="queued",
        error=None,
        started_at=None,
        finished_at=None,
        popo_status=None,
        popo_url=None,
        popo_doc_id=None,
        popo_error=None,
        popo_requested_at=None,
        popo_finished_at=None,
    )
    _wake.put(job_id)
    return _get_job(job_id)


@app.post("/api/jobs/{job_id}/popo/retry")
def retry_popo_delivery(job_id: str):
    job = _get_job(job_id)
    if job["status"] != "succeeded":
        raise HTTPException(409, "只有已成功生成的报告才能交付 POPO")
    if job.get("popo_status") == "processing":
        return job
    _set_job(
        job_id,
        popo_status="pending",
        popo_error=None,
        popo_requested_at=_now_cn(),
        popo_finished_at=None,
    )
    _popo_wake.put(job_id)
    return _get_job(job_id)


@app.get("/api/jobs/{job_id}/log")
def get_log(job_id: str):
    _get_job(job_id)
    log_path = JOBS_DIR / job_id / "run.log"
    text = (
        log_path.read_text(encoding="utf-8", errors="replace")[-50000:]
        if log_path.exists()
        else ""
    )
    return {"text": text}


@app.get("/api/jobs/{job_id}/files/{filename}")
def download(job_id: str, filename: str):
    _get_job(job_id)
    output_dir = (JOBS_DIR / job_id / "outputs").resolve()
    target = (output_dir / Path(filename).name).resolve()
    if not target.is_relative_to(output_dir) or not target.is_file():
        raise HTTPException(404, "文件不存在")
    media_type = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        if target.suffix.lower() == ".docx"
        else "application/json"
    )
    return FileResponse(target, filename=target.name, media_type=media_type)


app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
