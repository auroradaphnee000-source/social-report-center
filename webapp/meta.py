from __future__ import annotations

import hashlib
import hmac
import json
import math
import os
import secrets
import sqlite3
import time
import urllib.error
import urllib.parse
import urllib.request
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable, Iterable


UTC = timezone.utc
DEFAULT_SCOPES = (
    "pages_show_list",
    "pages_read_engagement",
    "read_insights",
    "instagram_basic",
    "instagram_manage_insights",
    "business_management",
)
METRIC_LABELS = {
    "facebook.posts": "Facebook 发布内容数",
    "facebook.views": "Facebook 浏览量",
    "facebook.reach": "Facebook 覆盖人数",
    "facebook.viewers": "Facebook 浏览人数",
    "facebook.interactions": "Facebook 互动次数",
    "facebook.likes": "Facebook 点赞数",
    "instagram.posts": "Instagram 发布内容数",
    "instagram.views": "Instagram 浏览量",
    "instagram.reach": "Instagram 覆盖人数",
    "instagram.interactions": "Instagram 互动次数",
    "instagram.likes": "Instagram 点赞数",
}


def _env(name: str, default: str = "") -> str:
    return (os.getenv(name, default) or "").strip()


@dataclass(frozen=True)
class MetaSettings:
    app_id: str
    app_secret: str
    redirect_uri: str
    token_key: str
    graph_version: str
    login_config_id: str
    scopes: tuple[str, ...]

    @classmethod
    def from_env(cls) -> "MetaSettings":
        scopes = tuple(
            item.strip()
            for item in _env("META_SCOPES", ",".join(DEFAULT_SCOPES)).split(",")
            if item.strip()
        )
        return cls(
            app_id=_env("META_APP_ID"),
            app_secret=_env("META_APP_SECRET"),
            redirect_uri=_env(
                "META_REDIRECT_URI",
                "",
            ),
            token_key=_env("META_TOKEN_ENCRYPTION_KEY"),
            graph_version=_env("META_GRAPH_VERSION", "v25.0"),
            login_config_id=_env("META_LOGIN_CONFIG_ID"),
            scopes=scopes,
        )

    @property
    def configured(self) -> bool:
        return bool(
            self.app_id and self.app_secret and self.redirect_uri and self.token_key
        )

    @property
    def missing_config(self) -> list[str]:
        missing = []
        for label, value in (
            ("META_APP_ID", self.app_id),
            ("META_APP_SECRET", self.app_secret),
            ("META_REDIRECT_URI", self.redirect_uri),
            ("META_TOKEN_ENCRYPTION_KEY", self.token_key),
        ):
            if not value:
                missing.append(label)
        return missing


class MetaError(RuntimeError):
    def __init__(
        self,
        message: str,
        *,
        code: int | None = None,
        subcode: int | None = None,
        reconnect: bool = False,
    ):
        super().__init__(message)
        self.code = code
        self.subcode = subcode
        self.reconnect = reconnect


class MetaSyncTimeout(MetaError):
    """A task deadline must escape optional-metric and highlight fallbacks."""


# Context-local: concurrent OAuth/status calls and later syncs share no deadline.
_SYNC_DEADLINE: ContextVar[tuple[float, float] | None] = ContextVar(
    "meta_sync_deadline", default=None
)


def _sync_seconds_remaining() -> float | None:
    budget = _SYNC_DEADLINE.get()
    if budget is None:
        return None
    deadline, limit = budget
    remaining = deadline - time.monotonic()
    if remaining <= 0:
        raise MetaSyncTimeout(
            f"Meta 数据同步超过总时限（{limit:g} 秒），已停止读取。"
            "请稍后重试，或联系管理员调整同步时限。"
        )
    return remaining


class TokenVault:
    """Encrypt Meta tokens at rest. The encryption key only lives in service env."""

    def __init__(self, key: str):
        if not key:
            raise MetaError("Meta token 加密密钥尚未配置")
        try:
            from cryptography.fernet import Fernet

            self._fernet = Fernet(key.encode("ascii"))
        except ImportError as exc:
            raise MetaError("服务器缺少 token 加密组件") from exc
        except Exception as exc:
            raise MetaError("META_TOKEN_ENCRYPTION_KEY 格式无效") from exc

    def encrypt(self, value: str) -> str:
        return self._fernet.encrypt(value.encode("utf-8")).decode("ascii")

    def decrypt(self, value: str) -> str:
        try:
            return self._fernet.decrypt(value.encode("ascii")).decode("utf-8")
        except Exception as exc:
            raise MetaError("已保存的 Meta 授权无法解密，请重新连接", reconnect=True) from exc


class MetaRepository:
    def __init__(self, db_path: Path, settings: MetaSettings):
        self.db_path = db_path
        self.settings = settings

    @contextmanager
    def _db(self):
        conn = sqlite3.connect(self.db_path, timeout=30)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _vault(self) -> TokenVault:
        return TokenVault(self.settings.token_key)

    def init_db(self):
        with self._db() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS meta_oauth_states (
                    state_hash TEXT PRIMARY KEY,
                    created_at INTEGER NOT NULL
                );
                CREATE TABLE IF NOT EXISTS meta_connection (
                    singleton INTEGER PRIMARY KEY CHECK(singleton=1),
                    user_id TEXT NOT NULL,
                    user_name TEXT NOT NULL,
                    token_encrypted TEXT NOT NULL,
                    token_expires_at INTEGER,
                    scopes_json TEXT NOT NULL,
                    connected_at TEXT NOT NULL,
                    last_verified_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS meta_assets (
                    page_id TEXT PRIMARY KEY,
                    page_name TEXT NOT NULL,
                    page_token_encrypted TEXT NOT NULL,
                    instagram_id TEXT,
                    instagram_username TEXT,
                    tasks_json TEXT NOT NULL,
                    refreshed_at TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS meta_syncs (
                    id TEXT PRIMARY KEY,
                    page_id TEXT NOT NULL,
                    instagram_id TEXT,
                    report_start TEXT NOT NULL,
                    report_end TEXT NOT NULL,
                    result_json TEXT NOT NULL,
                    created_at TEXT NOT NULL
                );
                """
            )

    def create_state(self) -> str:
        raw = secrets.token_urlsafe(32)
        digest = hashlib.sha256(raw.encode("utf-8")).hexdigest()
        now = int(datetime.now(UTC).timestamp())
        with self._db() as conn:
            conn.execute(
                "DELETE FROM meta_oauth_states WHERE created_at < ?", (now - 900,)
            )
            conn.execute(
                "INSERT INTO meta_oauth_states(state_hash,created_at) VALUES(?,?)",
                (digest, now),
            )
        return raw

    def consume_state(self, raw: str) -> bool:
        digest = hashlib.sha256((raw or "").encode("utf-8")).hexdigest()
        now = int(datetime.now(UTC).timestamp())
        with self._db() as conn:
            row = conn.execute(
                "SELECT created_at FROM meta_oauth_states WHERE state_hash=?",
                (digest,),
            ).fetchone()
            conn.execute(
                "DELETE FROM meta_oauth_states WHERE state_hash=?", (digest,)
            )
        return bool(row and int(row["created_at"]) >= now - 900)

    def save_connection(
        self,
        *,
        user_id: str,
        user_name: str,
        token: str,
        expires_in: int | None,
        scopes: Iterable[str],
    ):
        encrypted = self._vault().encrypt(token)
        expires_at = (
            int(datetime.now(UTC).timestamp()) + int(expires_in)
            if expires_in
            else None
        )
        now = datetime.now(UTC).isoformat(timespec="seconds")
        with self._db() as conn:
            conn.execute(
                """
                INSERT INTO meta_connection(
                    singleton,user_id,user_name,token_encrypted,token_expires_at,
                    scopes_json,connected_at,last_verified_at
                ) VALUES(1,?,?,?,?,?,?,?)
                ON CONFLICT(singleton) DO UPDATE SET
                    user_id=excluded.user_id,
                    user_name=excluded.user_name,
                    token_encrypted=excluded.token_encrypted,
                    token_expires_at=excluded.token_expires_at,
                    scopes_json=excluded.scopes_json,
                    connected_at=excluded.connected_at,
                    last_verified_at=excluded.last_verified_at
                """,
                (
                    user_id,
                    user_name,
                    encrypted,
                    expires_at,
                    json.dumps(sorted(set(scopes)), ensure_ascii=False),
                    now,
                    now,
                ),
            )
            conn.execute("DELETE FROM meta_assets")

    def get_connection(self, include_token: bool = False) -> dict[str, Any] | None:
        with self._db() as conn:
            row = conn.execute(
                "SELECT * FROM meta_connection WHERE singleton=1"
            ).fetchone()
        if not row:
            return None
        result = dict(row)
        result["scopes"] = json.loads(result.pop("scopes_json") or "[]")
        encrypted = result.pop("token_encrypted")
        if include_token:
            result["token"] = self._vault().decrypt(encrypted)
        return result

    def touch_connection(self):
        with self._db() as conn:
            conn.execute(
                "UPDATE meta_connection SET last_verified_at=? WHERE singleton=1",
                (datetime.now(UTC).isoformat(timespec="seconds"),),
            )

    def disconnect(self):
        with self._db() as conn:
            conn.execute("DELETE FROM meta_assets")
            conn.execute("DELETE FROM meta_connection")

    def save_assets(self, pages: list[dict[str, Any]]):
        vault = self._vault()
        now = datetime.now(UTC).isoformat(timespec="seconds")
        with self._db() as conn:
            conn.execute("DELETE FROM meta_assets")
            for page in pages:
                ig = page.get("instagram") or {}
                conn.execute(
                    """
                    INSERT INTO meta_assets(
                        page_id,page_name,page_token_encrypted,instagram_id,
                        instagram_username,tasks_json,refreshed_at
                    ) VALUES(?,?,?,?,?,?,?)
                    """,
                    (
                        page["id"],
                        page["name"],
                        vault.encrypt(page["access_token"]),
                        ig.get("id"),
                        ig.get("username") or ig.get("name"),
                        json.dumps(page.get("tasks") or []),
                        now,
                    ),
                )

    def public_assets(self) -> list[dict[str, Any]]:
        with self._db() as conn:
            rows = conn.execute(
                """
                SELECT page_id,page_name,instagram_id,instagram_username,tasks_json,
                       refreshed_at
                FROM meta_assets ORDER BY page_name COLLATE NOCASE
                """
            ).fetchall()
        return [
            {
                "page_id": row["page_id"],
                "page_name": row["page_name"],
                "instagram_id": row["instagram_id"],
                "instagram_username": row["instagram_username"],
                "tasks": json.loads(row["tasks_json"] or "[]"),
                "refreshed_at": row["refreshed_at"],
            }
            for row in rows
        ]

    def get_asset(self, page_id: str) -> dict[str, Any]:
        with self._db() as conn:
            row = conn.execute(
                "SELECT * FROM meta_assets WHERE page_id=?", (page_id,)
            ).fetchone()
        if not row:
            raise MetaError("所选 Facebook 页面不在当前授权账号中，请刷新账号列表")
        result = dict(row)
        result["page_token"] = self._vault().decrypt(
            result.pop("page_token_encrypted")
        )
        result["tasks"] = json.loads(result.pop("tasks_json") or "[]")
        return result

    def save_sync(
        self,
        *,
        page_id: str,
        instagram_id: str | None,
        report_start: str,
        report_end: str,
        result: dict[str, Any],
    ) -> str:
        sync_id = secrets.token_hex(6)
        with self._db() as conn:
            conn.execute(
                """
                INSERT INTO meta_syncs(
                    id,page_id,instagram_id,report_start,report_end,result_json,created_at
                ) VALUES(?,?,?,?,?,?,?)
                """,
                (
                    sync_id,
                    page_id,
                    instagram_id,
                    report_start,
                    report_end,
                    json.dumps(result, ensure_ascii=False),
                    datetime.now(UTC).isoformat(timespec="seconds"),
                ),
            )
        return sync_id

    def get_sync(self, sync_id: str) -> dict[str, Any] | None:
        if not sync_id:
            return None
        with self._db() as conn:
            row = conn.execute(
                "SELECT * FROM meta_syncs WHERE id=?", (sync_id,)
            ).fetchone()
        if not row:
            return None
        result = dict(row)
        result["result"] = json.loads(result.pop("result_json"))
        return result


class MetaGraphClient:
    ALLOWED_HOSTS = {"graph.facebook.com"}

    def __init__(
        self,
        settings: MetaSettings,
        timeout: int = 35,
        *,
        sync_timeout: float | None = None,
    ):
        self.settings = settings
        self.timeout = timeout
        self.sync_timeout = float(
            sync_timeout
            if sync_timeout is not None
            else _env("META_SYNC_TIMEOUT_SECONDS", "1200")
        )
        if not math.isfinite(self.sync_timeout) or self.sync_timeout <= 0:
            raise ValueError("META_SYNC_TIMEOUT_SECONDS 必须是大于 0 的秒数")

    def authorization_url(self, state: str) -> str:
        params = {
            "client_id": self.settings.app_id,
            "redirect_uri": self.settings.redirect_uri,
            "state": state,
            "response_type": "code",
            "auth_type": "rerequest",
        }
        if self.settings.login_config_id:
            params["config_id"] = self.settings.login_config_id
        else:
            params["scope"] = ",".join(self.settings.scopes)
        return (
            f"https://www.facebook.com/{self.settings.graph_version}/dialog/oauth?"
            + urllib.parse.urlencode(params)
        )

    def _request_json(self, url: str) -> dict[str, Any]:
        remaining = _sync_seconds_remaining()
        request_timeout = (
            min(self.timeout, remaining) if remaining is not None else self.timeout
        )
        parsed = urllib.parse.urlparse(url)
        if parsed.scheme != "https" or parsed.hostname not in self.ALLOWED_HOSTS:
            raise MetaError("Meta API 地址校验失败")
        request = urllib.request.Request(
            url,
            headers={
                "Accept": "application/json",
                "User-Agent": "SocialReportCenter/1.2",
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=request_timeout) as response:
                payload = response.read(8 * 1024 * 1024)
        except urllib.error.HTTPError as exc:
            payload = exc.read(1024 * 1024)
            try:
                body = json.loads(payload.decode("utf-8", "replace"))
            except Exception:
                body = {}
            error = body.get("error") or {}
            code = error.get("code")
            subcode = error.get("error_subcode")
            message = error.get("message") or f"Meta API HTTP {exc.code}"
            raise MetaError(
                f"Meta API：{message}",
                code=code,
                subcode=subcode,
                reconnect=code == 190,
            ) from exc
        except (urllib.error.URLError, TimeoutError) as exc:
            _sync_seconds_remaining()
            raise MetaError("连接 Meta API 超时或网络不可用") from exc
        _sync_seconds_remaining()
        try:
            result = json.loads(payload.decode("utf-8"))
        except Exception as exc:
            raise MetaError("Meta API 返回了无法解析的数据") from exc
        if "error" in result:
            error = result["error"] or {}
            code = error.get("code")
            raise MetaError(
                f"Meta API：{error.get('message') or '请求失败'}",
                code=code,
                subcode=error.get("error_subcode"),
                reconnect=code == 190,
            )
        return result

    def _graph(
        self,
        path: str,
        *,
        token: str | None = None,
        params: dict[str, Any] | None = None,
        versioned: bool = True,
    ) -> dict[str, Any]:
        query = dict(params or {})
        if token:
            query["access_token"] = token
            query["appsecret_proof"] = hmac.new(
                self.settings.app_secret.encode("utf-8"),
                token.encode("utf-8"),
                hashlib.sha256,
            ).hexdigest()
        prefix = f"/{self.settings.graph_version}" if versioned else ""
        url = (
            "https://graph.facebook.com"
            + prefix
            + "/"
            + path.lstrip("/")
            + "?"
            + urllib.parse.urlencode(query)
        )
        return self._request_json(url)

    def exchange_code(self, code: str) -> dict[str, Any]:
        short = self._graph(
            "oauth/access_token",
            params={
                "client_id": self.settings.app_id,
                "client_secret": self.settings.app_secret,
                "redirect_uri": self.settings.redirect_uri,
                "code": code,
            },
            versioned=False,
        )
        token = short.get("access_token")
        if not token:
            raise MetaError("Meta 授权未返回 access token")
        try:
            long_lived = self._graph(
                "oauth/access_token",
                params={
                    "grant_type": "fb_exchange_token",
                    "client_id": self.settings.app_id,
                    "client_secret": self.settings.app_secret,
                    "fb_exchange_token": token,
                },
                versioned=False,
            )
            if long_lived.get("access_token"):
                return long_lived
        except MetaError:
            pass
        return short

    def get_profile(self, token: str) -> dict[str, Any]:
        return self._graph("me", token=token, params={"fields": "id,name"})

    @staticmethod
    def _page_fields() -> str:
        return (
            "id,name,access_token,tasks,"
            "instagram_business_account{id,username,name}"
        )

    def _normalize_page(
        self, item: dict[str, Any], *, user_token: str
    ) -> dict[str, Any] | None:
        page_id = item.get("id")
        if not page_id:
            return None
        page_token = item.get("access_token")
        if not page_token:
            try:
                detail = self._graph(
                    str(page_id),
                    token=user_token,
                    params={"fields": self._page_fields()},
                )
            except MetaError as exc:
                if exc.reconnect:
                    raise
                return None
            item = {**item, **detail}
            page_token = item.get("access_token")
        if not page_token:
            return None
        ig = item.get("instagram_business_account")
        if not ig:
            try:
                detail = self._graph(
                    str(page_id),
                    token=page_token,
                    params={
                        "fields": (
                            "instagram_business_account{id,username,name}"
                        )
                    },
                )
                ig = detail.get("instagram_business_account")
            except MetaError:
                ig = None
        return {
            "id": str(page_id),
            "name": item.get("name") or str(page_id),
            "access_token": page_token,
            "tasks": item.get("tasks") or [],
            "instagram": ig,
        }

    @staticmethod
    def _merge_page(
        pages_by_id: dict[str, dict[str, Any]], page: dict[str, Any]
    ) -> None:
        existing = pages_by_id.get(page["id"])
        if not existing:
            pages_by_id[page["id"]] = page
            return
        if not existing.get("instagram") and page.get("instagram"):
            existing["instagram"] = page["instagram"]
        existing["tasks"] = sorted(
            set(existing.get("tasks") or []) | set(page.get("tasks") or [])
        )

    def _business_pages(self, token: str) -> list[dict[str, Any]]:
        try:
            businesses = self._paginate(
                "me/businesses",
                token=token,
                params={"fields": "id,name", "limit": 100},
            )
        except MetaError as exc:
            if exc.reconnect:
                raise
            return []

        pages: list[dict[str, Any]] = []
        for business in businesses:
            business_id = business.get("id")
            if not business_id:
                continue
            for edge in ("owned_pages", "client_pages"):
                try:
                    items = self._paginate(
                        f"{business_id}/{edge}",
                        token=token,
                        params={"fields": self._page_fields(), "limit": 100},
                    )
                except MetaError as exc:
                    if exc.reconnect:
                        raise
                    continue
                for item in items:
                    page = self._normalize_page(item, user_token=token)
                    if page:
                        pages.append(page)
        return pages

    def list_pages(self, token: str) -> list[dict[str, Any]]:
        pages_by_id: dict[str, dict[str, Any]] = {}
        direct_pages = self._paginate(
            "me/accounts",
            token=token,
            params={"fields": self._page_fields(), "limit": 100},
        )
        for item in direct_pages:
            page = self._normalize_page(item, user_token=token)
            if page:
                self._merge_page(pages_by_id, page)
        for page in self._business_pages(token):
            self._merge_page(pages_by_id, page)
        return sorted(
            pages_by_id.values(),
            key=lambda page: (page["name"].casefold(), page["id"]),
        )

    def _paginate(
        self, path: str, *, token: str, params: dict[str, Any]
    ) -> list[dict[str, Any]]:
        items: list[dict[str, Any]] = []
        result = self._graph(path, token=token, params=params)
        while True:
            items.extend(result.get("data") or [])
            next_url = (result.get("paging") or {}).get("next")
            if not next_url:
                return items
            result = self._request_json(next_url)

    def _insight(
        self,
        object_id: str,
        metric: str,
        *,
        token: str,
        since: str,
        until: str,
        period: str = "day",
        breakdown: str | None = None,
        metric_type: str | None = "total_value",
    ) -> dict[str, Any]:
        base = {
            "metric": metric,
            "period": period,
            "since": since,
            "until": until,
        }
        if metric_type:
            base["metric_type"] = metric_type
        if not breakdown:
            return self._graph(
                f"{object_id}/insights", token=token, params=base
            )
        last_error = None
        for key in ("breakdown", "metric_breakdowns"):
            try:
                params = dict(base)
                params[key] = breakdown
                return self._graph(
                    f"{object_id}/insights", token=token, params=params
                )
            except MetaError as exc:
                if exc.reconnect or isinstance(exc, MetaSyncTimeout):
                    raise
                last_error = exc
        raise last_error or MetaError(f"{metric} breakdown 请求失败")

    def _safe_insight(
        self,
        object_id: str,
        metric: str,
        *,
        token: str,
        since: str,
        until: str,
        period: str = "day",
        breakdown: str | None = None,
        metric_type: str | None = "total_value",
        progress_callback: Callable[[int, int, str], None] | None = None,
    ) -> tuple[dict[str, Any] | None, str | None]:
        try:
            start = date.fromisoformat(since)
            end = date.fromisoformat(until)
        except ValueError:
            windows = [(since, until)]
        else:
            windows: list[tuple[str, str]] = []
            cursor = start
            # Instagram account insights reject a since/until span above
            # 30 days. Bounded Page-insight requests are more reliable too.
            while cursor < end:
                window_end = min(cursor + timedelta(days=30), end)
                windows.append((cursor.isoformat(), window_end.isoformat()))
                cursor = window_end

        combined: list[dict[str, Any]] = []
        errors: list[str] = []
        for index, (window_since, window_until) in enumerate(windows, 1):
            _sync_seconds_remaining()
            try:
                display_end = (
                    date.fromisoformat(window_until) - timedelta(days=1)
                ).isoformat()
            except ValueError:
                display_end = window_until
            detail = f"第 {index}/{len(windows)} 段 · {window_since}—{display_end}"
            if progress_callback:
                progress_callback(index - 1, len(windows), detail + " · 读取中")
            outcome = "已读取"
            try:
                payload = self._insight(
                    object_id,
                    metric,
                    token=token,
                    since=window_since,
                    until=window_until,
                    period=period,
                    breakdown=breakdown,
                    metric_type=metric_type,
                )
                _sync_seconds_remaining()
                combined.extend(payload.get("data") or [])
            except MetaError as exc:
                if exc.reconnect or isinstance(exc, MetaSyncTimeout):
                    raise
                errors.append(f"{window_since}—{window_until}: {exc}")
                outcome = "未取得"
            if progress_callback:
                progress_callback(index, len(windows), detail + " · " + outcome)

        if not combined:
            return None, errors[0] if errors else None
        warning = f"部分日期区间读取失败（{errors[0]}）" if errors else None
        return {"data": combined}, warning

    def sync(
        self,
        asset: dict[str, Any],
        *,
        instagram_id: str | None,
        report_start: str,
        report_end: str,
        progress_callback: Callable[[int, str], None] | None = None,
    ) -> dict[str, Any]:
        context_token = _SYNC_DEADLINE.set(
            (time.monotonic() + self.sync_timeout, self.sync_timeout)
        )
        try:
            result = self._sync(
                asset,
                instagram_id=instagram_id,
                report_start=report_start,
                report_end=report_end,
                progress_callback=progress_callback,
            )
            _sync_seconds_remaining()
            return result
        finally:
            _SYNC_DEADLINE.reset(context_token)

    def _sync(
        self,
        asset: dict[str, Any],
        *,
        instagram_id: str | None,
        report_start: str,
        report_end: str,
        progress_callback: Callable[[int, str], None] | None = None,
    ) -> dict[str, Any]:
        def progress(value: int, message: str):
            _sync_seconds_remaining()
            if progress_callback:
                progress_callback(value, message)

        def read_insight(
            object_id: str,
            metric: str,
            *,
            label: str,
            progress_start: int,
            progress_end: int,
            **kwargs,
        ):
            def window_progress(completed: int, total: int, detail: str):
                value = progress_start + (
                    (progress_end - progress_start) * completed // total
                )
                progress(value, f"正在读取 {label}（{detail}）")

            return self._safe_insight(
                object_id, metric, progress_callback=window_progress, **kwargs
            )

        start = date.fromisoformat(report_start)
        end = date.fromisoformat(report_end)
        # Meta's "until" is exclusive on list and insights endpoints.
        until = (end + timedelta(days=1)).isoformat()
        page_id = asset["page_id"]
        token = asset["page_token"]
        warnings: list[str] = []

        progress(15, "正在读取 Facebook 内容")
        page_posts = self._paginate(
            f"{page_id}/posts",
            token=token,
            params={
                "fields": "id,message,created_time,permalink_url",
                "since": report_start,
                "until": until,
                "limit": 100,
            },
        )
        page_posts = [
            item
            for item in page_posts
            if _date_in_range(item.get("created_time"), start, end)
        ]
        fb_monthly = _publication_monthly(
            page_posts, "created_time", start, end
        )

        progress(30, "正在读取 Facebook 概览指标")
        fb_values: dict[str, int | None] = {
            "posts": len(page_posts),
            "views": None,
            "reach": None,
            "viewers": None,
            "interactions": None,
            "likes": None,
        }

        total_views_raw, error = read_insight(
            page_id,
            "page_media_view",
            label="Facebook 浏览量",
            progress_start=30,
            progress_end=34,
            token=token,
            since=report_start,
            until=until,
            metric_type=None,
        )
        if error:
            warnings.append(f"Facebook 浏览量：{error}")
        total_views = _metric_total(total_views_raw, additive=True)
        breakdown_raw, breakdown_error = read_insight(
            page_id,
            "page_media_view",
            label="Facebook 自然/付费浏览拆分",
            progress_start=34,
            progress_end=38,
            token=token,
            since=report_start,
            until=until,
            breakdown="is_from_ads",
            metric_type=None,
        )
        organic_views = _breakdown_total(
            breakdown_raw, "is_from_ads", {"0", "false", "organic", "no"}
        )
        paid_views = _breakdown_total(
            breakdown_raw, "is_from_ads", {"1", "true", "paid", "yes"}
        )
        if organic_views is not None:
            fb_values["views"] = organic_views
        elif total_views is not None and paid_views == 0:
            fb_values["views"] = total_views
        elif total_views is not None:
            fb_values["views"] = total_views
            warnings.append(
                "Facebook 已取得总浏览量，但 API 未返回自然/付费拆分；请在手动入口补充自然浏览量。"
            )
        if fb_values["views"] is not None:
            # Same daily response powers both the overview and monthly series.
            _merge_monthly_metric(
                fb_monthly, "views", _metric_monthly(total_views_raw)
            )

        viewers_raw, viewers_error = read_insight(
            page_id,
            "page_total_media_view_unique",
            label="Facebook 浏览人数与日度触达",
            progress_start=38,
            progress_end=42,
            token=token,
            since=report_start,
            until=until,
            metric_type=None,
        )
        viewers = _metric_total(viewers_raw, additive=False)
        if viewers is not None:
            fb_values["viewers"] = viewers
        elif viewers_error:
            warnings.append(f"Facebook 浏览人数：{viewers_error}")
        if paid_views in (None, 0):
            _merge_monthly_metric(
                fb_monthly,
                "daily_reach_sum",
                _metric_monthly(viewers_raw),
            )
            if viewers_error:
                warnings.append(
                    "Facebook 月度日度触达累计未返回："
                    f"{viewers_error}"
                )

        for key, metric, label, progress_start in (
            ("interactions", "page_post_engagements", "互动次数", 42),
            ("likes", "page_actions_post_reactions_like_total", "点赞数", 46),
        ):
            raw, metric_error = read_insight(
                page_id,
                metric,
                label=f"Facebook {label}",
                progress_start=progress_start,
                progress_end=progress_start + 4,
                token=token,
                since=report_start,
                until=until,
                metric_type=None,
            )
            value = _metric_total(raw, additive=True)
            if value is not None and paid_views in (None, 0):
                fb_values[key] = value
            elif metric_error:
                warnings.append(f"Facebook {key}：{metric_error}")
            if value is not None and paid_views in (None, 0):
                _merge_monthly_metric(
                    fb_monthly, key, _metric_monthly(raw)
                )
                if metric_error:
                    warnings.append(
                        f"Facebook 月度{key}趋势未返回：{metric_error}"
                    )

        fb_values["monthly"] = fb_monthly

        progress(50, "正在分析 Facebook 高表现内容")
        fb_highlight = self._facebook_highlight(page_posts, token)

        ig_values: dict[str, int | None] = {
            "posts": None,
            "views": None,
            "reach": None,
            "interactions": None,
            "likes": None,
        }
        ig_monthly = _calendar_month_rows(start, end)
        ig_highlight = ""
        if not instagram_id:
            warnings.append("所选 Facebook 页面没有关联 Instagram 专业账号。")
            progress(85, "Facebook 数据已读取，正在整理结果")
        elif instagram_id != asset.get("instagram_id"):
            raise MetaError("Instagram 账号与所选 Facebook 页面不匹配")
        else:
            progress(60, "正在读取 Instagram 内容")
            ig_media = self._paginate(
                f"{instagram_id}/media",
                token=token,
                params={
                    "fields": (
                        "id,caption,media_type,media_product_type,timestamp,"
                        "permalink,like_count,comments_count"
                    ),
                    "since": report_start,
                    "until": until,
                    "limit": 100,
                },
            )
            ig_media = [
                item
                for item in ig_media
                if _date_in_range(item.get("timestamp"), start, end)
            ]
            ig_monthly = _publication_monthly(
                ig_media, "timestamp", start, end
            )
            ig_values["posts"] = len(ig_media)
            progress(72, "正在读取 Instagram 概览指标")
            for key, metric, additive, label, progress_start in (
                ("views", "views", True, "浏览量", 72),
                ("reach", "reach", False, "覆盖人数", 76),
                ("interactions", "total_interactions", True, "互动次数", 80),
                ("likes", "likes", True, "点赞数", 84),
            ):
                raw, metric_error = read_insight(
                    instagram_id,
                    metric,
                    label=f"Instagram {label}",
                    progress_start=progress_start,
                    progress_end=progress_start + 2,
                    token=token,
                    since=report_start,
                    until=until,
                )
                value = _metric_total(raw, additive=additive)
                if value is not None:
                    ig_values[key] = value
                elif metric_error:
                    warnings.append(f"Instagram {key}：{metric_error}")
                if additive or key == "reach":
                    daily_raw, daily_error = read_insight(
                        instagram_id,
                        metric,
                        label=f"Instagram {label}日度趋势",
                        progress_start=progress_start + 2,
                        progress_end=progress_start + 4,
                        token=token,
                        since=report_start,
                        until=until,
                        metric_type=None,
                    )
                    _merge_monthly_metric(
                        ig_monthly,
                        "daily_reach_sum" if key == "reach" else key,
                        _metric_monthly(daily_raw),
                    )
                    if daily_error:
                        warnings.append(
                            f"Instagram 月度{key}趋势未返回：{daily_error}"
                        )
            progress(88, "正在分析 Instagram 高表现内容")
            ig_highlight = self._instagram_highlight(ig_media, token)

        ig_values["monthly"] = ig_monthly

        progress(92, "正在整理 Meta 同步结果")
        missing = _missing_metrics(fb_values, ig_values)
        if (end - start).days >= 90:
            warnings.append(
                "报告周期超过90天；Instagram账号级指标受 Meta 历史数据窗口限制，"
                "缺失项请使用手动入口补充。"
            )
        return {
            "source": "meta_graph_api",
            "graph_version": self.settings.graph_version,
            "page": {
                "id": page_id,
                "name": asset["page_name"],
            },
            "instagram_account": {
                "id": instagram_id,
                "username": asset.get("instagram_username"),
            }
            if instagram_id
            else None,
            "report_start": report_start,
            "report_end": report_end,
            "facebook": fb_values,
            "instagram": ig_values,
            "facebook_highlight": fb_highlight,
            "instagram_highlight": ig_highlight,
            "missing_metrics": missing,
            "missing_labels": [METRIC_LABELS[item] for item in missing],
            "warnings": _dedupe(warnings),
            "complete": not missing,
            "synced_at": datetime.now(UTC).isoformat(timespec="seconds"),
        }

    def _facebook_highlight(
        self, posts: list[dict[str, Any]], token: str
    ) -> str:
        ranked: list[tuple[int, str]] = []
        for post in posts[:25]:
            message = _clean_highlight(post.get("message"))
            if not message:
                continue
            try:
                raw = self._graph(
                    f"{post['id']}/insights",
                    token=token,
                    params={"metric": "post_media_view", "period": "lifetime"},
                )
                views = _metric_total(raw, additive=True) or 0
            except MetaError as exc:
                if exc.reconnect or isinstance(exc, MetaSyncTimeout):
                    raise
                views = 0
            ranked.append((views, message))
        return max(ranked, default=(0, ""), key=lambda item: item[0])[1]

    def _instagram_highlight(
        self, media: list[dict[str, Any]], token: str
    ) -> str:
        ranked: list[tuple[int, str]] = []
        for item in media[:25]:
            caption = _clean_highlight(item.get("caption"))
            if not caption:
                continue
            try:
                raw = self._graph(
                    f"{item['id']}/insights",
                    token=token,
                    params={"metric": "views", "period": "lifetime"},
                )
                views = _metric_total(raw, additive=True) or 0
            except MetaError as exc:
                if exc.reconnect or isinstance(exc, MetaSyncTimeout):
                    raise
                views = int(item.get("like_count") or 0)
            ranked.append((views, caption))
        return max(ranked, default=(0, ""), key=lambda item: item[0])[1]


def _date_in_range(value: Any, start: date, end: date) -> bool:
    if not value:
        return False
    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00")).date()
    except ValueError:
        return False
    return start <= parsed <= end


def _calendar_month_rows(start: date, end: date) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    cursor = date(start.year, start.month, 1)
    last = date(end.year, end.month, 1)
    while cursor <= last:
        rows.append(
            {
                "month": cursor.strftime("%Y-%m"),
                "posts": 0,
                "views": None,
                "interactions": None,
                "likes": None,
            }
        )
        cursor = (
            date(cursor.year + 1, 1, 1)
            if cursor.month == 12
            else date(cursor.year, cursor.month + 1, 1)
        )
    return rows


def _publication_monthly(
    items: list[dict[str, Any]], date_field: str, start: date, end: date
) -> list[dict[str, Any]]:
    rows = _calendar_month_rows(start, end)
    by_month = {row["month"]: row for row in rows}
    for item in items:
        value = item.get(date_field)
        if not value:
            continue
        try:
            item_date = datetime.fromisoformat(
                str(value).replace("Z", "+00:00")
            ).date()
        except ValueError:
            continue
        month = item_date.strftime("%Y-%m")
        if month in by_month:
            by_month[month]["posts"] += 1
    return rows


def _metric_monthly(payload: dict[str, Any] | None) -> dict[str, int]:
    """Collapse a Meta daily-insight response into calendar-month totals."""
    monthly: dict[str, int] = {}
    if not payload:
        return monthly
    for item in payload.get("data") or []:
        for point in item.get("values") or []:
            if not isinstance(point, dict):
                continue
            raw_date = point.get("end_time") or point.get("date")
            value = _scalar_from(point.get("value"))
            if not raw_date or value is None:
                continue
            try:
                point_date = datetime.fromisoformat(
                    str(raw_date).replace("Z", "+00:00")
                ).date()
            except ValueError:
                continue
            # Meta's daily end_time is the exclusive end of the interval.
            if point.get("end_time"):
                point_date -= timedelta(days=1)
            month = point_date.strftime("%Y-%m")
            monthly[month] = monthly.get(month, 0) + value
    return monthly


def _merge_monthly_metric(
    rows: list[dict[str, Any]], key: str, values: dict[str, int]
) -> None:
    for row in rows:
        if row["month"] in values:
            row[key] = values[row["month"]]


def _number(value: Any) -> int | None:
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return int(round(value))
    if isinstance(value, str):
        try:
            return int(round(float(value.replace(",", ""))))
        except ValueError:
            return None
    return None


def _scalar_from(value: Any) -> int | None:
    direct = _number(value)
    if direct is not None:
        return direct
    if isinstance(value, dict):
        if "value" in value:
            direct = _number(value["value"])
            if direct is not None:
                return direct
        numeric = [_number(item) for item in value.values()]
        numeric = [item for item in numeric if item is not None]
        if numeric:
            return sum(numeric)
    return None


def _metric_total(payload: dict[str, Any] | None, *, additive: bool) -> int | None:
    if not payload:
        return None
    items = payload.get("data") or []
    totals: list[int] = []
    daily: list[int] = []
    for item in items:
        total_value = item.get("total_value")
        value = _scalar_from(total_value)
        if value is not None:
            totals.append(value)
            continue
        values = item.get("values") or []
        for point in values:
            value = _scalar_from(point.get("value") if isinstance(point, dict) else point)
            if value is not None:
                daily.append(value)
    if totals:
        if additive:
            return sum(totals)
        # Unique values from separate windows cannot be safely summed.
        return totals[0] if len(totals) == 1 else None
    if additive and daily:
        return sum(daily)
    if not additive and len(daily) == 1:
        return daily[0]
    return None


def _breakdown_total(
    payload: dict[str, Any] | None, dimension: str, desired: set[str]
) -> int | None:
    if not payload:
        return None
    desired = {item.lower() for item in desired}
    found: list[int] = []

    def walk(value: Any):
        if isinstance(value, dict):
            keys = value.get("dimension_keys")
            results = value.get("results")
            if isinstance(keys, list) and dimension in keys and isinstance(results, list):
                index = keys.index(dimension)
                for result in results:
                    dimension_values = result.get("dimension_values") or []
                    if index < len(dimension_values) and str(
                        dimension_values[index]
                    ).lower() in desired:
                        number = _scalar_from(result.get("value"))
                        if number is not None:
                            found.append(number)
            if dimension in value and isinstance(value[dimension], dict):
                for key, child in value[dimension].items():
                    if str(key).lower() in desired:
                        number = _scalar_from(child)
                        if number is not None:
                            found.append(number)
            if (
                str(value.get("dimension") or value.get("breakdown") or "").lower()
                == dimension.lower()
                and str(
                    value.get("dimension_value")
                    or value.get("breakdown_value")
                    or value.get("key")
                    or ""
                ).lower()
                in desired
            ):
                number = _scalar_from(value.get("value"))
                if number is not None:
                    found.append(number)
            for child in value.values():
                walk(child)
        elif isinstance(value, list):
            for child in value:
                walk(child)

    walk(payload)
    return sum(found) if found else None


def _clean_highlight(value: Any) -> str:
    text = " ".join(str(value or "").replace("\n", " ").split())
    text = " ".join(
        token
        for token in text.split()
        if not token.startswith(("http://", "https://", "#"))
    )
    return text[:180]


def _missing_metrics(
    facebook: dict[str, int | None], instagram: dict[str, int | None]
) -> list[str]:
    missing = []
    for platform, values in (("facebook", facebook), ("instagram", instagram)):
        for key, value in values.items():
            if value is None:
                missing.append(f"{platform}.{key}")
    # Facebook v25 replaces reach with viewers. One of the two is sufficient.
    if facebook.get("viewers") is not None and "facebook.reach" in missing:
        missing.remove("facebook.reach")
    return missing


def _dedupe(values: Iterable[str]) -> list[str]:
    return list(dict.fromkeys(value for value in values if value))
