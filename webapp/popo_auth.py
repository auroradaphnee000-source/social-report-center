"""Server-owned Fabric OpenID authorization for POPO delivery."""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import threading
import time
from pathlib import Path


class PopoAuthError(RuntimeError):
    def __init__(self, message: str, *, code: str = "", retryable: bool = False):
        super().__init__(message)
        self.code = code
        self.retryable = retryable


_RUN_LOCK = threading.RLock()
_CACHE_LOCK = threading.Lock()
_STATUS_CACHE: dict[str, object] = {"at": 0.0, "value": None}
_STATUS_TTL = max(5, int(os.getenv("POPO_STATUS_TTL", "60")))
_PROXY_VARS = {
    "http_proxy",
    "https_proxy",
    "all_proxy",
    "ftp_proxy",
    "socks_proxy",
    "no_proxy",
}


def fabric_cli() -> str:
    explicit = os.getenv("FABRIC_BIN", "").strip()
    if explicit:
        return explicit
    appdata = os.getenv("APPDATA", "")
    if appdata:
        for name in ("fabric.cmd", "fabric.exe", "fabric"):
            candidate = os.path.join(appdata, "npm", name)
            if os.path.isfile(candidate):
                return candidate
    return shutil.which("fabric") or ""


def fabric_env() -> dict[str, str]:
    """POPO uses the injected TCP bridge and must bypass the Google proxy."""
    result = {
        key: value
        for key, value in os.environ.items()
        if key.lower() not in _PROXY_VARS
    }
    dns_patch = os.getenv("POPO_DNS_PATCH", "").strip()
    tunnel_ip = os.getenv("POPO_TUNNEL_IP", "").strip()
    if dns_patch and tunnel_ip and Path(dns_patch).is_file():
        existing = result.get("NODE_OPTIONS", "").strip()
        require_option = f"--require={dns_patch}"
        if require_option not in existing.split():
            result["NODE_OPTIONS"] = " ".join(
                item for item in (existing, require_option) if item
            )
    return result


def fabric_lock():
    return _RUN_LOCK


def _json_payload(output: str):
    match = re.search(r"\{.*\}", output or "", re.DOTALL)
    if not match:
        raise PopoAuthError("Fabric CLI 未返回可识别的 JSON")
    try:
        return json.loads(match.group())
    except ValueError as exc:
        raise PopoAuthError("Fabric CLI 返回了无效 JSON") from exc


def _run(*args: str, timeout: int = 150):
    executable = fabric_cli()
    if not executable:
        raise PopoAuthError("服务器尚未安装 Fabric CLI")
    try:
        with _RUN_LOCK:
            proc = subprocess.run(
                [executable, *args],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                env=fabric_env(),
            )
    except FileNotFoundError as exc:
        raise PopoAuthError("服务器尚未安装 Fabric CLI") from exc
    except subprocess.TimeoutExpired as exc:
        raise PopoAuthError("Fabric 授权请求超时，请稍后重试") from exc
    output = (proc.stdout or "") + "\n" + (proc.stderr or "")
    return _json_payload(output)


def _deep_value(node, key: str):
    if isinstance(node, dict):
        value = node.get(key)
        if value not in (None, ""):
            return value
        for child in node.values():
            found = _deep_value(child, key)
            if found not in (None, ""):
                return found
    elif isinstance(node, list):
        for child in node:
            found = _deep_value(child, key)
            if found not in (None, ""):
                return found
    return None


def invalidate_status_cache():
    with _CACHE_LOCK:
        _STATUS_CACHE["value"] = None


def _public_status(payload) -> dict:
    data = payload.get("data") if isinstance(payload, dict) else None
    data = data if isinstance(data, dict) else {}
    active = data.get("active_auth_method")
    connected = bool(
        isinstance(payload, dict)
        and payload.get("ok")
        and active in ("openid", "sk")
        and (data.get("has_access_token") or active == "sk")
    )
    return {
        "available": bool(fabric_cli()),
        "connected": connected,
        "auth_method": active,
        "account": data.get("uid") if connected else None,
        "status": data.get("status"),
        "gateway_configured": bool(os.getenv("FABRIC_GATEWAY_URL")),
        "credential_store_persistent": _credential_store_persistent(),
    }


def _credential_store_persistent() -> bool:
    app_data = os.getenv("APP_DATA_DIR", "").strip()
    if not app_data:
        return False
    try:
        home = Path.home().resolve()
        data_dir = Path(app_data).resolve()
        return home == data_dir or data_dir in home.parents
    except (OSError, RuntimeError):
        return False


def _auth_failure(payload, *, prefix: str) -> PopoAuthError:
    code = str(_deep_value(payload, "code") or "FABRIC_AUTH_UNKNOWN")
    original = str(_deep_value(payload, "message") or "Fabric 未返回错误详情")
    lowered = original.lower()
    if code == "FABRIC_INVALID_RESPONSE" and (
        "403" in lowered or "non-json" in lowered
    ):
        message = (
            "POPO 内网授权通道被拒绝（HTTP 403）。"
            "请确认服务通过 needs.popo 部署且 FABRIC_GATEWAY_URL 已注入；"
            "这不是浏览器登录问题。"
        )
    else:
        message = original
    return PopoAuthError(f"{prefix}（{code}）：{message}", code=code)


def _compute_status(auto_login: bool) -> dict:
    if not fabric_cli():
        return {
            "available": False,
            "connected": False,
            "message": "服务器尚未安装 Fabric CLI",
        }
    try:
        payload = _run("auth", "status", timeout=30)
        data = payload.get("data") if isinstance(payload, dict) else None
        data = data if isinstance(data, dict) else {}
        if (
            auto_login
            and isinstance(payload, dict)
            and payload.get("ok")
            and data.get("active_auth_method") == "openid"
            and data.get("has_refresh_token")
            and not data.get("has_access_token")
        ):
            refreshed = _run("auth", "refresh", timeout=60)
            if not isinstance(refreshed, dict) or not refreshed.get("ok"):
                result = _public_status(payload)
                result["connected"] = False
                result["message"] = str(
                    _deep_value(refreshed, "message")
                    or "POPO 已保存授权，但自动刷新访问令牌失败"
                )
                return result
            payload = _run("auth", "status", timeout=30)
        result = _public_status(payload)
        if not result["connected"] and auto_login and os.getenv("FABRIC_SK"):
            logged_in = _run(
                "auth", "login", "--sk-env", "FABRIC_SK", timeout=60
            )
            if not isinstance(logged_in, dict) or not logged_in.get("ok"):
                result["message"] = str(
                    _deep_value(logged_in, "message")
                    or "服务器 Fabric Secret Key 自动登录失败"
                )
                return result
            return _compute_status(False)
        return result
    except PopoAuthError as exc:
        return {
            "available": True,
            "connected": False,
            "message": str(exc),
            "gateway_configured": bool(os.getenv("FABRIC_GATEWAY_URL")),
            "credential_store_persistent": _credential_store_persistent(),
        }


def popo_status(*, auto_login: bool = True, force: bool = False) -> dict:
    if not force:
        with _CACHE_LOCK:
            cached = _STATUS_CACHE["value"]
            if (
                cached is not None
                and time.monotonic() - float(_STATUS_CACHE["at"]) < _STATUS_TTL
            ):
                return dict(cached)
    result = _compute_status(auto_login)
    with _CACHE_LOCK:
        _STATUS_CACHE["at"] = time.monotonic()
        _STATUS_CACHE["value"] = dict(result)
    return result


def begin_popo_login() -> dict:
    payload = _run("auth", "login", "--openid", "--no-browser")
    if not isinstance(payload, dict) or not payload.get("ok"):
        raise _auth_failure(payload, prefix="POPO 授权初始化失败")
    url = _deep_value(payload, "authorization_url")
    login_id = _deep_value(payload, "login_id")
    if not url:
        invalidate_status_cache()
        status = popo_status(auto_login=False, force=True)
        if status.get("connected"):
            return {"connected": True, "status": status}
        _run("auth", "logout", timeout=30)
        payload = _run("auth", "login", "--openid", "--no-browser")
        if not isinstance(payload, dict) or not payload.get("ok"):
            raise _auth_failure(payload, prefix="POPO 授权初始化失败")
        url = _deep_value(payload, "authorization_url")
        login_id = _deep_value(payload, "login_id")
    if not url:
        raise _auth_failure(payload, prefix="POPO 授权初始化失败")
    return {
        "connected": False,
        "authorization_url": str(url),
        "login_id": str(login_id or ""),
    }


def _authorization_pending(payload, message: str) -> bool:
    status = str(_deep_value(payload, "status") or "").strip().lower()
    code = str(_deep_value(payload, "code") or "").strip().lower()
    text = str(message or "").lower()
    pending = {
        "pending",
        "pending_browser",
        "authorization_pending",
        "auth_pending",
        "fabric_auth_pending",
    }
    return (
        status in pending
        or code in pending
        or "authorization pending" in text
        or "waiting for authorization" in text
        or (
            ("timeout" in status or "timeout" in code)
            and ("openid" in status or "openid" in code or "auth" in code)
        )
        or ("openid" in text and ("timeout" in text or "超时" in text))
        or ("授权" in text and ("等待" in text or "尚未" in text))
    )


def resume_popo_login() -> dict:
    payload = _run("auth", "login", "--openid", "--resume")
    if not isinstance(payload, dict) or not payload.get("ok"):
        message = str(_deep_value(payload, "message") or "POPO 授权尚未完成")
        retryable = _authorization_pending(payload, message)
        code = str(
            _deep_value(payload, "code")
            or _deep_value(payload, "status")
            or ("FABRIC_AUTH_PENDING" if retryable else "FABRIC_AUTH_FAILED")
        )
        raise PopoAuthError(message, code=code, retryable=retryable)
    invalidate_status_cache()
    status = popo_status(auto_login=False, force=True)
    if not status.get("connected"):
        raise PopoAuthError("POPO 仍未连接，请先在授权页面完成登录")
    return status


def refresh_popo_auth():
    payload = _run("auth", "refresh", timeout=60)
    if not isinstance(payload, dict) or not payload.get("ok"):
        message = str(_deep_value(payload, "message") or "Fabric 未确认刷新成功")
        raise PopoAuthError(f"POPO 授权刷新失败：{message}")
    invalidate_status_cache()
    return payload


def _refresh_loop(interval: int):
    while True:
        try:
            status = popo_status(auto_login=True, force=True)
            if status.get("connected") and status.get("auth_method") == "openid":
                refresh_popo_auth()
        except Exception as exc:
            print(f"[popo-refresh] {type(exc).__name__}: {exc}", flush=True)
        time.sleep(interval)


def start_auth_refresher():
    interval = max(600, int(os.getenv("POPO_REFRESH_SECONDS", str(12 * 3600))))
    print(
        "[popo-auth] "
        f"gateway_configured={bool(os.getenv('FABRIC_GATEWAY_URL'))} "
        f"credential_store_persistent={_credential_store_persistent()} "
        f"dns_fallback_configured="
        f"{bool(os.getenv('POPO_DNS_PATCH') and os.getenv('POPO_TUNNEL_IP'))}",
        flush=True,
    )
    threading.Thread(
        target=_refresh_loop,
        args=(interval,),
        daemon=True,
        name="popo-auth-refresh",
    ).start()
