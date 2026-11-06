from __future__ import annotations

import unittest
from pathlib import Path
from unittest.mock import patch

from webapp.popo_auth import (
    PopoAuthError,
    _authorization_pending,
    _compute_status,
    _credential_store_persistent,
    begin_popo_login,
    fabric_env,
)


class AuthorizationPendingTests(unittest.TestCase):
    def test_openid_timeout_message_remains_retryable(self):
        self.assertTrue(
            _authorization_pending(
                {"code": "OPENID_LOGIN_TIMEOUT"},
                "OpenID 登录超时，请完成浏览器登录后重试",
            )
        )

    def test_openid_timeout_status_remains_retryable(self):
        self.assertTrue(
            _authorization_pending(
                {"status": "openid_timeout"},
                "authorization has not completed",
            )
        )

    def test_unrelated_auth_failure_is_not_retryable(self):
        self.assertFalse(
            _authorization_pending(
                {"code": "INVALID_CREDENTIALS"},
                "invalid credentials",
            )
        )


class FabricEnvironmentTests(unittest.TestCase):
    def test_google_proxy_is_removed_but_popo_gateway_is_preserved(self):
        with patch.dict(
            "os.environ",
            {
                "HTTP_PROXY": "http://proxy.invalid",
                "HTTPS_PROXY": "http://proxy.invalid",
                "FABRIC_GATEWAY_URL": "https://cli.popo.netease.com:4443",
            },
            clear=True,
        ):
            env = fabric_env()
        self.assertNotIn("HTTP_PROXY", env)
        self.assertNotIn("HTTPS_PROXY", env)
        self.assertEqual(
            env["FABRIC_GATEWAY_URL"],
            "https://cli.popo.netease.com:4443",
        )

    def test_fabric_subprocess_gets_tunnel_dns_patch(self):
        patch_path = Path(__file__).parents[1] / "webapp" / "popo_dns_patch.cjs"
        with patch.dict(
            "os.environ",
            {
                "POPO_DNS_PATCH": str(patch_path),
                "POPO_TUNNEL_IP": "172.17.0.1",
                "NODE_OPTIONS": "--trace-warnings",
            },
            clear=True,
        ):
            env = fabric_env()
        self.assertIn("--trace-warnings", env["NODE_OPTIONS"])
        self.assertIn(f"--require={patch_path}", env["NODE_OPTIONS"])

    def test_fabric_home_below_app_data_is_persistent(self):
        with (
            patch.dict(
                "os.environ",
                {"APP_DATA_DIR": "/socialapp/runtime"},
                clear=True,
            ),
            patch(
                "webapp.popo_auth.Path.home",
                return_value=Path("/socialapp/runtime/fabric-home"),
            ),
        ):
            self.assertTrue(_credential_store_persistent())


class FabricInitializationTests(unittest.TestCase):
    @patch("webapp.popo_auth.fabric_cli", return_value="/usr/local/bin/fabric")
    @patch("webapp.popo_auth._run")
    def test_saved_refresh_token_is_refreshed_on_startup(self, run, _cli):
        run.side_effect = [
            {
                "ok": True,
                "data": {
                    "active_auth_method": "openid",
                    "has_refresh_token": True,
                    "has_access_token": False,
                },
            },
            {"ok": True},
            {
                "ok": True,
                "data": {
                    "active_auth_method": "openid",
                    "has_refresh_token": True,
                    "has_access_token": True,
                    "uid": "user@example.com",
                },
            },
        ]
        status = _compute_status(auto_login=True)
        self.assertTrue(status["connected"])
        self.assertEqual(run.call_args_list[1].args, ("auth", "refresh"))

    @patch("webapp.popo_auth._run")
    def test_http_403_is_reported_as_server_tunnel_problem(self, run):
        run.return_value = {
            "ok": False,
            "code": "FABRIC_INVALID_RESPONSE",
            "message": "HTTP 403: Non-JSON response received",
        }
        with self.assertRaises(PopoAuthError) as caught:
            begin_popo_login()
        self.assertIn("内网授权通道被拒绝", str(caught.exception))
        self.assertIn("不是浏览器登录问题", str(caught.exception))
        self.assertEqual(run.call_count, 1)

    @patch("webapp.popo_auth.fabric_cli", return_value="/usr/local/bin/fabric")
    @patch("webapp.popo_auth._run", return_value=["unexpected"])
    def test_unexpected_status_payload_is_safe(self, _run, _cli):
        status = _compute_status(auto_login=True)
        self.assertFalse(status["connected"])


if __name__ == "__main__":
    unittest.main()
