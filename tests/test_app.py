from __future__ import annotations

import json
import os
import tempfile
import unittest
from io import BytesIO
from urllib.parse import parse_qs, urlparse
from unittest.mock import patch

from PIL import Image


TEST_DATA_DIR = tempfile.TemporaryDirectory()
os.environ["APP_DATA_DIR"] = TEST_DATA_DIR.name
os.environ["POPO_TEAMSPACE_ID"] = "test-team"
os.environ["POPO_SOCIAL_FOLDER_ID"] = "test-folder"
os.environ["META_APP_ID"] = "123456789"
os.environ["META_APP_SECRET"] = "test-app-secret"
os.environ["META_REDIRECT_URI"] = "https://example.test/api/meta/callback"
os.environ["META_GRAPH_VERSION"] = "v25.0"
os.environ["META_TOKEN_ENCRYPTION_KEY"] = (
    "MDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA="
)

from fastapi.testclient import TestClient  # noqa: E402

from webapp import app as app_module  # noqa: E402
from webapp.meta import MetaError, MetaSyncTimeout  # noqa: E402


class MetaApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.client_context = TestClient(app_module.app)
        cls.client = cls.client_context.__enter__()

    @classmethod
    def tearDownClass(cls):
        cls.client_context.__exit__(None, None, None)
        TEST_DATA_DIR.cleanup()

    def test_health_reports_meta_configuration_without_exposing_secrets(self):
        response = self.client.get("/api/health")
        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertTrue(payload["meta_configured"])
        self.assertFalse(payload["meta_connected"])
        self.assertTrue(payload["popo_configured"])
        self.assertIn("popo_available", payload)
        self.assertNotIn("app_secret", payload)
        self.assertNotIn("token", payload)

    def test_status_requires_connection_before_sync(self):
        status = self.client.get("/api/meta/status")
        self.assertEqual(status.status_code, 200)
        self.assertEqual(status.json()["reason"], "not_connected")

        sync = self.client.post(
            "/api/meta/sync",
            json={
                "page_id": "page-1",
                "instagram_id": "ig-1",
                "report_start": "2026-06-30",
                "report_end": "2026-07-25",
            },
        )
        self.assertEqual(sync.status_code, 401)

    def test_connect_redirect_uses_state_and_configured_callback(self):
        response = self.client.get("/api/meta/connect", follow_redirects=False)
        self.assertEqual(response.status_code, 302)
        parsed = urlparse(response.headers["location"])
        query = parse_qs(parsed.query)
        self.assertEqual(query["client_id"], ["123456789"])
        self.assertEqual(
            query["redirect_uri"], ["https://example.test/api/meta/callback"]
        )
        self.assertEqual(query["response_type"], ["code"])
        self.assertEqual(len(query["state"][0]), 43)

    def test_callback_reports_meta_network_failure(self):
        state = app_module.META_REPOSITORY.create_state()
        with patch.object(
            app_module.META_CLIENT,
            "exchange_code",
            side_effect=MetaError("连接 Meta API 超时或网络不可用"),
        ):
            response = self.client.get(
                f"/api/meta/callback?state={state}&code=test-code",
                follow_redirects=False,
            )
        self.assertEqual(response.status_code, 302)
        self.assertEqual(
            response.headers["location"],
            "/?meta=error&reason=meta_network",
        )

    def test_ui_keeps_manual_fields_in_a_conditional_details_element(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        html = response.text
        self.assertIn('id="manualFallback" hidden', html)
        self.assertIn("手动补充 / 修正数据（非必填）", html)
        self.assertIn('id="metaPreview"', html)
        self.assertIn('id="popoConnect"', html)
        self.assertIn('id="popoAuthLink"', html)
        self.assertIn("window.open(started.authorization_url", html)
        self.assertNotIn("window.open('about:blank'", html)
        self.assertIn("POPO 云空间交付", html)
        self.assertIn("打开 POPO 原生报告", html)
        self.assertIn('id="linkedinMode"', html)
        self.assertIn("个人账号——人工 Top 10 汇总/截图", html)
        self.assertIn("/api/templates/linkedin-personal-top10.xlsx", html)
        self.assertNotIn("<legend>Facebook 概览</legend>", html)
        self.assertNotIn("<legend>Instagram 概览</legend>", html)
        self.assertIn('id="metaSyncProgress"', html)
        self.assertIn("/api/meta/sync/${encodeURIComponent(syncJobId)}", html)
        self.assertIn("Meta 数据仍在后台同步", html)
        self.assertIn("正在跳转到 Meta 授权页面", html)
        self.assertIn("服务器连接 Meta 超时", html)
        self.assertIn("重试生成", html)

    def test_failed_report_can_retry_with_saved_inputs(self):
        job_id = "retryreport1"
        job_dir = app_module.JOBS_DIR / job_id
        (job_dir / "inputs").mkdir(parents=True, exist_ok=True)
        (job_dir / "request.json").write_text("{}", encoding="utf-8")
        with app_module._db() as conn:
            conn.execute(
                """
                INSERT OR REPLACE INTO jobs(
                    id,brand,report_type,report_start,report_end,status,error,
                    created_at,created_by
                ) VALUES(?,?,?,?,?,'failed',?,?,?)
                """,
                (
                    job_id, "Retry Brand", "annual", "2025-08-01", "2026-07-31",
                    "old parser error", app_module._now_cn(), "tester",
                ),
            )
        with patch.object(app_module._wake, "put") as wake:
            response = self.client.post(f"/api/jobs/{job_id}/retry")
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.json()["status"], "queued")
        self.assertIsNone(response.json()["error"])
        wake.assert_called_once_with(job_id)

    def test_meta_sync_submission_returns_a_background_job_immediately(self):
        with (
            patch.object(
                app_module,
                "_meta_status_payload",
                return_value={"connected": True},
            ),
            patch.object(
                app_module.META_REPOSITORY,
                "get_asset",
                return_value={
                    "page_id": "page-background",
                    "page_name": "Background Page",
                    "page_token": "page-token",
                    "instagram_id": "ig-background",
                },
            ),
            patch.object(app_module._meta_sync_wake, "put"),
            patch.object(app_module.META_CLIENT, "sync") as sync,
        ):
            response = self.client.post(
                "/api/meta/sync",
                json={
                    "page_id": "page-background",
                    "instagram_id": "ig-background",
                    "report_start": "2026-05-01",
                    "report_end": "2026-07-31",
                },
            )
            self.assertEqual(response.status_code, 202, response.text)
            payload = response.json()
            self.assertEqual(payload["status"], "queued")
            self.assertEqual(payload["progress"], 5)
            self.assertNotIn("result", payload)
            sync.assert_not_called()
            app_module._set_meta_sync_job(
                payload["id"],
                status="failed",
                error="test cleanup",
                finished_at=app_module._now_cn(),
            )

    def test_meta_sync_worker_saves_result_and_exposes_progress(self):
        sync_job_id = "metaworkerok1"
        with app_module._db() as conn:
            conn.execute(
                """
                INSERT OR REPLACE INTO meta_sync_jobs(
                    id,page_id,instagram_id,report_start,report_end,status,
                    progress,message,created_at
                ) VALUES(?,?,?,?,?,'queued',5,?,?)
                """,
                (
                    sync_job_id,
                    "page-worker",
                    "ig-worker",
                    "2026-05-01",
                    "2026-07-31",
                    "queued",
                    app_module._now_cn(),
                ),
            )
            row = conn.execute(
                "SELECT * FROM meta_sync_jobs WHERE id=?", (sync_job_id,)
            ).fetchone()

        result = {
            "complete": True,
            "page": {"id": "page-worker", "name": "Worker Page"},
            "instagram_account": {"id": "ig-worker", "username": "worker_ig"},
            "facebook": {"posts": 1, "views": 2},
            "instagram": {"posts": 1, "views": 3},
            "missing_metrics": [],
            "missing_labels": [],
            "warnings": [],
        }

        def fake_sync(*_args, progress_callback=None, **_kwargs):
            segment_message = "正在读取 Facebook 浏览量（第 2/12 段）"
            progress_callback(32, segment_message)
            running = app_module._get_meta_sync_job(sync_job_id)
            self.assertEqual(running["progress"], 32)
            self.assertEqual(running["message"], segment_message)
            self.assertNotIn("result", running)
            progress_callback(60, "正在读取 Instagram 内容")
            return result

        with (
            patch.object(
                app_module.META_REPOSITORY,
                "get_asset",
                return_value={
                    "page_id": "page-worker",
                    "page_name": "Worker Page",
                    "page_token": "page-token",
                    "instagram_id": "ig-worker",
                },
            ),
            patch.object(app_module.META_CLIENT, "sync", side_effect=fake_sync),
        ):
            app_module._run_meta_sync_job(row)

        payload = app_module._get_meta_sync_job(sync_job_id)
        self.assertEqual(payload["status"], "succeeded")
        self.assertEqual(payload["progress"], 100)
        self.assertEqual(payload["result"]["sync_id"], payload["sync_id"])
        self.assertEqual(payload["result"]["facebook"]["views"], 2)

    def test_meta_sync_worker_returns_exact_meta_error(self):
        sync_job_id = "metaworkerfail"
        with app_module._db() as conn:
            conn.execute(
                """
                INSERT OR REPLACE INTO meta_sync_jobs(
                    id,page_id,instagram_id,report_start,report_end,status,
                    progress,message,created_at
                ) VALUES(?,?,?,?,?,'queued',5,?,?)
                """,
                (
                    sync_job_id,
                    "page-fail",
                    None,
                    "2026-07-01",
                    "2026-07-31",
                    "queued",
                    app_module._now_cn(),
                ),
            )
            row = conn.execute(
                "SELECT * FROM meta_sync_jobs WHERE id=?", (sync_job_id,)
            ).fetchone()

        with (
            patch.object(
                app_module.META_REPOSITORY,
                "get_asset",
                return_value={
                    "page_id": "page-fail",
                    "page_name": "Fail Page",
                    "page_token": "page-token",
                    "instagram_id": None,
                },
            ),
            patch.object(
                app_module.META_CLIENT,
                "sync",
                side_effect=MetaError(
                    "Meta API：访问令牌已过期",
                    code=190,
                    reconnect=True,
                ),
            ),
        ):
            app_module._run_meta_sync_job(row)

        payload = app_module._get_meta_sync_job(sync_job_id)
        self.assertEqual(payload["status"], "failed")
        self.assertEqual(payload["error_code"], 190)
        self.assertTrue(payload["reconnect"])
        self.assertIn("访问令牌已过期", payload["error"])

    def test_timed_out_meta_job_has_no_snapshot_and_can_be_submitted_again(self):
        sync_job_id = "metaworkertimeout"
        asset = {
            "page_id": "page-timeout",
            "page_name": "Timeout Page",
            "page_token": "synthetic",
            "instagram_id": None,
        }
        with app_module._db() as conn:
            conn.execute(
                """
                INSERT OR REPLACE INTO meta_sync_jobs(
                    id,page_id,report_start,report_end,status,progress,message,created_at
                ) VALUES(?,?,?,?,'running',30,'reading',?)
                """,
                (sync_job_id, asset["page_id"], "2026-07-01", "2026-07-31", app_module._now_cn()),
            )
            row = conn.execute(
                "SELECT * FROM meta_sync_jobs WHERE id=?", (sync_job_id,)
            ).fetchone()

        with (
            patch.object(app_module.META_REPOSITORY, "get_asset", return_value=asset),
            patch.object(app_module.META_REPOSITORY, "save_sync") as save,
            patch.object(app_module.META_CLIENT, "sync", side_effect=MetaSyncTimeout("总时限已到")),
        ):
            app_module._run_meta_sync_job(row)
        save.assert_not_called()
        response = self.client.get(f"/api/meta/sync/{sync_job_id}")
        payload = response.json()
        self.assertEqual(payload["status"], "failed")
        self.assertEqual(payload["message"], "Meta 数据同步超时")
        self.assertIn("总时限", payload["error"])
        self.assertFalse(payload["detail"]["reconnect"])
        self.assertIsNone(payload["sync_id"])
        self.assertIsNotNone(payload["finished_at"])
        self.assertNotIn("result", payload)

        with (
            patch.object(app_module, "_meta_status_payload", return_value={"connected": True}),
            patch.object(app_module.META_REPOSITORY, "get_asset", return_value=asset),
            patch.object(app_module._meta_sync_wake, "put"),
        ):
            retried = self.client.post("/api/meta/sync", json={
                "page_id": asset["page_id"], "report_start": "2026-07-01",
                "report_end": "2026-07-31",
            })
            self.assertEqual(retried.status_code, 202, retried.text)
            self.assertNotEqual(retried.json()["id"], sync_job_id)
            app_module._set_meta_sync_job(
                retried.json()["id"], status="failed", finished_at=app_module._now_cn()
            )

    def test_linkedin_personal_template_downloads_xlsx(self):
        response = self.client.get("/api/templates/linkedin-personal-top10.xlsx")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.content[:2], b"PK")
        self.assertIn(
            "LinkedIn_personal_Top10_template.xlsx",
            response.headers["content-disposition"],
        )

    @staticmethod
    def _png() -> bytes:
        stream = BytesIO()
        Image.new("RGB", (50, 30), color=(230, 240, 248)).save(
            stream, format="PNG"
        )
        return stream.getvalue()

    def test_personal_screenshot_only_is_accepted_by_job_api(self):
        image = self._png()
        with patch.object(app_module._wake, "put"):
            response = self.client.post(
                "/api/jobs",
                data={
                    "brand": "Personal API QA",
                    "report_type": "monthly",
                    "report_start": "2026-07-01",
                    "report_end": "2026-07-28",
                    "created_by": "Tester",
                    "linkedin_mode": "personal",
                },
                files=[
                    ("facebook_screenshot", ("facebook.png", image, "image/png")),
                    ("instagram_screenshot", ("instagram.png", image, "image/png")),
                    (
                        "linkedin_personal_screenshots",
                        ("linkedin.png", image, "image/png"),
                    ),
                ],
            )
        self.assertEqual(response.status_code, 200, response.text)
        job_id = response.json()["id"]
        request = (
            app_module.JOBS_DIR / job_id / "request.json"
        ).read_text(encoding="utf-8")
        self.assertIn('"linkedin_mode": "personal"', request)
        self.assertIn("linkedin_personal_screenshots", request)

    def test_personal_mode_requires_table_or_screenshot(self):
        image = self._png()
        response = self.client.post(
            "/api/jobs",
            data={
                "brand": "Missing Personal Input",
                "report_type": "monthly",
                "report_start": "2026-07-01",
                "report_end": "2026-07-28",
                "created_by": "Tester",
                "linkedin_mode": "personal",
            },
            files={
                "facebook_screenshot": ("facebook.png", image, "image/png"),
                "instagram_screenshot": ("instagram.png", image, "image/png"),
            },
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("至少需要上传", response.json()["detail"])

    def test_stale_meta_sync_id_uses_submitted_api_snapshot(self):
        image = self._png()
        with patch.object(app_module._wake, "put"):
            response = self.client.post(
                "/api/jobs",
                data={
                    "brand": "Stale Meta Snapshot QA",
                    "account_label": "Snapshot Page",
                    "report_type": "monthly",
                    "report_start": "2026-07-01",
                    "report_end": "2026-07-28",
                    "created_by": "Tester",
                    "meta_sync_id": "sync-from-replaced-container",
                    "facebook_posts": "12",
                    "facebook_views": "345",
                    "instagram_posts": "9",
                    "instagram_views": "210",
                    "linkedin_mode": "none",
                },
                files={
                    "facebook_screenshot": ("facebook.png", image, "image/png"),
                    "instagram_screenshot": ("instagram.png", image, "image/png"),
                },
            )
        self.assertEqual(response.status_code, 200, response.text)
        job_id = response.json()["id"]
        request = json.loads(
            (app_module.JOBS_DIR / job_id / "request.json").read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(request["meta"]["source"], "cached_api")
        self.assertEqual(request["meta"]["sync_id"], "sync-from-replaced-container")
        self.assertEqual(request["facebook"]["views"], 345)
        self.assertEqual(request["instagram"]["views"], 210)
        self.assertIn("API 同步快照", request["meta"]["warnings"][0])

    def test_popo_failure_does_not_change_successful_report_status(self):
        job_id = "popoisolated1"
        job_dir = app_module.JOBS_DIR / job_id
        (job_dir / "outputs").mkdir(parents=True)
        (job_dir / "request.json").write_text("{}", encoding="utf-8")
        with app_module._db() as conn:
            conn.execute(
                "INSERT INTO jobs("
                "id,brand,report_type,report_start,report_end,status,"
                "created_at,created_by,popo_status"
                ") VALUES(?,?,?,?,?,?,?,?,?)",
                (
                    job_id,
                    "Isolation QA",
                    "monthly",
                    "2026-07-01",
                    "2026-07-28",
                    "succeeded",
                    app_module._now_cn(),
                    "Tester",
                    "pending",
                ),
            )
        with patch(
            "webapp.app.deliver_job",
            side_effect=RuntimeError("simulated POPO outage"),
        ):
            with app_module._db() as conn:
                row = conn.execute(
                    "SELECT * FROM jobs WHERE id=?", (job_id,)
                ).fetchone()
            app_module._run_popo_delivery(row)
        job = app_module._get_job(job_id)
        self.assertEqual(job["status"], "succeeded")
        self.assertEqual(job["popo_status"], "failed")
        self.assertIn("simulated POPO outage", job["popo_error"])


if __name__ == "__main__":
    unittest.main()
