from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from unittest.mock import patch

from webapp.meta import (
    MetaError,
    MetaGraphClient,
    MetaRepository,
    MetaSettings,
    MetaSyncTimeout,
    _breakdown_total,
    _metric_monthly,
    _metric_total,
    _missing_metrics,
)


def settings(**overrides):
    values = {
        "app_id": "123",
        "app_secret": "secret",
        "redirect_uri": "https://example.test/api/meta/callback",
        "token_key": "unused-in-these-tests",
        "graph_version": "v25.0",
        "login_config_id": "",
        "scopes": ("pages_show_list", "read_insights"),
    }
    values.update(overrides)
    return MetaSettings(**values)


def total(value):
    return {"data": [{"total_value": {"value": value}}]}


class MetricParsingTests(unittest.TestCase):
    def test_daily_metric_end_time_is_grouped_into_its_actual_month(self):
        payload = {
            "data": [
                {
                    "values": [
                        {"value": 7, "end_time": "2026-08-01T00:00:00+0000"},
                        {"value": 3, "end_time": "2026-08-02T00:00:00+0000"},
                    ]
                }
            ]
        }
        self.assertEqual(_metric_monthly(payload), {"2026-07": 7, "2026-08": 3})

    def test_total_value_is_preferred_for_unique_metric(self):
        payload = {
            "data": [
                {
                    "total_value": {"value": 12},
                    "values": [{"value": 8}, {"value": 9}],
                }
            ]
        }
        self.assertEqual(_metric_total(payload, additive=False), 12)

    def test_unique_metric_does_not_sum_daily_values(self):
        payload = {"data": [{"values": [{"value": 8}, {"value": 9}]}]}
        self.assertIsNone(_metric_total(payload, additive=False))
        self.assertEqual(_metric_total(payload, additive=True), 17)

    def test_paid_organic_breakdown(self):
        payload = {
            "data": [
                {
                    "total_value": {
                        "value": 12,
                        "breakdowns": [
                            {
                                "dimension_keys": ["is_from_ads"],
                                "results": [
                                    {"dimension_values": ["false"], "value": 10},
                                    {"dimension_values": ["true"], "value": 2},
                                ],
                            }
                        ],
                    }
                }
            ]
        }
        self.assertEqual(
            _breakdown_total(payload, "is_from_ads", {"0", "false"}), 10
        )
        self.assertEqual(
            _breakdown_total(payload, "is_from_ads", {"1", "true"}), 2
        )

    def test_facebook_viewers_replaces_missing_reach(self):
        missing = _missing_metrics(
            {
                "posts": 1,
                "views": 2,
                "reach": None,
                "viewers": 2,
                "interactions": 0,
                "likes": 0,
            },
            {
                "posts": 1,
                "views": 2,
                "reach": 2,
                "interactions": 0,
                "likes": 0,
            },
        )
        self.assertNotIn("facebook.reach", missing)


class OAuthTests(unittest.TestCase):
    def test_authorization_url_contains_state_and_scopes(self):
        url = MetaGraphClient(settings()).authorization_url("state-token")
        query = parse_qs(urlparse(url).query)
        self.assertEqual(query["state"], ["state-token"])
        self.assertEqual(query["client_id"], ["123"])
        self.assertIn("pages_show_list", query["scope"][0])
        self.assertNotIn("secret", url)

    def test_login_for_business_config_id_replaces_scope_parameter(self):
        url = MetaGraphClient(
            settings(login_config_id="config-456")
        ).authorization_url("state-token")
        query = parse_qs(urlparse(url).query)
        self.assertEqual(query["config_id"], ["config-456"])
        self.assertNotIn("scope", query)

    def test_oauth_state_is_single_use(self):
        with tempfile.TemporaryDirectory() as folder:
            repo = MetaRepository(Path(folder) / "test.sqlite3", settings())
            repo.init_db()
            state = repo.create_state()
            self.assertTrue(repo.consume_state(state))
            self.assertFalse(repo.consume_state(state))


class PageEnumerationClient(MetaGraphClient):
    def _paginate(self, path, *, token, params):
        values = {
            "me/accounts": [
                {
                    "id": "direct",
                    "name": "Direct Page",
                    "access_token": "direct-token",
                    "tasks": ["ANALYZE"],
                },
                {
                    "id": "duplicate",
                    "name": "Duplicate Page",
                    "access_token": "duplicate-token",
                    "tasks": ["ANALYZE"],
                },
            ],
            "me/businesses": [
                {"id": "business-1", "name": "Portfolio One"},
                {"id": "business-2", "name": "Portfolio Two"},
            ],
            "business-1/owned_pages": [
                {
                    "id": "owned",
                    "name": "Owned Page",
                    "access_token": "owned-token",
                    "tasks": ["ANALYZE"],
                    "instagram_business_account": {
                        "id": "ig-owned",
                        "username": "owned_ig",
                    },
                },
                {
                    "id": "duplicate",
                    "name": "Duplicate Page",
                    "access_token": "duplicate-token",
                    "tasks": ["CREATE_CONTENT"],
                },
            ],
            "business-1/client_pages": [
                {
                    "id": "client",
                    "name": "Client Page",
                    "access_token": "client-token",
                    "tasks": ["ANALYZE"],
                }
            ],
            "business-2/owned_pages": [],
        }
        if path == "business-2/client_pages":
            raise MetaError("edge unavailable", code=100)
        return values[path]

    def _graph(self, path, *, token=None, params=None, versioned=True):
        return {}


class PageEnumerationTests(unittest.TestCase):
    def test_direct_and_business_pages_are_merged_and_deduplicated(self):
        pages = PageEnumerationClient(settings()).list_pages("user-token")
        self.assertEqual(
            [page["id"] for page in pages],
            ["client", "direct", "duplicate", "owned"],
        )
        duplicate = next(page for page in pages if page["id"] == "duplicate")
        self.assertEqual(
            duplicate["tasks"], ["ANALYZE", "CREATE_CONTENT"]
        )
        owned = next(page for page in pages if page["id"] == "owned")
        self.assertEqual(owned["instagram"]["username"], "owned_ig")


class InsightWindowClient(MetaGraphClient):
    def __init__(self):
        super().__init__(settings())
        self.windows = []

    def _insight(
        self,
        object_id,
        metric,
        *,
        token,
        since,
        until,
        period="day",
        breakdown=None,
        metric_type="total_value",
    ):
        self.windows.append((since, until, metric_type))
        return {
            "data": [
                {
                    "name": metric,
                    "values": [
                        {"value": 1, "end_time": f"{until}T00:00:00+0000"}
                    ],
                }
            ]
        }


class InsightWindowTests(unittest.TestCase):
    def test_annual_insight_is_split_into_at_most_30_day_windows(self):
        client = InsightWindowClient()
        payload, error = client._safe_insight(
            "ig",
            "views",
            token="token",
            since="2025-08-26",
            until="2026-08-11",
            metric_type=None,
        )
        self.assertIsNone(error)
        self.assertGreater(len(client.windows), 1)
        for since, until, metric_type in client.windows:
            self.assertLessEqual(
                (date.fromisoformat(until) - date.fromisoformat(since)).days,
                30,
            )
            self.assertIsNone(metric_type)
        self.assertEqual(len(payload["data"]), len(client.windows))


class FakeMetaClient(MetaGraphClient):
    def _paginate(self, path, *, token, params):
        if path.endswith("/posts"):
            return [
                {
                    "id": "page_post",
                    "message": "Product application demo",
                    "created_time": "2026-07-02T00:00:00+0000",
                }
            ]
        return [
            {
                "id": "ig_media",
                "caption": "Ingredient application demo",
                "timestamp": "2026-07-03T00:00:00+0000",
            }
        ]

    def _safe_insight(
        self,
        object_id,
        metric,
        *,
        token,
        since,
        until,
        period="day",
        breakdown=None,
        metric_type="total_value",
        progress_callback=None,
    ):
        if breakdown == "is_from_ads":
            return (
                {
                    "data": [
                        {
                            "total_value": {
                                "value": 20,
                                "breakdowns": [
                                    {
                                        "dimension_keys": ["is_from_ads"],
                                        "results": [
                                            {
                                                "dimension_values": ["false"],
                                                "value": 20,
                                            },
                                            {
                                                "dimension_values": ["true"],
                                                "value": 0,
                                            },
                                        ],
                                    }
                                ],
                            }
                        }
                    ]
                },
                None,
            )
        values = {
            "page_media_view": 20,
            "page_total_media_view_unique": 12,
            "page_post_engagements": 4,
            "page_actions_post_reactions_like_total": 3,
            "views": 30,
            "reach": 15,
            "total_interactions": 5,
            "likes": 4,
        }
        if metric_type is None:
            return (
                {
                    "data": [
                        {
                            "values": [
                                {
                                    "value": values[metric],
                                    "end_time": "2026-07-03T00:00:00+0000",
                                }
                            ]
                        }
                    ]
                },
                None,
            )
        return total(values[metric]), None

    def _facebook_highlight(self, posts, token):
        return "Product application demo"

    def _instagram_highlight(self, media, token):
        return "Ingredient application demo"


class WindowedSyncClient(FakeMetaClient):
    """Real windowing/aggregation with deterministic, network-free responses."""

    _safe_insight = MetaGraphClient._safe_insight

    def __init__(self, **kwargs):
        super().__init__(settings(), **kwargs)
        self.insight_calls = []

    def _insight(self, object_id, metric, **kwargs):
        self.insight_calls.append(
            (object_id, metric, tuple(sorted(kwargs.items())))
        )
        payload, _ = FakeMetaClient._safe_insight(
            self, object_id, metric, **kwargs
        )
        return payload


class SyncTests(unittest.TestCase):
    @staticmethod
    def run_sync(client, **kwargs):
        options = {
            "instagram_id": "ig",
            "report_start": "2025-11-25",
            "report_end": "2026-11-04",
        }
        options.update(kwargs)
        return client.sync(
            {
                "page_id": "page",
                "page_name": "Example Page",
                "page_token": "synthetic",
                "instagram_id": "ig",
            },
            **options,
        )

    def test_annual_sync_reuses_facebook_responses_with_unchanged_metrics(self):
        client = WindowedSyncClient()
        result = self.run_sync(client)
        fb_calls = [c for c in client.insight_calls if c[0] == "page"]
        self.assertEqual(len(fb_calls), 60)
        self.assertEqual(len(set(fb_calls)), 60)
        self.assertEqual(len(client.insight_calls), 156)
        self.assertEqual(result["facebook"]["views"], 240)
        self.assertEqual(result["facebook"]["interactions"], 48)
        self.assertEqual(result["facebook"]["monthly"][8]["month"], "2026-07")
        self.assertEqual(result["facebook"]["monthly"][8]["views"], 240)
        self.assertEqual(result["facebook"]["monthly"][8]["daily_reach_sum"], 144)
        self.assertIsNone(result["facebook"]["viewers"])
        # Distinct IG total_value and daily requests must not be conflated.
        self.assertEqual(result["instagram"]["views"], 360)
        self.assertEqual(result["instagram"]["monthly"][8]["views"], 360)
        self.assertIsNone(result["instagram"]["reach"])

        first_run_count = len(client.insight_calls)
        self.run_sync(client)
        self.assertEqual(len(client.insight_calls), first_run_count * 2)

    def test_progress_changes_within_each_platform_and_covers_date_windows(self):
        client = WindowedSyncClient()
        events = []
        self.run_sync(client, progress_callback=lambda *event: events.append(event))
        values = [value for value, _ in events]
        self.assertEqual(values, sorted(values))
        self.assertTrue(any(30 < value < 50 for value in values))
        self.assertTrue(any(72 < value < 88 for value in values))
        for platform in ("Facebook", "Instagram"):
            self.assertTrue(any(
                platform in message and "第 1/12 段" in message
                and "2025-11-25—2025-12-24" in message
                for _, message in events
            ))
            self.assertTrue(any(
                platform in message and "第 12/12 段" in message
                and "2026-10-21—2026-11-04" in message
                and "已读取" in message
                for _, message in events
            ))

    def test_failed_windows_remain_missing_and_progress_continues(self):
        client = WindowedSyncClient()
        original = client._insight

        def fail_views(object_id, metric, **kwargs):
            if object_id == "page" and metric == "page_media_view":
                raise MetaError("unavailable metric", code=100)
            return original(object_id, metric, **kwargs)

        events = []
        with patch.object(client, "_insight", side_effect=fail_views):
            result = self.run_sync(client, progress_callback=lambda *e: events.append(e))
        self.assertIsNone(result["facebook"]["views"])
        self.assertIsNone(result["facebook"]["monthly"][8]["views"])
        self.assertIn("facebook.views", result["missing_metrics"])
        self.assertTrue(any("未取得" in message for _, message in events))
        self.assertEqual(events[-1][0], 92)

    def test_deadline_stops_mid_metric_and_next_sync_has_a_fresh_budget(self):
        client = WindowedSyncClient(sync_timeout=10)
        clock = [0.0]
        original = client._insight

        def slow_insight(*args, **kwargs):
            payload = original(*args, **kwargs)
            clock[0] += 6
            return payload

        with patch("webapp.meta.time.monotonic", side_effect=lambda: clock[0]):
            with patch.object(client, "_insight", side_effect=slow_insight):
                with self.assertRaises(MetaSyncTimeout):
                    self.run_sync(client)
            self.assertEqual(len(client.insight_calls), 2)
            result = self.run_sync(client)
        self.assertIn("facebook", result)

    def test_http_wait_is_limited_to_remaining_budget_and_late_result_rejected(self):
        client = MetaGraphClient(settings(), sync_timeout=2)
        clock = [0.0]

        def read_payload(*_args):
            clock[0] = 3
            return b"{}"

        with (
            patch("webapp.meta.time.monotonic", side_effect=lambda: clock[0]),
            patch("webapp.meta.urllib.request.urlopen") as open_request,
        ):
            open_request.return_value.__enter__.return_value.read.side_effect = read_payload
            with patch.object(
                client, "_sync",
                side_effect=lambda *a, **k: client._request_json("https://graph.facebook.com/test"),
            ):
                with self.assertRaises(MetaSyncTimeout):
                    self.run_sync(client)
            self.assertEqual(open_request.call_args.kwargs["timeout"], 2)
            # The failed run must not shorten a later non-sync/OAuth request.
            self.assertEqual(client._request_json("https://graph.facebook.com/test"), {})
            self.assertEqual(open_request.call_args.kwargs["timeout"], 35)

    def test_timeout_escapes_breakdown_and_content_fallbacks(self):
        client = MetaGraphClient(settings())
        actions = (
            lambda: client._insight(
                "page", "views", token="synthetic", since="2026-07-01",
                until="2026-07-02", breakdown="is_from_ads",
            ),
            lambda: client._safe_insight(
                "page", "views", token="synthetic", since="2026-07-01",
                until="2026-09-01",
            ),
            lambda: client._facebook_highlight(
                [{"id": "post", "message": "Example"}], "synthetic"
            ),
            lambda: client._instagram_highlight(
                [{"id": "media", "caption": "Example"}], "synthetic"
            ),
        )
        for action in actions:
            with self.subTest(action=action), patch.object(
                client, "_graph", side_effect=MetaSyncTimeout("deadline exceeded")
            ) as graph:
                with self.assertRaises(MetaSyncTimeout):
                    action()
                self.assertEqual(graph.call_count, 1)

    def test_sync_timeout_configuration_rejects_unbounded_values(self):
        for value in (0, -1, float("inf"), float("nan")):
            with self.subTest(value=value), self.assertRaises(ValueError):
                MetaGraphClient(settings(), sync_timeout=value)

    def test_sync_returns_complete_server_side_snapshot(self):
        client = FakeMetaClient(settings())
        result = client.sync(
            {
                "page_id": "page",
                "page_name": "Example Page",
                "page_token": "page-token",
                "instagram_id": "ig",
                "instagram_username": "example_ig",
            },
            instagram_id="ig",
            report_start="2026-07-01",
            report_end="2026-07-28",
        )
        self.assertTrue(result["complete"])
        self.assertEqual(result["facebook"]["views"], 20)
        self.assertEqual(result["facebook"]["viewers"], 12)
        self.assertEqual(result["instagram"]["views"], 30)
        self.assertEqual(result["facebook"]["monthly"][0]["posts"], 1)
        self.assertEqual(result["facebook"]["monthly"][0]["views"], 20)
        self.assertEqual(
            result["facebook"]["monthly"][0]["daily_reach_sum"], 12
        )
        self.assertEqual(result["instagram"]["monthly"][0]["posts"], 1)
        self.assertEqual(result["instagram"]["monthly"][0]["views"], 30)
        self.assertEqual(
            result["instagram"]["monthly"][0]["daily_reach_sum"], 15
        )
        self.assertEqual(result["missing_metrics"], [])


if __name__ == "__main__":
    unittest.main()
