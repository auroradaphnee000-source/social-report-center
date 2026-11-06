from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from webapp.meta import (
    MetaError,
    MetaGraphClient,
    MetaRepository,
    MetaSettings,
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


class SyncTests(unittest.TestCase):
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
