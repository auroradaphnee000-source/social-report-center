from __future__ import annotations

import io
import json
import tempfile
import unittest
from copy import deepcopy
from pathlib import Path
from unittest.mock import MagicMock, patch

from webapp.popo import (
    PopoDeliveryError,
    _native_html,
    _title,
    _unwrap_popo_result,
    deliver_job,
)


def sample_analysis():
    return {
        "report_title": "社媒运营月报",
        "request": {
            "brand": "Example Brand",
            "account_label": "Example",
            "report_type": "monthly",
            "report_start": "2026-07-01",
            "report_end": "2026-07-28",
            "created_by": "Tester",
            "facebook": {
                "posts": 8,
                "views": 100,
                "reach": 80,
                "viewers": 70,
                "interactions": 12,
                "likes": 7,
            },
            "instagram": {
                "posts": 6,
                "views": 120,
                "reach": 90,
                "interactions": 15,
                "likes": 9,
            },
            "facebook_highlight": "应用演示",
            "instagram_highlight": "产品知识",
            "meta": {"source": "manual"},
            "files": {
                "facebook_screenshot": "facebook.png",
                "instagram_screenshot": "instagram.png",
            },
        },
        "linkedin": {
            "content": {
                "impressions_organic": 1200,
                "clicks_organic": 75,
                "reactions_organic": 42,
                "comments_organic": 8,
                "reposts_organic": 3,
                "engagement_rate": 0.1067,
            },
            "visitors": {
                "page_views_total": 98,
                "unique_total": 51,
            },
            "followers": {"organic": 22},
            "posts": {
                "count": 9,
                "top_impressions": [
                    {
                        "title": "Application demo",
                        "url": "https://www.linkedin.com/posts/example",
                        "created_date": "2026-07-10",
                        "impressions": 320,
                        "clicks": 21,
                        "ctr": 0.065625,
                        "likes": 10,
                        "comments": 2,
                        "reposts": 1,
                    }
                ],
            },
        },
    }


class PopoEnvelopeTests(unittest.TestCase):
    def test_nested_fabric_envelope_is_unwrapped(self):
        payload = {
            "ok": True,
            "data": {
                "data": {
                    "status": 1,
                    "data": {"docId": "doc-1", "url": "https://docs.popo/doc-1"},
                }
            },
        }
        self.assertEqual(_unwrap_popo_result(payload)["docId"], "doc-1")

    def test_business_failure_returns_none(self):
        payload = {"ok": True, "data": {"status": 0, "message": "denied"}}
        self.assertIsNone(_unwrap_popo_result(payload))


class NativeReportTests(unittest.TestCase):
    def test_native_html_contains_complete_sections_images_and_word(self):
        uploaded = {
            "facebook": {"url": "https://files/fb.png", "size": 1, "name": "fb.png"},
            "instagram": {"url": "https://files/ig.png", "size": 1, "name": "ig.png"},
            "word": {
                "url": "https://files/report.docx",
                "size": 2048,
                "name": "最终报告.docx",
            },
        }
        content = _native_html(sample_analysis(), uploaded)
        self.assertTrue(content.startswith("<doc>"))
        self.assertTrue(content.endswith("</doc>"))
        self.assertIn("第一部分　Facebook", content)
        self.assertIn("第二部分　Instagram", content)
        self.assertIn("第三部分　LinkedIn", content)
        self.assertIn("第四部分　综合结论与下阶段优化建议", content)
        self.assertIn("1.2　自然内容排行", content)
        self.assertIn("1.3　当月成效分析", content)
        self.assertNotIn("月度自然内容趋势", content)
        self.assertIn('<img src="https://files/fb.png" />', content)
        self.assertIn('<img src="https://files/ig.png" />', content)
        self.assertIn('<attachment src="https://files/report.docx"', content)
        self.assertIn("https://www.linkedin.com/posts/example", content)
        self.assertIn(
            "分析原则：本报告以自然内容运营为分析对象，不把总曝光或粉丝增长作为唯一标准。",
            content,
        )
        self.assertNotIn("Meta Graph API 自动同步", content)

    def test_native_html_uses_quarterly_stage_logic(self):
        analysis = deepcopy(sample_analysis())
        analysis["report_title"] = "社媒运营季度报"
        analysis["request"]["report_type"] = "quarterly"
        content = _native_html(analysis, {})
        self.assertIn("1.2　自然内容排行", content)
        self.assertIn("1.3　阶段成效分析", content)
        self.assertIn("这一阶段验证了什么", content)
        self.assertIn("下一季度策略", content)
        self.assertNotIn("月度自然内容趋势", content)

    def test_native_html_uses_full_annual_framework(self):
        analysis = deepcopy(sample_analysis())
        analysis["report_title"] = "社媒运营年报"
        analysis["request"]["report_type"] = "annual"
        analysis["request"]["facebook"]["monthly"] = [
            {"month": "2026-06", "views": 40, "daily_reach_sum": 28},
            {"month": "2026-07", "views": 60, "daily_reach_sum": 42},
        ]
        analysis["request"]["instagram"]["monthly"] = [
            {"month": "2026-06", "views": 50, "daily_reach_sum": 32},
            {"month": "2026-07", "views": 70, "daily_reach_sum": 48},
        ]
        analysis["linkedin"]["content"]["monthly"] = [
            {"month": "2026-06", "impressions_organic": 500, "unique_impressions_organic": 320},
            {"month": "2026-07", "impressions_organic": 700, "unique_impressions_organic": 460},
        ]
        analysis["linkedin"]["followers"]["monthly"] = [
            {"month": "2026-06", "organic": 8},
            {"month": "2026-07", "organic": 14},
        ]
        content = _native_html(
            analysis,
            {
                "facebook_monthly_trend": {"url": "https://files/fb-trend.png"},
                "instagram_monthly_trend": {"url": "https://files/ig-trend.png"},
                "linkedin_monthly_trend": {"url": "https://files/li-trend.png"},
            },
        )
        self.assertIn("1.2　月度自然内容趋势", content)
        self.assertIn("1.3　自然内容排行", content)
        self.assertIn("1.4　高表现内容分析", content)
        self.assertIn("3.2 月度自然内容趋势", content)
        self.assertIn("3.4 高表现内容分析", content)
        self.assertIn("这一年积累了什么，以及下一年应该怎么做", content)
        self.assertIn("下一年度策略", content)
        self.assertIn('<img src="https://files/fb-trend.png" />', content)
        self.assertIn('<img src="https://files/ig-trend.png" />', content)
        self.assertIn('<img src="https://files/li-trend.png" />', content)
        self.assertIn("2026-06", content)
        self.assertIn("日度触达累计", content)
        self.assertNotIn("发布内容</th>", content)
        self.assertNotIn("当前线上表单未提供", content)

    def test_native_html_supports_personal_profile_top10_and_screenshots(self):
        analysis = deepcopy(sample_analysis())
        analysis["request"]["linkedin_mode"] = "personal"
        analysis["linkedin"] = {
            "mode": "personal",
            "content": {"impressions_organic": None},
            "visitors": {"page_views_total": 18, "unique_total": None},
            "followers": {"organic": None, "total": 230},
            "posts": {
                "count": 1,
                "top_impressions": [
                    {
                        "title": "Factory process",
                        "url": "https://www.linkedin.com/posts/personal",
                        "created_date": "2026-07-10",
                        "impressions": 320,
                        "clicks": None,
                        "ctr": None,
                        "likes": None,
                        "comments": None,
                        "reposts": None,
                    }
                ],
            },
            "personal": {
                "top10_impressions": 320,
                "top10_interactions": None,
                "period_impressions": None,
                "profile_views": 18,
                "current_followers": 230,
                "highlight": "工艺演示表现较好",
            },
            "warnings": ["个人账号人工筛选样本"],
        }
        content = _native_html(
            analysis,
            {
                "linkedin_personal_screenshots": [
                    {
                        "url": "https://files/linkedin.png",
                        "size": 10,
                        "name": "linkedin.png",
                    }
                ]
            },
        )
        self.assertIn("LinkedIn 个人账号内容表现", content)
        self.assertIn('<img src="https://files/linkedin.png" />', content)
        self.assertIn("该数值不等同于账号报告期全量", content)
        self.assertIn("https://www.linkedin.com/posts/personal", content)

    def test_title_is_stable_for_idempotency(self):
        self.assertEqual(
            _title(sample_analysis()),
            "Example Brand_社媒运营月报_2026.07.01-2026.07.28",
        )


class DeliverJobTests(unittest.TestCase):
    def _job_dir(self, root: str) -> Path:
        job_dir = Path(root)
        inputs = job_dir / "inputs"
        outputs = job_dir / "outputs"
        inputs.mkdir()
        outputs.mkdir()
        analysis = sample_analysis()
        (outputs / "analysis.json").write_text(
            json.dumps(analysis, ensure_ascii=False),
            encoding="utf-8",
        )
        (outputs / "report.docx").write_bytes(b"word")
        (inputs / "facebook.png").write_bytes(b"fb")
        (inputs / "instagram.png").write_bytes(b"ig")
        return job_dir

    def test_existing_same_title_is_updated_without_creating_duplicate(self):
        client = MagicMock()
        client.find_document.return_value = ("doc-existing", "https://docs.popo/existing")
        client.upload.return_value = {
            "url": "https://files/item",
            "size": 10,
            "name": "item",
        }
        with (
            tempfile.TemporaryDirectory() as tmp,
            patch("webapp.popo.popo_status", return_value={"connected": True}),
            patch("webapp.popo.PopoClient", return_value=client),
        ):
            result = deliver_job(self._job_dir(tmp), io.StringIO())
        self.assertFalse(result["created"])
        self.assertEqual(result["url"], "https://docs.popo/existing")
        client.create_document.assert_not_called()
        self.assertEqual(client.upload.call_count, 3)
        client.replace_page.assert_called_once()
        client.delete_document.assert_not_called()

    def test_failed_update_preserves_existing_document(self):
        client = MagicMock()
        client.find_document.return_value = ("doc-existing", "https://docs.popo/existing")
        client.upload.return_value = {
            "url": "https://files/item",
            "size": 10,
            "name": "item",
        }
        client.replace_page.side_effect = PopoDeliveryError("write failed")
        with (
            tempfile.TemporaryDirectory() as tmp,
            patch("webapp.popo.popo_status", return_value={"connected": True}),
            patch("webapp.popo.PopoClient", return_value=client),
        ):
            with self.assertRaisesRegex(PopoDeliveryError, "write failed"):
                deliver_job(self._job_dir(tmp), io.StringIO())
        client.delete_document.assert_not_called()

    def test_failed_write_cleans_new_document(self):
        client = MagicMock()
        client.find_document.return_value = (None, None)
        client.create_document.return_value = ("doc-new", "https://docs.popo/new")
        client.upload.return_value = {
            "url": "https://files/item",
            "size": 10,
            "name": "item",
        }
        client.replace_page.side_effect = PopoDeliveryError("write failed")
        with (
            tempfile.TemporaryDirectory() as tmp,
            patch("webapp.popo.popo_status", return_value={"connected": True}),
            patch("webapp.popo.PopoClient", return_value=client),
        ):
            with self.assertRaisesRegex(PopoDeliveryError, "write failed"):
                deliver_job(self._job_dir(tmp), io.StringIO())
        client.delete_document.assert_called_once_with("doc-new")


if __name__ == "__main__":
    unittest.main()
