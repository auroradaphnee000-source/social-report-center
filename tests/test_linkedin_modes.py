from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from docx import Document
from openpyxl import Workbook
from PIL import Image

from social_report.parser import parse_linkedin_personal_top10
from social_report.report import generate_report


def make_top10(path: Path):
    book = Workbook()
    sheet = book.active
    sheet.append(
        [
            "帖子发布日期",
            "帖子标题或内容摘要",
            "帖子链接",
            "浏览量/展示量",
            "点赞数",
            "评论数",
            "转发数",
            "点击数",
        ]
    )
    sheet.append(
        [
            "2026-07-02",
            "Application demo",
            "https://www.linkedin.com/posts/demo",
            120,
            4,
            1,
            0,
            "",
        ]
    )
    sheet.append(["2026-07-08", "Factory process", "", 360, 12, 2, 1, 20])
    book.save(path)


def make_image(path: Path):
    Image.new("RGB", (1200, 700), color=(230, 240, 248)).save(path)


def request_for(mode: str) -> dict:
    return {
        "brand": "Personal LinkedIn QA",
        "account_label": "Example",
        "report_type": "monthly",
        "report_start": "2026-07-01",
        "report_end": "2026-07-28",
        "created_by": "Tester",
        "facebook": {
            "posts": 4,
            "views": 100,
            "reach": 80,
            "viewers": 70,
            "interactions": 12,
            "likes": 7,
        },
        "instagram": {
            "posts": 5,
            "views": 120,
            "reach": 90,
            "interactions": 15,
            "likes": 9,
        },
        "facebook_highlight": "应用演示",
        "instagram_highlight": "产品知识",
        "meta": {"source": "manual"},
        "linkedin_mode": mode,
        "linkedin_manual": {
            "period_impressions": None,
            "profile_views": 18,
            "current_followers": 230,
            "highlight": "工艺演示内容表现较好",
        },
        "files": {
            "facebook_screenshot": "facebook.png",
            "instagram_screenshot": "instagram.png",
        },
    }


def document_text(path: Path) -> str:
    doc = Document(path)
    chunks = [paragraph.text for paragraph in doc.paragraphs]
    for table in doc.tables:
        chunks.extend(cell.text for row in table.rows for cell in row.cells)
    return "\n".join(chunks)


class PersonalTop10ParserTests(unittest.TestCase):
    def test_template_rows_are_sorted_and_missing_metrics_stay_unknown(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "top10.xlsx"
            make_top10(path)
            parsed = parse_linkedin_personal_top10(path)
        self.assertEqual(parsed["top10_impressions"], 480)
        self.assertEqual(parsed["posts"][0]["title"], "Factory process")
        self.assertEqual(parsed["posts"][1]["clicks"], None)
        self.assertEqual(parsed["posts"][1]["ctr"], None)
        self.assertEqual(parsed["top10_interactions"], 20)


class LinkedInModeReportTests(unittest.TestCase):
    def _job(self, root: str) -> Path:
        job_dir = Path(root)
        inputs = job_dir / "inputs"
        inputs.mkdir()
        make_image(inputs / "facebook.png")
        make_image(inputs / "instagram.png")
        return job_dir

    def test_personal_table_generates_personal_profile_chapter(self):
        with tempfile.TemporaryDirectory() as tmp:
            job_dir = self._job(tmp)
            make_top10(job_dir / "inputs" / "top10.xlsx")
            request = request_for("personal")
            request["files"]["linkedin_personal_top10"] = "top10.xlsx"
            output, analysis = generate_report(job_dir, request)
            text = document_text(output)
        self.assertEqual(analysis["linkedin"]["mode"], "personal")
        self.assertEqual(analysis["linkedin"]["personal"]["top10_impressions"], 480)
        self.assertIn("LinkedIn 个人账号内容表现", text)
        self.assertIn("人工筛选样本", text)
        self.assertIn("不作为报告期账号全量", text)

    def test_personal_screenshot_only_does_not_invent_numbers(self):
        with tempfile.TemporaryDirectory() as tmp:
            job_dir = self._job(tmp)
            make_image(job_dir / "inputs" / "linkedin.png")
            request = request_for("personal")
            request["files"]["linkedin_personal_screenshots"] = ["linkedin.png"]
            output, analysis = generate_report(job_dir, request)
            text = document_text(output)
        self.assertEqual(analysis["linkedin"]["posts"]["count"], 0)
        self.assertIsNone(
            analysis["linkedin"]["personal"]["top10_impressions"]
        )
        self.assertIn("未对截图数字做 OCR 或推测", text)

    def test_none_mode_keeps_explicit_missing_data_note(self):
        with tempfile.TemporaryDirectory() as tmp:
            job_dir = self._job(tmp)
            output, analysis = generate_report(job_dir, request_for("none"))
            text = document_text(output)
        self.assertEqual(analysis["linkedin"]["mode"], "none")
        self.assertIn("本期未提供 LinkedIn", text)
        self.assertIn("不将缺失数据记为 0", text)

    def test_word_framework_changes_by_report_type(self):
        cases = {
            "monthly": {
                "included": ["1.2　自然内容排行", "1.3　当月成效分析", "本周期发生了什么"],
                "excluded": ["月度自然内容趋势"],
            },
            "quarterly": {
                "included": ["1.2　自然内容排行", "1.3　阶段成效分析", "这一阶段验证了什么"],
                "excluded": ["月度自然内容趋势"],
            },
            "annual": {
                "included": ["1.2　月度自然内容趋势", "1.3　自然内容排行", "1.4　高表现内容分析", "下一年度策略"],
                "excluded": [],
            },
        }
        for report_type, expected in cases.items():
            with self.subTest(report_type=report_type):
                with tempfile.TemporaryDirectory() as tmp:
                    job_dir = self._job(tmp)
                    request = request_for("none")
                    request["report_type"] = report_type
                    output, _ = generate_report(job_dir, request)
                    text = document_text(output)
                self.assertIn("第四部分　综合结论与下阶段优化建议", text)
                for item in expected["included"]:
                    self.assertIn(item, text)
                for item in expected["excluded"]:
                    self.assertNotIn(item, text)

    def test_annual_word_embeds_meta_monthly_charts_and_editable_tables(self):
        with tempfile.TemporaryDirectory() as tmp:
            job_dir = self._job(tmp)
            request = request_for("none")
            request["report_type"] = "annual"
            request["facebook"]["monthly"] = [
                {"month": "2026-06", "views": 40, "daily_reach_sum": 28},
                {"month": "2026-07", "views": 60, "daily_reach_sum": 42},
            ]
            request["instagram"]["monthly"] = [
                {"month": "2026-06", "views": 50, "daily_reach_sum": 32},
                {"month": "2026-07", "views": 70, "daily_reach_sum": 48},
            ]
            output, _ = generate_report(job_dir, request)
            doc = Document(output)
            text = document_text(output)

            self.assertTrue((job_dir / "charts" / "facebook_monthly_trend.png").is_file())
            self.assertTrue((job_dir / "charts" / "instagram_monthly_trend.png").is_file())
            self.assertGreaterEqual(len(doc.inline_shapes), 4)
            self.assertIn("Facebook 月度趋势明细（POPO 中可编辑）", text)
            self.assertIn("Instagram 月度趋势明细（POPO 中可编辑）", text)
            self.assertIn("日度触达累计", text)
            self.assertNotIn("Published posts", text)
            self.assertNotIn("请重新同步平台数据后再生成", text)


if __name__ == "__main__":
    unittest.main()
