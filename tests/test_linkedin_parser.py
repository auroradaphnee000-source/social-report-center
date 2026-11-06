from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from social_report.parser import parse_linkedin_exports


def _save_book(path: Path, sheets: list[tuple[str, list[list[object]]]]) -> None:
    book = Workbook()
    book.remove(book.active)
    for name, rows in sheets:
        sheet = book.create_sheet(name)
        for row in rows:
            sheet.append(row)
    book.save(path)
    book.close()


class LinkedInParserTests(unittest.TestCase):
    def test_accepts_current_english_export_headers(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            content = root / "content.xls"
            followers = root / "followers.xls"
            visitors = root / "visitors.xls"

            # The files may carry an .xls name while their payload is OOXML.
            _save_book(
                content,
                [
                    (
                        "CONTENT",
                        [
                            ["LinkedIn Page analytics export"],
                            [], [], [], [], [], [], [], [],
                            [
                                "Date", "Impressions (organic)",
                                "Impressions (sponsored)",
                                "Unique impressions (organic)",
                                "Clicks (organic)", "Reactions (organic)",
                                "Comments (organic)", "Reposts (organic)",
                            ],
                            ["08/01/2026", 100, 10, 80, 5, 3, 2, 1],
                            ["08/02/2026", 150, 0, 120, 8, 4, 1, 2],
                        ],
                    ),
                    (
                        "POSTS",
                        [
                            [
                                "Post title", "Post link", "Created date",
                                "Post type", "Impressions", "Clicks", "CTR",
                                "Likes", "Comments", "Reposts", "Follows",
                            ],
                            [
                                "Brake test", "https://example.com/post", "08/02/2026",
                                "Organic", 150, 8, 0.053, 4, 1, 2, 1,
                            ],
                        ],
                    ),
                ],
            )
            _save_book(
                followers,
                [
                    (
                        "FOLLOWERS",
                        [
                            ["Date", "Sponsored followers", "Organic followers", "Auto-invited followers", "Total followers"],
                            ["08/01/2026", 1, 2, 3, 6],
                            ["08/02/2026", 0, 4, 1, 5],
                        ],
                    )
                ],
            )
            _save_book(
                visitors,
                [
                    (
                        "VISITORS",
                        [
                            [
                                "Visitor metrics: Date",
                                "Overview page views (desktop)",
                                "Overview page views (mobile)",
                                "Overview page views (total)",
                                "Overview unique visitors (desktop)",
                                "Overview unique visitors (mobile)",
                                "Overview unique visitors (total)",
                            ],
                            ["08/01/2026", 5, 3, 8, 4, 2, 6],
                            ["08/02/2026", 7, 2, 9, 6, 1, 7],
                        ],
                    )
                ],
            )

            result = parse_linkedin_exports(content, followers, visitors)

            self.assertEqual(result["content"]["impressions_organic"], 250)
            self.assertEqual(result["content"]["clicks_organic"], 13)
            self.assertEqual(result["posts"]["count"], 1)
            self.assertEqual(result["visitors"]["page_views_total"], 17)
            self.assertEqual(result["followers"]["organic"], 6)
            self.assertEqual(
                result["content"]["monthly"],
                [
                    {
                        "month": "2026-08",
                        "impressions_organic": 250,
                        "unique_impressions_organic": 200,
                        "clicks_organic": 13,
                        "reactions_organic": 7,
                        "comments_organic": 3,
                        "reposts_organic": 3,
                    }
                ],
            )
            self.assertEqual(
                result["visitors"]["monthly"][0]["page_views_total"], 17
            )
            self.assertEqual(
                result["followers"]["monthly"][0]["organic"], 6
            )

    def test_reports_detected_headers_when_format_is_unknown(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            content = root / "content.xlsx"
            followers = root / "followers.xlsx"
            visitors = root / "visitors.xlsx"
            _save_book(content, [("Sheet1", [["Day", "Mystery metric", "Value"]])])
            _save_book(followers, [("Sheet1", [["Date", "Organic followers", "Total followers"]])])
            _save_book(visitors, [("Sheet1", [["Date", "Page views (total)", "Unique visitors (total)"]])])

            with self.assertRaisesRegex(ValueError, "实际检测到.*Mystery metric"):
                parse_linkedin_exports(content, followers, visitors)


if __name__ == "__main__":
    unittest.main()
