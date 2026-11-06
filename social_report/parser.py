from __future__ import annotations

import re
import unicodedata
from datetime import datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any

import xlrd
from openpyxl import load_workbook


DATE_FORMATS = ("%m/%d/%Y", "%Y-%m-%d", "%Y/%m/%d")


def _compact(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _normalize_header(value: Any) -> str:
    """Normalize localized/export-version header differences for matching."""
    text = unicodedata.normalize("NFKC", _compact(value)).casefold()
    text = text.replace("\ufeff", "").replace("\u200b", "")
    return re.sub(r"[\s_\-–—:：/\\()\[\]{}]+", "", text)


def _header_preview(sheets: list[dict[str, Any]]) -> str:
    previews: list[str] = []
    for sheet in sheets:
        candidates = []
        for row in sheet["rows"][:25]:
            values = [_compact(cell) for cell in row if _compact(cell)]
            if len(values) >= 3:
                candidates.append("、".join(values[:12]))
        if candidates:
            previews.append(f"{sheet['name']}：{candidates[0]}")
    return "；".join(previews[:4]) or "未识别到有效表头"


def _find_aliased_sheet(
    sheets: list[dict[str, Any]],
    aliases: dict[str, tuple[str, ...]],
    required_keys: tuple[str, ...],
    label: str,
) -> tuple[dict[str, Any], int, dict[str, int]]:
    normalized_aliases = {
        key: {_normalize_header(alias) for alias in names}
        for key, names in aliases.items()
    }
    for sheet in sheets:
        for row_idx, row in enumerate(sheet["rows"][:25]):
            headers = [_normalize_header(cell) for cell in row]
            columns: dict[str, int] = {}
            for key, names in normalized_aliases.items():
                for column_idx, header in enumerate(headers):
                    if header and header in names:
                        columns[key] = column_idx
                        break
            if all(key in columns for key in required_keys):
                return sheet, row_idx, columns
    expected = "、".join(required_keys)
    raise ValueError(
        f"无法识别 LinkedIn {label}的表头（缺少：{expected}）。"
        f"实际检测到：{_header_preview(sheets)}"
    )


def _sum_aliased(
    rows: list[list[Any]], columns: dict[str, int], key: str, *, required: bool = False
) -> float:
    index = columns.get(key, -1)
    if index < 0 and required:
        raise ValueError(f"LinkedIn 导出表缺少必需字段：{key}")
    return _sum(rows, index)


def _number(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").replace("%", "").strip()
    try:
        return float(text)
    except ValueError:
        return 0.0


def _date(value: Any, datemode: int = 0) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)) and value:
        try:
            return xlrd.xldate_as_datetime(value, datemode)
        except (ValueError, TypeError):
            return None
    text = _compact(value)
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _read_workbook(path: Path) -> list[dict[str, Any]]:
    magic = path.read_bytes()[:8]
    sheets: list[dict[str, Any]] = []
    if magic.startswith(b"\xD0\xCF\x11\xE0"):
        book = xlrd.open_workbook(path)
        for sheet in book.sheets():
            rows = [
                [sheet.cell_value(row, col) for col in range(sheet.ncols)]
                for row in range(sheet.nrows)
            ]
            sheets.append({"name": sheet.name, "rows": rows, "datemode": book.datemode})
        return sheets
    if magic.startswith(b"PK"):
        # Read from bytes so an OOXML workbook exported with an .xls filename
        # is accepted based on its real payload rather than its extension.
        book = load_workbook(BytesIO(path.read_bytes()), read_only=True, data_only=True)
        try:
            for sheet in book.worksheets:
                rows = [list(row) for row in sheet.iter_rows(values_only=True)]
                sheets.append({"name": sheet.title, "rows": rows, "datemode": 0})
        finally:
            book.close()
        return sheets
    raise ValueError(f"{path.name} 不是可识别的 .xls/.xlsx 文件")


def _find_sheet(
    sheets: list[dict[str, Any]], required_headers: tuple[str, ...]
) -> tuple[dict[str, Any], int, list[str]]:
    for sheet in sheets:
        for row_idx, row in enumerate(sheet["rows"][:8]):
            headers = [_compact(cell) for cell in row]
            if all(header in headers for header in required_headers):
                return sheet, row_idx, headers
    raise ValueError(f"找不到包含这些字段的工作表：{', '.join(required_headers)}")


def _column(headers: list[str], *names: str, required: bool = True) -> int:
    for name in names:
        if name in headers:
            return headers.index(name)
    if required:
        raise ValueError(f"导出表缺少字段：{' / '.join(names)}")
    return -1


def _sum(rows: list[list[Any]], index: int) -> float:
    if index < 0:
        return 0.0
    return sum(_number(row[index]) if index < len(row) else 0.0 for row in rows)


def _dated_rows(
    sheet: dict[str, Any], header_row: int, date_col: int = 0
) -> list[tuple[datetime, list[Any]]]:
    result = []
    for row in sheet["rows"][header_row + 1 :]:
        value = row[date_col] if date_col < len(row) else None
        parsed = _date(value, sheet["datemode"])
        if parsed:
            result.append((parsed, row))
    result.sort(key=lambda item: item[0])
    return result


def _monthly_aliased_sums(
    dated_rows: list[tuple[datetime, list[Any]]],
    columns: dict[str, int],
    keys: tuple[str, ...],
) -> list[dict[str, Any]]:
    """Keep the verifiable daily export series as calendar-month totals."""
    monthly: dict[str, dict[str, Any]] = {}
    for row_date, row in dated_rows:
        month = row_date.strftime("%Y-%m")
        values = monthly.setdefault(month, {"month": month})
        for key in keys:
            index = columns.get(key, -1)
            if index < 0:
                values.setdefault(key, None)
                continue
            current = values.get(key)
            number = _number(row[index] if index < len(row) else None)
            values[key] = _int((current or 0) + number)
    return [monthly[month] for month in sorted(monthly)]


def _int(value: float) -> int:
    return int(round(value))


def _find_flexible_sheet(
    sheets: list[dict[str, Any]], aliases: dict[str, tuple[str, ...]]
) -> tuple[dict[str, Any], int, dict[str, int]]:
    for sheet in sheets:
        for row_idx, row in enumerate(sheet["rows"][:10]):
            headers = [_compact(cell).lower() for cell in row]
            columns: dict[str, int] = {}
            for key, names in aliases.items():
                for name in names:
                    lowered = name.lower()
                    if lowered in headers:
                        columns[key] = headers.index(lowered)
                        break
            if {"date", "title", "impressions"}.issubset(columns):
                return sheet, row_idx, columns
    raise ValueError(
        "LinkedIn Top 10 表缺少必填字段：帖子发布日期、帖子标题或内容摘要、浏览量/展示量"
    )


def parse_linkedin_personal_top10(path: Path) -> dict[str, Any]:
    """Parse the standard/manual Top 10 workbook used for personal profiles."""
    sheet, header_row, columns = _find_flexible_sheet(
        _read_workbook(path),
        {
            "date": (
                "帖子发布日期",
                "发布日期",
                "创建日期",
                "日期",
                "post date",
                "date",
            ),
            "title": (
                "帖子标题或内容摘要",
                "帖子标题",
                "内容摘要",
                "内容主题",
                "动态标题",
                "标题",
                "post title",
                "title",
            ),
            "url": ("帖子链接", "动态链接", "链接", "post url", "url"),
            "impressions": (
                "浏览量/展示量",
                "浏览量",
                "展示量",
                "展示次数",
                "impressions",
                "views",
            ),
            "clicks": ("点击数", "点击次数", "clicks"),
            "likes": ("点赞数", "点赞量", "点赞", "回应量", "reactions", "likes"),
            "comments": ("评论数", "评论", "comments"),
            "reposts": ("转发数", "转发", "reposts", "shares"),
        },
    )

    def cell(row: list[Any], key: str) -> Any:
        index = columns.get(key, -1)
        return row[index] if index >= 0 and index < len(row) else None

    posts: list[dict[str, Any]] = []
    for row in sheet["rows"][header_row + 1 :]:
        created = _date(cell(row, "date"), sheet["datemode"])
        title = _compact(cell(row, "title"))
        impressions_value = cell(row, "impressions")
        if not created and not title and impressions_value in (None, ""):
            continue
        if not created or not title:
            raise ValueError("LinkedIn Top 10 表中存在缺少发布日期或内容摘要的记录")
        impressions = _number(impressions_value)
        if impressions < 0:
            raise ValueError("LinkedIn Top 10 表中的浏览量/展示量不能为负数")

        optional_metrics: dict[str, int | None] = {}
        for key in ("clicks", "likes", "comments", "reposts"):
            raw = cell(row, key)
            if raw in (None, ""):
                optional_metrics[key] = None
                continue
            number = _number(raw)
            if number < 0:
                raise ValueError(f"LinkedIn Top 10 表中的{key}不能为负数")
            optional_metrics[key] = _int(number)

        clicks = optional_metrics["clicks"]
        posts.append(
            {
                "title": title,
                "url": _compact(cell(row, "url")),
                "created_date": created.date().isoformat(),
                "content_type": "人工筛选",
                "impressions": _int(impressions),
                "views": _int(impressions),
                "clicks": clicks,
                "ctr": (
                    clicks / impressions
                    if clicks is not None and impressions
                    else None
                ),
                "likes": optional_metrics["likes"],
                "comments": optional_metrics["comments"],
                "reposts": optional_metrics["reposts"],
                "follows": None,
            }
        )
    if not posts:
        raise ValueError("LinkedIn Top 10 表没有有效帖子记录")

    posts.sort(key=lambda post: post["impressions"], reverse=True)
    posts = posts[:10]
    interaction_values = [
        int(post[key])
        for post in posts
        for key in ("likes", "comments", "reposts")
        if post.get(key) is not None
    ]
    return {
        "source_file": path.name,
        "posts": posts,
        "top10_impressions": sum(post["impressions"] for post in posts),
        "top10_interactions": (
            sum(interaction_values) if interaction_values else None
        ),
    }


def parse_linkedin_exports(
    content_path: Path, followers_path: Path, visitors_path: Path
) -> dict[str, Any]:
    content_sheets = _read_workbook(content_path)
    follower_sheets = _read_workbook(followers_path)
    visitor_sheets = _read_workbook(visitors_path)

    metrics_sheet, metrics_header_row, mh = _find_aliased_sheet(
        content_sheets,
        {
            "date": ("日期", "Date"),
            "impressions_organic": (
                "展示量 (原生内容)", "展示次数 (原生内容)", "展示量 (自然内容)",
                "展示次数 (自然内容)", "展示次数 (自然)", "自然展示次数",
                "Impressions (organic)", "Organic impressions",
            ),
            "impressions_sponsored": (
                "展示量 (付费内容)", "展示次数 (付费内容)", "展示次数 (推广)",
                "付费展示次数", "Impressions (sponsored)", "Sponsored impressions",
            ),
            "unique_impressions_organic": (
                "独立展示量 (原生内容)", "独立展示次数 (原生内容)",
                "独立展示次数 (自然)", "自然独立展示次数",
                "Unique impressions (organic)", "Organic unique impressions",
            ),
            "clicks_organic": (
                "点击量 (原生内容)", "点击次数 (原生内容)", "点击次数 (自然)",
                "自然点击次数", "Clicks (organic)", "Organic clicks",
            ),
            "reactions_organic": (
                "回应量 (原生内容)", "回应次数 (原生内容)", "反应次数 (自然)",
                "自然回应次数", "Reactions (organic)", "Organic reactions",
            ),
            "comments_organic": (
                "评论量 (原生内容)", "评论次数 (原生内容)", "评论次数 (自然)",
                "自然评论次数", "Comments (organic)", "Organic comments",
            ),
            "reposts_organic": (
                "转发 (原生)", "转发量 (原生内容)", "转发次数 (原生内容)",
                "转发次数 (自然)", "自然转发次数", "Reposts (organic)",
                "Shares (organic)", "Organic reposts", "Organic shares",
            ),
        },
        ("date", "impressions_organic", "clicks_organic"),
        "内容趋势表",
    )
    metric_rows_with_dates = _dated_rows(
        metrics_sheet, metrics_header_row, mh["date"]
    )
    if not metric_rows_with_dates:
        raise ValueError("LinkedIn 内容表没有有效的每日数据")
    metric_rows = [row for _, row in metric_rows_with_dates]

    impressions_organic = _sum_aliased(
        metric_rows, mh, "impressions_organic", required=True
    )
    impressions_sponsored = _sum_aliased(metric_rows, mh, "impressions_sponsored")
    clicks_organic = _sum_aliased(metric_rows, mh, "clicks_organic", required=True)
    reactions_organic = _sum_aliased(metric_rows, mh, "reactions_organic")
    comments_organic = _sum_aliased(metric_rows, mh, "comments_organic")
    reposts_organic = _sum_aliased(metric_rows, mh, "reposts_organic")
    engagement_actions = (
        clicks_organic + reactions_organic + comments_organic + reposts_organic
    )

    content = {
        "start": metric_rows_with_dates[0][0].date().isoformat(),
        "end": metric_rows_with_dates[-1][0].date().isoformat(),
        "days": len(metric_rows_with_dates),
        "impressions_organic": _int(impressions_organic),
        "impressions_sponsored": _int(impressions_sponsored),
        "unique_impressions_daily_sum": _int(
            _sum_aliased(metric_rows, mh, "unique_impressions_organic")
        ),
        "clicks_organic": _int(clicks_organic),
        "reactions_organic": _int(reactions_organic),
        "comments_organic": _int(comments_organic),
        "reposts_organic": _int(reposts_organic),
        "engagement_rate": (
            engagement_actions / impressions_organic if impressions_organic else 0.0
        ),
        "monthly": _monthly_aliased_sums(
            metric_rows_with_dates,
            mh,
            (
                "impressions_organic",
                "unique_impressions_organic",
                "clicks_organic",
                "reactions_organic",
                "comments_organic",
                "reposts_organic",
            ),
        ),
    }

    posts_sheet, posts_header_row, ph = _find_aliased_sheet(
        content_sheets,
        {
            "title": ("动态标题", "帖子标题", "内容标题", "Post title", "Update title"),
            "url": ("动态链接", "帖子链接", "Post link", "Post URL", "Update link"),
            "created": ("创建日期", "发布日期", "发布时间", "Created date", "Post date"),
            "post_type": ("动态类型", "帖子类型", "Post type", "Update type"),
            "content_type": ("内容类型", "Content type"),
            "impressions": ("展示次数", "展示量", "Impressions"),
            "views": ("浏览量", "视频观看次数", "Video views", "Views"),
            "clicks": ("点击次数", "点击量", "Clicks"),
            "ctr": ("点击率 (CTR)", "点击率", "Click-through rate (CTR)", "CTR"),
            "likes": ("点赞量", "点赞次数", "回应量", "Likes", "Reactions"),
            "comments": ("评论", "评论量", "Comments"),
            "reposts": ("转发", "转发量", "Reposts", "Shares"),
            "follows": ("关注量", "关注次数", "Follows"),
        },
        ("title", "created", "impressions"),
        "内容明细表",
    )
    created_col = ph["created"]
    type_col = ph.get("post_type", -1)
    posts = []
    for created, row in _dated_rows(posts_sheet, posts_header_row, created_col):
        post_type = _compact(row[type_col]) if type_col >= 0 and type_col < len(row) else ""
        if post_type and _normalize_header(post_type) not in {
            _normalize_header(value) for value in ("原生", "自然", "organic")
        }:
            continue

        def cell(key: str) -> Any:
            idx = ph.get(key, -1)
            return row[idx] if idx >= 0 and idx < len(row) else ""

        impressions = _number(cell("impressions"))
        clicks = _number(cell("clicks"))
        ctr = _number(cell("ctr"))
        if ctr > 1:
            ctr /= 100
        if not ctr and impressions:
            ctr = clicks / impressions
        posts.append(
            {
                "title": _compact(cell("title")),
                "url": _compact(cell("url")),
                "created_date": created.date().isoformat(),
                "content_type": _compact(cell("content_type") or cell("post_type")),
                "impressions": _int(impressions),
                "views": _int(_number(cell("views"))),
                "clicks": _int(clicks),
                "ctr": ctr,
                "likes": _int(_number(cell("likes"))),
                "comments": _int(_number(cell("comments"))),
                "reposts": _int(_number(cell("reposts"))),
                "follows": _int(_number(cell("follows"))),
            }
        )
    posts.sort(key=lambda post: (post["impressions"], post["clicks"]), reverse=True)

    visitor_sheet, visitor_header_row, vh = _find_aliased_sheet(
        visitor_sheets,
        {
            "date": ("日期", "Date", "Visitor metrics: Date"),
            "page_views_web": (
                "主页访问总量 (网页版)",
                "主页浏览量 (桌面端)",
                "Page views (desktop)",
                "Overview page views (desktop)",
                "Desktop page views",
            ),
            "page_views_mobile": (
                "主页访问总量 (移动版)",
                "主页浏览量 (移动端)",
                "Page views (mobile)",
                "Overview page views (mobile)",
                "Mobile page views",
            ),
            "page_views_total": (
                "主页访问总量 (总计)",
                "主页浏览量 (总计)",
                "Page views (total)",
                "Overview page views (total)",
                "Total page views",
            ),
            "unique_web": (
                "主页独立访问总量 (网页版)",
                "独立访客 (桌面端)",
                "Unique visitors (desktop)",
                "Overview unique visitors (desktop)",
                "Desktop unique visitors",
            ),
            "unique_mobile": (
                "主页独立访问总量 (移动版)",
                "独立访客 (移动端)",
                "Unique visitors (mobile)",
                "Overview unique visitors (mobile)",
                "Mobile unique visitors",
            ),
            "unique_total": (
                "主页独立访问总量 (总计)",
                "独立访客 (总计)",
                "Unique visitors (total)",
                "Overview unique visitors (total)",
                "Total unique visitors",
            ),
        },
        ("date", "page_views_total", "unique_total"),
        "访客表",
    )
    visitor_rows_with_dates = _dated_rows(
        visitor_sheet, visitor_header_row, vh["date"]
    )
    if not visitor_rows_with_dates:
        raise ValueError("LinkedIn 访客表没有有效的每日数据")
    visitor_rows = [row for _, row in visitor_rows_with_dates]
    visitors = {
        "start": visitor_rows_with_dates[0][0].date().isoformat(),
        "end": visitor_rows_with_dates[-1][0].date().isoformat(),
        "page_views_web": _int(_sum_aliased(visitor_rows, vh, "page_views_web")),
        "page_views_mobile": _int(_sum_aliased(visitor_rows, vh, "page_views_mobile")),
        "page_views_total": _int(_sum_aliased(visitor_rows, vh, "page_views_total", required=True)),
        "unique_web": _int(_sum_aliased(visitor_rows, vh, "unique_web")),
        "unique_mobile": _int(_sum_aliased(visitor_rows, vh, "unique_mobile")),
        "unique_total": _int(_sum_aliased(visitor_rows, vh, "unique_total", required=True)),
        "monthly": _monthly_aliased_sums(
            visitor_rows_with_dates,
            vh,
            ("page_views_total", "unique_total"),
        ),
    }

    follower_sheet, follower_header_row, fh = _find_aliased_sheet(
        follower_sheets,
        {
            "date": ("日期", "Date"),
            "sponsored": ("推广所得关注者", "付费关注者", "Sponsored followers", "Paid followers"),
            "organic": ("主动型关注者", "自然关注者", "Organic followers"),
            "auto_invited": ("已自动邀请关注者", "自动邀请关注者", "Auto-invited followers", "Auto invited followers"),
            "total": ("关注者总数", "总关注者", "Total followers"),
        },
        ("date", "organic", "total"),
        "关注者表",
    )
    follower_rows_with_dates = _dated_rows(
        follower_sheet, follower_header_row, fh["date"]
    )
    if not follower_rows_with_dates:
        raise ValueError("LinkedIn 关注者表没有有效的每日数据")
    follower_rows = [row for _, row in follower_rows_with_dates]
    followers = {
        "start": follower_rows_with_dates[0][0].date().isoformat(),
        "end": follower_rows_with_dates[-1][0].date().isoformat(),
        "sponsored": _int(_sum_aliased(follower_rows, fh, "sponsored")),
        "organic": _int(_sum_aliased(follower_rows, fh, "organic", required=True)),
        "auto_invited": _int(_sum_aliased(follower_rows, fh, "auto_invited")),
        "total": _int(_sum_aliased(follower_rows, fh, "total", required=True)),
        "monthly": _monthly_aliased_sums(
            follower_rows_with_dates,
            fh,
            ("sponsored", "organic", "auto_invited"),
        ),
    }

    ranges = {
        (content["start"], content["end"]),
        (visitors["start"], visitors["end"]),
        (followers["start"], followers["end"]),
    }
    warnings = []
    if len(ranges) > 1:
        warnings.append("LinkedIn 三份导出的日期范围不完全一致，报告已分别按各表实际范围汇总。")
    if impressions_sponsored:
        warnings.append("检测到 LinkedIn 付费展示；报告仅采用原生内容指标。")

    return {
        "mode": "company",
        "source_files": {
            "content": content_path.name,
            "followers": followers_path.name,
            "visitors": visitors_path.name,
        },
        "content": content,
        "posts": {
            "count": len(posts),
            "top_impressions": posts[:8],
            "all": posts,
        },
        "visitors": visitors,
        "followers": followers,
        "warnings": warnings,
    }
