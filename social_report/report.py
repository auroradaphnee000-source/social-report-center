from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

from .parser import parse_linkedin_exports, parse_linkedin_personal_top10


NAVY = "111111"
BLUE = "111111"
LIGHT_BLUE = "F7F7F7"
PALE_BLUE = "F2F2F2"
INK = "111111"
MUTED = "5F6368"
WHITE = "FFFFFF"
LINE = "D9D9D9"
GREEN = "2D8C68"
ORANGE = "D27A25"
RED = "FF3030"
FONT = "Microsoft YaHei"
REPORT_TYPE_TITLES = {
    "monthly": "社媒运营月报",
    "quarterly": "社媒运营季度报",
    "annual": "社媒运营年报",
    "custom": "社媒运营报告",
}
FIXED_ENDING = (
    "另：如果社媒平台没有推广引流的计划的话，建议把社媒作为网站内容附加和增加客户信任度的工具手段，"
    "做好内容输出和发帖频率，数据积累和转化还是网站为主。"
)
FIXED_ANALYSIS_PRINCIPLE = (
    "分析原则：本报告以自然内容运营为分析对象，不把总曝光或粉丝增长作为唯一标准。"
    "重点观察账号内部不同主题的相对表现、内容是否建立物流服务认知，以及浏览能否进一步"
    "转化为主页访问、关注、链接点击或询盘。"
)
REPORT_MODE_PROFILES = {
    "monthly": {
        "framework": "period",
        "question": "本周期发生了什么",
        "analysis_heading": "当月成效分析",
        "logic": "数据 → 现象 → 初步判断 → 下阶段动作",
        "conclusion_scope": "本周期结论",
        "recommendation_scope": "下阶段建议",
    },
    "quarterly": {
        "framework": "period",
        "question": "这一阶段验证了什么",
        "analysis_heading": "阶段成效分析",
        "logic": "数据 → 现象 → 阶段规律 → 下一季度动作",
        "conclusion_scope": "阶段性结论",
        "recommendation_scope": "下一季度策略",
    },
    "annual": {
        "framework": "annual",
        "question": "这一年积累了什么，以及下一年应该怎么做",
        "analysis_heading": "高表现内容分析",
        "logic": "数据 → 长期趋势 → 内容模型 → 平台价值 → 账号阶段 → 年度战略",
        "conclusion_scope": "年度核心结论",
        "recommendation_scope": "下一年度策略",
    },
    "custom": {
        "framework": "period",
        "question": "本阶段发生了什么",
        "analysis_heading": "阶段成效分析",
        "logic": "数据 → 现象 → 初步判断 → 下阶段动作",
        "conclusion_scope": "阶段性结论",
        "recommendation_scope": "下阶段建议",
    },
}


def report_mode_profile(report_type: str) -> dict[str, str]:
    """Return the fixed framework and analysis depth for a report type."""
    key = str(report_type or "monthly").strip().lower()
    return REPORT_MODE_PROFILES.get(key, REPORT_MODE_PROFILES["custom"])


def analysis_principle(report_type: str) -> str:
    profile = report_mode_profile(report_type)
    return (
        f"{FIXED_ANALYSIS_PRINCIPLE} 本次报告重点回答“{profile['question']}”，"
        f"分析链路固定为：{profile['logic']}。"
    )


def report_recommendations(report_type: str) -> list[tuple[str, str]]:
    profile = report_mode_profile(report_type)
    scope = profile["recommendation_scope"]
    return [
        (
            f"{scope}一｜持续深化高表现内容方向：",
            "优先延续高浏览内容的主题，用连续栏目而不是单次孤立发帖积累认知。",
        ),
        (
            f"{scope}二｜从“介绍能力”升级为“提供证据”：",
            "增加真实场景、流程、规格、案例和结果，让潜在客户更直观地判断合作能力。",
        ),
        (
            f"{scope}三｜优化平台内容形式：",
            "Facebook 使用案例长图、清单和行业传播；Instagram 使用 Reels、步骤卡片、现场特写和短字幕。",
        ),
        (
            f"{scope}四｜增强内容转化设计：",
            "每条内容设置一个明确 CTA，如索取目录、获取方案、查看案例、访问官网或私信咨询，避免多个动作分散注意力。",
        ),
    ]


def _xml(tag: str, **attrs: str) -> OxmlElement:
    el = OxmlElement(tag)
    for key, value in attrs.items():
        el.set(qn(key), value)
    return el


def _set_run_font(run, size: float = 10.5, bold: bool = False, color: str = INK):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)


def _set_cell_shading(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    old = tc_pr.find(qn("w:shd"))
    if old is not None:
        tc_pr.remove(old)
    tc_pr.append(_xml("w:shd", **{"w:fill": fill, "w:val": "clear"}))


def _set_cell_margins(cell, top=90, start=110, bottom=90, end=110):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _set_table_borders(table, color: str = LINE, size: int = 6):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        node = borders.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), str(size))
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)


def _set_fixed_table(table, widths_in: list[float]):
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_pr = table._tbl.tblPr
    layout = tbl_pr.first_child_found_in("w:tblLayout")
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")
    widths_twips = [round(width * 1440) for width in widths_in]
    total = sum(widths_twips)
    tbl_w = tbl_pr.first_child_found_in("w:tblW")
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_ind = tbl_pr.first_child_found_in("w:tblInd")
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), "110")
    tbl_ind.set(qn("w:type"), "dxa")
    grid_cols = list(table._tbl.tblGrid)
    for idx, grid_col in enumerate(grid_cols):
        width_twips = widths_twips[min(idx, len(widths_twips) - 1)]
        grid_col.set(qn("w:w"), str(width_twips))
    for row in table.rows:
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for idx, cell in enumerate(row.cells):
            width = Inches(widths_in[min(idx, len(widths_in) - 1)])
            cell.width = width
            tc_w = cell._tc.get_or_add_tcPr().first_child_found_in("w:tcW")
            tc_w.set(qn("w:w"), str(widths_twips[min(idx, len(widths_twips) - 1)]))
            tc_w.set(qn("w:type"), "dxa")
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            _set_cell_margins(cell)


def _paragraph_border(paragraph, color: str = BLUE, size: int = 18):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "5")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)


def _add_field(paragraph, instruction: str):
    run = paragraph.add_run()
    fld_char = OxmlElement("w:fldChar")
    fld_char.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_char, instr, separate, text, end])
    _set_run_font(run, 8.5, color=MUTED)


def _format_date(value: str) -> str:
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%Y.%m.%d")
    except ValueError:
        return value


def _format_number(value: Any) -> str:
    if value is None or value == "":
        return "—"
    return f"{int(round(float(value))):,}"


def _format_percent(value: float | None) -> str:
    if value is None:
        return "—"
    return f"{value * 100:.2f}%"


def _safe_name(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", value).strip(" .")
    return cleaned[:80] or "Social_Report"


def _add_heading(doc: Document, text: str, level: int = 1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.space_before = Pt(8 if level == 1 else 5)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    _set_run_font(run, 17 if level == 1 else 12.5, True, INK)
    return p


def _add_body(doc: Document, text: str, *, color: str = INK, bold: bool = False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.35
    run = p.add_run(text)
    _set_run_font(run, 10.5, bold, color)
    return p


def _add_lead(doc: Document, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(7)
    p.paragraph_format.line_spacing = 1.3
    p_pr = p._p.get_or_add_pPr()
    p_pr.append(_xml("w:shd", **{"w:fill": WHITE, "w:val": "clear"}))
    p_pr.append(
        _xml(
            "w:ind",
            **{"w:left": "160", "w:right": "160", "w:firstLine": "0"},
        )
    )
    p_bdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "12")
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), LINE)
    p_bdr.append(left)
    p_pr.append(p_bdr)
    run = p.add_run(text)
    _set_run_font(run, 10.5, False, MUTED)
    return p


def _add_bullet(doc: Document, label: str, text: str):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(0.65)
    p.paragraph_format.first_line_indent = Cm(-0.25)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.25
    run = p.add_run(label)
    _set_run_font(run, 10.3, True, INK)
    run = p.add_run(text)
    _set_run_font(run, 10.3, False, INK)
    return p


def _add_detail_table(
    doc: Document,
    rows: list[tuple[str, str, str]],
):
    table = doc.add_table(rows=1, cols=3)
    widths = [1.55, 1.35, 3.8]
    headers = ("指标", "本期数据", "说明")
    for index, label in enumerate(headers):
        cell = table.rows[0].cells[index]
        _set_cell_shading(cell, PALE_BLUE)
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = (
            WD_ALIGN_PARAGRAPH.CENTER if index == 1 else WD_ALIGN_PARAGRAPH.LEFT
        )
        run = p.add_run(label)
        _set_run_font(run, 9.5, True, INK)
    for metric, value, note in rows:
        cells = table.add_row().cells
        values = (metric, value, note)
        for index, text in enumerate(values):
            cells[index].text = ""
            p = cells[index].paragraphs[0]
            p.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER if index == 1 else WD_ALIGN_PARAGRAPH.LEFT
            )
            run = p.add_run(str(text))
            _set_run_font(run, 9.2, index in (0, 1), INK)
    _set_fixed_table(table, widths)
    _set_table_borders(table, LINE, 4)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return table


def _add_kpi_table(doc: Document, items: list[tuple[str, str]], columns: int = 4):
    rows = (len(items) + columns - 1) // columns
    table = doc.add_table(rows=rows, cols=columns)
    widths = [6.7 / columns] * columns
    _set_fixed_table(table, widths)
    _set_table_borders(table, "D8E4ED", 6)
    for idx in range(rows * columns):
        cell = table.cell(idx // columns, idx % columns)
        _set_cell_shading(cell, PALE_BLUE if idx % 2 == 0 else WHITE)
        cell.text = ""
        if idx >= len(items):
            continue
        label, value = items[idx]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(value)
        _set_run_font(r, 15, True, NAVY)
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(label)
        _set_run_font(r2, 8.5, False, MUTED)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)
    return table


def _fit_image(path: Path, max_width_in: float = 6.7, max_height_in: float = 4.15):
    with Image.open(path) as image:
        width_px, height_px = image.size
    ratio = min(max_width_in / width_px, max_height_in / height_px)
    return width_px * ratio, height_px * ratio


def _add_image(doc: Document, path: Path, alt: str):
    width, height = _fit_image(path)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run()
    run.add_picture(str(path), width=Inches(width), height=Inches(height))
    doc_pr = run._r.xpath(".//wp:docPr")
    if doc_pr:
        doc_pr[0].set("title", alt)
        doc_pr[0].set("descr", alt)
    return p


def _add_caption(doc: Document, text: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(7)
    r = p.add_run(text)
    _set_run_font(r, 8.3, False, MUTED)
    return p


def _chart_font(size: int, *, bold: bool = False):
    candidates = [
        Path("C:/Windows/Fonts/msyhbd.ttc" if bold else "C:/Windows/Fonts/msyh.ttc"),
        Path(
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
            if bold
            else "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"
        ),
        Path(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
            if bold
            else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
        ),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size=size)
    return ImageFont.load_default()


def _chart_number(value: int | float | None) -> str:
    if value is None:
        return "-"
    number = float(value)
    if abs(number) >= 1_000_000:
        return f"{number / 1_000_000:.1f}m"
    if abs(number) >= 1_000:
        return f"{number / 1_000:.1f}k"
    return f"{number:,.0f}"


def _normalize_monthly_rows(value: Any) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if isinstance(value, dict):
        for month, metrics in value.items():
            row = {"month": str(month)}
            if isinstance(metrics, dict):
                row.update(metrics)
            rows.append(row)
    elif isinstance(value, list):
        rows = [dict(row) for row in value if isinstance(row, dict)]
    return sorted(
        [row for row in rows if re.fullmatch(r"\d{4}-\d{2}", str(row.get("month") or ""))],
        key=lambda row: row["month"],
    )


def _render_monthly_chart(
    path: Path,
    platform: str,
    rows: list[dict[str, Any]],
    specs: list[tuple[str, str, str, str]],
) -> None:
    available = [
        spec
        for spec in specs
        if any(isinstance(row.get(spec[0]), (int, float)) for row in rows)
    ]
    width = 1400
    height = 560
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    title_font = _chart_font(30, bold=True)
    subtitle_font = _chart_font(18)
    legend_font = _chart_font(18, bold=True)
    axis_font = _chart_font(16)
    draw.text(
        (60, 24),
        f"{platform} Organic Content Monthly Trend",
        font=title_font,
        fill="#111111",
    )
    months = [str(row["month"]) for row in rows]
    draw.text(
        (60, 66),
        f"{months[0]} - {months[-1]} | "
        + " and ".join(spec[2] for spec in available),
        font=subtitle_font,
        fill="#677386",
    )
    chart_left, chart_right = 150, width - 65
    plot_top, plot_bottom = 142, 450
    all_values = [
        float(row[key])
        for key, _label_cn, _label_en, _color in available
        for row in rows
        if isinstance(row.get(key), (int, float))
    ]
    maximum = max(all_values, default=0)
    scale_max = maximum if maximum > 0 else 1
    for grid_index in range(6):
        ratio = grid_index / 5
        y = plot_bottom - (plot_bottom - plot_top) * ratio
        draw.line((chart_left, y, chart_right, y), fill="#E2E7ED", width=2)
        label = _chart_number(scale_max * ratio)
        box = draw.textbbox((0, 0), label, font=axis_font)
        draw.text(
            (chart_left - 16 - (box[2] - box[0]), y - 9),
            label,
            font=axis_font,
            fill="#677386",
        )
    draw.line((chart_left, plot_top, chart_left, plot_bottom), fill="#9AA7B6", width=2)
    draw.line((chart_left, plot_bottom, chart_right, plot_bottom), fill="#9AA7B6", width=2)

    legend_x = chart_right - 390
    for index, (_key, _label_cn, label_en, color) in enumerate(available):
        x = legend_x + index * 210
        draw.line((x, 82, x + 28, 82), fill=color, width=6)
        draw.text((x + 38, 70), label_en, font=legend_font, fill="#344054")

    step = (chart_right - chart_left) / max(len(rows) - 1, 1)
    for key, _label_cn, _label_en, color in available:
        values = [
            float(row[key]) if isinstance(row.get(key), (int, float)) else None
            for row in rows
        ]
        points: list[tuple[float, float] | None] = []
        for index, value in enumerate(values):
            if value is None:
                points.append(None)
                continue
            x = chart_left + index * step
            y = plot_bottom - (value / scale_max) * (plot_bottom - plot_top)
            points.append((x, y))
        previous = None
        for point in points:
            if point is not None and previous is not None:
                draw.line((*previous, *point), fill=color, width=5)
            previous = point
        for index, point in enumerate(points):
            if point is None:
                continue
            x, y = point
            draw.ellipse((x - 6, y - 6, x + 6, y + 6), fill=color, outline="white", width=2)
    label_every = max(1, (len(months) + 11) // 12)
    for index, month in enumerate(months):
        if index % label_every and index != len(months) - 1:
            continue
        x = chart_left + index * step
        short = month[2:]
        box = draw.textbbox((0, 0), short, font=axis_font)
        draw.text(
            (x - (box[2] - box[0]) / 2, plot_bottom + 14),
            short,
            font=axis_font,
            fill="#555F6D",
        )
    path.parent.mkdir(parents=True, exist_ok=True)
    image.save(path, format="PNG", optimize=True)


def _add_monthly_data_table(
    doc: Document,
    rows: list[dict[str, Any]],
    specs: list[tuple[str, str, str, str]],
) -> None:
    available = [
        spec
        for spec in specs
        if any(isinstance(row.get(spec[0]), (int, float)) for row in rows)
    ]
    table = doc.add_table(rows=1, cols=1 + len(available))
    headers = ["月份"] + [spec[1] for spec in available]
    for index, label in enumerate(headers):
        cell = table.rows[0].cells[index]
        _set_cell_shading(cell, PALE_BLUE)
        cell.text = ""
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_run_font(paragraph.add_run(label), 9.0, True, INK)
    for row in rows:
        cells = table.add_row().cells
        values = [row["month"]] + [row.get(spec[0]) for spec in available]
        for index, value in enumerate(values):
            cells[index].text = ""
            paragraph = cells[index].paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            text = str(value) if index == 0 else _format_number(value)
            _set_run_font(paragraph.add_run(text), 8.8, index == 0, INK)
    metric_width = (6.7 - 1.35) / max(len(available), 1)
    _set_fixed_table(table, [1.35] + [metric_width] * len(available))
    _set_table_borders(table, LINE, 4)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def _monthly_trend_summary(
    rows: list[dict[str, Any]],
    specs: list[tuple[str, str, str, str]],
) -> str:
    for key, label_cn, _label_en, _color in specs:
        points = [
            (row["month"], float(row[key]))
            for row in rows
            if isinstance(row.get(key), (int, float))
        ]
        if not points:
            continue
        peak_month, peak_value = max(points, key=lambda item: item[1])
        first_month, first_value = points[0]
        last_month, last_value = points[-1]
        if last_value > first_value:
            change = "高于"
        elif last_value < first_value:
            change = "低于"
        else:
            change = "持平于"
        return (
            f"可核对月度序列覆盖 {rows[0]['month']}—{rows[-1]['month']}。"
            f"{label_cn}峰值出现在 {peak_month}（{_format_number(peak_value)}）；"
            f"期末 {last_month} 为 {_format_number(last_value)}，{change}期初 "
            f"{first_month} 的 {_format_number(first_value)}。"
        )
    return ""


def _add_monthly_trend(
    doc: Document,
    job_dir: Path,
    platform: str,
    monthly: Any,
    specs: list[tuple[str, str, str, str]],
    caption_number: int,
) -> bool:
    rows = _normalize_monthly_rows(monthly)
    available = [
        spec
        for spec in specs
        if any(isinstance(row.get(spec[0]), (int, float)) for row in rows)
    ]
    if not rows or not available:
        _add_body(
            doc,
            "本次任务没有保存可核对的月度序列，因此不绘制估算趋势图。"
            "请重新同步平台数据后再生成；系统不会把汇总总数平均拆分到各月。",
            color=MUTED,
        )
        return False
    if platform == "LinkedIn":
        note = (
            "下图展示 LinkedIn 月度自然展示量及日度触达累计。日度触达累计为各日"
            "自然独立展示数之和，用于观察各月日度触达活跃程度，不等同于月度去重人数。"
        )
    else:
        note = (
            f"下图展示 {platform} 月度浏览量及日度触达累计。日度触达累计为各日"
            "去重触达值之和，用于观察各月日度触达活跃程度，不等同于月度去重人数。"
        )
    _add_body(doc, note)
    chart_path = job_dir / "charts" / f"{platform.lower()}_monthly_trend.png"
    _render_monthly_chart(chart_path, platform, rows, available)
    _add_image(doc, chart_path, f"{platform} 月度自然内容趋势图")
    _add_caption(doc, f"图{caption_number}　{platform} 月度自然内容趋势")
    _add_body(doc, _monthly_trend_summary(rows, available), color=MUTED)
    _add_monthly_data_table(doc, rows, available)
    _add_caption(doc, f"表　{platform} 月度趋势明细（POPO 中可编辑）")
    return True


def _linkedin_monthly_rows(linkedin: dict[str, Any]) -> list[dict[str, Any]]:
    combined: dict[str, dict[str, Any]] = {}
    for group_name, key_map in (
        (
            "content",
            {
                "impressions_organic": "impressions_organic",
                "unique_impressions_organic": "daily_reach_sum",
            },
        ),
    ):
        for source in linkedin.get(group_name, {}).get("monthly", []) or []:
            month = str(source.get("month") or "")
            if not re.fullmatch(r"\d{4}-\d{2}", month):
                continue
            target = combined.setdefault(month, {"month": month})
            for source_key, target_key in key_map.items():
                target[target_key] = source.get(source_key)
    return [combined[month] for month in sorted(combined)]


def _theme_label(title: str) -> str:
    text = title.lower()
    rules = [
        (("water", "solub", "dissolv"), "原料水溶性与应用演示"),
        (("pms", "menstrual", "cycle"), "PMS 女性健康应用"),
        (("menopause", "midlife"), "更年期情绪与睡眠应用"),
        (("powder", "extract"), "原粉与标准化提取物差异"),
        (("crocin", "picrocrocin", "safranal"), "标准化活性指标说明"),
        (("brain", "cognitive", "focus"), "脑健康与认知健康趋势"),
        (("beauty", "skincare"), "美容与护肤配方应用"),
        (("eye health", "vision"), "眼健康配方应用"),
    ]
    for keywords, label in rules:
        if all(keyword in text for keyword in keywords[:2]) or any(
            keyword in text for keyword in keywords
        ):
            return label
    clean = re.sub(r"[\r\n#]+", " ", title)
    clean = re.sub(r"\s+", " ", clean).strip(" .:-")
    return clean[:42] + ("…" if len(clean) > 42 else "") if clean else "LinkedIn 原生动态"


def _report_title(request: dict[str, Any]) -> tuple[str, str]:
    report_type = str(request.get("report_type") or "monthly").strip().lower()
    if report_type not in REPORT_TYPE_TITLES:
        report_type = "custom"
    return report_type, REPORT_TYPE_TITLES[report_type]


def _post_interactions(post: dict[str, Any]) -> int | None:
    values = [
        post.get(key)
        for key in ("likes", "comments", "reposts")
        if post.get(key) is not None
    ]
    return sum(int(value) for value in values) if values else None


def _add_linkedin_posts_table(
    doc: Document,
    posts: list[dict[str, Any]],
    *,
    limit: int = 6,
    impressions_label: str = "展示",
):
    headers = ["日期", "内容主题", impressions_label, "点击", "CTR", "互动"]
    widths = [0.75, 3.05, 0.7, 0.55, 0.65, 1.0]
    table = doc.add_table(rows=1, cols=len(headers))
    _set_fixed_table(table, widths)
    _set_table_borders(table)
    table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
    for idx, header in enumerate(headers):
        cell = table.rows[0].cells[idx]
        _set_cell_shading(cell, NAVY)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(header)
        _set_run_font(r, 8.6, True, WHITE)
    for post in posts[:limit]:
        cells = table.add_row().cells
        row = cells[0]._tc.getparent()
        row.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        values = [
            datetime.strptime(post["created_date"], "%Y-%m-%d").strftime("%m/%d"),
            _theme_label(post["title"]),
            _format_number(post["impressions"]),
            _format_number(post.get("clicks")),
            _format_percent(post.get("ctr")),
            _format_number(_post_interactions(post)),
        ]
        for idx, value in enumerate(values):
            cell = cells[idx]
            _set_cell_shading(cell, WHITE if len(table.rows) % 2 else PALE_BLUE)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if idx == 1 else WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(value)
            _set_run_font(r, 8.3, idx == 1, NAVY if idx == 1 else INK)
    _set_fixed_table(table, widths)
    return table


def _platform_analysis(
    platform: str,
    metrics: dict[str, Any],
    highlight: str,
    report_type: str = "monthly",
) -> list[tuple[str, str]]:
    def optional_int(key: str) -> int | None:
        value = metrics.get(key)
        return int(value) if value is not None and value != "" else None

    views = optional_int("views")
    reach = optional_int("reach")
    viewers = optional_int("viewers")
    interactions = optional_int("interactions")
    likes = optional_int("likes")
    posts = optional_int("posts")
    per_post = views / posts if views is not None and posts else None
    if views is None:
        distribution = (
            "Meta API 未返回本期自然浏览量，报告未将缺失数据推测为 0；"
            "建议在下个周期补齐授权或后台汇总。"
        )
    elif views == 0:
        distribution = "本期尚未形成有效自然浏览，当前重点应放在稳定发布和账号基础信息完善。"
    elif per_post is not None and per_post < 5:
        distribution = (
            f"本期共记录 {_format_number(views)} 次浏览，单帖平均约 {per_post:.1f} 次，"
            "自然分发仍处于冷启动阶段。"
        )
    elif per_post is not None:
        distribution = (
            f"本期共记录 {_format_number(views)} 次浏览，单帖平均约 {per_post:.1f} 次，"
            "已经出现可继续验证的自然分发信号。"
        )
    else:
        distribution = f"本期共记录 {_format_number(views)} 次浏览。"
    if interactions is None:
        engagement = (
            "Meta API 未返回完整互动指标，报告保留为空，不据此判断互动强弱。"
        )
    else:
        engagement = (
            f"本期记录 {_format_number(interactions)} 次互动、{_format_number(likes)} 次点赞。"
            + (
                "互动尚未形成，下一阶段应优先测试更直接的问题式标题、应用场景和行动引导。"
                if interactions == 0
                else "已有少量互动反馈，建议围绕产生互动的内容主题连续做系列化测试。"
            )
        )
    theme = (
        f"内容主题：{highlight.strip()}。建议延续该方向，并用案例、对比和应用演示提高可读性。"
        if highlight.strip()
        else "内容主题以后台截图为准；建议后续持续标记高浏览帖的主题和形式，形成可复用的选题清单。"
    )
    audience = viewers if viewers is not None else reach
    audience_label = "浏览人数" if viewers is not None else "覆盖人数"
    distribution_with_audience = distribution + (
        f" 本期{audience_label} {_format_number(audience)}。"
        if audience is not None
        else ""
    )
    if views is None or posts in (None, 0):
        account_stage = "数据基础仍不完整，暂不对账号阶段作强判断。"
    elif views == 0 or (per_post is not None and per_post < 5):
        account_stage = "账号仍处于冷启动或基础积累阶段，首先要稳定发布并补齐主页承接。"
    elif interactions in (None, 0):
        account_stage = "账号已出现自然分发，但互动与转化承接尚未验证。"
    else:
        account_stage = "账号已进入内容方向验证阶段，可围绕高表现主题做连续测试。"
    platform_character = (
        "Facebook 更适合承载案例长图、清单和行业信息，并通过主页与官网完成进一步承接。"
        if platform == "Facebook"
        else "Instagram 更适合用 Reels、步骤卡片、现场特写和短字幕强化视觉化表达。"
    )
    profile = report_mode_profile(report_type)
    if profile["framework"] == "annual":
        return [
            (
                "长期趋势：",
                "当前数据仅包含报告期汇总，未提供可核对的月度序列，因此不推测全年涨跌。"
                + distribution_with_audience,
            ),
            ("年度内容方向：", theme),
            (
                "年度内容模型：",
                f"应把已验证主题沉淀为“问题/场景—方法/流程—证据/结果—单一 CTA”的稳定结构。{engagement}",
            ),
            ("平台价值：", platform_character),
            ("账号阶段：", account_stage),
            (
                "年度战略：",
                f"下一年度以可复用栏目、可核对案例和统一主页承接为主线，持续验证{platform}从浏览到询盘的链路。",
            ),
        ]
    if str(report_type).strip().lower() == "quarterly":
        return [
            ("持续有效的内容方向：", theme),
            (
                "内容表现变化：",
                "当前数据未包含上阶段同口径对照，不推测环比升降；本阶段可确认的表现为："
                + distribution_with_audience,
            ),
            (
                "用户关注点变化：",
                "需结合连续三个周期的同口径排行验证变化；本期截图中的高浏览主题作为下一季度复测基线。",
            ),
            ("平台表现趋势：", f"{distribution_with_audience}{engagement}"),
            ("当前平台定位：", f"{platform_character}{account_stage}"),
        ]
    return [
        ("本周期表现较好的内容类型：", theme),
        ("用户关注方向：", "以内容排行截图中的高浏览主题为本周期用户关注信号，不把单篇表现外推为长期规律。"),
        ("内容结构表现：", f"{distribution_with_audience}{engagement}"),
        ("平台表现特点：", platform_character),
        ("当前账号阶段：", account_stage),
    ]


def _linkedin_analysis(
    data: dict[str, Any], report_type: str = "monthly"
) -> list[tuple[str, str]]:
    c = data["content"]
    v = data["visitors"]
    f = data["followers"]
    top = data["posts"]["top_impressions"][0] if data["posts"]["top_impressions"] else None
    base_result = [
        (
            "自然分发：",
            f"本期原生内容获得 {_format_number(c['impressions_organic'])} 次展示，"
            f"共发布 {_format_number(data['posts']['count'])} 条动态，说明账号已经形成稳定的内容供给。",
        ),
        (
            "互动效率：",
            f"原生点击 {_format_number(c['clicks_organic'])} 次，回应 {_format_number(c['reactions_organic'])} 次，"
            f"评论 {_format_number(c['comments_organic'])} 次，转发 {_format_number(c['reposts_organic'])} 次；"
            f"按上述动作合计计算的互动率为 {_format_percent(c['engagement_rate'])}。",
        ),
        (
            "主页承接：",
            f"主页访问 {_format_number(v['page_views_total'])} 次、独立访客 {_format_number(v['unique_total'])} 人，"
            f"新增关注者 {_format_number(f['organic'])} 人。建议保持主页产品定位、检测能力和联系入口清晰可见。",
        ),
    ]
    if top:
        base_result.append(
            (
                "高表现内容：",
                f"本期展示最高的是“{_theme_label(top['title'])}”，获得 "
                f"{_format_number(top['impressions'])} 次展示和 {_format_number(top['clicks'])} 次点击。"
                "这类带有具体应用、规格或演示信息的内容值得继续系列化。",
            )
        )
    profile = report_mode_profile(report_type)
    if profile["framework"] == "annual":
        top_text = (
            f"年度样本中展示最高的是“{_theme_label(top['title'])}”，获得 "
            f"{_format_number(top['impressions'])} 次展示和 {_format_number(top['clicks'])} 次点击。"
            if top
            else "本期没有可用于年度内容模型判断的帖子排行。"
        )
        return [
            ("长期趋势：", "当前导出仅提供报告期汇总，未提供月度序列，因此不推测全年涨跌。"),
            ("年度内容方向：", top_text),
            ("年度内容模型：", "优先沉淀具体应用、规格、演示、案例与结果证明类内容，并保持清晰 CTA。"),
            (
                "平台价值：",
                f"LinkedIn 本期获得 {_format_number(c['impressions_organic'])} 次原生展示、"
                f"{_format_number(v['page_views_total'])} 次主页访问和 {_format_number(f['organic'])} 名新增关注者，"
                "适合承担专业能力证明与商机前置筛选。",
            ),
            ("账号阶段：", base_result[1][1] + base_result[2][1]),
        ]
    if str(report_type).strip().lower() == "quarterly":
        return [
            ("持续有效的内容方向：", base_result[-1][1] if top else "本阶段暂无可核对的高表现帖子。"),
            ("内容表现变化：", "当前导出未包含上一阶段同口径对照，不推测环比升降。"),
            ("用户关注点变化：", "以本阶段 Top 内容作为下一季度复测基线，连续观察主题与点击变化。"),
            ("平台表现趋势：", base_result[0][1] + base_result[1][1]),
            ("当前平台定位：", base_result[2][1]),
        ]
    return [
        ("本周期表现较好的内容类型：", base_result[-1][1] if top else "本周期暂无可核对的高表现帖子。"),
        ("用户关注方向：", "以 Top 内容的展示与点击作为本周期关注信号，不外推为长期规律。"),
        ("内容结构表现：", base_result[0][1] + base_result[1][1]),
        ("平台表现特点：", "LinkedIn 更适合承载专业能力、应用证据与行业判断。"),
        ("当前账号阶段：", base_result[2][1]),
    ]


def _linkedin_personal_analysis(
    data: dict[str, Any], report_type: str = "monthly"
) -> list[tuple[str, str]]:
    personal = data["personal"]
    posts = data["posts"]["top_impressions"]
    top = posts[0] if posts else None
    result = [
        (
            "数据范围：",
            (
                f"本章节基于人工筛选的 {_format_number(len(posts))} 条高浏览帖子进行样本分析，"
                "不代表账号全部帖子或 LinkedIn 公共主页后台的全量数据。"
                if posts
                else "本章节基于人工提供的 Top 10 截图进行定性分析，系统未对截图数字做 OCR 或推测。"
            ),
        )
    ]
    if posts:
        result.append(
            (
                "样本表现：",
                f"已整理帖子的浏览量/展示量合计为 "
                f"{_format_number(personal.get('top10_impressions'))}，"
                f"可识别互动合计为 {_format_number(personal.get('top10_interactions'))}。"
                "以上仅为入选样本累计，不作为报告期账号全量。"
            )
        )
    if top:
        result.append(
            (
                "高表现内容：",
                f"样本中浏览量/展示量最高的是“{_theme_label(top['title'])}”，"
                f"记录 {_format_number(top['impressions'])}。"
                "建议围绕同类主题继续做案例、对比、工艺或应用场景的系列化内容。",
            )
        )
    highlight = str(personal.get("highlight") or "").strip()
    if highlight:
        result.append(
            (
                "人工补充：",
                f"{highlight}。后续周期建议沿用相同筛选标准，保证 Top 10 样本可横向追踪。",
            )
        )
    result.append(
        (
            "口径限制：",
            "个人账号不提供公共主页同等的内容、访客和关注者完整导出能力；"
            "缺失指标在报告中保留为“—”，不会按 0 处理。",
        )
    )
    profile = report_mode_profile(report_type)
    if profile["framework"] == "annual":
        return [
            ("长期趋势：", "个人账号仅提供人工筛选样本，没有月度全量序列，因此不推测全年涨跌。"),
            ("年度内容方向：", result[2][1] if top and len(result) > 2 else result[0][1]),
            ("年度内容模型：", "将样本中高浏览主题沉淀为案例、对比、工艺或应用场景系列，并沿用同一筛选口径。"),
            ("平台价值：", "个人账号适合作为专家观点与信任建立渠道，但样本数据不能替代公共主页全量分析。"),
            ("账号阶段：", result[-1][1]),
        ]
    if str(report_type).strip().lower() == "quarterly":
        return [
            ("持续有效的内容方向：", result[2][1] if top and len(result) > 2 else result[0][1]),
            ("内容表现变化：", "缺少上一阶段同口径 Top 10 样本，不推测环比升降。"),
            ("用户关注点变化：", "沿用相同筛选标准后，下一季度再判断关注主题是否迁移。"),
            ("平台表现趋势：", result[1][1] if posts and len(result) > 1 else result[0][1]),
            ("当前平台定位：", "作为个人专业影响力样本渠道使用，不替代公共主页全量数据。"),
        ]
    return [
        ("本周期表现较好的内容类型：", result[2][1] if top and len(result) > 2 else result[0][1]),
        ("用户关注方向：", "仅以人工入选 Top 10 样本识别本周期关注信号。"),
        ("内容结构表现：", result[1][1] if posts and len(result) > 1 else result[0][1]),
        ("平台表现特点：", "个人账号适合承载专家观点与信任建立，不能替代公共主页全量分析。"),
        ("当前账号阶段：", result[-1][1]),
    ]


def _personal_linkedin_data(
    inputs: Path, request: dict[str, Any]
) -> dict[str, Any]:
    files = request.get("files") or {}
    manual = request.get("linkedin_manual") or {}
    top10_name = files.get("linkedin_personal_top10")
    parsed = (
        parse_linkedin_personal_top10(inputs / top10_name)
        if top10_name
        else {
            "source_file": None,
            "posts": [],
            "top10_impressions": None,
            "top10_interactions": None,
        }
    )
    screenshots = list(files.get("linkedin_personal_screenshots") or [])
    warnings = [
        "本章节基于 LinkedIn 个人账号人工筛选样本，不包含公共主页的完整内容、访客和关注者导出。"
    ]
    if not top10_name:
        warnings.append("本次仅提供 Top 10 截图，系统未对截图数字做 OCR 或推测。")
    return {
        "mode": "personal",
        "source_files": {
            "top10": top10_name,
            "screenshots": screenshots,
        },
        "content": {
            "start": request["report_start"],
            "end": request["report_end"],
            "impressions_organic": manual.get("period_impressions"),
            "impressions_sponsored": None,
            "clicks_organic": None,
            "reactions_organic": None,
            "comments_organic": None,
            "reposts_organic": None,
            "engagement_rate": None,
        },
        "posts": {
            "count": len(parsed["posts"]),
            "top_impressions": parsed["posts"],
            "all": parsed["posts"],
        },
        "visitors": {
            "page_views_total": manual.get("profile_views"),
            "unique_total": None,
        },
        "followers": {
            "organic": None,
            "total": manual.get("current_followers"),
        },
        "personal": {
            "top10_impressions": parsed["top10_impressions"],
            "top10_interactions": parsed["top10_interactions"],
            "period_impressions": manual.get("period_impressions"),
            "profile_views": manual.get("profile_views"),
            "current_followers": manual.get("current_followers"),
            "highlight": str(manual.get("highlight") or "").strip(),
        },
        "warnings": warnings,
    }


def _empty_linkedin_data(request: dict[str, Any]) -> dict[str, Any]:
    return {
        "mode": "none",
        "source_files": {},
        "content": {
            "start": request["report_start"],
            "end": request["report_end"],
            "impressions_organic": None,
            "impressions_sponsored": None,
            "clicks_organic": None,
            "reactions_organic": None,
            "comments_organic": None,
            "reposts_organic": None,
            "engagement_rate": None,
        },
        "posts": {"count": 0, "top_impressions": [], "all": []},
        "visitors": {"page_views_total": None, "unique_total": None},
        "followers": {"organic": None, "total": None},
        "warnings": ["本期未提供 LinkedIn 数据，报告不对缺失指标作 0 值推测。"],
    }


def _setup_document(brand: str, report_title: str) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.68)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    section.header_distance = Inches(0.3)
    section.footer_distance = Inches(0.3)

    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string(INK)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    for style_name in ("List Bullet", "List Number"):
        style = doc.styles[style_name]
        style.font.name = FONT
        style.font.size = Pt(10.3)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)

    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run(f"{brand} · {report_title}  |  ")
    _set_run_font(r, 8.5, False, MUTED)
    _add_field(p, "PAGE")
    return doc


def generate_report(job_dir: Path, request: dict[str, Any]) -> tuple[Path, dict[str, Any]]:
    inputs = job_dir / "inputs"
    linkedin_mode = str(request.get("linkedin_mode") or "company").strip().lower()
    if linkedin_mode == "company":
        linkedin = parse_linkedin_exports(
            inputs / request["files"]["linkedin_content"],
            inputs / request["files"]["linkedin_followers"],
            inputs / request["files"]["linkedin_visitors"],
        )
    elif linkedin_mode == "personal":
        linkedin = _personal_linkedin_data(inputs, request)
    else:
        linkedin = _empty_linkedin_data(request)
    brand = request["brand"].strip()
    report_start = request["report_start"]
    report_end = request["report_end"]
    account_label = request.get("account_label", "").strip()
    meta_info = request.get("meta") or {}
    meta_source = meta_info.get("source") or "manual"
    report_type, report_title = _report_title(request)
    profile = report_mode_profile(report_type)
    annual_framework = profile["framework"] == "annual"
    period_label = f"{_format_date(report_start)}—{_format_date(report_end)}"
    fb = request["facebook"]
    ig = request["instagram"]
    fb_image = inputs / request["files"]["facebook_screenshot"]
    ig_image = inputs / request["files"]["instagram_screenshot"]

    doc = _setup_document(brand, report_title)
    figure_number = 1

    # Cover and report framing
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(brand)
    _set_run_font(r, 26, True, NAVY)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(report_title)
    _set_run_font(r, 20, True, BLUE)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run(f"Facebook / Instagram / LinkedIn\n报告区间：{period_label}")
    _set_run_font(r, 11.5, False, MUTED)
    _paragraph_border(p, LIGHT_BLUE, 12)
    _add_lead(doc, analysis_principle(report_type))

    _add_heading(doc, "第一部分　Facebook", 1)
    _add_heading(doc, "1.1　数据概览", 2)
    fb_audience_label = "浏览人数" if fb.get("viewers") is not None else "覆盖人数"
    fb_audience_value = (
        fb.get("viewers") if fb.get("viewers") is not None else fb.get("reach")
    )
    _add_detail_table(
        doc,
        [
            ("内容浏览量", _format_number(fb["views"]), "报告期内可核对的内容浏览量"),
            (
                fb_audience_label,
                _format_number(fb_audience_value),
                "看到内容的去重人数；缺失时保留为“—”",
            ),
            ("发布内容", _format_number(fb["posts"]), "报告期内发布的内容数量"),
            ("内容互动次数", _format_number(fb["interactions"]), "点赞、评论、分享等可计互动"),
        ],
    )
    source_note = {
        "api": "以上概览来自 Meta Graph API 自动同步",
        "mixed": "以上概览以 Meta Graph API 为主，缺失项由页面手动补充",
        "cached_api": "以上概览来自页面保留的最近一次 Meta Graph API 同步快照",
        "manual": "以上概览来自页面手动补充",
    }.get(meta_source, "以上概览来自本次任务数据")
    _add_body(
        doc,
        f"{source_note}。指标用于描述本报告期的自然运营基础；不同平台不直接比较绝对量级。",
        color=MUTED,
    )
    if meta_info.get("missing_metrics"):
        _add_lead(
            doc,
            "Meta 未取得的指标已保留为“—”，没有按 0 推测："
            + "、".join(meta_info["missing_metrics"])
            + "。",
        )

    if annual_framework:
        _add_heading(doc, "1.2　月度自然内容趋势", 2)
        if _add_monthly_trend(
            doc,
            job_dir,
            "Facebook",
            fb.get("monthly"),
            [
                ("views", "浏览量", "Views", "#2F5A9E"),
                ("daily_reach_sum", "日度触达累计", "Daily reach sum", "#5AA1D8"),
            ],
            figure_number,
        ):
            figure_number += 1
    fb_ranking_number = "1.3" if annual_framework else "1.2"
    fb_analysis_number = "1.4" if annual_framework else "1.3"
    _add_heading(doc, f"{fb_ranking_number}　自然内容排行", 2)
    _add_image(doc, fb_image, f"{brand} Facebook 内容排行后台截图")
    _add_caption(doc, f"图{figure_number}　Facebook 自然内容排行（报告区间：{period_label}）")
    figure_number += 1
    _add_heading(doc, f"{fb_analysis_number}　{profile['analysis_heading']}", 2)
    for label, text in _platform_analysis(
        "Facebook", fb, request.get("facebook_highlight", ""), report_type
    ):
        _add_bullet(doc, label, text)

    _add_heading(doc, "第二部分　Instagram", 1)
    _add_heading(doc, "2.1　数据概览", 2)
    _add_detail_table(
        doc,
        [
            ("内容浏览量", _format_number(ig["views"]), "报告期内可核对的内容浏览量"),
            ("覆盖人数", _format_number(ig["reach"]), "看到内容的去重账号数"),
            ("发布内容", _format_number(ig["posts"]), "报告期内发布的内容数量"),
            ("内容互动次数", _format_number(ig["interactions"]), "点赞、评论、分享等可计互动"),
        ],
    )
    _add_body(doc, f"{source_note}。缺失指标保留为“—”，不按 0 推测。", color=MUTED)
    if annual_framework:
        _add_heading(doc, "2.2　月度自然内容趋势", 2)
        if _add_monthly_trend(
            doc,
            job_dir,
            "Instagram",
            ig.get("monthly"),
            [
                ("views", "浏览量", "Views", "#2F5A9E"),
                ("daily_reach_sum", "日度触达累计", "Daily reach sum", "#5AA1D8"),
            ],
            figure_number,
        ):
            figure_number += 1
    ig_ranking_number = "2.3" if annual_framework else "2.2"
    ig_analysis_number = "2.4" if annual_framework else "2.3"
    _add_heading(doc, f"{ig_ranking_number}　自然内容排行", 2)
    _add_image(doc, ig_image, f"{brand} Instagram 内容排行后台截图")
    _add_caption(doc, f"图{figure_number}　Instagram 自然内容排行（报告区间：{period_label}）")
    figure_number += 1

    _add_heading(doc, f"{ig_analysis_number}　{profile['analysis_heading']}", 2)
    for label, text in _platform_analysis(
        "Instagram", ig, request.get("instagram_highlight", ""), report_type
    ):
        _add_bullet(doc, label, text)

    c, v, f = linkedin["content"], linkedin["visitors"], linkedin["followers"]
    doc.add_page_break()
    if linkedin_mode == "company":
        _add_heading(doc, "第三部分　LinkedIn", 1)
        _add_heading(doc, "3.1 数据概览", 2)
        _add_kpi_table(
            doc,
            [
                ("原生展示", _format_number(c["impressions_organic"])),
                ("原生点击", _format_number(c["clicks_organic"])),
                ("互动率", _format_percent(c["engagement_rate"])),
                ("发布动态", _format_number(linkedin["posts"]["count"])),
                ("主页访问", _format_number(v["page_views_total"])),
                ("独立访客", _format_number(v["unique_total"])),
                ("新增关注者", _format_number(f["organic"])),
                ("付费展示", _format_number(c["impressions_sponsored"])),
            ],
        )
        _add_body(
            doc,
            f"LinkedIn 导出实际数据区间：{_format_date(c['start'])}—{_format_date(c['end'])}，"
            "平台注明数据最多延迟两天；互动率按（点击＋回应＋评论＋转发）÷原生展示计算。",
            color=MUTED,
        )
        if linkedin["warnings"]:
            _add_lead(doc, "；".join(linkedin["warnings"]))
        doc.add_page_break()
        if annual_framework:
            _add_heading(doc, "3.2 月度自然内容趋势", 2)
            if _add_monthly_trend(
                doc,
                job_dir,
                "LinkedIn",
                _linkedin_monthly_rows(linkedin),
                [
                    ("impressions_organic", "自然展示量", "Organic impressions", "#2F5A9E"),
                    ("daily_reach_sum", "日度触达累计", "Daily reach sum", "#5AA1D8"),
                ],
                figure_number,
            ):
                figure_number += 1
        linkedin_ranking_number = "3.3" if annual_framework else "3.2"
        linkedin_analysis_number = "3.4" if annual_framework else "3.3"
        _add_heading(doc, f"{linkedin_ranking_number} 自然内容排行", 2)
        _add_linkedin_posts_table(doc, linkedin["posts"]["top_impressions"])
        _add_caption(doc, "表1　LinkedIn 原生内容按展示量排序（Top 6）")
        _add_heading(
            doc,
            f"{linkedin_analysis_number} {profile['analysis_heading']}",
            2,
        )
        for label, text in _linkedin_analysis(linkedin, report_type):
            _add_bullet(doc, label, text)
    elif linkedin_mode == "personal":
        personal = linkedin["personal"]
        _add_heading(doc, "第三部分　LinkedIn 个人账号内容表现", 1)
        _add_heading(doc, "3.1 数据口径与概览", 2)
        _add_lead(doc, "；".join(linkedin["warnings"]))
        _add_kpi_table(
            doc,
            [
                ("Top 帖子样本", _format_number(linkedin["posts"]["count"])),
                ("Top 10 浏览/展示", _format_number(personal["top10_impressions"])),
                ("Top 10 互动", _format_number(personal["top10_interactions"])),
                ("报告期总浏览", _format_number(personal["period_impressions"])),
                ("主页访问", _format_number(personal["profile_views"])),
                ("当前关注者", _format_number(personal["current_followers"])),
            ],
            columns=3,
        )
        _add_body(
            doc,
            "“Top 10 浏览/展示”仅为人工入选样本累计；“报告期总浏览、主页访问、当前关注者”"
            "仅在人工提供时展示，空值保留为“—”。",
            color=MUTED,
        )
        doc.add_page_break()
        if annual_framework:
            _add_heading(doc, "3.2 月度自然内容趋势", 2)
            _add_body(
                doc,
                "个人账号仅提供人工筛选样本，没有可核对的月度全量序列，因此本节不绘制估算趋势图。",
                color=MUTED,
            )
        linkedin_ranking_number = "3.3" if annual_framework else "3.2"
        linkedin_analysis_number = "3.4" if annual_framework else "3.3"
        _add_heading(doc, f"{linkedin_ranking_number} 自然内容排行", 2)
        if linkedin["posts"]["top_impressions"]:
            _add_linkedin_posts_table(
                doc,
                linkedin["posts"]["top_impressions"],
                limit=10,
                impressions_label="浏览/展示",
            )
            _add_caption(doc, "表1　LinkedIn 个人账号人工筛选帖子（最多 Top 10）")
        for index, filename in enumerate(
            request["files"].get("linkedin_personal_screenshots") or [], start=1
        ):
            _add_image(
                doc,
                inputs / filename,
                f"{brand} LinkedIn 个人账号 Top 10 人工截图 {index}",
            )
            _add_caption(doc, f"图{figure_number}　LinkedIn 个人账号 Top 10 人工截图")
            figure_number += 1
        _add_heading(
            doc,
            f"{linkedin_analysis_number} {profile['analysis_heading']}",
            2,
        )
        for label, text in _linkedin_personal_analysis(linkedin, report_type):
            _add_bullet(doc, label, text)
    else:
        _add_heading(doc, "第三部分　LinkedIn", 1)
        _add_heading(doc, "3.1 本期数据说明", 2)
        _add_lead(
            doc,
            "本期未提供 LinkedIn 公共主页导出、个人账号 Top 10 汇总表或截图。"
            "因此本章节不展示指标，也不将缺失数据记为 0。",
        )

    if linkedin_mode != "none":
        doc.add_page_break()
    else:
        doc.add_paragraph().paragraph_format.space_after = Pt(3)
    _add_heading(doc, "第四部分　综合结论与下阶段优化建议", 1)
    _add_heading(doc, "4.1　综合结论", 2)
    _add_lead(
        doc,
        f"{profile['conclusion_scope']}：本报告重点回答“{profile['question']}”，"
        f"结论仅基于本次可核对的数据与内容排行形成。",
    )
    if linkedin_mode == "company":
        _add_body(
            doc,
            f"本期三端均保持自然内容发布。Facebook 与 Instagram 仍以内容同步、品牌信息沉淀和客户信任建立为主要作用；"
            f"LinkedIn 取得 {_format_number(c['impressions_organic'])} 次原生展示、"
            f"{_format_number(c['clicks_organic'])} 次点击和 {_format_number(f['organic'])} 名新增关注者，"
            "已形成更清晰的内容反馈与主页承接链路。",
        )
    elif linkedin_mode == "personal":
        personal = linkedin["personal"]
        _add_body(
            doc,
            "Facebook 与 Instagram 继续承担品牌信息沉淀和客户信任建立；"
            f"LinkedIn 个人账号本期基于 {_format_number(linkedin['posts']['count'])} 条人工筛选帖子进行样本分析，"
            f"样本浏览量/展示量合计 {_format_number(personal['top10_impressions'])}。"
            "该数值不等同于账号报告期全量，访客和关注者缺失项未作推测。",
        )
    else:
        _add_body(
            doc,
            "本期结论基于 Facebook 与 Instagram 数据形成；LinkedIn 未提供数据，"
            "因此未纳入本期效果判断，也没有将缺失指标记为 0。",
        )
    _add_body(
        doc,
        "从 Facebook 与 Instagram 的表现判断，后续应优先复用已经获得相对高浏览的主题，"
        "并把企业介绍转化为具体场景、流程、案例和结果证明。两个平台可以共用选题，"
        "但应分别使用适合平台的表达形式，并在内容结尾设置单一、明确的下一步动作。",
    )
    _add_heading(doc, "4.2　下阶段优化建议", 2)
    for label, text in report_recommendations(report_type):
        _add_bullet(doc, label, text)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.35
    r = p.add_run(FIXED_ENDING)
    _set_run_font(r, 10.8, True, INK)

    core = doc.core_properties
    core.title = f"{brand} {report_title} {period_label}"
    core.subject = f"Facebook、Instagram、LinkedIn 自然流量{report_title}"
    core.author = request.get("created_by", "") or "Social Report Center"
    core.comments = "由 Social Report Center 根据用户上传数据自动生成"

    output_dir = job_dir / "outputs"
    output_dir.mkdir(parents=True, exist_ok=True)
    filename = (
        f"{_safe_name(brand)}_{report_title}_"
        f"{report_start.replace('-', '.')}-{report_end.replace('-', '.')}.docx"
    )
    output_path = output_dir / filename
    doc.save(output_path)

    analysis = {
        "request": request,
        "report_type": report_type,
        "report_title": report_title,
        "linkedin": linkedin,
        "generated_file": filename,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
    }
    (output_dir / "analysis.json").write_text(
        json.dumps(analysis, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return output_path, analysis
