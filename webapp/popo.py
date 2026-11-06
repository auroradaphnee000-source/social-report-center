"""Deliver a completed social report as a native POPO document."""

from __future__ import annotations

import html
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
import uuid
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import quote

from webapp.popo_auth import fabric_cli, fabric_env, fabric_lock, popo_status


class PopoDeliveryError(RuntimeError):
    pass


def _resolve_popo_js() -> str:
    """Locate the legacy POPO CLI used by local/admin delivery runs.

    Production keeps using Fabric and its server-owned OpenID credential.  The
    fallback mirrors gsc-analyzer so a logged-in administrator workstation can
    run one-off native document repairs without moving credentials to disk in
    this project.
    """
    explicit = os.getenv("POPO_CLI_JS", "").strip()
    if explicit and os.path.isfile(explicit):
        return explicit
    shim = shutil.which("popo-cli")
    if shim:
        candidate = Path(shim).parent / "node_modules" / "@popo" / "cli" / "bin" / "popo.js"
        if candidate.is_file():
            return str(candidate)
    appdata = os.getenv("APPDATA", "").strip()
    if appdata:
        candidate = (
            Path(appdata)
            / "npm"
            / "node_modules"
            / "@popo"
            / "cli"
            / "bin"
            / "popo.js"
        )
        if candidate.is_file():
            return str(candidate)
    return ""


@dataclass(frozen=True)
class PopoSettings:
    teamspace_id: str
    folder_id: str
    folder_url: str

    @classmethod
    def from_env(cls):
        return cls(
            teamspace_id=os.getenv(
                "POPO_TEAMSPACE_ID", ""
            ).strip(),
            folder_id=os.getenv(
                "POPO_SOCIAL_FOLDER_ID", ""
            ).strip(),
            folder_url=os.getenv(
                "POPO_SOCIAL_FOLDER_URL",
                "",
            ).strip(),
        )

    @property
    def configured(self) -> bool:
        return bool(self.teamspace_id and self.folder_id)


def _loads_lenient(text: str):
    try:
        return json.loads(text)
    except Exception:
        match = re.search(r"(\{.*\}|\[.*\])", text or "", re.DOTALL)
        if not match:
            return None
        try:
            return json.loads(match.group(1))
        except Exception:
            return None


def _unwrap_popo_result(obj):
    current = obj
    for _ in range(12):
        if not isinstance(current, dict):
            return current
        if current.get("ok") is False:
            return None
        if "status" in current:
            if current.get("status") != 1:
                return None
            return current.get("data") if "data" in current else current
        if "data" not in current:
            return current
        current = current.get("data")
        if current is None:
            return None
    return None


class PopoClient:
    def __init__(self, settings: PopoSettings):
        self.settings = settings
        self.executable = fabric_cli()
        self.popo_js = _resolve_popo_js()

    @property
    def enabled(self) -> bool:
        return bool(
            (self.executable or self.popo_js)
            and self.settings.teamspace_id
            and self.settings.folder_id
        )

    def call(self, tool: str, **params):
        if self.executable:
            args = [self.executable, f"popo_{tool}"]
        elif self.popo_js:
            args = [
                shutil.which("node") or "node",
                self.popo_js,
                "popo",
                tool,
            ]
        else:
            raise PopoDeliveryError("服务器尚未安装 Fabric CLI / POPO CLI")
        args.extend(f"{key}={value}" for key, value in params.items())
        last_detail = ""
        for attempt in range(3):
            try:
                with fabric_lock():
                    proc = subprocess.run(
                        args,
                        capture_output=True,
                        text=True,
                        encoding="utf-8",
                        errors="replace",
                        timeout=300,
                        env=fabric_env(),
                    )
            except subprocess.TimeoutExpired as exc:
                raise PopoDeliveryError(f"POPO {tool} 请求超时") from exc
            except OSError as exc:
                raise PopoDeliveryError(f"POPO {tool} 调用失败：{exc}") from exc
            output = (proc.stdout or "").strip()
            error = (proc.stderr or "").strip()
            if proc.returncode != 0:
                detail = (error or output or "无错误详情").replace("\n", " ")
                raise PopoDeliveryError(
                    f"POPO {tool} 调用失败（exit={proc.returncode}）：{detail[:300]}"
                )
            if output:
                parsed = _loads_lenient(output)
                result = _unwrap_popo_result(parsed)
                if result is not None:
                    return result
                last_detail = (error or output).replace("\n", " ")[:300]
                break
            last_detail = (error or "POPO 返回空结果").replace("\n", " ")[:300]
            if attempt < 2:
                time.sleep(2)
        raise PopoDeliveryError(f"POPO {tool} 未成功：{last_detail}")

    def children(self) -> list[dict]:
        items: list[dict] = []
        cursor = None
        for _ in range(50):
            params = {
                "folderId": self.settings.folder_id,
                "teamSpaceId": self.settings.teamspace_id,
            }
            if cursor:
                params["cursor"] = cursor
            result = self.call("doc_get_folder_children", **params)
            if not isinstance(result, dict):
                break
            items.extend(result.get("list") or [])
            if not result.get("more"):
                break
            cursor = result.get("nextCursor")
            if not cursor:
                break
        return items

    def find_document(self, title: str) -> tuple[str | None, str | None]:
        for item in self.children():
            if item.get("name") == title and item.get("type") == 1:
                return item.get("docId"), item.get("url")
        return None, None

    def create_document(self, title: str) -> tuple[str, str]:
        result = self.call(
            "doc_create_doc",
            docType=1,
            title=title,
            parentId=self.settings.folder_id,
            teamSpaceId=self.settings.teamspace_id,
        )
        doc_id = result.get("docId") if isinstance(result, dict) else None
        doc_url = result.get("url") if isinstance(result, dict) else None
        if not doc_id:
            raise PopoDeliveryError(f"POPO 创建文档失败：{title}")
        if not doc_url:
            doc_url = (
                "https://docs.popo.netease.com/team/pc/pageDetail/"
                + quote(str(doc_id))
            )
        return str(doc_id), str(doc_url)

    def delete_document(self, doc_id: str):
        try:
            self.call(
                "doc_delete_doc",
                docId=doc_id,
                teamSpaceId=self.settings.teamspace_id,
            )
        except Exception:
            pass

    def upload(self, path: Path, doc_id: str) -> dict:
        temp_path = None
        name_path = None
        try:
            suffix = path.suffix.lower()
            temp_path = Path(tempfile.gettempdir()) / (
                f"popo_{uuid.uuid4().hex[:12]}{suffix}"
            )
            shutil.copyfile(path, temp_path)
            file_name: str = path.name
            if not file_name.isascii():
                name_path = Path(tempfile.gettempdir()) / (
                    f"popo_{uuid.uuid4().hex[:12]}.txt"
                )
                name_path.write_text(file_name, encoding="utf-8")
                file_name = f"@text:{name_path}"
            result = self.call(
                "doc_upload_local_file",
                file=f"@upload:{temp_path}",
                fileName=file_name,
                docId=doc_id,
                teamSpaceId=self.settings.teamspace_id,
            )
            url = result if isinstance(result, str) else (
                result.get("url") if isinstance(result, dict) else None
            )
            if not isinstance(url, str) or not url.startswith("http"):
                raise PopoDeliveryError(f"POPO 附件上传失败：{path.name}")
            return {"url": url, "size": path.stat().st_size, "name": path.name}
        finally:
            if temp_path:
                temp_path.unlink(missing_ok=True)
            if name_path:
                name_path.unlink(missing_ok=True)

    def replace_page(self, doc_id: str, content: str):
        command_path = Path(tempfile.gettempdir()) / (
            f"popo_{uuid.uuid4().hex[:12]}.json"
        )
        command_path.write_text(
            json.dumps(
                {"type": "doc.replace_page", "content": content},
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        try:
            self.call(
                "doc_update_doc",
                docId=doc_id,
                teamSpaceId=self.settings.teamspace_id,
                command=f"@file:{command_path}",
            )
        finally:
            command_path.unlink(missing_ok=True)


def _e(value) -> str:
    return html.escape(str(value if value is not None else "—"), quote=False)


def _n(value) -> str:
    if value is None or value == "":
        return "—"
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return str(value)


def _pct(value) -> str:
    if value is None:
        return "—"
    return f"{float(value):.2%}"


def _table(headers: list[str], rows: list[list[object]], widths=None) -> str:
    widths = widths or [150] * len(headers)
    colgroup = "".join(
        f'<col style="width: {int(width)}px;" />' for width in widths
    )
    head = "".join(f"<th>{_e(item)}</th>" for item in headers)
    body = "".join(
        "<tr>" + "".join(f"<td>{_e(item)}</td>" for item in row) + "</tr>"
        for row in rows
    )
    return (
        "<table title-line-open><colgroup>"
        + colgroup
        + "</colgroup><tbody><tr>"
        + head
        + "</tr>"
        + body
        + "</tbody></table>"
    )


def _analysis_list(items: list[tuple[str, str]], list_id: str) -> str:
    return "".join(
        f'<doc-li list-id="{list_id}" list-type="unordered">'
        f"<b>{_e(label)}</b>{_e(text)}</doc-li>"
        for label, text in items
    )


def _native_html(
    analysis: dict,
    uploaded: dict[str, dict],
) -> str:
    from social_report.report import (
        FIXED_ENDING,
        _linkedin_analysis,
        _linkedin_monthly_rows,
        _linkedin_personal_analysis,
        _monthly_trend_summary,
        _normalize_monthly_rows,
        _platform_analysis,
        _post_interactions,
        _theme_label,
        analysis_principle,
        report_mode_profile,
        report_recommendations,
    )

    request = analysis["request"]
    linkedin = analysis["linkedin"]
    linkedin_mode = linkedin.get("mode") or request.get("linkedin_mode") or "company"
    content = linkedin["content"]
    visitors = linkedin["visitors"]
    followers = linkedin["followers"]
    fb = request.get("facebook") or {}
    ig = request.get("instagram") or {}
    report_type = str(request.get("report_type") or "monthly").strip().lower()
    profile = report_mode_profile(report_type)
    annual_framework = profile["framework"] == "annual"
    period = f"{request['report_start']}—{request['report_end']}"
    parts = [
        "<doc>",
        f"<h1>{_e(request['brand'])} {_e(analysis['report_title'])}</h1>",
        f"<p><b>报告区间：</b>{_e(period)}　｜　<b>数据来源：</b>Meta Business Suite</p>",
        f"<blockquote><p>{_e(analysis_principle(report_type))}</p></blockquote>",
        "<h2>第一部分　Facebook</h2>",
        "<h3>1.1　数据概览</h3>",
        _table(
            ["指标", "本期数据", "说明"],
            [
                ["内容浏览量", _n(fb.get("views")), "报告期内可核对的内容浏览量"],
                [
                    "浏览人数",
                    _n(fb.get("viewers") if fb.get("viewers") is not None else fb.get("reach")),
                    "看到内容的去重人数；缺失时保留为“—”",
                ],
                ["发布内容", _n(fb.get("posts")), "报告期内发布的内容数量"],
                ["内容互动次数", _n(fb.get("interactions")), "点赞、评论、分享等可计互动"],
            ],
            [160, 120, 420],
        ),
    ]

    def append_monthly_trend(
        platform: str,
        monthly,
        specs: list[tuple[str, str, str, str]],
    ) -> None:
        rows = _normalize_monthly_rows(monthly)
        available = [
            spec
            for spec in specs
            if any(isinstance(row.get(spec[0]), (int, float)) for row in rows)
        ]
        if not rows or not available:
            parts.append(
                "<p>本次任务没有保存可核对的月度序列，因此不绘制估算趋势图。"
                "请重新同步平台数据后再生成；系统不会把汇总总数平均拆分到各月。</p>"
            )
            return
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
        parts.append(f"<p>{_e(note)}</p>")
        chart = uploaded.get(f"{platform.lower()}_monthly_trend")
        if chart:
            parts.append(f'<img src="{chart["url"]}" />')
            parts.append(f"<p>{_e(platform)} 月度自然内容趋势</p>")
        parts.append(f"<p>{_e(_monthly_trend_summary(rows, available))}</p>")
        parts.append(
            _table(
                ["月份"] + [spec[1] for spec in available],
                [
                    [row["month"]]
                    + [_n(row.get(spec[0])) for spec in available]
                    for row in rows
                ],
                [110] + [160] * len(available),
            )
        )
    if annual_framework:
        parts.append("<h3>1.2　月度自然内容趋势</h3>")
        append_monthly_trend(
            "Facebook",
            fb.get("monthly"),
            [
                ("views", "浏览量", "Views", "#2F5A9E"),
                ("daily_reach_sum", "日度触达累计", "Daily reach sum", "#5AA1D8"),
            ],
        )
    fb_ranking_number = "1.3" if annual_framework else "1.2"
    fb_analysis_number = "1.4" if annual_framework else "1.3"
    parts.append(f"<h3>{fb_ranking_number}　自然内容排行</h3>")
    if uploaded.get("facebook"):
        parts.append(f'<img src="{uploaded["facebook"]["url"]}" />')
        parts.append(f"<p>Facebook 自然内容排行（报告区间：{_e(period)}）</p>")
    parts.extend([
        f"<h3>{fb_analysis_number}　{_e(profile['analysis_heading'])}</h3>",
        _analysis_list(
            _platform_analysis(
                "Facebook", fb, request.get("facebook_highlight", ""), report_type
            ),
            "facebook-analysis",
        ),
        "<h2>第二部分　Instagram</h2>",
        "<h3>2.1　数据概览</h3>",
        _table(
            ["指标", "本期数据", "说明"],
            [
                ["内容浏览量", _n(ig.get("views")), "报告期内可核对的内容浏览量"],
                ["覆盖人数", _n(ig.get("reach")), "看到内容的去重账号数"],
                ["发布内容", _n(ig.get("posts")), "报告期内发布的内容数量"],
                ["内容互动次数", _n(ig.get("interactions")), "点赞、评论、分享等可计互动"],
            ],
            [160, 120, 420],
        ),
    ])
    if annual_framework:
        parts.append("<h3>2.2　月度自然内容趋势</h3>")
        append_monthly_trend(
            "Instagram",
            ig.get("monthly"),
            [
                ("views", "浏览量", "Views", "#2F5A9E"),
                ("daily_reach_sum", "日度触达累计", "Daily reach sum", "#5AA1D8"),
            ],
        )
    ig_ranking_number = "2.3" if annual_framework else "2.2"
    ig_analysis_number = "2.4" if annual_framework else "2.3"
    parts.append(f"<h3>{ig_ranking_number}　自然内容排行</h3>")
    if uploaded.get("instagram"):
        parts.append(f'<img src="{uploaded["instagram"]["url"]}" />')
        parts.append(f"<p>Instagram 自然内容排行（报告区间：{_e(period)}）</p>")
    parts.extend([
        f"<h3>{ig_analysis_number}　{_e(profile['analysis_heading'])}</h3>",
        _analysis_list(
            _platform_analysis(
                "Instagram", ig, request.get("instagram_highlight", ""), report_type
            ),
            "instagram-analysis",
        ),
    ])

    def append_posts_table(limit: int, impressions_label: str):
        top_rows = []
        for post in (linkedin.get("posts") or {}).get("top_impressions", [])[:limit]:
            title = _theme_label(post.get("title") or "")
            url = str(post.get("url") or "")
            if url.startswith(("http://", "https://")):
                safe_url = html.escape(url, quote=True)
                title = f'<a href="{safe_url}">{_e(title)}</a>'
            top_rows.append([
                title,
                post.get("created_date") or "—",
                _n(post.get("impressions")),
                _n(post.get("clicks")),
                _pct(post.get("ctr")),
                _n(_post_interactions(post)),
            ])
        if not top_rows:
            return
        rows_html = []
        for row in top_rows:
            rows_html.append(
                "<tr><td>"
                + row[0]
                + "</td>"
                + "".join(f"<td>{_e(item)}</td>" for item in row[1:])
                + "</tr>"
            )
        parts.append(
            '<table title-line-open><colgroup><col style="width: 320px;" />'
            '<col style="width: 105px;" /><col style="width: 100px;" />'
            '<col style="width: 80px;" /><col style="width: 80px;" />'
            '<col style="width: 80px;" /></colgroup><tbody>'
            f"<tr><th>内容主题</th><th>日期</th><th>{_e(impressions_label)}</th>"
            "<th>点击</th><th>CTR</th><th>互动</th></tr>"
            + "".join(rows_html)
            + "</tbody></table>"
        )

    if linkedin_mode == "company":
        parts.extend([
            "<h2>第三部分　LinkedIn</h2>",
            "<h3>3.1 数据概览</h3>",
            _table(
                ["原生展示", "点击", "互动率", "主页访问", "独立访客", "新增关注"],
                [[
                    _n(content.get("impressions_organic")),
                    _n(content.get("clicks_organic")),
                    _pct(content.get("engagement_rate")),
                    _n(visitors.get("page_views_total")),
                    _n(visitors.get("unique_total")),
                    _n(followers.get("organic")),
                ]],
                [125, 100, 100, 110, 110, 110],
            ),
        ])
        if annual_framework:
            parts.append("<h3>3.2 月度自然内容趋势</h3>")
            append_monthly_trend(
                "LinkedIn",
                _linkedin_monthly_rows(linkedin),
                [
                    ("impressions_organic", "自然展示量", "Organic impressions", "#2F5A9E"),
                    ("daily_reach_sum", "日度触达累计", "Daily reach sum", "#5AA1D8"),
                ],
            )
        linkedin_ranking_number = "3.3" if annual_framework else "3.2"
        linkedin_analysis_number = "3.4" if annual_framework else "3.3"
        parts.append(f"<h3>{linkedin_ranking_number} 自然内容排行</h3>")
        append_posts_table(6, "展示")
        parts.extend([
            f"<h3>{linkedin_analysis_number} {_e(profile['analysis_heading'])}</h3>",
            _analysis_list(
                _linkedin_analysis(linkedin, report_type), "linkedin-analysis"
            ),
        ])
        linkedin_conclusion = (
            "<p>本期三端均保持自然内容发布。Facebook 与 Instagram 仍以内容同步、"
            "品牌信息沉淀和客户信任建立为主要作用；LinkedIn 取得 "
            f"{_e(_n(content.get('impressions_organic')))} 次原生展示、"
            f"{_e(_n(content.get('clicks_organic')))} 次点击和 "
            f"{_e(_n(followers.get('organic')))} 名新增关注者，"
            "已形成更清晰的内容反馈与主页承接链路。</p>"
        )
        record_advice = (
            "每个报告周期保存 Meta 内容排行截图和 LinkedIn 三份公共主页导出，"
            "保持报告口径、日期和源文件可追溯。"
        )
    elif linkedin_mode == "personal":
        personal = linkedin["personal"]
        parts.extend([
            "<h2>第三部分　LinkedIn 个人账号内容表现</h2>",
            "<h3>3.1 数据口径与概览</h3>",
            f"<blockquote><p>{_e('；'.join(linkedin.get('warnings') or []))}</p></blockquote>",
            _table(
                ["Top 帖子样本", "Top 10 浏览/展示", "Top 10 互动", "报告期总浏览", "主页访问", "当前关注者"],
                [[
                    _n(linkedin["posts"]["count"]),
                    _n(personal.get("top10_impressions")),
                    _n(personal.get("top10_interactions")),
                    _n(personal.get("period_impressions")),
                    _n(personal.get("profile_views")),
                    _n(personal.get("current_followers")),
                ]],
                [115, 140, 115, 130, 110, 120],
            ),
            "<p>Top 10 累计仅代表人工入选样本；人工未提供的指标保留为“—”。</p>",
        ])
        if annual_framework:
            parts.extend([
                "<h3>3.2 月度自然内容趋势</h3>",
                "<p>个人账号仅提供人工筛选样本，没有可核对的月度全量序列，因此本节不绘制估算趋势图。</p>",
            ])
        linkedin_ranking_number = "3.3" if annual_framework else "3.2"
        linkedin_analysis_number = "3.4" if annual_framework else "3.3"
        parts.append(f"<h3>{linkedin_ranking_number} 自然内容排行</h3>")
        append_posts_table(10, "浏览/展示")
        for item in uploaded.get("linkedin_personal_screenshots") or []:
            parts.append(f'<img src="{item["url"]}" />')
            parts.append("<p>LinkedIn 个人账号 Top 10 人工截图</p>")
        parts.extend([
            f"<h3>{linkedin_analysis_number} {_e(profile['analysis_heading'])}</h3>",
            _analysis_list(
                _linkedin_personal_analysis(linkedin, report_type),
                "linkedin-personal-analysis",
            ),
        ])
        linkedin_conclusion = (
            "<p>Facebook 与 Instagram 继续承担品牌信息沉淀和客户信任建立；"
            f"LinkedIn 个人账号本期基于 {_e(_n(linkedin['posts']['count']))} 条"
            "人工筛选帖子进行样本分析，样本浏览量/展示量合计 "
            f"{_e(_n(personal.get('top10_impressions')))}。该数值不等同于账号报告期全量，"
            "访客和关注者缺失项未作推测。</p>"
        )
        record_advice = (
            "每个报告周期沿用同一份 LinkedIn 个人账号 Top 10 模板和筛选标准，"
            "并保留原始截图。"
        )
    else:
        parts.extend([
            "<h2>第三部分　LinkedIn</h2>",
            "<h3>3.1 本期数据说明</h3>",
            "<blockquote><p>本期未提供 LinkedIn 公共主页导出、个人账号 Top 10 "
            "汇总表或截图，因此不展示指标，也不将缺失数据记为 0。</p></blockquote>",
        ])
        linkedin_conclusion = (
            "<p>本期结论基于 Facebook 与 Instagram 数据形成；LinkedIn 未提供数据，"
            "因此未纳入本期效果判断，也没有将缺失指标记为 0。</p>"
        )
        record_advice = (
            "下个周期如有 LinkedIn 数据，请选择公共主页导出或个人账号 Top 10 "
            "模式留存原始依据。"
        )

    parts.extend([
        "<h2>第四部分　综合结论与下阶段优化建议</h2>",
        "<h3>4.1　综合结论</h3>",
        f"<blockquote><p>{_e(profile['conclusion_scope'])}：本报告重点回答“"
        f"{_e(profile['question'])}”，结论仅基于本次可核对的数据与内容排行形成。</p></blockquote>",
        linkedin_conclusion,
        (
            "<p>后续应优先复用已经获得相对高浏览的主题，并把企业介绍转化为具体场景、"
            "流程、案例和结果证明。两个平台可以共用选题，但应分别使用适合平台的表达"
            "形式，并在内容结尾设置单一、明确的下一步动作。</p>"
        ),
        "<h3>4.2　下阶段优化建议</h3>",
        _analysis_list(
            report_recommendations(report_type),
            "recommendations",
        ),
        f"<blockquote><p><b>{_e(FIXED_ENDING)}</b></p></blockquote>",
        "<h2>附件</h2>",
    ])
    if uploaded.get("word"):
        word = uploaded["word"]
        parts.append(
            f"<p>最终 Word 报告：{_e(word['name'])}</p>"
            f'<attachment src="{word["url"]}" file-size="{word["size"]}" />'
        )
    parts.append("</doc>")
    return "".join(parts)


def _title(analysis: dict) -> str:
    request = analysis["request"]
    return (
        f"{request['brand']}_{analysis['report_title']}_"
        f"{request['report_start'].replace('-', '.')}-"
        f"{request['report_end'].replace('-', '.')}"
    )


def deliver_job(job_dir: Path, log) -> dict:
    status = popo_status(auto_login=True, force=True)
    if not status.get("connected"):
        raise PopoDeliveryError(
            status.get("message") or "POPO 尚未连接，请先在平台完成授权"
        )
    settings = PopoSettings.from_env()
    if not settings.configured:
        raise PopoDeliveryError("POPO 团队空间或社媒报告文件夹尚未配置")
    analysis_path = job_dir / "outputs" / "analysis.json"
    if not analysis_path.is_file():
        raise PopoDeliveryError("报告分析数据不存在，无法生成 POPO 原生文档")
    analysis = json.loads(analysis_path.read_text(encoding="utf-8"))
    title = _title(analysis)
    client = PopoClient(settings)
    existing_id, existing_url = client.find_document(title)
    created = not bool(existing_id)
    if existing_id:
        doc_id = existing_id
        if not existing_url:
            existing_url = (
                "https://docs.popo.netease.com/team/pc/"
                f"pageDetail/{existing_id}"
            )
        doc_url = existing_url
        log.write(f"[POPO] 同名文档已存在，将原位更新：{existing_url}\n")
    else:
        doc_id, doc_url = client.create_document(title)
        log.write(f"[POPO] 已创建原生文档：{title}\n")
    try:
        request = analysis["request"]
        inputs = job_dir / "inputs"
        outputs = job_dir / "outputs"
        docx_files = sorted(outputs.glob("*.docx"))
        if not docx_files:
            raise PopoDeliveryError("最终 Word 报告不存在")
        files = {
            "facebook": inputs / request["files"]["facebook_screenshot"],
            "instagram": inputs / request["files"]["instagram_screenshot"],
            "word": docx_files[0],
        }
        uploaded = {}
        for key, path in files.items():
            if not path.is_file():
                raise PopoDeliveryError(f"POPO 交付资源不存在：{path.name}")
            uploaded[key] = client.upload(path, doc_id)
            log.write(f"[POPO] 已上传：{path.name}\n")
        uploaded["linkedin_personal_screenshots"] = []
        for filename in request["files"].get("linkedin_personal_screenshots") or []:
            path = inputs / filename
            if not path.is_file():
                raise PopoDeliveryError(f"POPO 交付资源不存在：{path.name}")
            item = client.upload(path, doc_id)
            uploaded["linkedin_personal_screenshots"].append(item)
            log.write(f"[POPO] 已上传：{path.name}\n")
        charts = job_dir / "charts"
        for platform in ("facebook", "instagram", "linkedin"):
            path = charts / f"{platform}_monthly_trend.png"
            if not path.is_file():
                continue
            key = f"{platform}_monthly_trend"
            uploaded[key] = client.upload(path, doc_id)
            log.write(f"[POPO] 已上传：{path.name}\n")
        client.replace_page(doc_id, _native_html(analysis, uploaded))
        log.write(f"[POPO] 原生正文写入完成：{doc_url}\n")
    except Exception:
        if created:
            client.delete_document(doc_id)
            log.write("[POPO] 正文或附件交付失败，已清理空文档，可安全重试\n")
        else:
            log.write("[POPO] 同名文档更新失败，已保留原文档，可安全重试\n")
        raise
    return {
        "url": doc_url,
        "doc_id": doc_id,
        "created": created,
        "title": title,
    }
