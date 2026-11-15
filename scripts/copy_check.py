#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""copy_check.py —— 文案机检（Copy QA）

查的是「可计算的」部分：槽位字数、空话词表、AI 句式、句长、同构排比、数字无出处、
同页重复短语、标题是否命题、审校残留标记。**不判断写得好不好**，那部分见
`references/copywriting-antislop.md` 与 `workflows/06` 的 §6.5b。

用法：
    python scripts/copy_check.py --deck source/deck.json
    python scripts/copy_check.py --deck source/deck.json --out build/copy-report.md --json build/copy.json
    python scripts/copy_check.py --deck source/deck.json --copy records/copy-deck.md
    python scripts/copy_check.py --deck source/deck.json --lenient   # 仅 P0 才非零退出

退出码：
    0  没有 P0，且（默认模式下）没有 P1
    1  存在 P0，或默认模式下存在 P1
    2  参数、文件或 deck 校验错误

阈值全部读 schema/thresholds.json 的 `copy` 段与 `type.maxTitleChars`、`density.maxCharsPerSlide`；
本文件不写死词表与数字。
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

try:  # Windows 控制台默认 GBK，中文报告会乱码
    sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    sys.stderr.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
except (AttributeError, OSError):
    pass

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent
DEFAULT_THRESHOLDS = REPO_ROOT / "schema" / "thresholds.json"

COPY_SCHEMA_VERSION = "1.0.0"
SEVERITIES = ("P0", "P1", "P2", "P3")
SEVERITY_RANK = {level: index for index, level in enumerate(SEVERITIES)}

# 槽位 → thresholds.copy.limits 的键（不在此列的角色不查字数）
ROLE_LIMIT_KEYS = {
    "coverTitle": "coverTitle",
    "lead": "lead",
    "h1": "h1",
    "h2": "h2",
    "h3": "h3",
    "body": "body",
    "kpi": "kpi",
    "label": "label",
    "caption": "caption",
    "source": "source",
    "quote": "quote",
    "note": "note",
}

# 标题里出现这些字（或数字）就认为它是命题句，不再提示
CLAIM_HINTS = (
    "是", "要", "需", "应", "先", "再", "才", "不", "未", "已", "将", "让", "给", "把",
    "降", "升", "涨", "跌", "减", "增", "超", "低于", "高于", "等于", "达到", "决定",
    "导致", "带来", "取决", "优于", "不如", "没有", "有", "从", "到", "每", "只", "仍",
    "如何", "为什么", "怎样",
)

# 残留标记：出现在成片文案里就是事故
UNREVIEWED_MARKERS = ("待确认", "TBD", "TODO", "xxx", "[假设]", "【待定】", "？？")

CJK_RE = re.compile(r"[\u4e00-\u9fa5]")
DIGIT_RE = re.compile(r"[0-9]")
SENT_SPLIT_RE = re.compile(r"[。！？!?；;]+")
CLAUSE_SPLIT_RE = re.compile(r"[，,。！？!?；;、：:]+")
CJK_RUN_RE = re.compile(r"[\u4e00-\u9fa5]{4}")


# ---------------------------------------------------------------------------
# 动态加载同目录 lib（与 qa_check.py 一致的做法）
# ---------------------------------------------------------------------------


def _load_lib(module_name: str, filename: str):
    path = SCRIPT_DIR / "lib" / filename
    if not path.is_file():
        raise SystemExit("错误：缺少依赖文件 %s（应与 copy_check.py 同仓库）" % path)
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise SystemExit("错误：无法加载 %s" % path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


deckload = _load_lib("copy_deckload", "deckload.py")


# ---------------------------------------------------------------------------
# 阈值读取
# ---------------------------------------------------------------------------


class Thresholds:
    def __init__(self, doc: Dict[str, Any]):
        self.doc = doc or {}

    def get(self, dotted: str, default: Any = None) -> Any:
        node: Any = self.doc
        for part in dotted.split("."):
            if not isinstance(node, dict) or part not in node:
                return default
            node = node[part]
        return node

    def num(self, dotted: str, default: float) -> float:
        value = self.get(dotted, default)
        try:
            return float(value)
        except (TypeError, ValueError):
            return float(default)


def load_thresholds(path: Optional[str]) -> Thresholds:
    target = Path(path).expanduser() if path else DEFAULT_THRESHOLDS
    if not target.is_file():
        raise SystemExit("错误：找不到 thresholds.json：%s" % target)
    with target.open("r", encoding="utf-8") as handle:
        return Thresholds(json.load(handle))


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------


@dataclass
class Issue:
    level: str
    check: str
    page: str
    where: str
    message: str
    fix: str

    def as_dict(self) -> Dict[str, str]:
        return {
            "level": self.level,
            "check": self.check,
            "page": self.page,
            "where": self.where,
            "message": self.message,
            "fix": self.fix,
        }


@dataclass
class PageData:
    index: int
    slide_id: str
    layout: str
    texts: List[Tuple[str, str]] = field(default_factory=list)  # (role, text)

    @property
    def label(self) -> str:
        sid = self.slide_id
        if re.fullmatch(r"s\d{2,3}", sid or ""):
            return sid
        return "s%02d %s" % (self.index + 1, sid)

    @property
    def full_text(self) -> str:
        return "\n".join(text for _role, text in self.texts)

    def roles(self) -> List[str]:
        return [role for role, _text in self.texts]

    def has_role(self, role: str) -> bool:
        return role in self.roles()


def element_text(el: Dict[str, Any]) -> str:
    """取元素的可读文本。

    优先级：paragraphs > runs > text。三者同时存在时只取最高优先级的那个，
    否则同一段文字会被拼接两次，字数与重复短语检查都会翻倍。
    """
    paragraphs = el.get("paragraphs") or []
    if paragraphs:
        chunks: List[str] = []
        for para in paragraphs:
            runs = para.get("runs") or []
            if runs:
                chunks.append("".join(str(r.get("text", "")) for r in runs if isinstance(r.get("text"), str)))
            elif isinstance(para.get("text"), str):
                chunks.append(para["text"])
        return "\n".join(chunk for chunk in chunks if chunk)
    runs = el.get("runs") or []
    if runs:
        return "".join(str(r.get("text", "")) for r in runs if isinstance(r.get("text"), str))
    if isinstance(el.get("text"), str):
        return el["text"]
    return ""


def visible_chars(text: str) -> int:
    """字数口径：去空白后逐字符计（中文 1 字，西文 1 字符）。"""
    return len(re.sub(r"\s+", "", text))


def collect_pages(deck) -> List[PageData]:
    pages: List[PageData] = []
    for si, slide in enumerate(deck.slides):
        page = PageData(
            index=si,
            slide_id=str(slide.get("id", "?")),
            layout=str(slide.get("layout", "")),
        )
        for el in slide.get("elements") or []:
            if el.get("kind") != "text":
                continue
            text = element_text(el).strip()
            if not text:
                continue
            role = str(el.get("role") or "")
            page.texts.append((role, text))
        pages.append(page)
    return pages


# ---------------------------------------------------------------------------
# 各项检查
# ---------------------------------------------------------------------------


def check_slot_length(pages: List[PageData], th: Thresholds) -> List[Issue]:
    """CP-01 槽位字数超限。"""
    issues: List[Issue] = []
    limits = th.get("copy.limits", {}) or {}
    for page in pages:
        for role, text in page.texts:
            key = ROLE_LIMIT_KEYS.get(role)
            if not key:
                continue
            limit = limits.get(key)
            if not isinstance(limit, (int, float)):
                continue
            count = visible_chars(text)
            if count > limit:
                # 核心槽位超标会把版面撑破；label/kpi/h2/h3 常被借用于卡片标题与说明，
                # 超标只作提示，避免把"角色借用"误报成"文案超长"。
                level = "P2" if role in ("label", "kpi", "h2", "h3") else "P1"
                issues.append(
                    Issue(
                        level,
                        "CP-01",
                        page.label,
                        "role=%s" % role,
                        "文案 %d 字，超过 %s 的 %d 字上限" % (count, role, int(limit)),
                        "删修饰词与重复句，不缩字号；或按 copywriting-patterns §3 拆到别的槽位",
                    )
                )
    return issues


def check_empty_phrases(pages: List[PageData], th: Thresholds) -> List[Issue]:
    """CP-02 空话短语。"""
    issues: List[Issue] = []
    phrases = [p for p in (th.get("copy.emptyPhrases", []) or []) if isinstance(p, str)]
    sev = str(th.get("copy.slopSeverity.emptyPhrase", "P1"))
    for page in pages:
        text = page.full_text
        hits = [phrase for phrase in phrases if phrase in text]
        for phrase in hits:
            issues.append(
                Issue(
                    sev,
                    "CP-02",
                    page.label,
                    "空话短语",
                    "命中空话短语「%s」" % phrase,
                    "整句删除，不要换一个同义词顶替（见 copywriting-antislop §2 特征 1）",
                )
            )
    return issues


def check_filler_adjectives(pages: List[PageData], th: Thresholds) -> List[Issue]:
    """CP-03 填充形容词密度（每页 >2 个才报）。"""
    issues: List[Issue] = []
    words = [w for w in (th.get("copy.fillerAdjectives", []) or []) if isinstance(w, str)]
    sev = str(th.get("copy.slopSeverity.fillerAdjective", "P2"))
    for page in pages:
        text = page.full_text
        hits = [word for word in words if word in text]
        if len(hits) > 2:
            issues.append(
                Issue(
                    sev,
                    "CP-03",
                    page.label,
                    "形容词密度",
                    "本页出现 %d 个填充形容词：%s" % (len(hits), "、".join(hits)),
                    "每个形容词换成可核对的量或动作；实在没有量就删",
                )
            )
    return issues


def check_ai_patterns(pages: List[PageData], th: Thresholds) -> List[Issue]:
    """CP-04 AI 句式（正则）。"""
    issues: List[Issue] = []
    patterns = [p for p in (th.get("copy.aiPatterns", []) or []) if isinstance(p, str)]
    sev = str(th.get("copy.slopSeverity.aiPattern", "P1"))
    compiled: List[Tuple[str, re.Pattern]] = []
    for pattern in patterns:
        try:
            compiled.append((pattern, re.compile(pattern)))
        except re.error:
            continue
    for page in pages:
        text = page.full_text
        for pattern, regex in compiled:
            match = regex.search(text)
            if match:
                issues.append(
                    Issue(
                        sev,
                        "CP-04",
                        page.label,
                        "AI 句式",
                        "命中模板句式「%s」：%s" % (pattern, match.group(0)),
                        "拆成两句直陈，或只保留最有力的一项（见 copywriting-antislop §2 特征 3–4）",
                    )
                )
    return issues


def check_long_clauses(pages: List[PageData], th: Thresholds) -> List[Issue]:
    """CP-05 单句过长。"""
    issues: List[Issue] = []
    limit = int(th.num("copy.sentenceMaxChars", 45))
    sev = str(th.get("copy.slopSeverity.sentenceTooLong", "P2"))
    for page in pages:
        clauses: List[str] = []
        for role, text in page.texts:
            if role in ("source", "pageNumber"):
                continue
            # 代码行 / 纯西文命令行：不是中文文案，不参与句长检查
            if not CJK_RE.search(text):
                continue
            if CJK_RE.findall(text) and len(CJK_RE.findall(text)) / max(1, visible_chars(text)) < 0.15:
                continue
            if re.search(r"[<>|/\\]\S*\s+-{1,2}\w", text):
                continue
            for raw in re.split(r"[\n]", text):
                for piece_inner in CLAUSE_SPLIT_RE.split(raw):
                    piece = piece_inner.strip()
                    if piece:
                        clauses.append(piece)
            continue
            for raw in CLAUSE_SPLIT_RE.split(text):  # pragma: no cover - 兼容旧逻辑
                piece = raw.strip()
                if piece:
                    clauses.append(piece)
        for piece in clauses:
            count = visible_chars(piece)
            if count > limit:
                issues.append(
                    Issue(
                        sev,
                        "CP-05",
                        page.label,
                        "长句",
                        "单句 %d 字（上限 %d）：%s" % (count, limit, piece[:24] + ("…" if len(piece) > 24 else "")),
                        "在逗号处断成两句；把状语（条件/范围）挪到单独一句",
                    )
                )
                break  # 每页只报第一处，避免刷屏
    return issues


def check_paragraph_uniform(pages: List[PageData], th: Thresholds) -> List[Issue]:
    """CP-06 同构排比：连续 3 段字数相同（±1）。"""
    issues: List[Issue] = []
    need = int(th.num("copy.paragraphRepeatMin", 3))
    sev = str(th.get("copy.slopSeverity.paragraphUniform", "P2"))
    for page in pages:
        for role, text in page.texts:
            paras = [p.strip() for p in text.split("\n") if p.strip()]
            if len(paras) < need:
                continue
            lengths_all = sorted(visible_chars(p) for p in paras)
            median = lengths_all[len(lengths_all) // 2]
            # 列表/读数清单本来就是等长的（"1200 rpm 74.8 ℃" ×5），不判排比
            if median < 14:
                continue
            # 编号清单（"0 需求探索…"）也是等长结构，同样不判排比
            numbered = sum(1 for para in paras if re.match(r"^\d+\s", para))
            if numbered >= max(2, int(len(paras) * 0.7)):
                continue
            run: List[int] = []
            lengths = [visible_chars(p) for p in paras]
            for idx in range(len(lengths) - need + 1):
                window = lengths[idx : idx + need]
                if max(window) - min(window) <= 1 and min(window) >= 8:
                    run = window
                    issues.append(
                        Issue(
                            sev,
                            "CP-06",
                            page.label,
                            "role=%s 第 %d–%d 段" % (role, idx + 1, idx + need),
                            "连续 %d 段字数几乎相同（%s），读起来是对仗排比" % (need, "、".join(str(n) for n in run)),
                            "打散其中一至两段：改成 8 字以内的短句，或扩到 30 字以上",
                        )
                    )
                    break
            if run:
                break
    return issues


QUANTITY_RE = re.compile(
    r"\d+(?:\.\d+)?\s*(?:℃|°C|%|mm|cm|m|km|g|kg|K|W|kW|dB|pt|px|in|s|ms|h|min|"
    r"次|条|页|倍|元|万|亿|台|个|家|组|轮|周|天|月|年)"
)


def check_digits_without_source(pages: List[PageData], th: Thresholds) -> List[Issue]:
    """CP-07 页面有"数字+单位"但没有出处。

    只查带单位的量（6.8 K、40.3 dB、25%），不查序号（"02 — 问题"、"0 需求探索"）。
    出处可能写在别的页（数据页），所以这一项是提示级，不是阻断级。
    """
    issues: List[Issue] = []
    sev = str(th.get("copy.slopSeverity.missingSource", "P2"))
    for page in pages:
        if page.layout in ("cover", "closing") or page.has_role("source"):
            continue
        text = page.full_text
        if not QUANTITY_RE.search(text):
            continue
        issues.append(
            Issue(
                sev,
                "CP-07",
                page.label,
                "出处缺失",
                "本页含数字但没有 source 槽位",
                "补一行出处的条件与口径（例：n=3，45 ℃ 环温，1200–3600 rpm）；虚构数据写「演示数据（虚构）」",
            )
        )
    return issues


def check_title_claims(pages: List[PageData], th: Thresholds) -> List[Issue]:
    """CP-08 标题不是命题句（启发式，P3）。"""
    issues: List[Issue] = []
    max_chars = int(th.num("type.maxTitleChars", 25))
    sev = str(th.get("copy.slopSeverity.titleNotClaim", "P3"))
    for page in pages:
        if page.layout in ("cover", "closing", "section"):
            continue
        for role, text in page.texts:
            if role not in ("h1", "coverTitle"):
                continue
            title = text.replace("\n", "").strip()
            if not title or visible_chars(title) > max_chars:
                continue
            if DIGIT_RE.search(title):
                continue
            if any(hint in title for hint in CLAIM_HINTS):
                continue
            if re.search(r"[，,：:—·]", title):
                continue
            issues.append(
                Issue(
                    sev,
                    "CP-08",
                    page.label,
                    "role=%s" % role,
                    "标题像话题名「%s」，读不出主张" % title,
                    "改成可判断真假的一句话：加数字，或改成「A 决定 B」结构",
                )
            )
    return issues


def check_repeated_phrase(pages: List[PageData], th: Thresholds) -> List[Issue]:
    """CP-09 同页重复短语（4 字窗口出现 ≥3 次）。"""
    issues: List[Issue] = []
    need = int(th.num("copy.repeatedPhraseMin", 4))
    sev = str(th.get("copy.slopSeverity.repeatedPhrase", "P3"))
    for page in pages:
        text = page.full_text
        counts: Dict[str, int] = {}
        for match in CJK_RUN_RE.finditer(text):
            gram = match.group(0)
            counts[gram] = counts.get(gram, 0) + 1
        repeated = [(gram, n) for gram, n in counts.items() if n >= need]
        if not repeated:
            continue
        repeated.sort(key=lambda item: -item[1])
        gram, n = repeated[0]
        issues.append(
            Issue(
                sev,
                "CP-09",
                page.label,
                "重复短语",
                "「%s」在本页出现 %d 次" % (gram, n),
                "同页只保留一次；其余改成指代或删掉",
            )
        )
    return issues


def check_unreviewed_markers(pages: List[PageData], th: Thresholds) -> List[Issue]:
    """CP-10 成片里残留审校标记。"""
    issues: List[Issue] = []
    sev = str(th.get("copy.slopSeverity.unreviewedMarker", "P1"))
    for page in pages:
        text = page.full_text
        for marker in UNREVIEWED_MARKERS:
            if marker in text:
                issues.append(
                    Issue(
                        sev,
                        "CP-10",
                        page.label,
                        "残留标记",
                        "文案里还有「%s」" % marker,
                        "文案未定稿不得进制作；确认后从 deck.json 删除标记（见 workflows/01b S9）",
                    )
                )
    return issues


def check_copy_md(path: Path, th: Thresholds) -> List[Issue]:
    """对 records/copy-deck.md 做状态检查：还有没有未确认行。"""
    issues: List[Issue] = []
    if not path.is_file():
        return issues
    text = path.read_text(encoding="utf-8", errors="replace")
    pending = [line.strip() for line in text.splitlines() if "[待确认]" in line]
    if pending:
        issues.append(
            Issue(
                "P1",
                "CP-11",
                "copy-deck.md",
                "第 %d 行起" % 1,
                "还有 %d 行文案处于 [待确认]" % len(pending),
                "交用户逐行确认，或写进「待确认区」的默认档并告知（workflows/01b C6）",
            )
        )
    if "状态" in text and "已冻结" not in text:
        issues.append(
            Issue(
                "P2",
                "CP-12",
                "copy-deck.md",
                "头部状态",
                "文案稿尚未冻结",
                "全部确认后把头部状态改为「已冻结」并写版本号（workflows/01b S9）",
            )
        )
    return issues


# ---------------------------------------------------------------------------
# 报告
# ---------------------------------------------------------------------------


def sort_issues(issues: Iterable[Issue]) -> List[Issue]:
    return sorted(issues, key=lambda item: (SEVERITY_RANK.get(item.level, 9), item.check, item.page))


def summarize(issues: List[Issue]) -> Dict[str, int]:
    summary = {level: 0 for level in SEVERITIES}
    for issue in issues:
        summary[issue.level] = summary.get(issue.level, 0) + 1
    return summary


def render_markdown(
    issues: List[Issue],
    summary: Dict[str, int],
    deck_path: str,
    thresholds_version: str,
    pages: List[PageData],
    exit_code: int,
) -> str:
    lines: List[str] = []
    lines.append("# 文案机检报告（Copy QA）")
    lines.append("")
    lines.append("| 项 | 值 |")
    lines.append("|---|---|")
    lines.append("| 检查对象 | `%s` |" % deck_path)
    lines.append("| 页数 | %d |" % len(pages))
    lines.append("| 阈值版本 | %s |" % thresholds_version)
    lines.append("| 生成时间 | %s |" % datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    lines.append("| 检查脚本 | `scripts/copy_check.py` (schema %s) |" % COPY_SCHEMA_VERSION)
    lines.append("")
    lines.append("**SUMMARY P0=%d P1=%d P2=%d P3=%d issues=%d pages=%d exit=%d**" % (
        summary.get("P0", 0),
        summary.get("P1", 0),
        summary.get("P2", 0),
        summary.get("P3", 0),
        len(issues),
        len(pages),
        exit_code,
    ))
    lines.append("")
    if not issues:
        lines.append("未发现可计算的问题。**这不等于文案达标**：句读、语气、是否像人话，仍需按")
        lines.append("`references/copywriting-antislop.md` 第 2 节人工过一遍。")
        lines.append("")
        return "\n".join(lines) + "\n"

    lines.append("## 问题清单")
    lines.append("")
    lines.append("| 级别 | 检查项 | 页 | 位置 | 问题 | 修复建议 |")
    lines.append("|---|---|---|---|---|---|")
    for issue in issues:
        lines.append(
            "| %s | %s | %s | %s | %s | %s |"
            % (
                issue.level,
                issue.check,
                issue.page,
                _escape(issue.where),
                _escape(issue.message),
                _escape(issue.fix),
            )
        )
    lines.append("")
    lines.append("## 分级含义")
    lines.append("")
    lines.append("| 级别 | 含义 | 处理 |")
    lines.append("|---|---|---|")
    lines.append("| P0 | 阻断（本脚本目前不产生） | 必须修 |")
    lines.append("| P1 | 进阶段 2 前必须修 | 修完重跑 |")
    lines.append("| P2 | 建议修 | 记录理由后可放行 |")
    lines.append("| P3 | 提示 | 复核时可忽略 |")
    lines.append("")
    lines.append("## 人工判定项（脚本不实现）")
    lines.append("")
    lines.append("1. 句子读起来像不像人写的（十二特征逐条数，见 copywriting-antislop §2）。")
    lines.append("2. 每页是否只有一个主张，听众能否在 3 秒内抓住。")
    lines.append("3. 事实是否成立：数字口径、比较对象、样本量。")
    lines.append("4. 与用户原话、原意是否一致（尤其结论方向有没有被改软）。")
    lines.append("")
    return "\n".join(lines) + "\n"


def _escape(text: str) -> str:
    return str(text).replace("|", "\\|").replace("\n", " ")


def build_json(
    issues: List[Issue],
    summary: Dict[str, int],
    deck_path: str,
    thresholds_version: str,
    pages: List[PageData],
    exit_code: int,
) -> Dict[str, Any]:
    return {
        "schemaVersion": COPY_SCHEMA_VERSION,
        "checkedAt": datetime.now().isoformat(timespec="seconds"),
        "deck": deck_path,
        "thresholdsVersion": thresholds_version,
        "pages": len(pages),
        "summary": summary,
        "issues": [issue.as_dict() for issue in issues],
        "exitCode": exit_code,
        "manualChecks": [
            "句子像不像人写的",
            "每页是否只有一个主张",
            "事实口径是否成立",
            "与用户原意是否一致",
        ],
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="文案机检：槽位字数、空话、AI 句式、数字出处")
    parser.add_argument("--deck", required=True, help="deck.json 路径")
    parser.add_argument("--copy", default="", help="records/copy-deck.md（检查未确认行）")
    parser.add_argument("--out", default="", help="输出 Markdown 报告路径")
    parser.add_argument("--json", dest="json_path", default="", help="输出 JSON 报告路径")
    parser.add_argument("--thresholds", default="", help="自定义 thresholds.json")
    parser.add_argument("--lenient", action="store_true", help="仅在存在 P0 时返回非零")
    parser.add_argument("--quiet", action="store_true", help="不打印问题清单，只打印 SUMMARY")
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)
    thresholds = load_thresholds(args.thresholds)
    deck_path = Path(args.deck).expanduser()

    if not deck_path.is_file():
        print("错误：找不到 deck 文件：%s" % deck_path, file=sys.stderr)
        return 2

    try:
        deck = deckload.load(str(deck_path))
    except SystemExit:
        raise
    except Exception as exc:  # noqa: BLE001 - deck 校验失败要保持可读信息
        print("错误：deck 校验失败 —— %s" % exc, file=sys.stderr)
        return 2

    pages = collect_pages(deck)
    issues: List[Issue] = []
    issues += check_slot_length(pages, thresholds)
    issues += check_empty_phrases(pages, thresholds)
    issues += check_filler_adjectives(pages, thresholds)
    issues += check_ai_patterns(pages, thresholds)
    issues += check_long_clauses(pages, thresholds)
    issues += check_paragraph_uniform(pages, thresholds)
    issues += check_digits_without_source(pages, thresholds)
    issues += check_title_claims(pages, thresholds)
    issues += check_repeated_phrase(pages, thresholds)
    issues += check_unreviewed_markers(pages, thresholds)
    if args.copy:
        issues += check_copy_md(Path(args.copy).expanduser(), thresholds)

    issues = sort_issues(issues)
    summary = summarize(issues)
    blocking = summary.get("P0", 0) > 0 or (not args.lenient and summary.get("P1", 0) > 0)
    exit_code = 1 if blocking else 0

    try:
        displayed_deck = str(deck_path.resolve().relative_to(REPO_ROOT))
    except ValueError:
        displayed_deck = str(deck_path)

    thresholds_version = str(thresholds.get("version", "unknown"))
    markdown = render_markdown(issues, summary, displayed_deck, thresholds_version, pages, exit_code)
    payload = build_json(issues, summary, displayed_deck, thresholds_version, pages, exit_code)

    if args.out:
        out_path = Path(args.out).expanduser()
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(markdown, encoding="utf-8", newline="\n")
        print("报告：%s" % out_path)
    if args.json_path:
        json_path = Path(args.json_path).expanduser()
        json_path.parent.mkdir(parents=True, exist_ok=True)
        json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")
        print("JSON：%s" % json_path)

    print(
        "SUMMARY P0=%d P1=%d P2=%d P3=%d issues=%d pages=%d exit=%d"
        % (
            summary.get("P0", 0),
            summary.get("P1", 0),
            summary.get("P2", 0),
            summary.get("P3", 0),
            len(issues),
            len(pages),
            exit_code,
        )
    )

    if not args.quiet and issues:
        print("")
        print("问题清单（级别  检查项  页  问题）")
        for issue in issues:
            print("  %s  %s  %s  %s" % (issue.level, issue.check, issue.page, issue.message))

    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
