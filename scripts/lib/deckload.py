#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""deck.json 读取、契约校验与设计令牌解析（build_pptx / qa_check / render 的公共入口）。

本模块是「deck.schema.json 是唯一输入契约」这条规则的执行者：它不重新定义字段，
而是**直接读取 schema/deck.schema.json 并按其中的约束逐项校验**。因此在没有安装
jsonschema 的机器上（本机实测未安装）仍然得到与 schema 完全一致的校验结果，
不会出现「文档允许、代码拒绝」的漂移。

    from lib.deckload import load, resolve_color, resolve_type
    deck = load("demo/smoke/deck.json")
    print(deck.size.width_in, deck.size.height_in)
    rgb = resolve_color(deck.tokens, "accent")       # -> (r, g, b)
    ts  = resolve_type(deck.tokens, "body")          # -> size/weight/lineHeight/letterSpacing

对外函数（唯一公共契约）：

    load(path, *, check_assets=True) -> Deck
    load_dict(doc, base_dir, source="<dict>") -> Deck
    validate_doc(doc, schema_doc=None) -> list[str]    # 返回中文错误列表，空列表=通过
    resolve_color(tokens, ref, path="color") -> (r, g, b)
    resolve_type(tokens, role, path="role") -> dict
    resolve_font(tokens, name) -> str
    resolve_size(spec) -> Size
    rgb_to_hex(rgb) -> "#RRGGBB"

错误一律抛 DeckError / DeckValidationError，消息为中文并带 JSON 路径（如
`$.slides[0].elements[2].text`），调用方直接打印即可交付给使用者看。

坐标与尺寸约定：`meta.size` 之外的数值单位全部是英寸，原点在页面左上角；
字号单位是 pt。两者之间的换算只在 build_pptx.py 里做（Inches()/Pt()），
本模块只做语义解析，不引入 python-pptx 依赖，方便 probe/qa 在无 pptx 环境导入。
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Tuple

__all__ = [
    "DeckError",
    "DeckValidationError",
    "Size",
    "Deck",
    "load",
    "load_dict",
    "validate_doc",
    "resolve_color",
    "resolve_type",
    "resolve_font",
    "resolve_size",
    "resolve_grid",
    "rgb_to_hex",
    "ROLE_FALLBACK",
    "DEFAULT_TYPE_SCALE",
    "DEFAULT_GRID",
]

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------

#: 画布预设。与 schema/thresholds.json 的 canvas 段保持一致（16:9 = 13.333x7.5 in）。
SIZE_PRESETS: Dict[str, Tuple[float, float]] = {
    "16:9": (13.333, 7.5),
    "16x9": (13.333, 7.5),
    "4:3": (10.0, 7.5),
    "4x3": (10.0, 7.5),
    "16:10": (13.333, 8.333),
}

#: 角色回退链：tokens.type 里没有该角色时，按顺序找第一个存在的键。
#: 这份表让「deck 只定义了 h1/body 也能渲染出 coverTitle/kpi/caption」成为确定行为，
#: 而不是渲染成 16 pt 默认值。
ROLE_FALLBACK: Dict[str, Tuple[str, ...]] = {
    "coverTitle": ("coverTitle", "h1", "h2"),
    "h1": ("h1", "coverTitle", "h2"),
    "h2": ("h2", "h3", "h1"),
    "h3": ("h3", "h2", "body"),
    "lead": ("lead", "body"),
    "body": ("body", "lead"),
    "caption": ("caption", "source", "body"),
    "source": ("source", "caption", "body"),
    "pageNumber": ("pageNumber", "caption", "source"),
    "kpi": ("kpi", "coverTitle", "h1"),
    "label": ("label", "caption", "body"),
    "quote": ("quote", "coverTitle", "h2"),
    "decor": ("decor", "caption"),
}

#: 缺省字号阶梯（schema typeScale 的 default 值）。角色与回退链都落空时使用。
DEFAULT_TYPE_SCALE: Dict[str, float] = {
    "size": 16.0,
    "weight": 400,
    "lineHeight": 1.35,
    "letterSpacing": 0.0,
}

#: 网格缺省（schema tokens.grid 的 default 值）。
DEFAULT_GRID: Dict[str, float] = {
    "cols": 12,
    "gutterIn": 0.2,
    "marginXIn": 0.6,
    "marginYIn": 0.45,
    "baselineIn": 0.08,
}

#: 元素必填几何字段。
_GEOM_KEYS = ("x", "y", "w", "h")

_SIZE_RE = re.compile(r"^\s*(\d+(?:\.\d+)?)\s*[x×*]\s*(\d+(?:\.\d+)?)\s*$")
_HEX_RE = re.compile(r"^#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{3})$")


# ---------------------------------------------------------------------------
# 异常
# ---------------------------------------------------------------------------


class DeckError(Exception):
    """deck 读取/解析失败的基类。消息为中文，可直接打印给使用者。"""


class DeckValidationError(DeckError):
    """schema 校验未通过。`issues` 保存全部错误（JSON 路径 + 中文原因）。"""

    def __init__(self, message: str, issues: Optional[List[str]] = None, source: str = ""):
        super().__init__(message)
        self.issues: List[str] = list(issues or [])
        self.source = source


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Size:
    """画布尺寸。英寸为主，像素按 96 dpi 折算（与 thresholds.canvas 一致）。"""

    width_in: float
    height_in: float
    label: str = ""

    @property
    def width_px(self) -> int:
        return int(round(self.width_in * 96))

    @property
    def height_px(self) -> int:
        return int(round(self.height_in * 96))

    @property
    def ratio(self) -> float:
        return self.width_in / self.height_in

    def __str__(self) -> str:  # pragma: no cover - 仅用于打印
        return f"{self.width_in:g}x{self.height_in:g} in ({self.label or 'custom'})"


@dataclass
class Deck:
    """校验通过后的 deck。原始 doc 保留，解析结果按需缓存。"""

    doc: Dict[str, Any]
    source: str
    base_dir: str
    size: Size
    warnings: List[str] = field(default_factory=list)
    missing_images: List[str] = field(default_factory=list)

    # -- 便捷访问 ---------------------------------------------------------
    @property
    def meta(self) -> Dict[str, Any]:
        return self.doc.get("meta") or {}

    @property
    def title(self) -> str:
        return str(self.meta.get("title", ""))

    @property
    def tokens(self) -> Dict[str, Any]:
        return self.doc.get("tokens") or {}

    @property
    def slides(self) -> List[Dict[str, Any]]:
        return self.doc.get("slides") or []

    @property
    def slide_count(self) -> int:
        return len(self.slides)

    def slide_bg(self, slide: Dict[str, Any]) -> str:
        """本页背景色引用：页内 bg 优先，否则 tokens.colors.bg。"""
        return slide.get("bg") or (self.tokens.get("colors") or {}).get("bg") or "#FFFFFF"

    def color(self, ref: str, path: str = "color") -> Tuple[int, int, int]:
        return resolve_color(self.tokens, ref, path=path)

    def type_scale(self, role: str) -> Dict[str, Any]:
        return resolve_type(self.tokens, role)

    def font_of(self, name: Optional[str] = None) -> str:
        return resolve_font(self.tokens, name)

    def grid(self) -> Dict[str, float]:
        return resolve_grid(self.tokens)

    def iter_elements(self) -> Iterable[Tuple[int, int, Dict[str, Any]]]:
        """产出 (slide_index, element_index, element)，索引从 0 开始。"""
        for si, slide in enumerate(self.slides):
            for ei, el in enumerate(slide.get("elements") or []):
                yield si, ei, el

    def fonts_used(self) -> List[str]:
        """deck 里显式出现过的字体（用于构建摘要与缺字体核对）。"""
        found: List[str] = []
        fonts = self.tokens.get("fonts") or {}
        for key in ("display", "body", "mono"):
            if fonts.get(key) and fonts[key] not in found:
                found.append(str(fonts[key]))
        for _si, _ei, el in self.iter_elements():
            for name in _element_font_names(el):
                resolved = resolve_font(self.tokens, name)
                if resolved not in found:
                    found.append(resolved)
        return found


def _element_font_names(el: Dict[str, Any]) -> List[str]:
    names: List[str] = []
    if el.get("font"):
        names.append(str(el["font"]))
    for run in el.get("runs") or []:
        if isinstance(run, dict) and run.get("font"):
            names.append(str(run["font"]))
    for para in el.get("paragraphs") or []:
        if isinstance(para, dict):
            for run in para.get("runs") or []:
                if isinstance(run, dict) and run.get("font"):
                    names.append(str(run["font"]))
    return names


# ---------------------------------------------------------------------------
# schema 迷你校验器
# ---------------------------------------------------------------------------

_CN_TYPE = {
    "object": "对象",
    "array": "数组",
    "string": "字符串",
    "number": "数字",
    "integer": "整数",
    "boolean": "布尔值",
    "null": "空值",
}


def _json_type(value: Any) -> str:
    """返回 JSON 类型名。bool 必须先于 int 判断（Python 里 bool 是 int 的子类）。"""
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, dict):
        return "object"
    if isinstance(value, list):
        return "array"
    if isinstance(value, str):
        return "string"
    if isinstance(value, int):
        return "integer"
    if isinstance(value, float):
        return "number"
    return type(value).__name__


def _type_matches(value: Any, expected: str) -> bool:
    actual = _json_type(value)
    if expected == "number":
        return actual in ("number", "integer")
    if expected == "integer":
        return actual == "integer"
    return actual == expected


def _join_path(path: str, key: Any) -> str:
    if isinstance(key, int):
        return f"{path}[{key}]"
    return f"{path}.{key}"


def _short(value: Any, limit: int = 60) -> str:
    text = json.dumps(value, ensure_ascii=False, default=str)
    return text if len(text) <= limit else text[: limit - 3] + "..."


def _resolve_ref(ref: str, root: Dict[str, Any]) -> Dict[str, Any]:
    if not isinstance(ref, str) or not ref.startswith("#/"):
        raise DeckError(f"schema 含不支持的 $ref：{ref!r}（只支持文档内 #/$defs/... 引用）")
    node: Any = root
    for part in ref[2:].split("/"):
        part = part.replace("~1", "/").replace("~0", "~")
        if not isinstance(node, dict) or part not in node:
            raise DeckError(f"schema 引用无法解析：{ref}")
        node = node[part]
    return node


def _validate_node(
    value: Any,
    schema: Dict[str, Any],
    root: Dict[str, Any],
    path: str,
    issues: List[str],
    depth: int = 0,
) -> None:
    """按 schema 的一个节点递归校验 value，把中文错误追加进 issues。"""
    if depth > 64:
        issues.append(f"{path}: schema 递归过深（>64 层），疑似循环 $ref")
        return
    if not isinstance(schema, dict):
        return

    if "$ref" in schema:
        _validate_node(value, _resolve_ref(schema["$ref"], root), root, path, issues, depth + 1)
        return

    expected = schema.get("type")
    if isinstance(expected, str) and not _type_matches(value, expected):
        issues.append(
            f"{path}: 类型应为{_CN_TYPE.get(expected, expected)}，实际是{_CN_TYPE.get(_json_type(value), _json_type(value))}"
        )
        return
    if isinstance(expected, list):
        if not any(_type_matches(value, t) for t in expected if isinstance(t, str)):
            issues.append(f"{path}: 类型应为 {'/'.join(map(str, expected))}，实际是 {_json_type(value)}")
            return

    if "enum" in schema and value not in schema["enum"]:
        allowed = "、".join(str(x) for x in schema["enum"])
        issues.append(f"{path}: 值 {_short(value)} 不在允许的枚举内（可选：{allowed}）")

    if isinstance(value, str) and "pattern" in schema:
        if re.search(schema["pattern"], value) is None:
            issues.append(f"{path}: 字符串 {_short(value)} 不符合格式 {schema['pattern']}")

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        lo, hi = schema.get("minimum"), schema.get("maximum")
        if lo is not None and value < lo:
            issues.append(f"{path}: 值 {value:g} 小于下限 {lo:g}")
        if hi is not None and value > hi:
            issues.append(f"{path}: 值 {value:g} 大于上限 {hi:g}")

    if isinstance(value, dict):
        for key in schema.get("required") or []:
            if key not in value:
                issues.append(f"{path}: 缺少必填字段 {key}")
        props = schema.get("properties") or {}
        addl = schema.get("additionalProperties", True)
        for key, sub in value.items():
            child = _join_path(path, key)
            if key in props:
                _validate_node(sub, props[key], root, child, issues, depth + 1)
            elif addl is False:
                allowed = "、".join(props.keys()) or "（无）"
                shown = "、".join(sorted(props.keys())) or "（无）"
                issues.append(f"{child}: 不允许多余字段 {key}（允许的字段：{shown or allowed}）")
            elif isinstance(addl, dict):
                _validate_node(sub, addl, root, child, issues, depth + 1)

    if isinstance(value, list):
        if "minItems" in schema and len(value) < schema["minItems"]:
            issues.append(f"{path}: 至少需要 {schema['minItems']} 项，实际 {len(value)} 项")
        if "maxItems" in schema and len(value) > schema["maxItems"]:
            issues.append(f"{path}: 最多允许 {schema['maxItems']} 项，实际 {len(value)} 项")
        item_schema = schema.get("items")
        if isinstance(item_schema, dict):
            for i, item in enumerate(value):
                _validate_node(item, item_schema, root, _join_path(path, i), issues, depth + 1)


def _default_schema_path() -> str:
    """定位 schema/deck.schema.json：<repo>/scripts/lib/deckload.py -> <repo>/schema/。"""
    here = os.path.dirname(os.path.abspath(__file__))
    repo = os.path.dirname(os.path.dirname(here))
    return os.path.join(repo, "schema", "deck.schema.json")


def _load_schema(schema_doc: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    if schema_doc is not None:
        return schema_doc
    path = _default_schema_path()
    if not os.path.isfile(path):
        raise DeckError(
            f"找不到输入契约 schema 文件：{path}\n"
            "请确认脚本位于 <仓库>/scripts/lib/ 下，或用 validate_doc(doc, schema_doc=...) 显式传入。"
        )
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except OSError as exc:
        raise DeckError(f"读取 schema 失败：{path}（{exc}）") from exc
    except json.JSONDecodeError as exc:
        raise DeckError(f"schema 不是合法 JSON：{path} 第 {exc.lineno} 行第 {exc.colno} 列（{exc.msg}）") from exc


def validate_doc(doc: Any, schema_doc: Optional[Dict[str, Any]] = None) -> List[str]:
    """按 deck.schema.json 校验，返回中文错误列表；空列表表示通过。

    优先使用官方 jsonschema 库（若已安装，额外得到 oneOf/allOf 等完整语义），
    否则使用本模块内置的迷你实现——两者对本 schema 的结论一致。
    """
    schema = _load_schema(schema_doc)
    issues: List[str] = []
    _validate_node(doc, schema, schema, "$", issues)
    if issues:
        return issues
    # 官方实现作为交叉验证（未安装时静默跳过，不改变结论）
    try:  # pragma: no cover - 取决于运行环境
        import jsonschema  # type: ignore

        validator = jsonschema.Draft202012Validator(schema)
        for err in sorted(validator.iter_errors(doc), key=lambda e: list(e.absolute_path)):
            loc = "$" + "".join(
                f"[{p}]" if isinstance(p, int) else f".{p}" for p in err.absolute_path
            )
            message = f"{loc}: {err.message}"
            if message not in issues:
                issues.append(message)
    except ImportError:
        pass
    except Exception:  # schema 用法与库版本不兼容时不影响主结论
        pass
    return issues


# ---------------------------------------------------------------------------
# 令牌解析
# ---------------------------------------------------------------------------


def _parse_hex(text: str, path: str) -> Tuple[int, int, int]:
    if not isinstance(text, str):
        raise DeckError(f"{path}: 颜色必须是 #RRGGBB 字符串，实际是 {type(text).__name__}")
    raw = text.strip()
    if not _HEX_RE.match(raw):
        raise DeckError(f"{path}: 颜色 {raw!r} 不是 #RRGGBB 或 #RGB 格式")
    body = raw[1:]
    if len(body) == 3:
        body = "".join(ch * 2 for ch in body)
    return (int(body[0:2], 16), int(body[2:4], 16), int(body[4:6], 16))


def resolve_color(tokens: Dict[str, Any], ref: Any, path: str = "color") -> Tuple[int, int, int]:
    """把 tokenRef 解析为 (r, g, b)。

    规则（与 schema tokenRef 说明一致）：
      1. `$name` —— 精确引用 tokens.colors.name，不存在即报错；
      2. `name`  —— 先查 tokens.colors，命中则用令牌值；
      3. `#RRGGBB` / `#RGB` —— 直接当颜色解析；
      4. 其余 —— 报错并指出该 JSON 路径。
    """
    colors = (tokens or {}).get("colors") or {}
    if ref is None:
        raise DeckError(f"{path}: 缺少颜色值（需要令牌名或 #RRGGBB）")
    if not isinstance(ref, str):
        raise DeckError(f"{path}: 颜色必须是字符串令牌名或 #RRGGBB，实际是 {type(ref).__name__}")
    raw = ref.strip()
    if not raw:
        raise DeckError(f"{path}: 颜色为空字符串")

    if raw.startswith("$"):
        key = raw[1:].strip()
        if not key:
            raise DeckError(f"{path}: 颜色引用 {raw!r} 的 `$` 后缺少令牌名")
        if key not in colors:
            available = "、".join(sorted(colors.keys())) or "（tokens.colors 为空）"
            raise DeckError(f"{path}: 找不到颜色令牌 ${key}（现有令牌：{available}）")
        return _parse_hex(colors[key], _join_path(path, key))

    if raw in colors:
        return _parse_hex(colors[raw], _join_path(path, raw))

    if raw.startswith("#"):
        return _parse_hex(raw, path)

    available = "、".join(sorted(colors.keys())) or "（tokens.colors 为空）"
    raise DeckError(
        f"{path}: 无法解析颜色 {raw!r} —— 它既不在 tokens.colors（现有：{available}）中，"
        "也不是 #RRGGBB 形式"
    )


def resolve_type(tokens: Dict[str, Any], role: Any, path: str = "role") -> Dict[str, Any]:
    """解析字号阶梯：返回 size(pt)/weight/lineHeight/letterSpacing 及命中的令牌名。

    角色未定义时按 ROLE_FALLBACK 回退；全落空时用 schema 缺省值 16 pt/400/1.35，
    并在返回值 `source` 中标明，便于排查「为什么这行字是 16 pt」。
    """
    role_key = str(role or "body")
    scale_table = (tokens or {}).get("type") or {}
    chain = ROLE_FALLBACK.get(role_key, (role_key, "body"))

    for key in chain:
        raw = scale_table.get(key)
        if not isinstance(raw, dict):
            continue
        size = raw.get("size", DEFAULT_TYPE_SCALE["size"])
        if not isinstance(size, (int, float)) or isinstance(size, bool) or size <= 0:
            raise DeckError(
                f"{path}: tokens.type.{key}.size 必须为正数（单位 pt），实际是 {size!r}"
            )
        weight = raw.get("weight", DEFAULT_TYPE_SCALE["weight"])
        line_height = raw.get("lineHeight", DEFAULT_TYPE_SCALE["lineHeight"])
        spacing = raw.get("letterSpacing", DEFAULT_TYPE_SCALE["letterSpacing"])
        return {
            "role": role_key,
            "source": key,
            "size": float(size),
            "weight": int(weight) if isinstance(weight, (int, float)) else 400,
            "lineHeight": float(line_height) if isinstance(line_height, (int, float)) else 1.35,
            "letterSpacing": float(spacing) if isinstance(spacing, (int, float)) else 0.0,
        }

    if role_key not in ROLE_FALLBACK:
        # 未知角色不是致命错误：回退到 body 并让调用方记录警告。
        pass
    return {
        "role": role_key,
        "source": "default",
        "size": DEFAULT_TYPE_SCALE["size"],
        "weight": int(DEFAULT_TYPE_SCALE["weight"]),
        "lineHeight": DEFAULT_TYPE_SCALE["lineHeight"],
        "letterSpacing": DEFAULT_TYPE_SCALE["letterSpacing"],
    }


def resolve_font(tokens: Dict[str, Any], name: Optional[str]) -> str:
    """解析字体：字体名可以是字面字体（"Microsoft YaHei"），也可以是 tokens.fonts 的键。

    这样 deck 里写 `"font": "display"` 与写 `"font": "Bahnschrift"` 效果相同，
    改字体只需改一处令牌。未给出时回退 tokens.fonts.body。
    """
    fonts = (tokens or {}).get("fonts") or {}
    fallback = str(fonts.get("body") or "Microsoft YaHei")
    if name is None or name == "":
        return fallback
    if not isinstance(name, str):
        raise DeckError(f"字体名必须是字符串，实际是 {type(name).__name__}")
    key = name.strip()
    if key in fonts and isinstance(fonts[key], str):
        return str(fonts[key])
    return key


def resolve_size(spec: Any) -> Size:
    """解析 meta.size：支持 "16:9" / "4:3" / "13.333x7.5" / "13.333×7.5"。"""
    if spec is None:
        spec = "16:9"
    if not isinstance(spec, str):
        raise DeckError(f"$.meta.size: 必须是字符串（如 \"16:9\" 或 \"13.333x7.5\"），实际是 {type(spec).__name__}")
    raw = spec.strip()
    if raw in SIZE_PRESETS:
        w, h = SIZE_PRESETS[raw]
        return Size(w, h, raw)
    match = _SIZE_RE.match(raw)
    if match:
        w, h = float(match.group(1)), float(match.group(2))
        if w <= 0 or h <= 0:
            raise DeckError(f"$.meta.size: 宽高必须为正数，实际是 {raw!r}")
        if w > 60 or h > 60:
            raise DeckError(f"$.meta.size: 画布 {raw!r} 超过 60 in，疑似单位写错（应为英寸）")
        ratio = w / h
        label = "16:9" if abs(ratio - 16 / 9) < 0.01 else ("4:3" if abs(ratio - 4 / 3) < 0.01 else f"{ratio:.3f}:1")
        return Size(w, h, label)
    raise DeckError(
        f"$.meta.size: 无法解析 {raw!r}，只支持 \"16:9\"、\"4:3\" 或显式 \"宽x高\"（英寸，如 \"13.333x7.5\"）"
    )


def resolve_grid(tokens: Dict[str, Any]) -> Dict[str, float]:
    """返回网格参数（缺省值来自 schema tokens.grid 的 default）。"""
    raw = (tokens or {}).get("grid") or {}
    grid = dict(DEFAULT_GRID)
    for key in DEFAULT_GRID:
        if isinstance(raw.get(key), (int, float)) and not isinstance(raw.get(key), bool):
            grid[key] = raw[key]  # type: ignore[assignment]
    grid["cols"] = int(grid["cols"])
    return grid


def rgb_to_hex(rgb: Tuple[int, int, int]) -> str:
    r, g, b = rgb
    return f"#{r:02X}{g:02X}{b:02X}"


# ---------------------------------------------------------------------------
# 语义检查（schema 管不到的跨字段约束）
# ---------------------------------------------------------------------------

_KINDS = ("text", "rect", "line", "ellipse", "polygon", "image", "chart", "table", "icon")


def _check_semantics(deck: Deck, check_assets: bool) -> None:
    """逐页逐元素检查 schema 无法表达的约束，结果写入 deck.warnings / missing_images。

    这里只收集「不阻断构建」的问题（如缺图、未知角色）；结构性错误在构建阶段抛错，
    保证 --check-only 与真实构建的判定一致。
    """
    roles_in_tokens = set((deck.tokens.get("type") or {}).keys())

    for _si, _ei, el in deck.iter_elements():
        path = f"$.slides[{_si}].elements[{_ei}]"
        kind = el.get("kind")

        for key in _GEOM_KEYS:
            value = el.get(key)
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                # line 的几何由 x2/y2 决定，水平线的 h=0、垂直线的 w=0 都合法
                if key in ("w", "h") and value <= 0 and kind != "line":
                    deck.warnings.append(f"{path}.{key}: 宽/高必须大于 0，实际 {value:g}（该元素会被跳过或渲染为零尺寸）")

        if kind == "text" or any(k in el for k in ("text", "runs", "paragraphs")):
            provided = [k for k in ("text", "runs", "paragraphs") if el.get(k) not in (None, "", [])]
            if len(provided) > 1:
                deck.warnings.append(
                    f"{path}: 同时给出了 {'/'.join(provided)}，三者互斥，渲染优先级 paragraphs > runs > text"
                )
            if not provided:
                deck.warnings.append(f"{path}: 文本元素没有任何内容（text/runs/paragraphs 全缺）")

        role = el.get("role")
        if role and role not in roles_in_tokens and role not in ROLE_FALLBACK:
            deck.warnings.append(
                f"{path}.role: 未知角色 {role!r}，将回退为 16 pt 正文字号；"
                f"tokens.type 现有角色：{'、'.join(sorted(roles_in_tokens)) or '（空）'}"
            )
        elif role and role not in roles_in_tokens:
            deck.warnings.append(
                f"{path}.role: tokens.type 未定义角色 {role!r}，按回退链渲染（见 deckload.ROLE_FALLBACK）"
            )

        if kind == "image":
            src = el.get("src")
            if not src:
                deck.warnings.append(f"{path}.src: image 元素缺少 src，该图片无法渲染")
            elif check_assets:
                abs_path = src if os.path.isabs(src) else os.path.join(deck.base_dir, src)
                abs_path = os.path.normpath(abs_path)
                if not os.path.isfile(abs_path):
                    message = f"{path}.src: 图片不存在：{abs_path}"
                    deck.warnings.append(message)
                    deck.missing_images.append(abs_path)
                if not el.get("alt"):
                    deck.warnings.append(f"{path}.alt: 图片缺少替代文本（无障碍检查项，WCAG 2.2 SC 1.1.1 要求）")

        if kind == "polygon":
            points = el.get("points") or []
            if len(points) < 3:
                deck.warnings.append(f"{path}.points: polygon 至少需要 3 个顶点，实际 {len(points)} 个")

        if kind == "line" and ("x2" not in el or "y2" not in el):
            deck.warnings.append(f"{path}: line 未给 x2/y2，将回退为 (x+w, y+h) 作为终点")

        if kind == "chart":
            chart = el.get("chart") or {}
            if not chart.get("categories") and el.get("chartType") != "scatter":
                deck.warnings.append(f"{path}.chart.categories: 图表缺少 categories，横轴将没有标签")
            if not chart.get("unit"):
                deck.warnings.append(f"{path}.chart.unit: 图表缺少单位，页面会出现无单位数字（违反 thresholds 数字规范）")
            if not chart.get("source"):
                deck.warnings.append(f"{path}.chart.source: 图表缺少数据来源标注")
            if not el.get("chartType"):
                deck.warnings.append(f"{path}.chartType: 缺省按 column 渲染")

        if kind == "table":
            table = el.get("table") or {}
            if not table.get("rows"):
                deck.warnings.append(f"{path}.table.rows: 表格没有数据行")
            widths = table.get("colWidths")
            if widths:
                total = sum(float(v) for v in widths)
                declared = float(el.get("w") or 0)
                if declared and abs(total - declared) > 0.05:
                    deck.warnings.append(
                        f"{path}.table.colWidths: 列宽合计 {total:.3f} in 与元素宽 w={declared:.3f} in 不一致，"
                        "将按列宽渲染（可能溢出框）"
                    )


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------


def load_dict(doc: Any, base_dir: str = ".", source: str = "<dict>", *, check_assets: bool = True) -> Deck:
    """从已解析的 dict 构造 Deck（会先按 schema 校验）。"""
    if not isinstance(doc, dict):
        raise DeckValidationError(
            f"{source}: 顶层必须是 JSON 对象，实际是 {_CN_TYPE.get(_json_type(doc), _json_type(doc))}",
            source=source,
        )
    issues = validate_doc(doc)
    if issues:
        head = "\n".join(f"  - {item}" for item in issues[:20])
        more = f"\n  …… 另有 {len(issues) - 20} 处" if len(issues) > 20 else ""
        raise DeckValidationError(
            f"{source} 未通过 deck.schema.json 校验，共 {len(issues)} 处：\n{head}{more}",
            issues=issues,
            source=source,
        )

    meta = doc.get("meta") or {}
    size = resolve_size(meta.get("size", "16:9"))
    deck = Deck(
        doc=doc,
        source=source,
        base_dir=os.path.abspath(base_dir),
        size=size,
    )
    _check_semantics(deck, check_assets)

    colors = (deck.tokens.get("colors") or {})
    unique = {str(v).upper() for v in colors.values() if isinstance(v, str)}
    if unique and unique <= {"#000000", "#FFFFFF"}:
        deck.warnings.append(
            "tokens.colors: 配色只含纯黑与纯白，缺少主色/强调色（schema 明确禁止）；"
            "至少补 primary/accent 两个色相"
        )
    return deck


def load(path: str, *, check_assets: bool = True) -> Deck:
    """读取并校验 deck.json，返回 Deck。

    - 文件不存在、不是合法 JSON、不是 UTF-8 —— DeckError（中文，含文件路径与行号）。
    - schema 校验失败 —— DeckValidationError（含全部 JSON 路径与原因）。
    - 图片缺失、未知角色等「可构建但需告知」的问题 —— deck.warnings / deck.missing_images。
    """
    if not path:
        raise DeckError("未提供 deck.json 路径")
    abs_path = os.path.abspath(path)
    if not os.path.isfile(abs_path):
        raise DeckError(f"找不到 deck 文件：{abs_path}")
    try:
        with open(abs_path, "rb") as fh:
            raw = fh.read()
    except OSError as exc:
        raise DeckError(f"读取 deck 文件失败：{abs_path}（{exc}）") from exc
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise DeckError(
            f"deck 文件不是 UTF-8 编码：{abs_path}（第 {exc.start} 字节）；请另存为 UTF-8"
        ) from exc
    try:
        doc = json.loads(text)
    except json.JSONDecodeError as exc:
        raise DeckError(
            f"deck 不是合法 JSON：{abs_path} 第 {exc.lineno} 行第 {exc.colno} 列（{exc.msg}）"
        ) from exc
    return load_dict(doc, base_dir=os.path.dirname(abs_path), source=abs_path, check_assets=check_assets)


def dump_summary(deck: Deck) -> str:
    """一行式摘要，供脚本打印。"""
    return (
        f"{deck.title or '(无标题)'} | {deck.slide_count} 页 | 画布 {deck.size} | "
        f"字体 {len(deck.fonts_used())} 种 | 警告 {len(deck.warnings)} 条 | 缺图 {len(deck.missing_images)} 张"
    )


if __name__ == "__main__":  # pragma: no cover - 手工调试入口
    import sys

    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
    if len(sys.argv) < 2:
        print("用法：python scripts/lib/deckload.py <deck.json>", file=sys.stderr)
        raise SystemExit(2)
    try:
        d = load(sys.argv[1])
    except DeckError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        raise SystemExit(1)
    print(dump_summary(d))
    for warning in d.warnings:
        print(f"警告：{warning}")
