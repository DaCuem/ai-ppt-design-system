#!/usr/bin/env python3
"""把 .pptx 渲染成 PNG 预览图，按优先级自动探测可用的渲染后端。

为什么需要这一步：pptx 是二进制包，模型看不到画面。只有渲染出 PNG，才能用视觉
能力检查版式是否错位、文字是否溢出、图表是否可读。于是「同一份 deck.json 同时
决定 PPTX 与预览图」这条契约才可被验证。

后端优先级（逐个探测，第一个可用的胜出）：
  1. PPT_RENDER_CMD   自定义命令模板，含 {input} {outdir} {pages} {dpi} 占位符
  2. dsh-native       本机 DSH 自带渲染器（node + office-cli.mjs，LibreOffice 光栅化）
  3. soffice          LibreOffice 先转 PDF，再 pdftoppm / Ghostscript 转 PNG
  4. powerpoint-com   Windows PowerPoint COM（需要 pywin32 + 已安装 PowerPoint）

退出码：0 成功 / 1 输入或参数有问题 / 2 所有渲染后端都不可用或全部失败。

用法：
    python scripts/render.py --pptx out/deck.pptx --out-dir preview
    python scripts/render.py --pptx out/deck.pptx --out-dir preview --pages "1,3" --dpi 96
    python scripts/render.py --pptx out/deck.pptx --out-dir preview --pages all --json

输出：
    <out-dir>/slide-01.png       每页一张，文件名编号 = 源页面号（1 起，左补零）
    <out-dir>/render-manifest.json  页数、尺寸、DPI、缺失字体、实际使用的后端
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional, Sequence, Tuple

DEFAULT_DPI = 144
MIN_DPI = 50
MAX_DPI = 600
TIMEOUT_S = 300
MANIFEST_NAME = "render-manifest.json"
MANIFEST_SCHEMA_VERSION = 1

BACKEND_PPT_CMD = "ppt-render-cmd"
BACKEND_DSH = "dsh-native"
BACKEND_SOFFICE = "soffice"
BACKEND_COM = "powerpoint-com"


class RenderError(Exception):
    """渲染失败（单个后端内部错误，可继续尝试下一个后端）。"""


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------


def _configure_stdio() -> None:
    """Windows 控制台默认不是 UTF-8，中文摘要会变成乱码，这里强制改掉。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[union-attr]
        except (AttributeError, ValueError):
            pass


def _first_line(text: str, limit: int = 400) -> str:
    """取输出的第一行有效内容，用于把子进程错误压缩成一句可读信息。"""
    for line in (text or "").splitlines():
        line = line.strip()
        if line:
            return line[:limit]
    return "（无输出）"


def _run(cmd: Sequence[str], *, timeout: int = TIMEOUT_S,
         env: Optional[Dict[str, str]] = None) -> subprocess.CompletedProcess:
    """执行子进程并捕获输出。

    Windows 上 node 往往是 .cmd 转发脚本，直接执行可能报 WinError 193，
    此时改用 cmd.exe /c 重试一次——这是本机实测踩到的坑。
    """
    kwargs = dict(capture_output=True, text=True, encoding="utf-8", errors="replace",
                  timeout=timeout, env=env)
    try:
        return subprocess.run(list(cmd), **kwargs)
    except FileNotFoundError:
        raise RenderError(f"找不到可执行文件：{cmd[0]}") from None
    except OSError as exc:
        if os.name == "nt" and not str(cmd[0]).lower().endswith((".exe",)):
            try:
                return subprocess.run(["cmd.exe", "/c", *cmd], **kwargs)
            except OSError:
                pass
        raise RenderError(f"无法执行 {cmd[0]}：{exc}") from None
    except subprocess.TimeoutExpired:
        raise RenderError(f"{os.path.basename(str(cmd[0]))} 超过 {timeout} 秒未结束，已放弃") from None


def _extract_json(text: str) -> Optional[dict]:
    """从子进程 stdout 里捞出 JSON 对象（渲染器可能先打印日志再打印 JSON）。"""
    text = (text or "").strip()
    if not text:
        return None
    try:
        data = json.loads(text)
        return data if isinstance(data, dict) else None
    except json.JSONDecodeError:
        pass
    start = text.find("{")
    end = text.rfind("}")
    if start < 0 or end <= start:
        return None
    try:
        data = json.loads(text[start:end + 1])
        return data if isinstance(data, dict) else None
    except json.JSONDecodeError:
        return None


def png_size(path: str) -> Tuple[int, int]:
    """只读 PNG 的 IHDR 头拿宽高（避免为此依赖 Pillow）。"""
    try:
        with open(path, "rb") as handle:
            head = handle.read(33)
    except OSError as exc:
        raise RenderError(f"读不到渲染结果 {path}：{exc}") from None
    if head[:8] != b"\x89PNG\r\n\x1a\n" or head[12:16] != b"IHDR":
        raise RenderError(f"{path} 不是有效 PNG（文件头不匹配）")
    width, height = struct.unpack(">II", head[16:24])
    if width <= 0 or height <= 0:
        raise RenderError(f"{path} 的 PNG 尺寸非法：{width}x{height}")
    return width, height


def pptx_page_count(path: str) -> int:
    """用 zip 目录数 ppt/slides/slideN.xml，不依赖 python-pptx。"""
    if not zipfile.is_zipfile(path):
        raise RenderError(f"{path} 不是有效的 .pptx（不是 zip 包）")
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
    slides = [n for n in names if re.match(r"^ppt/slides/slide\d+\.xml$", n)]
    if not slides:
        raise RenderError(f"{path} 里没有幻灯片（ppt/slides/slideN.xml），可能是空文档")
    return len(slides)


def parse_pages(spec: Optional[str], total: int) -> List[int]:
    """把 --pages 解析成 1 起的页号列表。

    支持 "all"、""（视为全部）、"1,2,3"、"1-3"、"3,1"；重复页号会去重并升序排列，
    因为 PNG 文件名按页号编号，顺序固定才可预测。
    """
    if total <= 0:
        raise RenderError("文档页数为 0，无法渲染")
    if spec is None or str(spec).strip() == "" or str(spec).strip().lower() == "all":
        return list(range(1, total + 1))
    text = str(spec).strip().replace("，", ",").replace("、", ",")
    pages: List[int] = []
    for chunk in text.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        match = re.fullmatch(r"(\d+)\s*[-~]\s*(\d+)", chunk)
        if match:
            begin, end = int(match.group(1)), int(match.group(2))
            if begin > end:
                begin, end = end, begin
            pages.extend(range(begin, end + 1))
            continue
        if not re.fullmatch(r"\d+", chunk):
            raise RenderError(f"--pages 里的 {chunk!r} 不是页号；写法：all / \"1,2,3\" / \"1-3\"")
        pages.append(int(chunk))
    if not pages:
        raise RenderError("--pages 没有解析出任何页号")
    bad = [p for p in pages if p < 1 or p > total]
    if bad:
        raise RenderError(f"--pages 里的页号 {bad} 超出范围（本文档共 {total} 页）")
    return sorted(set(pages))


# ---------------------------------------------------------------------------
# 后端探测
# ---------------------------------------------------------------------------


@dataclass
class Backend:
    """一个渲染后端的能力快照（探测结果，供 probe.py 与 render.py 共用）。"""

    id: str
    label: str
    available: bool
    detail: str
    hint: str = ""
    data: Dict[str, str] = field(default_factory=dict)


def _dedupe(items: Sequence[str]) -> List[str]:
    seen = set()
    result = []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            result.append(item)
    return result


def _appdata() -> str:
    return os.environ.get("APPDATA") or os.path.expanduser(os.path.join("~", "AppData", "Roaming"))


def _dsh_node_candidates() -> List[str]:
    """找 node 可执行文件：环境变量 → PATH → DSH 桌面端的 .cmd 转发脚本。"""
    candidates: List[str] = []
    env_node = os.environ.get("DSH_NODE")
    if env_node:
        candidates.append(env_node)
    for name in ("node", "node.exe", "node.cmd"):
        found = shutil.which(name)
        if found:
            candidates.append(found)
    candidates.append(os.path.join(_appdata(), "dsh-desktop", "harness", ".desktop-bin", "node.cmd"))
    return [p for p in _dedupe(candidates) if os.path.isfile(p)]


def _dsh_cli_candidates(nodes: Sequence[str]) -> List[str]:
    """找 office-cli.mjs：环境变量 → 从 node 转发脚本里读出的安装目录/resources/。"""
    candidates: List[str] = []
    env_cli = os.environ.get("DSH_OFFICE_CLI")
    if env_cli:
        candidates.append(env_cli)
    for node in nodes:
        if not node.lower().endswith((".cmd", ".bat")):
            continue
        try:
            with open(node, "r", encoding="utf-8", errors="replace") as handle:
                text = handle.read()
        except OSError:
            continue
        # node.cmd 里写着真正的 exe 路径（DSH Desktop.exe），同一安装目录下就是 resources/
        for match in re.finditer(r'"([A-Za-z]:\\[^"\r\n]+?\.exe)"', text):
            exe = match.group(1)
            candidates.append(os.path.join(os.path.dirname(exe), "resources", "office-cli.mjs"))
    return [p for p in _dedupe(candidates) if os.path.isfile(p)]


def probe_ppt_render_cmd() -> Backend:
    template = os.environ.get("PPT_RENDER_CMD", "").strip()
    if not template:
        return Backend(
            BACKEND_PPT_CMD, "自定义命令（PPT_RENDER_CMD）", False,
            "未设置 PPT_RENDER_CMD",
            "设置环境变量 PPT_RENDER_CMD，例如：mytool --in {input} --out {outdir} --pages {pages} --dpi {dpi}",
        )
    return Backend(BACKEND_PPT_CMD, "自定义命令（PPT_RENDER_CMD）", True, template, data={"template": template})


def probe_dsh() -> Backend:
    nodes = _dsh_node_candidates()
    clis = _dsh_cli_candidates(nodes)
    if not nodes:
        return Backend(
            BACKEND_DSH, "本机 DSH 渲染器", False, "找不到 node 可执行文件",
            "安装 Node.js 并加入 PATH，或设置 DSH_NODE 指向 node 可执行文件",
        )
    if not clis:
        return Backend(
            BACKEND_DSH, "本机 DSH 渲染器", False,
            f"找到 node（{nodes[0]}）但找不到 office-cli.mjs",
            "设置 DSH_OFFICE_CLI 指向 office-cli.mjs（DSH 桌面端安装目录 resources/ 下）",
        )
    return Backend(
        BACKEND_DSH, "本机 DSH 渲染器", True,
        f"node={nodes[0]}；office-cli={clis[0]}",
        data={"node": nodes[0], "cli": clis[0]},
    )


def _soffice_candidates() -> List[str]:
    candidates: List[str] = []
    env_soffice = os.environ.get("DSH_SOFFICE")
    if env_soffice:
        candidates.append(env_soffice)
    for name in ("soffice", "soffice.exe"):
        found = shutil.which(name)
        if found:
            candidates.append(found)
    candidates.extend([
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        "/opt/libreoffice/program/soffice",
    ])
    return [p for p in _dedupe(candidates) if os.path.isfile(p)]


def _pdftoppm_candidates() -> List[str]:
    candidates: List[str] = []
    env_pdftoppm = os.environ.get("DSH_PDFTOPPM")
    if env_pdftoppm:
        candidates.append(env_pdftoppm)
    for name in ("pdftoppm", "pdftoppm.exe"):
        found = shutil.which(name)
        if found:
            candidates.append(found)
    for pattern_root in (r"C:\poppler", os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "poppler")):
        if os.path.isdir(pattern_root):
            for root, _dirs, files in os.walk(pattern_root):
                if "pdftoppm.exe" in files:
                    candidates.append(os.path.join(root, "pdftoppm.exe"))
    return [p for p in _dedupe(candidates) if os.path.isfile(p)]


def _gs_candidates() -> List[str]:
    candidates: List[str] = []
    for name in ("gswin64c", "gswin32c", "gs"):
        found = shutil.which(name)
        if found:
            candidates.append(found)
    return _dedupe(candidates)


def probe_soffice() -> Backend:
    soffice = _soffice_candidates()
    if not soffice:
        return Backend(
            BACKEND_SOFFICE, "LibreOffice + pdftoppm", False, "未找到 soffice",
            "安装 LibreOffice（Windows 默认路径 C:\\Program Files\\LibreOffice）",
        )
    pdftoppm = _pdftoppm_candidates()
    gs = _gs_candidates()
    if not pdftoppm and not gs:
        return Backend(
            BACKEND_SOFFICE, "LibreOffice + pdftoppm", False,
            f"找到 soffice（{soffice[0]}）但缺少 PDF→PNG 工具",
            "安装 poppler（提供 pdftoppm）或 Ghostscript；Pillow 不能直接读 PDF",
            data={"soffice": soffice[0]},
        )
    return Backend(
        BACKEND_SOFFICE, "LibreOffice + pdftoppm", True,
        f"soffice={soffice[0]}；" + (f"pdftoppm={pdftoppm[0]}" if pdftoppm else f"gs={gs[0]}"),
        data={"soffice": soffice[0],
              "pdftoppm": pdftoppm[0] if pdftoppm else "",
              "gs": gs[0] if gs else ""},
    )


def probe_powerpoint_com() -> Backend:
    if os.name != "nt":
        return Backend(BACKEND_COM, "PowerPoint COM", False, "仅 Windows 可用", "改用 LibreOffice 或 DSH 渲染器")
    try:
        import win32com.client  # noqa: F401
    except Exception:
        return Backend(
            BACKEND_COM, "PowerPoint COM", False, "未安装 pywin32",
            "pip install pywin32（还需要本机安装 Microsoft PowerPoint）",
        )
    try:
        import winreg

        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application"))
    except OSError:
        return Backend(
            BACKEND_COM, "PowerPoint COM", False, "注册表里没有 PowerPoint.Application",
            "安装 Microsoft PowerPoint（WPS 不注册该 COM 接口）",
        )
    return Backend(BACKEND_COM, "PowerPoint COM", True, "pywin32 + PowerPoint 就绪")


def probe_backends() -> List[Backend]:
    """按优先级返回所有后端的能力快照（供 render.py 挑选、probe.py 报告）。"""
    return [probe_ppt_render_cmd(), probe_dsh(), probe_soffice(), probe_powerpoint_com()]


# ---------------------------------------------------------------------------
# 各后端的渲染实现
# ---------------------------------------------------------------------------


@dataclass
class RenderOutput:
    """一次成功渲染的结果（图片仍在临时目录里，等 main 搬运）。"""

    backend: str
    backend_detail: str
    images: List[Tuple[int, str, int, int]]  # (源页号, 临时 PNG 路径, 宽, 高)
    missing_fonts: List[str] = field(default_factory=list)
    reported_dpi: Optional[float] = None


def _require_fresh_dir(path: str) -> str:
    """准备一个「存在且为空」的目录（soffice / COM / 自定义命令需要目录先存在）。"""
    if os.path.exists(path):
        shutil.rmtree(path, ignore_errors=True)
    os.makedirs(path)
    return path


def _require_absent_path(path: str) -> str:
    """准备一个「不存在」的目录路径。

    本机 DSH 渲染器自己 mkdir，目标已存在时直接报
    "EEXIST: file already exists, mkdir ..."，所以这里只清理、不创建。
    """
    if os.path.exists(path):
        shutil.rmtree(path, ignore_errors=True)
    return path


def _render_ppt_cmd(pptx: str, pages: Sequence[int], dpi: int, backend: Backend, staging: str) -> RenderOutput:
    template = backend.data["template"]
    out_dir = _require_fresh_dir(os.path.join(staging, BACKEND_PPT_CMD))
    values = {
        "input": pptx,
        "outdir": out_dir,
        "pages": ",".join(str(p) for p in pages),
        "dpi": str(dpi),
    }
    expanded = template
    for key, value in values.items():
        expanded = expanded.replace("{" + key + "}", value)
    if "{" in expanded and "}" in expanded:
        raise RenderError(f"PPT_RENDER_CMD 里有无法识别的占位符：{template}")
    parts = [token.strip('"').strip("'") for token in re.findall(r'"[^"]*"|\S+', expanded)]
    proc = _run(parts)
    if proc.returncode != 0:
        raise RenderError(f"PPT_RENDER_CMD 退出码 {proc.returncode}：{_first_line(proc.stderr or proc.stdout)}")

    found = []
    for name in sorted(os.listdir(out_dir)):
        if name.lower().endswith(".png"):
            found.append(os.path.join(out_dir, name))
    if not found:
        raise RenderError(f"PPT_RENDER_CMD 没有在 {out_dir} 生成 PNG；请确认命令把图片写进 {{outdir}}")
    if len(found) < len(pages):
        raise RenderError(f"PPT_RENDER_CMD 只生成了 {len(found)} 张 PNG，但请求了 {len(pages)} 页")
    images = []
    for index, path in enumerate(found[:len(pages)]):
        width, height = png_size(path)
        images.append((pages[index], path, width, height))
    return RenderOutput(BACKEND_PPT_CMD, "自定义命令", images, [])


def _render_dsh(pptx: str, pages: Sequence[int], dpi: int, backend: Backend, staging: str) -> RenderOutput:
    node, cli = backend.data["node"], backend.data["cli"]
    out_dir = _require_absent_path(os.path.join(staging, BACKEND_DSH))
    cmd = [node, cli, "render", "--input", pptx, "--output-dir", out_dir,
           "--pages", ",".join(str(p) for p in pages), "--dpi", str(dpi)]
    proc = _run(cmd)
    if proc.returncode != 0:
        parsed = _extract_json(proc.stderr) or _extract_json(proc.stdout) or {}
        message = parsed.get("error") or _first_line(proc.stderr or proc.stdout)
        raise RenderError(f"本机 DSH 渲染器退出码 {proc.returncode}：{message}")
    data = _extract_json(proc.stdout)
    if not data:
        raise RenderError(f"本机 DSH 渲染器没有返回可解析的 JSON：{_first_line(proc.stdout)}")
    images: List[Tuple[int, str, int, int]] = []
    for item in data.get("images") or []:
        path = item.get("path") or ""
        page = int(item.get("page") or 0)
        if not path or not os.path.isfile(path):
            raise RenderError(f"渲染器报告了图片但文件不存在：{path}")
        width = int(item.get("width") or 0)
        height = int(item.get("height") or 0)
        if width <= 0 or height <= 0:
            width, height = png_size(path)
        images.append((page, path, width, height))
    if not images:
        raise RenderError("本机 DSH 渲染器返回的 images 为空")
    detail = f"{data.get('backend', 'native')} / {data.get('rasterEngine', '未知光栅引擎')}"
    return RenderOutput(
        BACKEND_DSH, detail, images,
        [str(f) for f in (data.get("missingFonts") or [])],
        float(data.get("dpi")) if data.get("dpi") else None,
    )


def _render_soffice(pptx: str, pages: Sequence[int], dpi: int, backend: Backend, staging: str) -> RenderOutput:
    soffice = backend.data["soffice"]
    pdftoppm = backend.data.get("pdftoppm") or ""
    gs = backend.data.get("gs") or ""
    work = _require_fresh_dir(os.path.join(staging, BACKEND_SOFFICE))
    proc = _run([soffice, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", work, pptx])
    pdfs = [os.path.join(work, n) for n in sorted(os.listdir(work)) if n.lower().endswith(".pdf")]
    if not pdfs:
        raise RenderError(
            "LibreOffice 没有生成 PDF：" + _first_line(proc.stderr or proc.stdout)
            + "（若 soffice 正在运行，先关闭它再重试）"
        )
    pdf = pdfs[0]
    images: List[Tuple[int, str, int, int]] = []
    if pdftoppm:
        prefix = os.path.join(work, "page")
        proc = _run([pdftoppm, "-r", str(dpi), "-png", pdf, prefix])
        produced = sorted(p for p in os.listdir(work) if p.startswith("page") and p.lower().endswith(".png"))
        if not produced:
            raise RenderError("pdftoppm 没有生成 PNG：" + _first_line(proc.stderr or proc.stdout))
        for index, name in enumerate(produced):
            page = index + 1
            path = os.path.join(work, name)
            width, height = png_size(path)
            images.append((page, path, width, height))
    else:
        pattern = os.path.join(work, "page-%03d.png")
        proc = _run([gs, "-dNOPAUSE", "-dBATCH", "-dSAFER", "-sDEVICE=png16m",
                     f"-r{dpi}", f"-sOutputFile={pattern}", pdf])
        produced = sorted(p for p in os.listdir(work) if p.startswith("page-") and p.lower().endswith(".png"))
        if not produced:
            raise RenderError("Ghostscript 没有生成 PNG：" + _first_line(proc.stderr or proc.stdout))
        for index, name in enumerate(produced):
            page = index + 1
            path = os.path.join(work, name)
            width, height = png_size(path)
            images.append((page, path, width, height))
    wanted = set(pages)
    selected = [item for item in images if item[0] in wanted]
    if not selected:
        raise RenderError(f"渲染出的页 {[i[0] for i in images]} 不包含请求的 {list(pages)}")
    return RenderOutput(BACKEND_SOFFICE, "LibreOffice → PDF → PNG", selected, [], float(dpi))


def _render_com(pptx: str, pages: Sequence[int], dpi: int, backend: Backend, staging: str) -> RenderOutput:
    import pythoncom  # type: ignore
    import win32com.client as win32  # type: ignore

    work = _require_fresh_dir(os.path.join(staging, BACKEND_COM))
    pythoncom.CoInitialize()
    app = None
    presentation = None
    images: List[Tuple[int, str, int, int]] = []
    try:
        app = win32.Dispatch("PowerPoint.Application")
        try:
            presentation = app.Presentations.Open(pptx, ReadOnly=True, WithWindow=False)
        except Exception:
            app.Visible = True
            presentation = app.Presentations.Open(pptx, ReadOnly=True, WithWindow=True)
        slide_w_pt = float(presentation.PageSetup.SlideWidth)
        slide_h_pt = float(presentation.PageSetup.SlideHeight)
        width_px = max(1, int(round(slide_w_pt / 72.0 * dpi)))
        height_px = max(1, int(round(slide_h_pt / 72.0 * dpi)))
        for page in pages:
            path = os.path.join(work, f"page-{page:04d}.png")
            presentation.Slides(page).Export(path, "PNG", width_px, height_px)
            width, height = png_size(path)
            images.append((page, path, width, height))
    except RenderError:
        raise
    except Exception as exc:  # COM 抛的是 pywintypes.com_error
        raise RenderError(f"PowerPoint COM 渲染失败：{exc}") from None
    finally:
        try:
            if presentation is not None:
                presentation.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()
    if not images:
        raise RenderError("PowerPoint COM 没有导出任何页")
    return RenderOutput(BACKEND_COM, "PowerPoint 原生导出", images, [], float(dpi))


_RENDERERS = {
    BACKEND_PPT_CMD: _render_ppt_cmd,
    BACKEND_DSH: _render_dsh,
    BACKEND_SOFFICE: _render_soffice,
    BACKEND_COM: _render_com,
}


# ---------------------------------------------------------------------------
# 落盘与清单
# ---------------------------------------------------------------------------


def materialize(images: Sequence[Tuple[int, str, int, int]], out_dir: str,
                pages: Sequence[int]) -> List[dict]:
    """把临时 PNG 按源页号改名成 slide-NN.png 并复制到 out-dir。"""
    by_page = {page: (path, width, height) for page, path, width, height in images}
    missing = [p for p in pages if p not in by_page]
    if missing:
        raise RenderError(f"渲染结果缺少第 {missing} 页，无法生成完整预览")
    try:
        os.makedirs(out_dir, exist_ok=True)
    except OSError as exc:
        raise RenderError(f"无法创建输出目录 {out_dir}：{exc}") from None

    width_pad = max(2, len(str(max(pages))))
    records: List[dict] = []
    for page in pages:
        source, width, height = by_page[page]
        filename = f"slide-{page:0{width_pad}d}.png"
        target = os.path.join(out_dir, filename)
        try:
            shutil.copyfile(source, target)
        except OSError as exc:
            raise RenderError(f"复制预览图到 {target} 失败：{exc}") from None
        records.append({
            "page": page,
            "filename": filename,
            "path": os.path.abspath(target),
            "width": width,
            "height": height,
        })
    _prune_stale(out_dir, {r["filename"] for r in records})
    return records


def _prune_stale(out_dir: str, keep: set) -> None:
    """删掉上一次渲染残留、本次没有重新生成的 slide-*.png（只动这个命名模式）。"""
    try:
        names = os.listdir(out_dir)
    except OSError:
        return
    for name in names:
        if re.fullmatch(r"slide-\d+\.png", name) and name not in keep:
            try:
                os.remove(os.path.join(out_dir, name))
            except OSError:
                pass


def write_manifest(out_dir: str, pptx: str, output: RenderOutput, records: Sequence[dict],
                   pages: Sequence[int], total_pages: int, dpi: float) -> dict:
    manifest = {
        "schemaVersion": MANIFEST_SCHEMA_VERSION,
        "pptx": os.path.abspath(pptx),
        "backend": output.backend,
        "backendDetail": output.backend_detail,
        "dpi": dpi,
        "pageCount": total_pages,
        "requestedPages": list(pages),
        "images": list(records),
        "missingFonts": list(output.missing_fonts),
        "generatedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
    }
    path = os.path.join(out_dir, MANIFEST_NAME)
    try:
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(manifest, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
    except OSError as exc:
        raise RenderError(f"写入清单失败 {path}：{exc}") from None
    return manifest


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="render.py",
        description="把 .pptx 渲染成 PNG 预览（逐后端探测，失败自动降级）。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python scripts/render.py --pptx out/deck.pptx --out-dir preview\n"
            "  python scripts/render.py --pptx out/deck.pptx --out-dir preview --pages \"1,3\" --dpi 96\n"
            "  python scripts/render.py --pptx out/deck.pptx --out-dir preview --pages all --json\n"
            "\n"
            "输出：<out-dir>/slide-01.png … 与 <out-dir>/render-manifest.json\n"
            "退出码：0 成功 / 1 输入参数有问题 / 2 所有渲染后端不可用或全部失败"
        ),
    )
    parser.add_argument("--pptx", required=True, help="输入 .pptx 路径（相对或绝对）")
    parser.add_argument("--out-dir", default="preview", help="输出目录（默认 preview/，会自动创建）")
    parser.add_argument("--pages", default="all",
                        help='要渲染的页：all（默认）、"1,2,3" 或 "1-3"；页号从 1 开始')
    parser.add_argument("--dpi", type=int, default=DEFAULT_DPI,
                        help=f"渲染 DPI（默认 {DEFAULT_DPI}；50-600）")
    parser.add_argument("--json", action="store_true", help="把 render-manifest 以 JSON 打到 stdout")
    parser.add_argument("--timeout", type=int, default=TIMEOUT_S, help=f"单次渲染超时秒数（默认 {TIMEOUT_S}）")
    return parser


def _print_failure(attempts: List[Tuple[Backend, str]]) -> None:
    sys.stderr.write("错误：所有渲染后端都不可用或渲染失败。\n")
    for backend, reason in attempts:
        sys.stderr.write(f"  - {backend.label}（{backend.id}）：{reason}\n")
    sys.stderr.write(
        "\n怎么修（按成本从低到高）：\n"
        "  1) 用本机自带的 DSH 渲染器：确认 node 在 PATH，且设置 DSH_OFFICE_CLI 指向 office-cli.mjs；\n"
        "  2) 安装 LibreOffice（提供 soffice）和 poppler（提供 pdftoppm）；\n"
        "  3) Windows 上装 Microsoft PowerPoint + pywin32（pip install pywin32）；\n"
        "  4) 或者用 PPT_RENDER_CMD 指定你们自己的渲染命令，模板里用 {input} {outdir} {pages} {dpi}。\n"
        "注意：Pillow 不能直接读 PDF，所以只用 LibreOffice、不装 poppler/ghostscript 时无法出 PNG。\n"
    )


def main(argv: Optional[Sequence[str]] = None) -> int:
    _configure_stdio()
    parser = build_parser()
    args = parser.parse_args(argv)

    global TIMEOUT_S
    TIMEOUT_S = max(30, int(args.timeout))

    pptx = os.path.abspath(args.pptx)
    if not os.path.isfile(pptx):
        sys.stderr.write(f"错误：找不到 pptx：{pptx}\n")
        return 1
    if os.path.getsize(pptx) == 0:
        sys.stderr.write(f"错误：pptx 是空文件：{pptx}\n")
        return 1
    if not MIN_DPI <= args.dpi <= MAX_DPI:
        sys.stderr.write(f"错误：--dpi 必须落在 {MIN_DPI}-{MAX_DPI}，实际是 {args.dpi}\n")
        return 1
    out_dir = os.path.abspath(args.out_dir)

    try:
        total_pages = pptx_page_count(pptx)
        pages = parse_pages(args.pages, total_pages)
    except RenderError as exc:
        sys.stderr.write(f"错误：{exc}\n")
        return 1

    backends = probe_backends()
    attempts: List[Tuple[Backend, str]] = []
    staging = tempfile.mkdtemp(prefix="ppt-render-")
    output: Optional[RenderOutput] = None
    try:
        for backend in backends:
            if not backend.available:
                attempts.append((backend, f"不可用：{backend.detail}"))
                continue
            try:
                output = _RENDERERS[backend.id](pptx, pages, args.dpi, backend, staging)
                break
            except RenderError as exc:
                attempts.append((backend, str(exc)))
                continue
            except Exception as exc:  # 后端内部异常不应中断降级流程
                attempts.append((backend, f"{type(exc).__name__}: {exc}"))
                continue

        if output is None:
            _print_failure(attempts)
            return 2

        try:
            records = materialize(output.images, out_dir, pages)
        except RenderError as exc:
            sys.stderr.write(f"错误：{exc}\n")
            return 1
        dpi_value = output.reported_dpi if output.reported_dpi else float(args.dpi)
        try:
            manifest = write_manifest(out_dir, pptx, output, records, pages, total_pages, dpi_value)
        except RenderError as exc:
            sys.stderr.write(f"错误：{exc}\n")
            return 1
    finally:
        shutil.rmtree(staging, ignore_errors=True)

    if args.json:
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
    else:
        print("=== 渲染摘要 ===")
        print(f"输出目录：{out_dir}")
        print(f"渲染后端：{output.backend}（{output.backend_detail}）")
        print(f"DPI：{dpi_value:g}")
        selected = "全部" if len(pages) == total_pages else f"{len(pages)}/{total_pages}"
        print(f"页数：{len(records)}（{selected}）")
        for record in records:
            size_kb = os.path.getsize(record["path"]) / 1024.0
            print(f"  {record['filename']:14s} {record['width']}x{record['height']}  {size_kb:.1f} KB")
        print(f"清单：{os.path.join(out_dir, MANIFEST_NAME)}")
        if output.missing_fonts:
            print(f"缺失字体（{len(output.missing_fonts)}）：" + "、".join(output.missing_fonts))
        else:
            print("缺失字体：无")
        skipped = [f"{b.label}={reason.split('：')[0]}" for b, reason in attempts if b.id != output.backend]
        if skipped:
            print("跳过的后端：" + "；".join(skipped))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
