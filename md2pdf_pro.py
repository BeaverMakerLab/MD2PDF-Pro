# -*- coding: utf-8 -*-
"""
MD2PDF Pro - Markdown 转 PDF 文档排版工具
- 原生拖拽引擎：集成官方 TkDND (tkinterdnd2) C 扩展，杜绝 Windows 64 位指针截断崩溃
- 商务科技美学：经典科技蓝视觉基调、精细表格与代码块卡片、柔和段落呼吸感
- 规范排版：A4 纵横向切换、标准页边距、Word 中文经典字号（小初至小五及任意磅值 pt）、行间距、段落间距、首行缩进、两端对齐
- 页面元素：规范页码、首页封面隐藏、自定义页眉、防泄密矢量背景水印
- 批处理流水线：多文件独立批量导出与长篇多章节顺序合并
"""

import os
import sys
import json
import re
import html
import subprocess
import tempfile
import urllib.parse
import threading

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES

import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from bs4 import BeautifulSoup, NavigableString, Tag

# ==============================================================================
# 路径与排版常量字典
# ==============================================================================

if getattr(sys, "frozen", False):
    TOOL_DIR = os.path.dirname(os.path.abspath(sys.executable))
    BUNDLE_DIR = getattr(sys, "_MEIPASS", TOOL_DIR)
    ASSETS_DIR = os.path.join(BUNDLE_DIR, "assets")
    if not os.path.exists(ASSETS_DIR):
        ASSETS_DIR = os.path.join(TOOL_DIR, "assets")
    JS_DIR = os.path.join(ASSETS_DIR, "js")
    CONFIG_PATH = os.path.join(TOOL_DIR, "config.json")
else:
    TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
    ASSETS_DIR = os.path.join(TOOL_DIR, "assets")
    JS_DIR = os.path.join(ASSETS_DIR, "js")
    CONFIG_PATH = os.path.join(TOOL_DIR, "config.json")

# 中文出版与 Word 标准字号与磅值 (pt) 映射表
FONT_SIZE_MAP = {
    "初号 (42 pt)": 42.0,
    "小初 (36 pt)": 36.0,
    "一号 (26 pt)": 26.0,
    "小一 (24 pt)": 24.0,
    "二号 (22 pt)": 22.0,
    "小二 (18 pt)": 18.0,
    "三号 (16 pt)": 16.0,
    "小三 (15 pt)": 15.0,
    "四号 (14 pt)": 14.0,
    "小四 (12 pt)": 12.0,
    "五号 (10.5 pt)": 10.5,
    "小五 (9 pt)": 9.0,
    "六号 (7.5 pt)": 7.5,
    "七号 (5.5 pt)": 5.5,
    "八号 (5 pt)": 5.0,
}

FONT_SIZE_CHOICES = [
    "小初 (36 pt)",
    "一号 (26 pt)",
    "二号 (22 pt)",
    "小二 (18 pt)",
    "三号 (16 pt)",
    "小三 (15 pt)",
    "四号 (14 pt)",
    "小四 (12 pt)",
    "五号 (10.5 pt)",
    "小五 (9 pt)",
]

# 水印字号与磅值 (pt) 映射表
WATERMARK_SIZE_MAP = {
    "较小 (14 pt)": 14,
    "标准 (18 pt)": 18,
    "中等 (22 pt)": 22,
    "较大 (26 pt)": 26,
    "醒目 (32 pt)": 32,
    "超大 (40 pt)": 40,
}
WATERMARK_SIZE_CHOICES = list(WATERMARK_SIZE_MAP.keys())

FONT_MAP = {
    "微软雅黑 (推荐·现代无衬线)": "Microsoft YaHei, -apple-system, sans-serif",
    "仿宋 (公文标准·严肃正式)": "'FangSong', '仿宋', 'SimSun', serif",
    "宋体 (报刊典雅·传统衬线)": "'SimSun', '宋体', serif",
    "楷体 (清秀书法)": "'KaiTi', '楷体', serif",
    "等宽 Consolas (技术文档)": "Consolas, 'Courier New', monospace"
}

MARGIN_MAP = {
    "standard": {
        "css": "20mm 16mm 20mm 16mm",
        "dict": {"top": "20mm", "bottom": "20mm", "left": "16mm", "right": "16mm"},
        "label": "标准 A4 (上下20 / 左右16mm)"
    },
    "compact": {
        "css": "12mm 12mm 12mm 12mm",
        "dict": {"top": "12mm", "bottom": "12mm", "left": "12mm", "right": "12mm"},
        "label": "紧凑窄边 (上下左右 12mm)"
    },
    "spacious": {
        "css": "25mm 22mm 25mm 22mm",
        "dict": {"top": "25mm", "bottom": "25mm", "left": "22mm", "right": "22mm"},
        "label": "宽裕留白 (上下25 / 左右22mm)"
    },
    "gov": {
        "css": "37mm 26mm 35mm 28mm",
        "dict": {"top": "37mm", "bottom": "35mm", "left": "28mm", "right": "26mm"},
        "label": "公文标准 (上37/下35/左28/右26mm)"
    }
}

COLOR_PRESETS = [
    ("#2b6cb0", "科技蓝"),
    ("#2f855a", "商务绿"),
    ("#c53030", "中国红"),
    ("#6b46c1", "典雅紫"),
    ("#1a202c", "墨黑")
]

PRESET_THEMES = {
    "tech_blue": {
        "name": "政企科技蓝 (默认推荐)",
        "font_family_key": "微软雅黑 (推荐·现代无衬线)",
        "font_size": "小四 (12 pt)",
        "accent_color": "#2b6cb0",
        "line_height": "1.65 倍 (推荐)",
        "para_spacing": "标准 (8 pt)",
        "page_margin": "standard",
        "page_orientation": "portrait",
        "text_align_justify": True,
        "indent_first_line": False,
        "h1_align": "center",
        "code_style": "blue_accent"
    },
    "gov": {
        "name": "严谨公文风 (仿宋·红黑·首行缩进)",
        "font_family_key": "仿宋 (公文标准·严肃正式)",
        "font_size": "三号 (16 pt)",
        "accent_color": "#c53030",
        "line_height": "1.75 倍",
        "para_spacing": "紧凑 (4 pt)",
        "page_margin": "gov",
        "page_orientation": "portrait",
        "text_align_justify": True,
        "indent_first_line": True,
        "h1_align": "center",
        "code_style": "minimal"
    },
    "github_clean": {
        "name": "GitHub 经典清爽 (开发者开源风)",
        "font_family_key": "微软雅黑 (推荐·现代无衬线)",
        "font_size": "五号 (10.5 pt)",
        "accent_color": "#24292e",
        "line_height": "1.50 倍",
        "para_spacing": "标准 (8 pt)",
        "page_margin": "standard",
        "page_orientation": "portrait",
        "text_align_justify": True,
        "indent_first_line": False,
        "h1_align": "left",
        "code_style": "github_clean"
    }
}

DEFAULT_CONFIG = {
    "output_mode": "separate",
    "export_format": "pdf",
    "last_output_dir": "",
    "preset": "tech_blue",
    "page_orientation": "portrait",
    "page_margin": "standard",
    "h1_align": "center",
    "enable_toc": True,

    "font_family": "Microsoft YaHei, -apple-system, sans-serif",
    "font_family_key": "微软雅黑 (推荐·现代无衬线)",
    "font_size": "小四 (12 pt)",
    "line_height": "1.65 倍 (推荐)",
    "para_spacing": "标准 (8 pt)",
    "text_align_justify": True,
    "indent_first_line": False,

    "enable_footer": True,
    "suppress_first_page": True,
    "enable_header": True,
    "header_text": "内部交流材料",
    "enable_watermark": False,
    "watermark_text": "内部资料 严禁外传",
    "watermark_font_size": "标准 (18 pt)",

    "accent_color": "#2b6cb0",
    "code_style": "blue_accent"
}


# ==============================================================================
# 工具函数：字号解析与配置存取
# ==============================================================================

def parse_font_size(size_str, default=12.0):
    """从中文规格字符串（如 '小四 (12 pt)'）或手动输入的数值（如 '12', '10.5'）中提取实际磅值"""
    if isinstance(size_str, (int, float)):
        return float(size_str)
    s = str(size_str).strip()
    if s in FONT_SIZE_MAP:
        return FONT_SIZE_MAP[s]
    for k, v in FONT_SIZE_MAP.items():
        if k.startswith(s) or s in k:
            return v
    m = re.search(r"([0-9]+(?:\.[0-9]+)?)", s)
    if m:
        try:
            val = float(m.group(1))
            if 5.0 <= val <= 72.0:
                return val
        except ValueError:
            pass
    return default

def parse_watermark_size(val, default=18):
    """从水印字号选项或手动输入文本中提取实际磅值"""
    if isinstance(val, (int, float)):
        return int(val)
    s = str(val).strip()
    if s in WATERMARK_SIZE_MAP:
        return WATERMARK_SIZE_MAP[s]
    for k, v in WATERMARK_SIZE_MAP.items():
        if s in k:
            return v
    m = re.search(r"([0-9]+)", s)
    if m:
        try:
            num = int(m.group(1))
            if 8 <= num <= 72:
                return num
        except ValueError:
            pass
    return default

def parse_line_height(lh_str, default=1.65):
    if isinstance(lh_str, (int, float)):
        return float(lh_str)
    m = re.search(r"([0-9]+(?:\.[0-9]+)?)", str(lh_str))
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    return default

def parse_para_spacing(ps_str, default=8.0):
    if isinstance(ps_str, (int, float)):
        return float(ps_str)
    s = str(ps_str)
    if "紧凑" in s or "4" in s:
        return 4.0
    elif "宽松" in s or "12" in s:
        return 12.0
    return default

def load_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                c = json.load(f)
                res = DEFAULT_CONFIG.copy()
                res.update(c)
                return res
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存配置失败: {e}")

def find_browser():
    paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    for name in ["msedge.exe", "chrome.exe"]:
        found = subprocess.run(["where", name], capture_output=True, text=True).stdout.strip()
        if found and os.path.exists(found.splitlines()[0]):
            return found.splitlines()[0]
    return None


# ==============================================================================
# 参数化 CSS 渲染引擎 (高保真科技与商务美感)
# ==============================================================================

def generate_parametric_css(cfg):
    """根据排版参数自动生成打印级专业 CSS"""
    font_val = cfg.get("font_family", "Microsoft YaHei, -apple-system, sans-serif")
    font_family = FONT_MAP.get(font_val, font_val)

    font_size_pt = parse_font_size(cfg.get("font_size", "小四 (12 pt)"), 12.0)
    line_height = parse_line_height(cfg.get("line_height", "1.65 倍 (推荐)"), 1.65)
    para_spacing_pt = parse_para_spacing(cfg.get("para_spacing", "标准 (8 pt)"), 8.0)

    accent_color = cfg.get("accent_color", "#2b6cb0")
    h1_align = cfg.get("h1_align", "center")
    indent_first_line = cfg.get("indent_first_line", False)
    text_align_justify = cfg.get("text_align_justify", True)
    code_style = cfg.get("code_style", "blue_accent")

    orientation = cfg.get("page_orientation", "portrait")
    margin_key = cfg.get("page_margin", "standard")
    margin_css = MARGIN_MAP.get(margin_key, MARGIN_MAP["standard"])["css"]

    # 标题按排版比例自动联动
    h1_size = round(font_size_pt * 1.55, 1)
    h2_size = round(font_size_pt * 1.25, 1)
    h3_size = round(font_size_pt * 1.10, 1)

    indent_css = "text-indent: 2em;" if indent_first_line else ""
    justify_css = """
        text-align: justify;
        text-justify: inter-ideograph;
        word-break: break-word;
    """ if text_align_justify else ""

    # 代码块风格分支
    if code_style == "blue_accent":
        code_block_css = f"""
        pre {{
            background-color: #f8fafc;
            border: 1px solid #e2e8f0;
            border-left: 3.5px solid {accent_color};
            padding: 9pt 12pt;
            border-radius: 4px;
            font-family: Consolas, 'Courier New', monospace;
            font-size: {round(font_size_pt * 0.88, 1)}pt;
            line-height: 1.48;
            page-break-inside: avoid;
            margin: 8pt 0 12pt 0;
            color: #2d3748;
            white-space: pre-wrap;
            word-break: break-all;
        }}
        pre code {{ background: transparent; border: none; padding: 0; color: #2d3748; }}
        """
    elif code_style == "github_clean":
        code_block_css = f"""
        pre {{
            background-color: #f6f8fa;
            border: 1px solid #d0d7de;
            padding: 9pt 12pt;
            border-radius: 6px;
            font-family: Consolas, monospace;
            font-size: {round(font_size_pt * 0.88, 1)}pt;
            line-height: 1.45;
            page-break-inside: avoid;
            margin: 8pt 0 12pt 0;
            color: #24292e;
            white-space: pre-wrap;
            word-break: break-all;
        }}
        pre code {{ background: transparent; border: none; padding: 0; color: #24292e; }}
        """
    elif code_style == "dark":
        code_block_css = f"""
        pre {{
            background-color: #1e293b;
            border: 1px solid #334155;
            padding: 9pt 12pt;
            border-radius: 6px;
            font-family: Consolas, monospace;
            font-size: {round(font_size_pt * 0.88, 1)}pt;
            line-height: 1.48;
            page-break-inside: avoid;
            margin: 8pt 0 12pt 0;
            color: #e2e8f0;
            white-space: pre-wrap;
            word-break: break-all;
        }}
        pre code {{ background: transparent; border: none; padding: 0; color: #e2e8f0; }}
        """
    else:  # minimal
        code_block_css = f"""
        pre {{
            background-color: #f7fafc;
            border: none;
            padding: 9pt 12pt;
            border-radius: 4px;
            font-family: Consolas, monospace;
            font-size: {round(font_size_pt * 0.88, 1)}pt;
            line-height: 1.48;
            page-break-inside: avoid;
            margin: 8pt 0 12pt 0;
            color: #2d3748;
            white-space: pre-wrap;
            word-break: break-all;
        }}
        pre code {{ background: transparent; border: none; padding: 0; color: #2d3748; }}
        """

    return f"""
    @page {{
        size: A4 {orientation};
        margin: {margin_css};
    }}
    body {{
        font-family: {font_family};
        font-size: {font_size_pt}pt;
        line-height: {line_height};
        color: #2d3748;
        background: #ffffff;
    }}
    h1 {{
        font-size: {h1_size}pt;
        border-bottom: 2px solid {accent_color};
        padding-bottom: 0.3em;
        margin-top: 0;
        margin-bottom: 14pt;
        color: #1a202c;
        text-align: {h1_align};
        font-weight: 700;
    }}
    h2 {{
        font-size: {h2_size}pt;
        border-bottom: 1px solid #e2e8f0;
        padding-bottom: 0.25em;
        margin-top: 18pt;
        margin-bottom: 8pt;
        color: {accent_color};
        page-break-after: avoid;
        font-weight: 600;
    }}
    h3 {{
        font-size: {h3_size}pt;
        margin-top: 14pt;
        margin-bottom: 6pt;
        color: #2d3748;
        page-break-after: avoid;
        font-weight: 600;
    }}
    h4 {{
        font-size: {font_size_pt}pt;
        margin-top: 10pt;
        margin-bottom: 4pt;
        color: #4a5568;
        page-break-after: avoid;
        font-weight: 600;
    }}
    p {{
        margin-top: 0;
        margin-bottom: {para_spacing_pt}pt;
        {indent_css}
        {justify_css}
    }}
    ul, ol {{
        padding-left: 20pt;
        margin-top: 0;
        margin-bottom: {para_spacing_pt}pt;
    }}
    li {{
        margin-bottom: 3pt;
        line-height: {line_height};
        {justify_css}
    }}
    code {{
        font-family: Consolas, "Courier New", monospace;
        background-color: #f7fafc;
        border: 1px solid #edf2f7;
        padding: 0.15em 0.35em;
        border-radius: 3px;
        font-size: 88%;
        color: #805ad5;
    }}
    {code_block_css}
    hr {{
        height: 1px;
        background-color: #e2e8f0;
        border: none;
        margin: 16pt 0;
    }}
    table {{
        border-collapse: collapse;
        width: 100%;
        margin: 10pt 0 14pt 0;
        page-break-inside: avoid;
        table-layout: fixed;
        word-wrap: break-word;
    }}
    table th, table td {{
        border: 1px solid #cbd5e0;
        padding: 6pt 9pt;
        font-size: {round(font_size_pt * 0.92, 1)}pt;
        line-height: 1.45;
        vertical-align: top;
    }}
    table th {{
        background-color: #ebf8ff;
        color: {accent_color};
        font-weight: 600;
    }}
    table tr:nth-child(2n) {{
        background-color: #f7fafc;
    }}
    blockquote {{
        margin: 0 0 {para_spacing_pt}pt 0;
        padding: 6pt 12pt;
        color: #4a5568;
        background-color: #f7fafc;
        border-left: 3.5px solid {accent_color};
        font-size: {round(font_size_pt * 0.95, 1)}pt;
        line-height: {line_height};
        {justify_css}
    }}
    /* GFM 删除线样式 */
    del {{
        color: #a0aec0;
        text-decoration: line-through;
    }}
    /* GFM 任务清单样式 */
    li.task-list-item {{
        list-style-type: none;
        margin-left: -14pt;
    }}
    input.task-list-checkbox {{
        margin-right: 6pt;
        vertical-align: middle;
        accent-color: {accent_color};
    }}
    /* 脚注样式 */
    .footnote {{
        font-size: 85%;
        color: #718096;
        border-top: 1px solid #e2e8f0;
        margin-top: 20pt;
        padding-top: 8pt;
    }}
    .footnote ol {{
        padding-left: 14pt;
    }}
    """

def build_watermark_css(watermark_text, font_size=18):
    if not watermark_text:
        return ""
    try:
        font_size = int(float(font_size))
    except (ValueError, TypeError):
        font_size = 18
    safe_text = html.escape(watermark_text)
    char_count = max(len(watermark_text), 4)
    tile_w = max(320, int(char_count * font_size * 1.35 + 80))
    tile_h = max(220, int(font_size * 8.5))
    cx = tile_w // 2
    cy = tile_h // 2
    svg = f"""<svg xmlns='http://www.w3.org/2000/svg' width='{tile_w}' height='{tile_h}' viewBox='0 0 {tile_w} {tile_h}'><text x='{cx}' y='{cy}' text-anchor='middle' dominant-baseline='middle' fill='%23000' font-size='{font_size}' font-family='sans-serif' font-weight='500' transform='rotate(-25 {cx} {cy})'>{safe_text}</text></svg>"""
    svg_uri = urllib.parse.quote(svg)
    return f"""
    .watermark-overlay {{
        position: fixed;
        top: 0; left: 0; width: 100%; height: 100%;
        z-index: 99999;
        pointer-events: none;
        opacity: 0.07;
        background-repeat: repeat;
        background-image: url("data:image/svg+xml;utf8,{svg_uri}");
    }}
    """


# ==============================================================================
# Markdown 解析与 Chromium 离线渲染内核
# ==============================================================================

def markdown_to_html(md_text, md_dir, config, doc_title=""):
    import markdown

    extensions = [
        "tables",        # GFM 表格支持 (| 表头 | 对齐 |)
        "fenced_code",   # GFM 围栏代码块 (```语言代码块)
        "nl2br",         # 自然换行 (符合主流 AI 与 Typora 体验)
        "footnotes",     # GFM 脚注支持 ([^1])
        "sane_lists",    # 严谨规范列表嵌套与混排
        "attr_list",     # 扩展属性支持
        "def_list",      # 术语定义列表
    ]
    if config.get("enable_toc", True):
        extensions.append("toc")

    raw_html = markdown.markdown(md_text, extensions=extensions)

    def replace_mermaid(match):
        code_content = match.group(1)
        unescaped = html.unescape(code_content).strip()
        return f'<div class="mermaid">\n{unescaped}\n</div>'

    processed_html = re.sub(
        r'<pre><code class="(?:language-)?mermaid">(.*?)</code></pre>',
        replace_mermaid,
        raw_html,
        flags=re.DOTALL
    )

    # GFM 规范扩充：删除线 ~~text~~ -> <del>text</del>
    processed_html = re.sub(r'~~([^~]+)~~', r'<del>\1</del>', processed_html)

    # GFM 规范扩充：任务清单复选框 - [ ] 与 - [x]
    processed_html = re.sub(
        r'<li>\[ \]\s*',
        r'<li class="task-list-item"><input type="checkbox" class="task-list-checkbox" disabled> ',
        processed_html
    )
    processed_html = re.sub(
        r'<li>\[x\]\s*',
        r'<li class="task-list-item"><input type="checkbox" class="task-list-checkbox" checked disabled> ',
        processed_html
    )

    def to_file_url(p):
        return "file:///" + p.replace("\\", "/")

    katex_css_url = to_file_url(os.path.join(JS_DIR, "katex.min.css"))
    katex_js_url = to_file_url(os.path.join(JS_DIR, "katex.min.js"))
    autorender_js_url = to_file_url(os.path.join(JS_DIR, "auto-render.min.js"))
    mermaid_js_url = to_file_url(os.path.join(JS_DIR, "mermaid.min.js"))
    base_url = to_file_url(md_dir) + "/"

    parametric_css = generate_parametric_css(config)
    wm_size = parse_watermark_size(config.get("watermark_font_size", 18), 18)
    watermark_css = build_watermark_css(config.get("watermark_text", ""), font_size=wm_size) if config.get("enable_watermark", False) else ""
    watermark_div = '<div class="watermark-overlay"></div>' if config.get("enable_watermark", False) else ""

    first_page_css = ""
    if config.get("suppress_first_page", True):
        first_page_css = """
        @page :first {
            margin-top: 15mm !important;
            margin-bottom: 15mm !important;
        }
        """

    full_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<base href="{base_url}">
<title>{html.escape(doc_title)}</title>
<link rel="stylesheet" href="{katex_css_url}">
<script src="{katex_js_url}"></script>
<script src="{autorender_js_url}"></script>
<script src="{mermaid_js_url}"></script>
<style>
{first_page_css}
{parametric_css}
{watermark_css}
</style>
</head>
<body>
{watermark_div}
{processed_html}

<script>
document.addEventListener("DOMContentLoaded", function() {{
    if (typeof renderMathInElement !== 'undefined') {{
        renderMathInElement(document.body, {{
            delimiters: [
                {{left: "$$", right: "$$", display: true}},
                {{left: "$", right: "$", display: false}}
            ],
            throwOnError: false
        }});
    }}
    if (typeof mermaid !== 'undefined') {{
        mermaid.initialize({{
            startOnLoad: true,
            theme: 'default',
            securityLevel: 'loose',
            flowchart: {{ useMaxWidth: true, htmlLabels: true }}
        }});
    }}
}});
</script>
</body>
</html>
"""
    return full_html


def render_html_to_pdf(html_path, output_pdf_path, config, doc_title=""):
    browser_exe = find_browser()
    if not browser_exe:
        raise RuntimeError("未检测到 Microsoft Edge 浏览器，无法完成 PDF 渲染。")

    from playwright.sync_api import sync_playwright

    enable_header = config.get("enable_header", True)
    header_text = config.get("header_text", "")
    if not header_text and doc_title:
        header_text = doc_title

    enable_footer = config.get("enable_footer", True)

    margin_key = config.get("page_margin", "standard")
    margin_dict = MARGIN_MAP.get(margin_key, MARGIN_MAP["standard"])["dict"]
    is_landscape = (config.get("page_orientation", "portrait") == "landscape")

    m_left = margin_dict.get("left", "16mm")

    header_template = ""
    if enable_header and header_text:
        header_template = f'''<div style="font-size: 8.5pt; width: 100%; text-align: center; color: #718096; font-family: -apple-system, 'Microsoft YaHei', sans-serif; border-bottom: 0.5px solid #e2e8f0; padding-bottom: 3px; margin: 0 {m_left};">{html.escape(header_text)}</div>'''
    else:
        header_template = '<div></div>'

    footer_template = ""
    if enable_footer:
        footer_template = f'''<div style="font-size: 8.5pt; width: 100%; text-align: center; color: #718096; font-family: -apple-system, 'Microsoft YaHei', sans-serif; margin: 0 {m_left};">第 <span class="pageNumber"></span> 页 / 共 <span class="totalPages"></span> 页</div>'''
    else:
        footer_template = '<div></div>'

    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=browser_exe, headless=True)
        page = browser.new_page()
        file_url = "file:///" + html_path.replace("\\", "/")
        page.goto(file_url, wait_until="load")
        page.wait_for_timeout(900)

        pdf_bytes = page.pdf(
            landscape=is_landscape,
            display_header_footer=bool(enable_header or enable_footer),
            header_template=header_template,
            footer_template=footer_template,
            margin=margin_dict,
            print_background=True,
            format='A4'
        )

        with open(output_pdf_path, "wb") as f:
            f.write(pdf_bytes)

        browser.close()

    if not os.path.exists(output_pdf_path) or os.path.getsize(output_pdf_path) == 0:
        raise RuntimeError("PDF 生成失败或文件大小为0。")
    return output_pdf_path


def convert_single_file(md_path, out_pdf_path=None, config=None):
    if config is None:
        config = load_config()

    if not os.path.exists(md_path):
        raise FileNotFoundError(f"文件不存在: {md_path}")

    md_dir = os.path.dirname(os.path.abspath(md_path))
    base_name = os.path.splitext(os.path.basename(md_path))[0]

    if not out_pdf_path:
        out_pdf_path = os.path.join(md_dir, f"{base_name}.pdf")

    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    doc_title = base_name
    for line in md_text.splitlines():
        if line.startswith("# "):
            doc_title = line[2:].strip()
            break

    full_html = markdown_to_html(md_text, md_dir, config, doc_title=doc_title)
    temp_html = os.path.join(tempfile.gettempdir(), f"_md2pdf_{os.getpid()}_{base_name}.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(full_html)

    try:
        return render_html_to_pdf(temp_html, out_pdf_path, config, doc_title=doc_title)
    finally:
        if os.path.exists(temp_html):
            try:
                os.remove(temp_html)
            except Exception:
                pass


def convert_merged_files(md_paths, out_pdf_path, config=None):
    if config is None:
        config = load_config()

    if not md_paths:
        raise ValueError("请至少选择一个 Markdown 文件进行合并。")

    merged_md_blocks = []
    base_dir = os.path.dirname(os.path.abspath(md_paths[0]))
    main_title = os.path.splitext(os.path.basename(out_pdf_path))[0]

    for i, p in enumerate(md_paths):
        if not os.path.exists(p):
            continue
        with open(p, "r", encoding="utf-8") as f:
            content = f.read()
        if i > 0:
            merged_md_blocks.append('\n\n<div style="page-break-before: always;"></div>\n\n')
        merged_md_blocks.append(content)

    combined_text = "".join(merged_md_blocks)
    full_html = markdown_to_html(combined_text, base_dir, config, doc_title=main_title)

    temp_html = os.path.join(tempfile.gettempdir(), f"_md2pdf_merged_{os.getpid()}.html")
    with open(temp_html, "w", encoding="utf-8") as f:
        f.write(full_html)

    try:
        return render_html_to_pdf(temp_html, out_pdf_path, config, doc_title=main_title)
    finally:
        if os.path.exists(temp_html):
            try:
                os.remove(temp_html)
            except Exception:
                pass


# ==============================================================================
# Word (.docx) 原生高保真排版生成引擎
# ==============================================================================

def hex_to_rgb(hex_str, default=(43, 108, 176)):
    if not hex_str:
        return RGBColor(*default)
    s = hex_str.lstrip('#')
    if len(s) == 3:
        s = ''.join([c*2 for c in s])
    if len(s) == 6:
        try:
            return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
        except ValueError:
            pass
    return RGBColor(*default)

def set_run_font(run, font_name="Microsoft YaHei", size_pt=12.0, color_rgb=None, bold=False, italic=False, strike=False, superscript=False):
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    if size_pt:
        run.font.size = Pt(size_pt)
    if color_rgb:
        run.font.color.rgb = color_rgb
    if bold:
        run.font.bold = True
    if italic:
        run.font.italic = True
    if strike:
        run.font.strike = True
    if superscript:
        run.font.superscript = True

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('shd'):
            tcPr.remove(child)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color.lstrip('#'))
    tcPr.append(shd)

def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('tcBorders'):
            tcPr.remove(child)
    tcBorders = OxmlElement('w:tcBorders')
    borders = {'top': top, 'bottom': bottom, 'left': left, 'right': right}
    for edge, border_args in borders.items():
        if border_args is not None:
            el = OxmlElement(f'w:{edge}')
            for k, v in border_args.items():
                el.set(qn(f'w:{k}'), str(v))
            tcBorders.append(el)
    tcPr.append(tcBorders)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('tcMar'):
            tcPr.remove(child)
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_docx_field(p, field_name, font_name="Microsoft YaHei", size_pt=8.5, color_rgb=None):
    run = p.add_run()
    r = run._r
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = field_name
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    r.append(fldChar1)
    r.append(instrText)
    r.append(fldChar2)
    r.append(fldChar3)
    set_run_font(run, font_name=font_name, size_pt=size_pt, color_rgb=color_rgb or RGBColor(0x71, 0x80, 0x96))

def render_html_to_docx(full_html, output_docx_path, config, doc_title="", md_dir=""):
    """将经过预处理的 HTML 及其渲染组件转换为高保真专业 Word (.docx) 文档"""
    has_mermaid = bool(re.search(r'class=["\']mermaid["\']', full_html))
    has_katex = bool("katex" in full_html or "$$" in full_html)

    temp_images = []
    processed_html = full_html

    # 若含有 Mermaid 架构图或 KaTeX 复杂公式，调度 Chromium 进行 2x Retina 截图
    if has_mermaid or has_katex:
        browser_exe = find_browser()
        if browser_exe:
            tmp_html = os.path.join(tempfile.gettempdir(), f"_md2docx_{os.getpid()}_{abs(hash(doc_title)) % 100000}.html")
            try:
                from playwright.sync_api import sync_playwright
                with open(tmp_html, "w", encoding="utf-8") as f:
                    f.write(full_html)

                with sync_playwright() as p:
                    browser = p.chromium.launch(executable_path=browser_exe, headless=True)
                    page = browser.new_page(device_scale_factor=2)
                    page.goto("file:///" + tmp_html.replace("\\", "/"), wait_until="load")
                    page.wait_for_timeout(900)

                    # 截图所有 Mermaid
                    mm_locs = page.locator(".mermaid").all()
                    mm_list = []
                    for i, el in enumerate(mm_locs):
                        img_p = os.path.join(tempfile.gettempdir(), f"_mm_docx_{os.getpid()}_{i}.png").replace("\\", "/")
                        el.screenshot(path=img_p)
                        mm_list.append(img_p)
                        temp_images.append(img_p)

                    # 截图所有 KaTeX display 块级公式
                    kd_locs = page.locator(".katex-display").all()
                    kd_list = []
                    for i, el in enumerate(kd_locs):
                        img_p = os.path.join(tempfile.gettempdir(), f"_kd_docx_{os.getpid()}_{i}.png").replace("\\", "/")
                        el.screenshot(path=img_p)
                        kd_list.append(img_p)
                        temp_images.append(img_p)

                    # 在浏览器 DOM 中替换为对应图片与行内公式
                    page.evaluate('''data => {
                        document.querySelectorAll('.mermaid').forEach((el, i) => {
                            if (data.mm[i]) {
                                const img = document.createElement('img');
                                img.src = data.mm[i];
                                img.className = 'rendered-diagram';
                                el.parentNode.replaceChild(img, el);
                            }
                        });
                        document.querySelectorAll('.katex-display').forEach((el, i) => {
                            if (data.kd[i]) {
                                const img = document.createElement('img');
                                img.src = data.kd[i];
                                img.className = 'rendered-math';
                                el.parentNode.replaceChild(img, el);
                            }
                        });
                        document.querySelectorAll('span.katex:not(.katex-display span.katex)').forEach(el => {
                            const ann = el.querySelector('annotation');
                            if (ann) {
                                const textNode = document.createTextNode('$' + ann.textContent + '$');
                                el.parentNode.replaceChild(textNode, el);
                            }
                        });
                    }''', {'mm': mm_list, 'kd': kd_list})

                    processed_html = page.content()
                    browser.close()
            except Exception as e:
                print(f"Playwright 截图预处理异常: {e}")
            finally:
                if os.path.exists(tmp_html):
                    try:
                        os.remove(tmp_html)
                    except Exception:
                        pass

    try:
        soup = BeautifulSoup(processed_html, "html.parser")
        body = soup.find("body") or soup

        doc = Document()

        # 页面尺寸与方向配置
        section = doc.sections[0]
        is_landscape = (config.get("page_orientation", "portrait") == "landscape")
        if is_landscape:
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = Inches(11.69)
            section.page_height = Inches(8.27)
        else:
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)

        margin_key = config.get("page_margin", "standard")
        if margin_key == "compact":
            section.top_margin = Inches(0.47)
            section.bottom_margin = Inches(0.47)
            section.left_margin = Inches(0.47)
            section.right_margin = Inches(0.47)
        elif margin_key == "spacious":
            section.top_margin = Inches(0.98)
            section.bottom_margin = Inches(0.98)
            section.left_margin = Inches(0.87)
            section.right_margin = Inches(0.87)
        elif margin_key == "gov":
            section.top_margin = Inches(1.46)
            section.bottom_margin = Inches(1.38)
            section.left_margin = Inches(1.10)
            section.right_margin = Inches(1.02)
        else:
            section.top_margin = Inches(0.79)
            section.bottom_margin = Inches(0.79)
            section.left_margin = Inches(0.63)
            section.right_margin = Inches(0.63)

        content_width_in = (section.page_width.inches - section.left_margin.inches - section.right_margin.inches)

        font_key = config.get("font_family_key", "微软雅黑 (推荐·现代无衬线)")
        font_name = "Microsoft YaHei"
        if "仿宋" in font_key:
            font_name = "FangSong"
        elif "宋体" in font_key:
            font_name = "SimSun"
        elif "楷体" in font_key:
            font_name = "KaiTi"
        elif "Consolas" in font_key:
            font_name = "Consolas"

        base_font_size = parse_font_size(config.get("font_size", "小四 (12 pt)"), 12.0)
        line_height = parse_line_height(config.get("line_height", "1.65 倍 (推荐)"), 1.65)
        word_line_spacing = max(1.15, min(1.5, line_height * 0.82))
        para_spacing = parse_para_spacing(config.get("para_spacing", "标准 (8 pt)"), 8.0)
        accent_hex = config.get("accent_color", "#2b6cb0")
        accent_rgb = hex_to_rgb(accent_hex, (43, 108, 176))
        text_align_justify = config.get("text_align_justify", True)
        indent_first_line = config.get("indent_first_line", False)
        h1_align = config.get("h1_align", "center")

        # 页眉页脚
        if config.get("enable_header", True):
            h_text = config.get("header_text", "").strip() or doc_title
            if h_text:
                header = section.header
                p_h = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
                p_h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                r_h = p_h.add_run(h_text)
                set_run_font(r_h, font_name=font_name, size_pt=8.5, color_rgb=RGBColor(0x71, 0x80, 0x96))

        if config.get("enable_footer", True):
            footer = section.footer
            p_f = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r1 = p_f.add_run("第 ")
            set_run_font(r1, font_name=font_name, size_pt=8.5, color_rgb=RGBColor(0x71, 0x80, 0x96))
            add_docx_field(p_f, "PAGE", font_name=font_name, size_pt=8.5)
            r2 = p_f.add_run(" 页 / 共 ")
            set_run_font(r2, font_name=font_name, size_pt=8.5, color_rgb=RGBColor(0x71, 0x80, 0x96))
            add_docx_field(p_f, "NUMPAGES", font_name=font_name, size_pt=8.5)
            r3 = p_f.add_run(" 页")
            set_run_font(r3, font_name=font_name, size_pt=8.5, color_rgb=RGBColor(0x71, 0x80, 0x96))

        if config.get("suppress_first_page", True):
            section.different_first_page_header_footer = True

        def process_inline(paragraph, node, inherit_bold=False, inherit_italic=False, inherit_color=None, inherit_size=None):
            if isinstance(node, NavigableString):
                text = str(node)
                if text:
                    run = paragraph.add_run(text)
                    set_run_font(
                        run,
                        font_name=font_name,
                        size_pt=inherit_size or base_font_size,
                        color_rgb=inherit_color or RGBColor(0x2D, 0x37, 0x48),
                        bold=inherit_bold,
                        italic=inherit_italic
                    )
                return

            if not isinstance(node, Tag):
                return

            tag_name = node.name.lower()

            if tag_name in ["strong", "b"]:
                for child in node.children:
                    process_inline(paragraph, child, inherit_bold=True, inherit_italic=inherit_italic, inherit_color=inherit_color, inherit_size=inherit_size)
            elif tag_name in ["em", "i"]:
                for child in node.children:
                    process_inline(paragraph, child, inherit_bold=inherit_bold, inherit_italic=True, inherit_color=inherit_color, inherit_size=inherit_size)
            elif tag_name == "code":
                code_text = node.get_text()
                run = paragraph.add_run(f" {code_text} ")
                set_run_font(
                    run,
                    font_name="Consolas",
                    size_pt=round((inherit_size or base_font_size) * 0.90, 1),
                    color_rgb=RGBColor(0x80, 0x5A, 0xD5),
                    bold=False,
                    italic=False
                )
            elif tag_name in ["del", "s", "strike"]:
                for child in node.children:
                    text = child.get_text() if isinstance(child, Tag) else str(child)
                    run = paragraph.add_run(text)
                    set_run_font(
                        run,
                        font_name=font_name,
                        size_pt=inherit_size or base_font_size,
                        color_rgb=RGBColor(0xA0, 0xAE, 0xC0),
                        bold=inherit_bold,
                        italic=inherit_italic,
                        strike=True
                    )
            elif tag_name == "a":
                link_text = node.get_text()
                run = paragraph.add_run(link_text)
                set_run_font(
                    run,
                    font_name=font_name,
                    size_pt=inherit_size or base_font_size,
                    color_rgb=accent_rgb,
                    bold=inherit_bold,
                    italic=inherit_italic
                )
                run.font.underline = True
            elif tag_name == "sup":
                sup_text = node.get_text()
                run = paragraph.add_run(sup_text)
                set_run_font(
                    run,
                    font_name=font_name,
                    size_pt=round((inherit_size or base_font_size) * 0.8, 1),
                    color_rgb=RGBColor(0x71, 0x80, 0x96),
                    superscript=True
                )
            elif tag_name == "input" and node.get("type") == "checkbox":
                checked = node.has_attr("checked")
                box_char = "☑ " if checked else "☐ "
                run = paragraph.add_run(box_char)
                set_run_font(
                    run,
                    font_name="Segoe UI Symbol" if os.name == "nt" else font_name,
                    size_pt=inherit_size or base_font_size,
                    color_rgb=accent_rgb if checked else RGBColor(0xA0, 0xAE, 0xC0),
                    bold=True
                )
            elif tag_name == "img":
                src = node.get("src", "")
                if src:
                    img_path = None
                    if os.path.isabs(src) and os.path.exists(src):
                        img_path = src
                    elif md_dir and os.path.exists(os.path.join(md_dir, src)):
                        img_path = os.path.join(md_dir, src)

                    if img_path and os.path.exists(img_path):
                        try:
                            run = paragraph.add_run()
                            max_w = min(content_width_in, 6.0)
                            run.add_picture(img_path, width=Inches(max_w))
                        except Exception as e:
                            print(f"插入行内图片失败: {e}")
            elif tag_name == "br":
                paragraph.add_run().add_break()
            else:
                for child in node.children:
                    process_inline(paragraph, child, inherit_bold=inherit_bold, inherit_italic=inherit_italic, inherit_color=inherit_color, inherit_size=inherit_size)

        for elem in body.children:
            if isinstance(elem, NavigableString):
                if not elem.strip():
                    continue
                p = doc.add_paragraph()
                p.paragraph_format.line_spacing = word_line_spacing
                p.paragraph_format.space_after = Pt(para_spacing)
                r = p.add_run(str(elem).strip())
                set_run_font(r, font_name=font_name, size_pt=base_font_size)
                continue

            if not isinstance(elem, Tag):
                continue

            tag = elem.name.lower()

            if tag in ["script", "style", "meta", "link", "title"]:
                continue

            # 分页符支持 (长报告章节合并)
            if tag == "div" and ("docx-page-break" in elem.get("class", []) or "page-break-before: always" in elem.get("style", "")):
                doc.add_page_break()
                continue

            # 1. 标题 H1 ~ H6
            if tag in ["h1", "h2", "h3", "h4", "h5", "h6"]:
                level = int(tag[1])
                p = doc.add_paragraph()
                p.paragraph_format.keep_with_next = True

                if level == 1:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(14)
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if h1_align == "center" else WD_ALIGN_PARAGRAPH.LEFT
                    h_size = round(base_font_size * 1.55, 1)
                    for child in elem.children:
                        process_inline(p, child, inherit_bold=True, inherit_color=RGBColor(0x1A, 0x20, 0x2C), inherit_size=h_size)
                elif level == 2:
                    p.paragraph_format.space_before = Pt(16)
                    p.paragraph_format.space_after = Pt(6)
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    h_size = round(base_font_size * 1.25, 1)
                    for child in elem.children:
                        process_inline(p, child, inherit_bold=True, inherit_color=accent_rgb, inherit_size=h_size)
                elif level == 3:
                    p.paragraph_format.space_before = Pt(12)
                    p.paragraph_format.space_after = Pt(4)
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    h_size = round(base_font_size * 1.10, 1)
                    for child in elem.children:
                        process_inline(p, child, inherit_bold=True, inherit_color=RGBColor(0x2D, 0x37, 0x48), inherit_size=h_size)
                else:
                    p.paragraph_format.space_before = Pt(8)
                    p.paragraph_format.space_after = Pt(3)
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    h_size = base_font_size
                    for child in elem.children:
                        process_inline(p, child, inherit_bold=True, inherit_color=RGBColor(0x4A, 0x55, 0x68), inherit_size=h_size)

            # 2. 独立图片或图表/公式
            elif tag == "img" or (elem.find("img", class_=["rendered-diagram", "rendered-math"])):
                rendered_img = elem if tag == "img" else elem.find("img", class_=["rendered-diagram", "rendered-math"])
                src = rendered_img.get("src", "")
                img_path = None
                if os.path.isabs(src) and os.path.exists(src):
                    img_path = src
                elif md_dir and os.path.exists(os.path.join(md_dir, src)):
                    img_path = os.path.join(md_dir, src)

                if img_path and os.path.exists(img_path):
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = Pt(8)
                    p.paragraph_format.space_after = Pt(12)
                    try:
                        is_math = "rendered-math" in rendered_img.get("class", [])
                        max_w = min(content_width_in, 5.8) if not is_math else min(content_width_in * 0.7, 4.5)
                        p.add_run().add_picture(img_path, width=Inches(max_w))
                    except Exception as e:
                        print(f"插入图片异常: {e}")

            # 3. 普通段落 <p>
            elif tag == "p":
                p = doc.add_paragraph()
                p.paragraph_format.line_spacing = word_line_spacing
                p.paragraph_format.space_after = Pt(para_spacing)
                if indent_first_line:
                    p.paragraph_format.first_line_indent = Pt(base_font_size * 2)
                if text_align_justify:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                for child in elem.children:
                    process_inline(p, child)

            # 4. 引用块 <blockquote> (专业科技蓝条卡片)
            elif tag == "blockquote":
                tbl = doc.add_table(rows=1, cols=1)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False
                cell = tbl.cell(0, 0)
                cell.width = Inches(content_width_in)

                set_cell_background(cell, "F7FAFC")
                set_cell_borders(cell,
                    left={'val': 'single', 'sz': '24', 'space': '0', 'color': accent_hex.lstrip('#')},
                    top={'val': 'none'},
                    bottom={'val': 'none'},
                    right={'val': 'none'}
                )
                set_cell_margins(cell, top=120, bottom=120, left=180, right=180)

                first_p = True
                for b_child in elem.children:
                    if isinstance(b_child, Tag) and b_child.name.lower() == "p":
                        p = cell.paragraphs[0] if first_p else cell.add_paragraph()
                        first_p = False
                        p.paragraph_format.line_spacing = word_line_spacing
                        p.paragraph_format.space_after = Pt(4)
                        for c in b_child.children:
                            process_inline(p, c, inherit_italic=False, inherit_color=RGBColor(0x4A, 0x55, 0x68), inherit_size=round(base_font_size * 0.95, 1))
                    elif isinstance(b_child, NavigableString) and b_child.strip():
                        p = cell.paragraphs[0] if first_p else cell.add_paragraph()
                        first_p = False
                        p.paragraph_format.line_spacing = word_line_spacing
                        p.paragraph_format.space_after = Pt(4)
                        r = p.add_run(str(b_child).strip())
                        set_run_font(r, font_name=font_name, size_pt=round(base_font_size * 0.95, 1), color_rgb=RGBColor(0x4A, 0x55, 0x68))

                p_spacer = doc.add_paragraph()
                p_spacer.paragraph_format.space_before = Pt(0)
                p_spacer.paragraph_format.space_after = Pt(4)

            # 5. 代码块 <pre><code> (1x1 优雅代码卡片)
            elif tag == "pre":
                code_elem = elem.find("code")
                code_text = code_elem.get_text() if code_elem else elem.get_text()

                tbl = doc.add_table(rows=1, cols=1)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False
                cell = tbl.cell(0, 0)
                cell.width = Inches(content_width_in)

                set_cell_background(cell, "F8FAFC")
                set_cell_borders(cell,
                    left={'val': 'single', 'sz': '24', 'space': '0', 'color': accent_hex.lstrip('#')},
                    top={'val': 'single', 'sz': '4', 'space': '0', 'color': 'E2E8F0'},
                    bottom={'val': 'single', 'sz': '4', 'space': '0', 'color': 'E2E8F0'},
                    right={'val': 'single', 'sz': '4', 'space': '0', 'color': 'E2E8F0'}
                )
                set_cell_margins(cell, top=140, bottom=140, left=180, right=180)

                code_lines = code_text.splitlines()
                if not code_lines:
                    code_lines = [""]
                for i, cl in enumerate(code_lines):
                    p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    r = p.add_run(cl if cl else " ")
                    set_run_font(r, font_name="Consolas", size_pt=9.5, color_rgb=RGBColor(0x2D, 0x37, 0x48))

                p_spacer = doc.add_paragraph()
                p_spacer.paragraph_format.space_before = Pt(0)
                p_spacer.paragraph_format.space_after = Pt(6)

            # 6. GFM 表格 <table> (淡蓝表头、精细边框、斑马纹)
            elif tag == "table":
                rows_data = []
                for tr in elem.find_all("tr"):
                    cells = tr.find_all(["th", "td"])
                    rows_data.append((tr, cells))

                if rows_data:
                    num_rows = len(rows_data)
                    num_cols = max(len(c) for _, c in rows_data)
                    tbl = doc.add_table(rows=num_rows, cols=num_cols)
                    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                    tbl.autofit = False

                    col_w = content_width_in / max(1, num_cols)

                    for r_idx, (tr, cells) in enumerate(rows_data):
                        row = tbl.rows[r_idx]
                        is_header = bool(tr.find("th")) or (r_idx == 0)

                        trPr = row._tr.get_or_add_trPr()
                        cantSplit = OxmlElement('w:cantSplit')
                        trPr.append(cantSplit)
                        if is_header:
                            tblHeader = OxmlElement('w:tblHeader')
                            trPr.append(tblHeader)

                        for c_idx, cell_elem in enumerate(cells):
                            if c_idx >= num_cols:
                                break
                            cell = row.cells[c_idx]
                            cell.width = Inches(col_w)
                            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                            set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
                            set_cell_borders(cell,
                                top={'val': 'single', 'sz': '4', 'space': '0', 'color': 'CBD5E0'},
                                bottom={'val': 'single', 'sz': '4', 'space': '0', 'color': 'CBD5E0'},
                                left={'val': 'single', 'sz': '4', 'space': '0', 'color': 'CBD5E0'},
                                right={'val': 'single', 'sz': '4', 'space': '0', 'color': 'CBD5E0'}
                            )

                            p = cell.paragraphs[0]
                            p.paragraph_format.line_spacing = 1.2
                            p.paragraph_format.space_before = Pt(0)
                            p.paragraph_format.space_after = Pt(0)

                            align_attr = cell_elem.get("align", "").lower()
                            if "center" in align_attr:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            elif "right" in align_attr:
                                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            else:
                                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                            if is_header:
                                set_cell_background(cell, "EBF8FF")
                                for ch in cell_elem.children:
                                    process_inline(p, ch, inherit_bold=True, inherit_color=accent_rgb, inherit_size=round(base_font_size * 0.92, 1))
                            else:
                                if r_idx % 2 == 0:
                                    set_cell_background(cell, "F7FAFC")
                                else:
                                    set_cell_background(cell, "FFFFFF")
                                for ch in cell_elem.children:
                                    process_inline(p, ch, inherit_bold=False, inherit_color=RGBColor(0x2D, 0x37, 0x48), inherit_size=round(base_font_size * 0.92, 1))

                    p_spacer = doc.add_paragraph()
                    p_spacer.paragraph_format.space_before = Pt(0)
                    p_spacer.paragraph_format.space_after = Pt(6)

            # 7. 列表 <ul> 与 <ol> (包括任务清单)
            elif tag in ["ul", "ol"]:
                is_ordered = (tag == "ol")
                item_idx = 1
                for li in elem.find_all("li", recursive=False):
                    p = doc.add_paragraph()
                    p.paragraph_format.line_spacing = word_line_spacing
                    p.paragraph_format.space_after = Pt(3)
                    p.paragraph_format.left_indent = Inches(0.28)

                    task_checkbox = li.find("input", class_="task-list-checkbox")
                    if task_checkbox:
                        checked = task_checkbox.has_attr("checked")
                        box_char = "☑ " if checked else "☐ "
                        r_box = p.add_run(box_char)
                        set_run_font(
                            r_box,
                            font_name="Segoe UI Symbol" if os.name == "nt" else font_name,
                            size_pt=base_font_size,
                            color_rgb=accent_rgb if checked else RGBColor(0xA0, 0xAE, 0xC0),
                            bold=True
                        )
                        for ch in li.children:
                            if ch != task_checkbox:
                                process_inline(p, ch)
                    elif is_ordered:
                        r_num = p.add_run(f"{item_idx}. ")
                        set_run_font(r_num, font_name=font_name, size_pt=base_font_size, bold=True, color_rgb=accent_rgb)
                        for ch in li.children:
                            process_inline(p, ch)
                        item_idx += 1
                    else:
                        r_bullet = p.add_run("• ")
                        set_run_font(r_bullet, font_name=font_name, size_pt=base_font_size, bold=True, color_rgb=accent_rgb)
                        for ch in li.children:
                            process_inline(p, ch)

            # 8. 水平分割线 <hr>
            elif tag == "hr":
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(12)
                r = p.add_run("―" * 45)
                set_run_font(r, font_name="Arial", size_pt=7.5, color_rgb=RGBColor(0xE2, 0xE8, 0xF0))

            # 9. 目录 [TOC]
            elif tag == "div" and "toc" in elem.get("class", []):
                p_toc_title = doc.add_paragraph()
                p_toc_title.paragraph_format.space_before = Pt(4)
                p_toc_title.paragraph_format.space_after = Pt(6)
                r_title = p_toc_title.add_run("目录索引")
                set_run_font(r_title, font_name=font_name, size_pt=round(base_font_size * 1.1, 1), bold=True, color_rgb=accent_rgb)

                for li in elem.find_all("li"):
                    p = doc.add_paragraph()
                    p.paragraph_format.line_spacing = 1.2
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.left_indent = Inches(0.20)
                    r_dot = p.add_run("▸ ")
                    set_run_font(r_dot, font_name=font_name, size_pt=base_font_size * 0.9, color_rgb=accent_rgb)
                    for ch in li.children:
                        process_inline(p, ch, inherit_size=round(base_font_size * 0.95, 1))

            # 10. 脚注区域
            elif tag == "div" and "footnote" in elem.get("class", []):
                p_fn_line = doc.add_paragraph()
                p_fn_line.paragraph_format.space_before = Pt(14)
                p_fn_line.paragraph_format.space_after = Pt(4)
                r_line = p_fn_line.add_run("―" * 20)
                set_run_font(r_line, font_name="Arial", size_pt=7.5, color_rgb=RGBColor(0xCB, 0xD5, 0xE0))

                for li in elem.find_all("li"):
                    p = doc.add_paragraph()
                    p.paragraph_format.line_spacing = 1.2
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.left_indent = Inches(0.20)
                    for ch in li.children:
                        process_inline(p, ch, inherit_size=round(base_font_size * 0.85, 1), inherit_color=RGBColor(0x71, 0x80, 0x96))

        doc.save(output_docx_path)
        return output_docx_path

    finally:
        for p_img in temp_images:
            if os.path.exists(p_img):
                try:
                    os.remove(p_img)
                except Exception:
                    pass


def convert_single_to_docx(md_path, out_docx_path=None, config=None):
    if config is None:
        config = load_config()

    if not os.path.exists(md_path):
        raise FileNotFoundError(f"文件不存在: {md_path}")

    md_dir = os.path.dirname(os.path.abspath(md_path))
    base_name = os.path.splitext(os.path.basename(md_path))[0]

    if not out_docx_path:
        out_docx_path = os.path.join(md_dir, f"{base_name}.docx")

    with open(md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    doc_title = base_name
    for line in md_text.splitlines():
        if line.startswith("# "):
            doc_title = line[2:].strip()
            break

    full_html = markdown_to_html(md_text, md_dir, config, doc_title=doc_title)
    return render_html_to_docx(full_html, out_docx_path, config, doc_title=doc_title, md_dir=md_dir)


def convert_merged_to_docx(md_paths, out_docx_path, config=None):
    if config is None:
        config = load_config()

    if not md_paths:
        raise ValueError("请至少选择一个 Markdown 文件进行合并。")

    merged_md_blocks = []
    base_dir = os.path.dirname(os.path.abspath(md_paths[0]))
    main_title = os.path.splitext(os.path.basename(out_docx_path))[0]

    for i, p in enumerate(md_paths):
        if not os.path.exists(p):
            continue
        with open(p, "r", encoding="utf-8") as f:
            content = f.read()
        if i > 0:
            merged_md_blocks.append('\n\n<div class="docx-page-break" style="page-break-before: always;"></div>\n\n')
        merged_md_blocks.append(content)

    combined_text = "".join(merged_md_blocks)
    full_html = markdown_to_html(combined_text, base_dir, config, doc_title=main_title)
    return render_html_to_docx(full_html, out_docx_path, config, doc_title=main_title, md_dir=base_dir)


# ==============================================================================
# 专业级 CustomTkinter 左右双栏工作台 (集成官方 TkDND 原生拖拽)
# ==============================================================================

def get_screen_metrics():
    try:
        import ctypes
        user32 = ctypes.windll.user32
        sw = user32.GetSystemMetrics(0)
        sh = user32.GetSystemMetrics(1)
        if sw > 0 and sh > 0:
            return sw, sh
    except Exception:
        pass
    return 1920, 1080


class ModernMD2PDFApp(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        # 初始化 TkDND 原生扩展 (杜绝任何底层 Windows 句柄截断 Access Violation)
        self.TkdndVersion = TkinterDnD._require(self)

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title("MD2PDF Pro · 高清文档排版与双格式导出中心 (PDF / Word)")

        # 屏幕分辨率动态自适应
        sw, sh = get_screen_metrics()
        win_w = min(1220, max(980, int(sw * 0.65)))
        win_h = min(860, max(680, int(sh * 0.74)))
        pos_x = max(0, (sw - win_w) // 2)
        pos_y = max(0, (sh - win_h) // 2)

        self.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        self.minsize(960, 640)
        self.win_w = win_w

        self.config = load_config()
        self.file_list = []
        self.last_output_pdf = None

        self.setup_ui()
        self.load_settings_into_ui()
        self.setup_drag_and_drop()

    def setup_ui(self):
        # 1. 顶部经典科技蓝 Banner
        top_bar = ctk.CTkFrame(self, height=44, corner_radius=0, fg_color="#2b6cb0")
        top_bar.pack(fill="x", side="top")

        lbl_app = ctk.CTkLabel(
            top_bar,
            text="MD2PDF Pro · 专业级 Markdown 转 PDF / Word 工作台",
            font=ctk.CTkFont(family="Microsoft YaHei", size=14, weight="bold"),
            text_color="#ffffff"
        )
        lbl_app.pack(side="left", padx=18, pady=8)

        lbl_sub = ctk.CTkLabel(
            top_bar,
            text="双格式原生导出 (PDF + Word) | 离线图表与公式 | 批处理与长报告合并",
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            text_color="#bee3f8"
        )
        lbl_sub.pack(side="right", padx=18, pady=8)

        # 2. 主工作区：双栏结构 (响应式占比)
        workspace = ctk.CTkFrame(self, fg_color="#f8fafc", corner_radius=0)
        workspace.pack(fill="both", expand=True, padx=12, pady=10)

        # ----------------------------------------------------------------------
        # 左栏：文件队列与任务控制 (约 35% 紧凑占比，不再过度挤占排版区)
        # ----------------------------------------------------------------------
        left_width = min(460, max(400, int(getattr(self, "win_w", 1220) * 0.36)))
        left_col = ctk.CTkFrame(workspace, width=left_width, fg_color="#ffffff", corner_radius=8, border_width=1, border_color="#e2e8f0")
        left_col.pack(side="left", fill="y", padx=(0, 10), pady=0)
        left_col.pack_propagate(False)

        # 待转换文档清单与拖拽投放区 (统一合并为一个开阔大区域)
        list_card = ctk.CTkFrame(left_col, fg_color="#ffffff", corner_radius=8, border_width=1, border_color="#e2e8f0")
        list_card.pack(fill="both", expand=True, padx=10, pady=(10, 6))

        # 卡片顶部标题栏
        card_header = ctk.CTkFrame(list_card, fg_color="transparent")
        card_header.pack(fill="x", padx=12, pady=(10, 6))

        self.lbl_list_title = ctk.CTkLabel(
            card_header,
            text="待转换文档与拖拽区",
            font=ctk.CTkFont(family="Microsoft YaHei", size=13, weight="bold"),
            text_color="#2d3748"
        )
        self.lbl_list_title.pack(side="left")

        self.lbl_file_count = ctk.CTkLabel(
            card_header,
            text="(0 篇)",
            font=ctk.CTkFont(family="Microsoft YaHei", size=12),
            text_color="#718096"
        )
        self.lbl_file_count.pack(side="left", padx=6)

        lbl_dnd_hint = ctk.CTkLabel(
            card_header,
            text="支持拖拽追加 / 调序",
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            text_color="#3182ce"
        )
        lbl_dnd_hint.pack(side="right")

        # 核心内容区：在空状态投放卡片与文档列表间智能切换，享受全部开阔纵向空间
        self.drop_container = ctk.CTkFrame(list_card, fg_color="transparent")
        self.drop_container.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        # 1. 空状态大尺寸拖拽投放区 (当没有文件时呈现)
        self.empty_drop_view = ctk.CTkFrame(
            self.drop_container,
            fg_color="#f8fafc",
            corner_radius=8,
            border_width=2,
            border_color="#cbd5e0",
            cursor="hand2"
        )

        center_box = ctk.CTkFrame(self.empty_drop_view, fg_color="transparent")
        center_box.place(relx=0.5, rely=0.5, anchor="center")

        lbl_empty_icon = ctk.CTkLabel(
            center_box,
            text="拖拽 Markdown 文件或文件夹到此区域",
            font=ctk.CTkFont(family="Microsoft YaHei", size=14, weight="bold"),
            text_color="#2b6cb0"
        )
        lbl_empty_icon.pack(pady=(0, 8))

        lbl_empty_sub = ctk.CTkLabel(
            center_box,
            text="支持单个、多个 .md 文件批量拖入\n或拖入包含文档的文件夹自动扫描\n\n也可点击此区域或下方按钮添加",
            font=ctk.CTkFont(family="Microsoft YaHei", size=12),
            text_color="#718096",
            justify="center"
        )
        lbl_empty_sub.pack()

        # 点击空状态区也可以直接选择文件
        for w in [self.empty_drop_view, center_box, lbl_empty_icon, lbl_empty_sub]:
            w.bind("<Button-1>", lambda e: self.add_files())

        # 2. 已有文档时的列表呈现区 (根据 DPI 缩放比智能计算大像素字号，彻底告别 96 DPI 字体偏小问题)
        self.list_view = ctk.CTkFrame(self.drop_container, fg_color="transparent")

        # 动态自适应 DPI 像素级放大：150% 下为 -24px (行高 31px)，100% 下为 -18px
        scale = ctk.ScalingTracker.get_widget_scaling(self)
        self.list_font_px = -max(18, int(round(16 * scale)))

        self.listbox = tk.Listbox(
            self.list_view,
            selectmode="extended",
            font=("Microsoft YaHei", self.list_font_px),
            bg="#ffffff",
            fg="#2d3748",
            selectbackground="#ebf8ff",
            selectforeground="#2b6cb0",
            activestyle="none",
            relief="solid",
            borderwidth=1,
            highlightthickness=0
        )

        self.sb = ctk.CTkScrollbar(self.list_view, command=self.listbox.yview, width=14)
        self.sb_x = ctk.CTkScrollbar(self.list_view, orientation="horizontal", command=self.listbox.xview, height=10)

        self.sb.pack(side="right", fill="y")
        self.sb_x.pack(side="bottom", fill="x")
        self.listbox.pack(side="left", fill="both", expand=True)

        self.listbox.config(yscrollcommand=self.sb.set, xscrollcommand=self.sb_x.set)

        # 鼠标与键盘交互增强：单击查看详细路径、双击直接打开文件、Shift+滚轮水平滚动
        self.listbox.bind("<<ListboxSelect>>", self.on_listbox_select)
        self.listbox.bind("<Double-Button-1>", self.on_listbox_double_click)
        self.listbox.bind("<Shift-MouseWheel>", lambda e: self.listbox.xview_scroll(int(-1 * (e.delta / 120)), "units"))

        # 默认呈现空状态投放卡片
        self.empty_drop_view.pack(fill="both", expand=True)

        # 列表操作工具栏
        tool_row = ctk.CTkFrame(list_card, fg_color="transparent")
        tool_row.pack(fill="x", padx=10, pady=(0, 10))

        def make_btn(parent, text, cmd, text_color="#2d3748", width=50):
            return ctk.CTkButton(
                parent, text=text, width=width, height=30,
                font=ctk.CTkFont(family="Microsoft YaHei", size=11),
                fg_color="#edf2f7", text_color=text_color,
                border_width=1, border_color="#e2e8f0",
                hover_color="#e2e8f0", command=cmd
            )

        make_btn(tool_row, "添加文件...", self.add_files, width=76).pack(side="left", padx=(0, 3))
        make_btn(tool_row, "添加文件夹...", self.add_folder, width=86).pack(side="left", padx=(0, 3))
        make_btn(tool_row, "上移", self.move_up, width=44).pack(side="left", padx=(0, 3))
        make_btn(tool_row, "下移", self.move_down, width=44).pack(side="left", padx=(0, 3))
        make_btn(tool_row, "移除", self.remove_selected, text_color="#c53030", width=44).pack(side="right", padx=(3, 0))
        make_btn(tool_row, "清空", self.clear_files, text_color="#718096", width=44).pack(side="right", padx=(3, 0))

        # 导出模式与格式配置
        mode_box = ctk.CTkFrame(left_col, fg_color="#f8fafc", corner_radius=8, border_width=1, border_color="#e2e8f0")
        mode_box.pack(fill="x", padx=10, pady=(0, 6))

        lbl_mode = ctk.CTkLabel(mode_box, text="转换模式与保存路径：", font=ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold"), text_color="#2d3748")
        lbl_mode.pack(anchor="w", padx=10, pady=(6, 4))

        self.seg_mode = ctk.CTkSegmentedButton(
            mode_box,
            values=["分别批量转换 (独立文件)", "按顺序合并为单一长报告"],
            command=self.on_mode_change,
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            height=32
        )
        self.seg_mode.pack(fill="x", padx=10, pady=(0, 6))

        lbl_format = ctk.CTkLabel(mode_box, text="导出格式选择：", font=ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold"), text_color="#2d3748")
        lbl_format.pack(anchor="w", padx=10, pady=(2, 4))

        self.seg_format = ctk.CTkSegmentedButton(
            mode_box,
            values=["导出 PDF", "导出 Word (.docx)", "同时导出 (PDF + Word)"],
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            height=32
        )
        self.seg_format.pack(fill="x", padx=10, pady=(0, 6))

        out_row = ctk.CTkFrame(mode_box, fg_color="transparent")
        out_row.pack(fill="x", padx=10, pady=(0, 8))
        self.entry_out_dir = ctk.CTkEntry(out_row, placeholder_text="默认保存在原文件所在目录（留空即可）", font=ctk.CTkFont(family="Microsoft YaHei", size=11), height=32)
        self.entry_out_dir.pack(side="left", fill="x", expand=True, padx=(0, 6))
        make_btn(out_row, "选择目录...", self.choose_out_dir, width=72).pack(side="right")

        # 底部状态栏与主行动按钮
        bottom_box = ctk.CTkFrame(left_col, fg_color="#ffffff")
        bottom_box.pack(fill="x", padx=10, pady=(0, 10), side="bottom")

        self.status_lbl = ctk.CTkLabel(bottom_box, text="状态：就绪。可直接将 .md 文件拖入窗口。", font=ctk.CTkFont(family="Microsoft YaHei", size=12), text_color="#718096", anchor="w")
        self.status_lbl.pack(fill="x", pady=(0, 6))

        action_row = ctk.CTkFrame(bottom_box, fg_color="transparent")
        action_row.pack(fill="x")

        self.btn_convert = ctk.CTkButton(
            action_row,
            text="开始转换",
            font=ctk.CTkFont(family="Microsoft YaHei", size=14, weight="bold"),
            height=42,
            fg_color="#2b6cb0",
            hover_color="#2c5282",
            command=self.start_conversion_thread
        )
        self.btn_convert.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.btn_open_dir = ctk.CTkButton(
            action_row,
            text="打开输出目录",
            font=ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold"),
            height=42,
            width=115,
            fg_color="#2f855a",
            hover_color="#276749",
            command=self.open_output_dir
        )
        self.btn_open_dir.pack(side="right")

        # ----------------------------------------------------------------------
        # 右栏：参数化排版设置面板 (约 65% 宽裕空间，清晰阅读)
        # ----------------------------------------------------------------------
        right_scroll = ctk.CTkScrollableFrame(workspace, fg_color="#ffffff", corner_radius=8, border_width=1, border_color="#e2e8f0")
        right_scroll.pack(side="right", fill="both", expand=True, padx=(0, 0), pady=0)

        # 快捷预设栏
        preset_bar = ctk.CTkFrame(right_scroll, fg_color="#f8fafc", corner_radius=8, border_width=1, border_color="#e2e8f0")
        preset_bar.pack(fill="x", padx=6, pady=(4, 8))

        lbl_p = ctk.CTkLabel(preset_bar, text="预设风格方案：", font=ctk.CTkFont(family="Microsoft YaHei", size=11, weight="bold"), text_color="#2d3748")
        lbl_p.pack(side="left", padx=10, pady=8)

        self.opt_preset = ctk.CTkOptionMenu(
            preset_bar,
            values=["政企科技蓝 (默认推荐)", "严谨公文风 (仿宋·红黑·首行缩进)", "GitHub 经典清爽 (开发者开源风)"],
            command=self.on_preset_selected,
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            height=28
        )
        self.opt_preset.pack(side="left", fill="x", expand=True, padx=(0, 10), pady=8)

        # ===== 分组 1：页面与版式规范 =====
        card_page = ctk.CTkFrame(right_scroll, fg_color="#f8fafc", corner_radius=8, border_width=1, border_color="#e2e8f0")
        card_page.pack(fill="x", padx=6, pady=4)

        ctk.CTkLabel(card_page, text="页面与版式规范", font=ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold"), text_color="#2d3748").pack(anchor="w", padx=12, pady=(8, 4))

        # 纸张方向
        orient_row = ctk.CTkFrame(card_page, fg_color="transparent")
        orient_row.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(orient_row, text="纸张方向：", font=ctk.CTkFont(family="Microsoft YaHei", size=11), text_color="#4a5568", width=85).pack(side="left")
        self.seg_orientation = ctk.CTkSegmentedButton(
            orient_row,
            values=["纵向 (Portrait)", "横向 (Landscape)"],
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            height=28
        )
        self.seg_orientation.pack(side="left", fill="x", expand=True)

        # 页边距
        margin_row = ctk.CTkFrame(card_page, fg_color="transparent")
        margin_row.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(margin_row, text="页边距规格：", font=ctk.CTkFont(family="Microsoft YaHei", size=11), text_color="#4a5568", width=85).pack(side="left")
        margin_options = [v["label"] for v in MARGIN_MAP.values()]
        self.opt_margin = ctk.CTkOptionMenu(margin_row, values=margin_options, font=ctk.CTkFont(family="Microsoft YaHei", size=11), height=28)
        self.opt_margin.pack(side="left", fill="x", expand=True)

        # 标题对齐
        h1_row = ctk.CTkFrame(card_page, fg_color="transparent")
        h1_row.pack(fill="x", padx=12, pady=(3, 8))
        ctk.CTkLabel(h1_row, text="一级标题对齐：", font=ctk.CTkFont(family="Microsoft YaHei", size=11), text_color="#4a5568", width=85).pack(side="left")
        self.seg_h1_align = ctk.CTkSegmentedButton(h1_row, values=["居中对齐", "居左对齐"], font=ctk.CTkFont(family="Microsoft YaHei", size=11), height=28)
        self.seg_h1_align.pack(side="left", fill="x", expand=True)

        # ===== 分组 2：字体与段落排版 (Word 交互体系) =====
        card_font = ctk.CTkFrame(right_scroll, fg_color="#f8fafc", corner_radius=8, border_width=1, border_color="#e2e8f0")
        card_font.pack(fill="x", padx=6, pady=4)

        ctk.CTkLabel(card_font, text="字体与段落排版", font=ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold"), text_color="#2d3748").pack(anchor="w", padx=12, pady=(8, 4))

        # 字体
        font_row = ctk.CTkFrame(card_font, fg_color="transparent")
        font_row.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(font_row, text="正文字体：", font=ctk.CTkFont(family="Microsoft YaHei", size=11), text_color="#4a5568", width=85).pack(side="left")
        self.opt_font = ctk.CTkOptionMenu(font_row, values=list(FONT_MAP.keys()), font=ctk.CTkFont(family="Microsoft YaHei", size=11), height=28)
        self.opt_font.pack(side="left", fill="x", expand=True)

        # 字号 (Word 规范 Combobox，支持小初-小五选择或直接输入数值)
        size_row = ctk.CTkFrame(card_font, fg_color="transparent")
        size_row.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(size_row, text="正文字号：", font=ctk.CTkFont(family="Microsoft YaHei", size=11), text_color="#4a5568", width=85).pack(side="left")
        self.combo_size = ctk.CTkComboBox(
            size_row,
            values=FONT_SIZE_CHOICES,
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            height=28
        )
        self.combo_size.pack(side="left", fill="x", expand=True)

        # 行间距与段落间距
        spacing_row = ctk.CTkFrame(card_font, fg_color="transparent")
        spacing_row.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(spacing_row, text="行间距：", font=ctk.CTkFont(family="Microsoft YaHei", size=11), text_color="#4a5568", width=85).pack(side="left")
        self.opt_line_height = ctk.CTkOptionMenu(
            spacing_row,
            values=["1.25 倍", "1.50 倍", "1.65 倍 (推荐)", "1.75 倍", "2.00 倍"],
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            height=28,
            width=120
        )
        self.opt_line_height.pack(side="left", padx=(0, 8))

        ctk.CTkLabel(spacing_row, text="段落间距：", font=ctk.CTkFont(family="Microsoft YaHei", size=11), text_color="#4a5568").pack(side="left")
        self.opt_para_spacing = ctk.CTkOptionMenu(
            spacing_row,
            values=["紧凑 (4 pt)", "标准 (8 pt)", "宽松 (12 pt)"],
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            height=28
        )
        self.opt_para_spacing.pack(side="left", fill="x", expand=True)

        # 段落排版开关 (两端对齐与首行缩进)
        toggle_row = ctk.CTkFrame(card_font, fg_color="transparent")
        toggle_row.pack(fill="x", padx=12, pady=(3, 8))
        self.chk_justify = ctk.CTkCheckBox(toggle_row, text="正文两端对齐 (避头尾)", font=ctk.CTkFont(family="Microsoft YaHei", size=11))
        self.chk_justify.pack(side="left", padx=(0, 16))
        self.chk_indent = ctk.CTkCheckBox(toggle_row, text="首行缩进 2 字符 (公文规范)", font=ctk.CTkFont(family="Microsoft YaHei", size=11))
        self.chk_indent.pack(side="left")

        # ===== 分组 3：页眉、页脚与水印 =====
        card_header = ctk.CTkFrame(right_scroll, fg_color="#f8fafc", corner_radius=8, border_width=1, border_color="#e2e8f0")
        card_header.pack(fill="x", padx=6, pady=4)

        ctk.CTkLabel(card_header, text="页眉、页脚与水印", font=ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold"), text_color="#2d3748").pack(anchor="w", padx=12, pady=(8, 4))

        self.chk_footer = ctk.CTkCheckBox(card_header, text="显示页码 (第 X 页 / 共 Y 页)", font=ctk.CTkFont(family="Microsoft YaHei", size=11))
        self.chk_footer.pack(anchor="w", padx=12, pady=2)

        self.chk_suppress_first = ctk.CTkCheckBox(card_header, text="首页隐藏页眉与页脚 (用于封面或主标题页)", font=ctk.CTkFont(family="Microsoft YaHei", size=11))
        self.chk_suppress_first.pack(anchor="w", padx=12, pady=2)

        self.chk_toc = ctk.CTkCheckBox(card_header, text="自动生成目录 [TOC]", font=ctk.CTkFont(family="Microsoft YaHei", size=11))
        self.chk_toc.pack(anchor="w", padx=12, pady=2)

        hdr_row = ctk.CTkFrame(card_header, fg_color="transparent")
        hdr_row.pack(fill="x", padx=12, pady=3)
        self.chk_header = ctk.CTkCheckBox(hdr_row, text="页眉文本：", font=ctk.CTkFont(family="Microsoft YaHei", size=11), width=95)
        self.chk_header.pack(side="left")
        self.entry_header_text = ctk.CTkEntry(hdr_row, placeholder_text="例如：内部技术方案", height=28, font=ctk.CTkFont(family="Microsoft YaHei", size=11))
        self.entry_header_text.pack(side="left", fill="x", expand=True)

        wm_row = ctk.CTkFrame(card_header, fg_color="transparent")
        wm_row.pack(fill="x", padx=12, pady=(3, 8))
        self.chk_watermark = ctk.CTkCheckBox(wm_row, text="防泄密水印：", font=ctk.CTkFont(family="Microsoft YaHei", size=11), width=95)
        self.chk_watermark.pack(side="left")
        self.entry_watermark_text = ctk.CTkEntry(wm_row, placeholder_text="例如：内部资料 严禁外传", height=28, font=ctk.CTkFont(family="Microsoft YaHei", size=11))
        self.entry_watermark_text.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkLabel(wm_row, text="水印字号：", font=ctk.CTkFont(family="Microsoft YaHei", size=11), text_color="#4a5568").pack(side="left", padx=(0, 4))
        self.combo_watermark_size = ctk.CTkComboBox(
            wm_row,
            values=WATERMARK_SIZE_CHOICES,
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            width=115,
            height=28
        )
        self.combo_watermark_size.pack(side="right")

        # ===== 分组 4：色彩与代码块样式 =====
        card_style = ctk.CTkFrame(right_scroll, fg_color="#f8fafc", corner_radius=8, border_width=1, border_color="#e2e8f0")
        card_style.pack(fill="x", padx=6, pady=(4, 8))

        ctk.CTkLabel(card_style, text="色彩与代码块样式", font=ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold"), text_color="#2d3748").pack(anchor="w", padx=12, pady=(8, 4))

        color_row = ctk.CTkFrame(card_style, fg_color="transparent")
        color_row.pack(fill="x", padx=12, pady=3)
        ctk.CTkLabel(color_row, text="主题色：", font=ctk.CTkFont(family="Microsoft YaHei", size=11), text_color="#4a5568", width=85).pack(side="left")

        self.var_accent_color = tk.StringVar(value="#2b6cb0")
        color_btns = ctk.CTkFrame(color_row, fg_color="transparent")
        color_btns.pack(side="left", fill="x", expand=True)

        for hex_code, cname in COLOR_PRESETS:
            ctk.CTkButton(
                color_btns,
                text=cname,
                width=54,
                height=26,
                fg_color=hex_code,
                hover_color=hex_code,
                font=ctk.CTkFont(family="Microsoft YaHei", size=10, weight="bold"),
                command=lambda h=hex_code: self.var_accent_color.set(h)
            ).pack(side="left", padx=2)

        self.entry_color_hex = ctk.CTkEntry(color_btns, textvariable=self.var_accent_color, width=75, height=26, font=ctk.CTkFont(family="Consolas", size=10))
        self.entry_color_hex.pack(side="left", padx=(4, 0))

        code_row = ctk.CTkFrame(card_style, fg_color="transparent")
        code_row.pack(fill="x", padx=12, pady=(3, 8))
        ctk.CTkLabel(code_row, text="代码块样式：", font=ctk.CTkFont(family="Microsoft YaHei", size=11), text_color="#4a5568", width=85).pack(side="left")
        self.opt_code_style = ctk.CTkOptionMenu(
            code_row,
            values=["科技蓝边框卡片 (推荐)", "GitHub 细灰边框", "深色高对比卡片", "极简浅灰背景"],
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            height=28
        )
        self.opt_code_style.pack(side="left", fill="x", expand=True)

    def on_mode_change(self, val):
        pass

    def on_preset_selected(self, choice):
        if "公文" in choice:
            self.apply_preset_dict(PRESET_THEMES["gov"])
        elif "GitHub" in choice:
            self.apply_preset_dict(PRESET_THEMES["github_clean"])
        else:
            self.apply_preset_dict(PRESET_THEMES["tech_blue"])

    def apply_preset_dict(self, p):
        # 字体
        f_key = p.get("font_family_key")
        if f_key in FONT_MAP:
            self.opt_font.set(f_key)

        # 字号
        self.combo_size.set(p.get("font_size", "小四 (12 pt)"))

        # 颜色
        self.var_accent_color.set(p.get("accent_color", "#2b6cb0"))

        # 行距与段距
        self.opt_line_height.set(p.get("line_height", "1.65 倍 (推荐)"))
        self.opt_para_spacing.set(p.get("para_spacing", "标准 (8 pt)"))

        # 页面方向与边距
        if p.get("page_orientation") == "landscape":
            self.seg_orientation.set("横向 (Landscape)")
        else:
            self.seg_orientation.set("纵向 (Portrait)")

        m_key = p.get("page_margin", "standard")
        if m_key in MARGIN_MAP:
            self.opt_margin.set(MARGIN_MAP[m_key]["label"])

        # 标题对齐
        if p.get("h1_align") == "left":
            self.seg_h1_align.set("居左对齐")
        else:
            self.seg_h1_align.set("居中对齐")

        # 开关项
        if p.get("indent_first_line", False):
            self.chk_indent.select()
        else:
            self.chk_indent.deselect()

        if p.get("text_align_justify", True):
            self.chk_justify.select()
        else:
            self.chk_justify.deselect()

        # 代码块
        cs = p.get("code_style", "blue_accent")
        if cs == "github_clean":
            self.opt_code_style.set("GitHub 细灰边框")
        elif cs == "dark":
            self.opt_code_style.set("深色高对比卡片")
        elif cs == "minimal":
            self.opt_code_style.set("极简浅灰背景")
        else:
            self.opt_code_style.set("科技蓝边框卡片 (推荐)")

    def load_settings_into_ui(self):
        c = self.config

        # 模式与输出格式
        if c.get("output_mode") == "merge":
            self.seg_mode.set("按顺序合并为单一长报告")
        else:
            self.seg_mode.set("分别批量转换 (独立文件)")

        fmt = c.get("export_format", "pdf")
        if fmt == "docx":
            self.seg_format.set("导出 Word (.docx)")
        elif fmt == "both":
            self.seg_format.set("同时导出 (PDF + Word)")
        else:
            self.seg_format.set("导出 PDF")

        if c.get("last_output_dir"):
            self.entry_out_dir.insert(0, c.get("last_output_dir"))

        # 页面与版式
        if c.get("page_orientation") == "landscape":
            self.seg_orientation.set("横向 (Landscape)")
        else:
            self.seg_orientation.set("纵向 (Portrait)")

        m_key = c.get("page_margin", "standard")
        if m_key in MARGIN_MAP:
            self.opt_margin.set(MARGIN_MAP[m_key]["label"])

        if c.get("h1_align") == "left":
            self.seg_h1_align.set("居左对齐")
        else:
            self.seg_h1_align.set("居中对齐")

        # 字体与排版
        f_key = c.get("font_family_key")
        if f_key and f_key in FONT_MAP:
            self.opt_font.set(f_key)
        else:
            self.opt_font.set("微软雅黑 (推荐·现代无衬线)")

        self.combo_size.set(str(c.get("font_size", "小四 (12 pt)")))
        self.opt_line_height.set(str(c.get("line_height", "1.65 倍 (推荐)")))
        self.opt_para_spacing.set(str(c.get("para_spacing", "标准 (8 pt)")))

        if c.get("text_align_justify", True):
            self.chk_justify.select()
        else:
            self.chk_justify.deselect()

        if c.get("indent_first_line", False):
            self.chk_indent.select()
        else:
            self.chk_indent.deselect()

        # 页眉页脚与水印
        if c.get("enable_footer", True):
            self.chk_footer.select()
        else:
            self.chk_footer.deselect()

        if c.get("suppress_first_page", True):
            self.chk_suppress_first.select()
        else:
            self.chk_suppress_first.deselect()

        if c.get("enable_toc", True):
            self.chk_toc.select()
        else:
            self.chk_toc.deselect()

        if c.get("enable_header", True):
            self.chk_header.select()
        else:
            self.chk_header.deselect()
        self.entry_header_text.insert(0, c.get("header_text", "内部交流材料"))

        if c.get("enable_watermark", False):
            self.chk_watermark.select()
        else:
            self.chk_watermark.deselect()
        self.entry_watermark_text.insert(0, c.get("watermark_text", "内部资料 严禁外传"))
        self.combo_watermark_size.set(str(c.get("watermark_font_size", "标准 (18 pt)")))

        # 颜色与代码块
        self.var_accent_color.set(c.get("accent_color", "#2b6cb0"))
        cs = c.get("code_style", "blue_accent")
        if cs == "github_clean":
            self.opt_code_style.set("GitHub 细灰边框")
        elif cs == "dark":
            self.opt_code_style.set("深色高对比卡片")
        elif cs == "minimal":
            self.opt_code_style.set("极简浅灰背景")
        else:
            self.opt_code_style.set("科技蓝边框卡片 (推荐)")

    def collect_settings_from_ui(self):
        c = {}
        # 模式与输出格式
        c["output_mode"] = "merge" if "合并" in self.seg_mode.get() else "separate"
        sel_fmt = self.seg_format.get()
        if "Word" in sel_fmt and "同时" not in sel_fmt:
            c["export_format"] = "docx"
        elif "同时" in sel_fmt:
            c["export_format"] = "both"
        else:
            c["export_format"] = "pdf"
        c["last_output_dir"] = self.entry_out_dir.get().strip()

        # 页面与版式
        c["page_orientation"] = "landscape" if "横向" in self.seg_orientation.get() else "portrait"

        sel_margin_label = self.opt_margin.get()
        m_key = "standard"
        for k, v in MARGIN_MAP.items():
            if v["label"] == sel_margin_label:
                m_key = k
                break
        c["page_margin"] = m_key
        c["h1_align"] = "left" if "居左" in self.seg_h1_align.get() else "center"

        # 字体与排版
        sel_font_key = self.opt_font.get()
        c["font_family_key"] = sel_font_key
        c["font_family"] = FONT_MAP.get(sel_font_key, "Microsoft YaHei, sans-serif")
        c["font_size"] = self.combo_size.get().strip()
        c["line_height"] = self.opt_line_height.get()
        c["para_spacing"] = self.opt_para_spacing.get()
        c["text_align_justify"] = bool(self.chk_justify.get())
        c["indent_first_line"] = bool(self.chk_indent.get())

        # 页眉页脚与水印
        c["enable_footer"] = bool(self.chk_footer.get())
        c["suppress_first_page"] = bool(self.chk_suppress_first.get())
        c["enable_toc"] = bool(self.chk_toc.get())
        c["enable_header"] = bool(self.chk_header.get())
        c["header_text"] = self.entry_header_text.get().strip()
        c["enable_watermark"] = bool(self.chk_watermark.get())
        c["watermark_text"] = self.entry_watermark_text.get().strip()
        c["watermark_font_size"] = self.combo_watermark_size.get().strip()

        # 色彩与代码块
        c["accent_color"] = self.var_accent_color.get().strip() or "#2b6cb0"
        cs_text = self.opt_code_style.get()
        if "GitHub" in cs_text:
            c["code_style"] = "github_clean"
        elif "深色" in cs_text:
            c["code_style"] = "dark"
        elif "极简" in cs_text:
            c["code_style"] = "minimal"
        else:
            c["code_style"] = "blue_accent"

        return c

    # --------------------------------------------------------------------------
    # 原生 TkDND 拖拽与文件清单管理
    # --------------------------------------------------------------------------

    def setup_drag_and_drop(self):
        """采用官方 TkDND 扩展，原生递归绑定 OLE 拖放事件，零底层崩溃风险"""
        def register_widget(w):
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind('<<Drop>>', self.on_drop_event)
            except Exception:
                pass
            try:
                for child in w.winfo_children():
                    register_widget(child)
            except Exception:
                pass

        register_widget(self)

    def on_drop_event(self, event):
        raw_data = event.data
        if not raw_data:
            return
        # 使用 Tkinter 内置 splitlist 准确解析包含空格与中文路径的文件序列
        files = self.tk.splitlist(raw_data)
        self.process_dropped_files(files)

    def process_dropped_files(self, files):
        added = 0
        skipped = 0
        for f in files:
            path_str = os.path.normpath(str(f))

            if os.path.isdir(path_str):
                for root_dir, _, sub_files in os.walk(path_str):
                    for sf in sub_files:
                        if sf.lower().endswith((".md", ".markdown")):
                            full = os.path.normpath(os.path.join(root_dir, sf))
                            if full not in self.file_list:
                                self.file_list.append(full)
                                added += 1
            elif os.path.isfile(path_str):
                if path_str.lower().endswith((".md", ".markdown")):
                    if path_str not in self.file_list:
                        self.file_list.append(path_str)
                        added += 1
                else:
                    skipped += 1

        if added > 0:
            self.refresh_listbox()
            self.status_lbl.configure(text=f"状态：已成功添加 {added} 个 Markdown 文件（共 {len(self.file_list)} 篇）。")
        elif skipped > 0:
            self.status_lbl.configure(text=f"提示：拖入的 {skipped} 个文件不是 Markdown 文件 (.md / .markdown)。")
        else:
            self.status_lbl.configure(text="提示：所选文件已在列表中或未扫描到有效 .md 文件。")

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 Markdown 文件 (支持多选)",
            filetypes=[("Markdown 文档", "*.md;*.markdown"), ("所有文件", "*.*")]
        )
        added = 0
        for p in paths:
            norm = os.path.normpath(p)
            if norm not in self.file_list:
                self.file_list.append(norm)
                added += 1
        if added > 0:
            self.refresh_listbox()
            self.status_lbl.configure(text=f"状态：已添加 {added} 个文件（共 {len(self.file_list)} 篇）。")

    def add_folder(self):
        d = filedialog.askdirectory(title="选择包含 Markdown 文件的文件夹")
        if d:
            added = 0
            for root_dir, _, files in os.walk(d):
                for f in files:
                    if f.lower().endswith((".md", ".markdown")):
                        full = os.path.normpath(os.path.join(root_dir, f))
                        if full not in self.file_list:
                            self.file_list.append(full)
                            added += 1
            if added > 0:
                self.refresh_listbox()
                self.status_lbl.configure(text=f"状态：扫描完成，已添加 {added} 个 Markdown 文件（共 {len(self.file_list)} 篇）。")
            else:
                self.status_lbl.configure(text="提示：该文件夹下未扫描到新的 .md 文件。")

    def move_up(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx > 0:
            item = self.file_list.pop(idx)
            self.file_list.insert(idx - 1, item)
            self.refresh_listbox(idx - 1)

    def move_down(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx < len(self.file_list) - 1:
            item = self.file_list.pop(idx)
            self.file_list.insert(idx + 1, item)
            self.refresh_listbox(idx + 1)

    def remove_selected(self):
        sel = list(self.listbox.curselection())
        for idx in reversed(sel):
            self.file_list.pop(idx)
        self.refresh_listbox()
        self.status_lbl.configure(text=f"状态：已移除选中项，剩余 {len(self.file_list)} 个文件。")

    def clear_files(self):
        self.file_list.clear()
        self.refresh_listbox()
        self.status_lbl.configure(text="状态：已清空待转换列表。")

    def refresh_listbox(self, select_idx=None):
        self.listbox.delete(0, "end")
        for i, p in enumerate(self.file_list):
            fname = os.path.basename(p)
            fdir = os.path.dirname(p)
            self.listbox.insert("end", f"  [{i+1}]  {fname}    ({fdir})")

        count = len(self.file_list)
        self.lbl_file_count.configure(text=f"({count} 篇)")

        if count == 0:
            self.list_view.pack_forget()
            self.empty_drop_view.pack(fill="both", expand=True)
        else:
            self.empty_drop_view.pack_forget()
            self.list_view.pack(fill="both", expand=True)

        if select_idx is not None and 0 <= select_idx < count:
            self.listbox.selection_set(select_idx)

    def on_listbox_select(self, event=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if 0 <= idx < len(self.file_list):
            p = self.file_list[idx]
            fname = os.path.basename(p)
            fdir = os.path.dirname(p)
            if len(sel) == 1:
                self.status_lbl.configure(text=f"选中 [{idx+1}/{len(self.file_list)}]：{fname}  (目录: {fdir})")
            else:
                self.status_lbl.configure(text=f"当前已多选 {len(sel)} 个文档（共 {len(self.file_list)} 篇）")

    def on_listbox_double_click(self, event=None):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if 0 <= idx < len(self.file_list):
            p = self.file_list[idx]
            if os.path.exists(p):
                try:
                    os.startfile(p)
                except Exception:
                    pass

    def choose_out_dir(self):
        d = filedialog.askdirectory(title="选择输出目录")
        if d:
            self.entry_out_dir.delete(0, "end")
            self.entry_out_dir.insert(0, d)

    # --------------------------------------------------------------------------
    # 异步转换执行流水线
    # --------------------------------------------------------------------------

    def start_conversion_thread(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先通过拖拽或按钮添加至少一个 Markdown 文件。")
            return

        cfg = self.collect_settings_from_ui()
        save_config(cfg)
        self.config = cfg

        custom_out_dir = cfg.get("last_output_dir", "").strip()
        if custom_out_dir and not os.path.exists(custom_out_dir):
            try:
                os.makedirs(custom_out_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录: {e}")
                return

        self.btn_convert.configure(state="disabled", text="正在排版转换中...")
        self.status_lbl.configure(text="状态：正在调度 Chromium 引擎渲染离线图表与公式...")

        threading.Thread(target=self._worker_convert, args=(cfg, custom_out_dir), daemon=True).start()

    def _worker_convert(self, cfg, custom_out_dir):
        mode = cfg.get("output_mode", "separate")
        export_format = cfg.get("export_format", "pdf")
        success_count = 0
        last_out = None
        error_msg = None

        try:
            if mode == "separate":
                for i, fpath in enumerate(self.file_list):
                    base_name = os.path.splitext(os.path.basename(fpath))[0]
                    target_dir = custom_out_dir if custom_out_dir else os.path.dirname(fpath)

                    if export_format in ["pdf", "both"]:
                        out_pdf = os.path.join(target_dir, f"{base_name}.pdf")
                        self._update_status_safe(f"状态：正在导出 PDF ({i+1}/{len(self.file_list)}): {os.path.basename(fpath)} ...")
                        convert_single_file(fpath, out_pdf, cfg)
                        last_out = out_pdf

                    if export_format in ["docx", "both"]:
                        out_docx = os.path.join(target_dir, f"{base_name}.docx")
                        self._update_status_safe(f"状态：正在导出 Word ({i+1}/{len(self.file_list)}): {os.path.basename(fpath)} ...")
                        convert_single_to_docx(fpath, out_docx, cfg)
                        last_out = out_docx

                    success_count += 1
            else:
                first_name = os.path.splitext(os.path.basename(self.file_list[0]))[0]
                target_dir = custom_out_dir if custom_out_dir else os.path.dirname(self.file_list[0])

                if export_format in ["pdf", "both"]:
                    merged_pdf = os.path.join(target_dir, f"{first_name}_合并长报告.pdf")
                    self._update_status_safe(f"状态：正在合并并生成长报告 PDF ({len(self.file_list)} 篇)...")
                    convert_merged_files(self.file_list, merged_pdf, cfg)
                    last_out = merged_pdf

                if export_format in ["docx", "both"]:
                    merged_docx = os.path.join(target_dir, f"{first_name}_合并长报告.docx")
                    self._update_status_safe(f"状态：正在合并并生成长报告 Word ({len(self.file_list)} 篇)...")
                    convert_merged_to_docx(self.file_list, merged_docx, cfg)
                    last_out = merged_docx

                success_count = len(self.file_list)
        except Exception as e:
            error_msg = str(e)

        self.last_output_pdf = last_out
        self.after(0, lambda: self._on_convert_finished(success_count, last_out, error_msg, mode, export_format))

    def _update_status_safe(self, text):
        self.after(0, lambda: self.status_lbl.configure(text=text))

    def _on_convert_finished(self, success_count, last_out, error_msg, mode, export_format="pdf"):
        self.btn_convert.configure(state="normal", text="开始转换")
        if error_msg:
            self.status_lbl.configure(text=f"状态异常：{error_msg}")
            messagebox.showerror("转换失败", f"处理过程中发生错误:\n{error_msg}")
        else:
            fmt_desc = "组文档 (PDF + Word)" if export_format == "both" else ("个 Word (.docx) 文档" if export_format == "docx" else "个 PDF 文档")
            if mode == "separate":
                self.status_lbl.configure(text=f"状态：完成。已成功生成 {success_count} {fmt_desc}。")
                messagebox.showinfo("完成", f"批量转换完成！\n已成功生成 {success_count} {fmt_desc}。")
            else:
                self.status_lbl.configure(text=f"状态：长报告合并完成: {os.path.basename(last_out)}")
                messagebox.showinfo("完成", f"长报告合并成功！\n文件保存至:\n{last_out}")

    def open_output_dir(self):
        target = None
        if self.last_output_pdf and os.path.exists(self.last_output_pdf):
            target = os.path.dirname(self.last_output_pdf)
        elif self.entry_out_dir.get().strip() and os.path.exists(self.entry_out_dir.get().strip()):
            target = self.entry_out_dir.get().strip()
        elif self.file_list:
            target = os.path.dirname(self.file_list[0])

        if target and os.path.exists(target):
            os.startfile(target)
        else:
            messagebox.showinfo("提示", "暂无输出目录可打开。")


def main():
    if len(sys.argv) > 1:
        cfg = load_config()
        args = sys.argv[1:]
        export_docx = any(a in ["--docx", "-w", "--word"] for a in args)
        export_both = any(a in ["--both", "-b"] for a in args)
        if export_both:
            cfg["export_format"] = "both"
        elif export_docx:
            cfg["export_format"] = "docx"

        md_files = [os.path.abspath(a) for a in args if os.path.isfile(a) and a.lower().endswith((".md", ".markdown"))]
        if not md_files:
            print("[WARN] 未检测到有效的 .md 文件。")
            return
        for f in md_files:
            print(f"正在处理: {os.path.basename(f)} ...")
            try:
                if cfg.get("export_format", "pdf") in ["pdf", "both"]:
                    res_p = convert_single_file(f, None, cfg)
                    print(f"[SUCCESS] PDF 导出成功: {res_p}")
                if cfg.get("export_format", "pdf") in ["docx", "both"]:
                    res_d = convert_single_to_docx(f, None, cfg)
                    print(f"[SUCCESS] Word 导出成功: {res_d}")
            except Exception as e:
                print(f"[FAILED] 失败: {e}", file=sys.stderr)
    else:
        app = ModernMD2PDFApp()
        app.mainloop()

if __name__ == "__main__":
    main()
