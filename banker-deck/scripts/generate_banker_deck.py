#!/usr/bin/env python3
"""Generate editable investment-banking-style PPTX files from an BankerDeck JSON spec."""

from __future__ import annotations

import argparse
import json
import math
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape


EMU_PER_INCH = 914400
SLIDE_W = int(13.333333 * EMU_PER_INCH)
SLIDE_H = int(7.5 * EMU_PER_INCH)


THEMES = {
    "ib-classic": {
        "bg": "FFFFFF",
        "ink": "1F2933",
        "muted": "5B6770",
        "primary": "17365D",
        "secondary": "4472C4",
        "accent": "70AD47",
        "light": "F3F5F7",
        "rule": "D9E2EC",
        "font": "Aptos",
        "title_font": "Aptos Display",
        "title_size": 18,
        "body_size": 12,
        "source_size": 7,
        "margin_x": 0.45,
        "title_y": 0.25,
        "rule_y": 0.92,
        "body_y": 1.25,
        "footer_y": 7.08,
        "line_width": 12700,
        "table_row_h": 0.48,
        "chart_style": "banker-grid",
        "tone": "action-title, quantified, diligence-oriented",
        "density": "banking-dense",
        "profile": "classic investment banking pitch book",
    },
    "mckinsey-inspired": {
        "bg": "FFFFFF",
        "ink": "202124",
        "muted": "5F6368",
        "primary": "003A70",
        "secondary": "00A3E0",
        "accent": "2F80ED",
        "light": "E6E9ED",
        "rule": "CFD7E2",
        "font": "Arial",
        "title_font": "Arial",
        "title_size": 20,
        "body_size": 12,
        "source_size": 7,
        "margin_x": 0.62,
        "title_y": 0.28,
        "rule_y": 0.98,
        "body_y": 1.32,
        "footer_y": 7.06,
        "line_width": 9525,
        "table_row_h": 0.56,
        "chart_style": "consulting-white-space",
        "tone": "answer-first, synthesized, executive",
        "density": "consulting-medium",
        "profile": "public McKinsey-style consulting report direction",
    },
    "goldman-inspired": {
        "bg": "FFFFFF",
        "ink": "1C2530",
        "muted": "637083",
        "primary": "0B1F3A",
        "secondary": "C9A227",
        "accent": "8A6D1D",
        "light": "E9F1FA",
        "rule": "D4DAE3",
        "font": "Arial",
        "title_font": "Arial",
        "title_size": 17,
        "body_size": 11,
        "source_size": 6,
        "margin_x": 0.42,
        "title_y": 0.22,
        "rule_y": 0.84,
        "body_y": 1.14,
        "footer_y": 7.12,
        "line_width": 12700,
        "table_row_h": 0.42,
        "chart_style": "capital-markets-compact",
        "tone": "formal, market-facing, valuation-aware",
        "density": "banking-very-dense",
        "profile": "public Goldman-style institutional presentation direction",
    },
    "cicc-inspired": {
        "bg": "FFFFFF",
        "ink": "1F1F1F",
        "muted": "666666",
        "primary": "B20D1E",
        "secondary": "B88900",
        "accent": "8C0014",
        "light": "F5F1EA",
        "rule": "E3D6C1",
        "font": "Microsoft YaHei",
        "title_font": "Microsoft YaHei",
        "title_size": 17,
        "body_size": 11,
        "source_size": 6,
        "margin_x": 0.42,
        "title_y": 0.22,
        "rule_y": 0.84,
        "body_y": 1.12,
        "footer_y": 7.12,
        "line_width": 15240,
        "table_row_h": 0.42,
        "chart_style": "china-research-table-forward",
        "tone": "formal, research-note, source-forward",
        "density": "research-very-dense",
        "profile": "public CICC-style China research briefing direction",
    },
    "boutique-dark": {
        "bg": "111827",
        "ink": "FFFFFF",
        "muted": "D1D5DB",
        "primary": "22D3EE",
        "secondary": "7C3AED",
        "accent": "22D3EE",
        "light": "374151",
        "rule": "4B5563",
        "font": "Aptos",
        "title_font": "Aptos Display",
        "title_size": 19,
        "body_size": 12,
        "source_size": 7,
        "margin_x": 0.55,
        "title_y": 0.28,
        "rule_y": 0.98,
        "body_y": 1.28,
        "footer_y": 7.06,
        "line_width": 12700,
        "table_row_h": 0.52,
        "chart_style": "growth-equity-dark",
        "tone": "sharp, founder-investor, thesis-led",
        "density": "modern-medium",
        "profile": "boutique growth advisory dark deck",
    },
    "board-clean": {
        "bg": "FFFFFF",
        "ink": "111111",
        "muted": "6B7280",
        "primary": "111111",
        "secondary": "2563EB",
        "accent": "B91C1C",
        "light": "F8FAFC",
        "rule": "E5E7EB",
        "font": "Arial",
        "title_font": "Arial",
        "title_size": 19,
        "body_size": 13,
        "source_size": 7,
        "margin_x": 0.62,
        "title_y": 0.3,
        "rule_y": 1.02,
        "body_y": 1.35,
        "footer_y": 7.05,
        "line_width": 9525,
        "table_row_h": 0.58,
        "chart_style": "board-readable",
        "tone": "decision-oriented, plain-English, management",
        "density": "board-readable",
        "profile": "board and management discussion paper",
    },
}


@dataclass
class Theme:
    bg: str
    ink: str
    muted: str
    primary: str
    secondary: str
    accent: str
    light: str
    rule: str
    font: str
    title_font: str
    title_size: int
    body_size: int
    source_size: int
    margin_x: float
    title_y: float
    rule_y: float
    body_y: float
    footer_y: float
    line_width: int
    table_row_h: float
    chart_style: str
    tone: str
    density: str
    profile: str


def emu(inches: float) -> int:
    return int(inches * EMU_PER_INCH)


def safe_text(value: Any) -> str:
    return escape(str(value or ""))


def load_theme(name: str) -> Theme:
    raw = THEMES.get(name, THEMES["ib-classic"])
    return Theme(**raw)


def shape_xml(
    shape_id: int,
    name: str,
    x: float,
    y: float,
    w: float,
    h: float,
    fill: str | None = None,
    line: str | None = None,
    text: str | None = None,
    font_size: int = 14,
    font_color: str | None = None,
    bold: bool = False,
    align: str = "l",
    font_face: str | None = None,
    radius: bool = False,
) -> str:
    fill_xml = "<a:noFill/>" if fill is None else f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    line_xml = "<a:ln><a:noFill/></a:ln>" if line is None else f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>'
    preset = "roundRect" if radius else "rect"
    tx = ""
    if text is not None:
        color = font_color or "000000"
        face = font_face or "Aptos"
        b = ' b="1"' if bold else ""
        paragraphs = str(text).split("\n")
        runs = []
        for paragraph in paragraphs:
            runs.append(
                f'<a:p><a:pPr algn="{align}"/>'
                f"<a:r><a:rPr lang=\"en-US\" sz=\"{font_size * 100}\"{b}>"
                f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                f'<a:latin typeface="{safe_text(face)}"/></a:rPr>'
                f"<a:t>{safe_text(paragraph)}</a:t></a:r></a:p>"
            )
        tx = (
            "<p:txBody><a:bodyPr wrap=\"square\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\"/>"
            "<a:lstStyle/>"
            + "".join(runs)
            + "</p:txBody>"
        )
    return f"""
    <p:sp>
      <p:nvSpPr><p:cNvPr id="{shape_id}" name="{safe_text(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>
        <a:prstGeom prst="{preset}"><a:avLst/></a:prstGeom>
        {fill_xml}
        {line_xml}
      </p:spPr>
      {tx}
    </p:sp>
    """


def text_box(shape_id: int, name: str, x: float, y: float, w: float, h: float, text: str, theme: Theme, size: int = 14, color: str | None = None, bold: bool = False, align: str = "l", font_face: str | None = None) -> str:
    return shape_xml(shape_id, name, x, y, w, h, None, None, text, size, color or theme.ink, bold, align, font_face or theme.font)


def line_xml(shape_id: int, name: str, x: float, y: float, w: float, color: str, width: int = 12700) -> str:
    return f"""
    <p:cxnSp>
      <p:nvCxnSpPr><p:cNvPr id="{shape_id}" name="{safe_text(name)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
      <p:spPr>
        <a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="0"/></a:xfrm>
        <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
        <a:ln w="{width}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>
      </p:spPr>
    </p:cxnSp>
    """


def bullet_text(items: list[Any]) -> str:
    return "\n".join(f"• {item}" for item in items)


def base_decor(slide_num: int, title: str, footer: str, theme: Theme) -> tuple[list[str], int]:
    mx = theme.margin_x
    parts = [
        shape_xml(2, "Background", 0, 0, 13.333, 7.5, theme.bg, None),
        text_box(3, "Action Title", mx, theme.title_y, 12.45 - mx, 0.62, title, theme, theme.title_size, theme.primary, True, font_face=theme.title_font),
        line_xml(4, "Title Rule", mx, theme.rule_y, 12.6 - mx, theme.rule, theme.line_width),
        text_box(5, "Footer", mx, theme.footer_y, 10.2, 0.22, footer, theme, theme.source_size, theme.muted),
        text_box(6, "Page Number", 12.15, theme.footer_y, 0.5, 0.22, str(slide_num), theme, theme.source_size, theme.muted, False, "r"),
        text_box(7, "Theme Profile Marker", 11.05, 7.25, 1.9, 0.13, f"{theme.profile} | {theme.density}", theme, 1, theme.bg),
    ]
    return parts, 8


def render_title(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    title = slide.get("title", "Untitled Deck")
    subtitle = slide.get("subtitle", "")
    parts = [shape_xml(2, "Background", 0, 0, 13.333, 7.5, theme.bg, None)]
    if theme.bg != "FFFFFF":
        title_color, subtitle_color = theme.ink, theme.muted
    else:
        title_color, subtitle_color = theme.primary, theme.muted
    parts.extend(
        [
            shape_xml(3, "Accent Bar", 0, 0, 0.18, 7.5, theme.secondary, None),
            text_box(4, "Deck Title", 0.75, 2.28, 10.8, 0.95, title, theme, theme.title_size + 11, title_color, True, font_face=theme.title_font),
            text_box(5, "Deck Subtitle", 0.78, 3.22, 10.0, 0.45, subtitle, theme, theme.body_size + 3, subtitle_color),
            line_xml(6, "Title Rule", 0.78, 3.98, 3.2, theme.secondary, 19050),
            text_box(7, "Footer", 0.78, 6.95, 10.2, 0.25, footer, theme, theme.source_size + 1, subtitle_color),
            text_box(8, "Theme Profile Marker", 10.2, 7.25, 2.8, 0.13, f"{theme.profile} | {theme.tone} | {theme.chart_style}", theme, 1, theme.bg),
        ]
    )
    return slide_xml(parts)


def render_summary(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    bullets = slide.get("bullets", [])
    cols = 1 if theme.density == "board-readable" else min(2, max(1, math.ceil(len(bullets) / 3)))
    card_w = (11.95 - theme.margin_x) if cols == 1 else 5.85
    card_h = 1.02 if theme.density in {"board-readable", "consulting-medium"} else 0.84
    row_gap = 1.24 if card_h > 0.9 else 1.02
    for idx, item in enumerate(bullets):
        col = idx % cols
        row = idx // cols
        x = theme.margin_x + 0.13 + col * 6.05
        y = theme.body_y + row * row_gap
        parts.append(shape_xml(sid, f"Message {idx + 1}", x, y, card_w, card_h, theme.light, theme.rule, str(item), theme.body_size, theme.ink, False, "l", theme.font, theme.density != "banking-very-dense"))
        parts.append(shape_xml(sid + 1, f"Message Accent {idx + 1}", x, y, 0.06, card_h, theme.secondary, None))
        sid += 2
    if slide.get("source"):
        parts.append(text_box(sid, "Source", theme.margin_x + 0.13, 6.72, 11.6, 0.24, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_two_column(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    columns = slide.get("columns", [])
    for idx, col in enumerate(columns[:2]):
        x = theme.margin_x + 0.13 + idx * 6.1
        parts.append(shape_xml(sid, f"Column Header {idx + 1}", x, theme.body_y - 0.08, 5.65, 0.42, theme.primary, None, col.get("heading", ""), theme.body_size, "FFFFFF", True, "l", theme.font))
        parts.append(shape_xml(sid + 1, f"Column Body {idx + 1}", x, theme.body_y + 0.44, 5.65, 4.75, theme.light, theme.rule, bullet_text(col.get("bullets", [])), theme.body_size, theme.ink, False, "l", theme.font))
        sid += 2
    if slide.get("source"):
        parts.append(text_box(sid, "Source", theme.margin_x + 0.13, 6.72, 11.6, 0.24, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_table(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    headers = slide.get("headers", [])
    rows = slide.get("rows", [])
    total_rows = max(1, len(rows) + 1)
    cols = max(1, len(headers))
    x0, y0, w, h = theme.margin_x + 0.13, theme.body_y, 12.0 - max(0, theme.margin_x - 0.45), min(5.35, theme.table_row_h * total_rows)
    col_w = w / cols
    row_h = h / total_rows
    for c, header in enumerate(headers):
        parts.append(shape_xml(sid, f"Header {c + 1}", x0 + c * col_w, y0, col_w, row_h, theme.primary, "FFFFFF", header, max(8, theme.body_size - 1), "FFFFFF", True, "l", theme.font))
        sid += 1
    for r, row in enumerate(rows):
        for c in range(cols):
            fill = "FFFFFF" if r % 2 == 0 else theme.light
            value = row[c] if c < len(row) else ""
            parts.append(shape_xml(sid, f"Cell {r + 1}-{c + 1}", x0 + c * col_w, y0 + (r + 1) * row_h, col_w, row_h, fill, theme.rule, value, max(7, theme.body_size - 2), theme.ink, False, "l", theme.font))
            sid += 1
    if slide.get("source"):
        parts.append(text_box(sid, "Source", theme.margin_x + 0.13, 6.72, 11.6, 0.24, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_chart(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    cats = slide.get("categories", [])
    series = slide.get("series", [])
    values = [float(v) for s in series for v in s.get("values", []) if isinstance(v, (int, float))]
    max_v = max(values) if values else 1
    x0, y0, chart_w, chart_h = theme.margin_x + 0.4, theme.body_y + 0.28, 10.9 - max(0, theme.margin_x - 0.45), 4.45
    parts.append(line_xml(sid, "Axis", x0, y0 + chart_h, chart_w, theme.rule))
    sid += 1
    group_w = chart_w / max(1, len(cats))
    bar_gap = 0.06
    colors = [theme.secondary, theme.accent, theme.primary]
    for ci, cat in enumerate(cats):
        for si, serie in enumerate(series[:3]):
            vals = serie.get("values", [])
            val = float(vals[ci]) if ci < len(vals) and isinstance(vals[ci], (int, float)) else 0
            bw = max(0.08, (group_w - 0.18) / max(1, len(series[:3])) - bar_gap)
            bh = (val / max_v) * (chart_h - 0.25)
            bx = x0 + ci * group_w + 0.09 + si * (bw + bar_gap)
            by = y0 + chart_h - bh
            parts.append(shape_xml(sid, f"Bar {ci + 1}-{si + 1}", bx, by, bw, bh, colors[si % len(colors)], None))
            sid += 1
        parts.append(text_box(sid, f"Category {ci + 1}", x0 + ci * group_w, y0 + chart_h + 0.1, group_w, 0.25, cat, theme, max(7, theme.source_size + 1), theme.muted, False, "ctr"))
        sid += 1
    legend_y = 1.12
    for si, serie in enumerate(series[:3]):
        parts.append(shape_xml(sid, f"Legend Swatch {si + 1}", 8.8 + si * 1.35, legend_y, 0.16, 0.16, colors[si % len(colors)], None))
        parts.append(text_box(sid + 1, f"Legend {si + 1}", 9.0 + si * 1.35, legend_y - 0.02, 1.05, 0.2, serie.get("name", f"Series {si + 1}"), theme, max(7, theme.source_size + 1), theme.muted))
        sid += 2
    if slide.get("source"):
        parts.append(text_box(sid, "Source", theme.margin_x + 0.13, 6.72, 11.6, 0.24, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_roadmap(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    phases = slide.get("phases", [])
    count = max(1, len(phases))
    col_w = 11.8 / count
    for idx, phase in enumerate(phases):
        x = theme.margin_x + 0.25 + idx * col_w
        parts.append(shape_xml(sid, f"Phase {idx + 1}", x, theme.body_y + 0.25, col_w - 0.18, 0.48, theme.primary, None, phase.get("name", ""), theme.body_size, "FFFFFF", True, "ctr", theme.font))
        parts.append(text_box(sid + 1, f"Period {idx + 1}", x, theme.body_y + 0.83, col_w - 0.18, 0.28, phase.get("period", ""), theme, max(8, theme.body_size - 3), theme.secondary, True, "ctr"))
        parts.append(shape_xml(sid + 2, f"Items {idx + 1}", x, theme.body_y + 1.23, col_w - 0.18, 2.8, theme.light, theme.rule, bullet_text(phase.get("items", [])), max(8, theme.body_size - 2), theme.ink, False, "l", theme.font, theme.density != "banking-very-dense"))
        sid += 3
    if slide.get("source"):
        parts.append(text_box(sid, "Source", theme.margin_x + 0.13, 6.72, 11.6, 0.24, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_section(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts = [shape_xml(2, "Background", 0, 0, 13.333, 7.5, theme.primary, None)]
    parts.append(text_box(3, "Section Title", 0.8, 3.05, 11.5, 0.85, slide.get("title", "Section"), theme, theme.title_size + 9, "FFFFFF", True, font_face=theme.title_font))
    parts.append(line_xml(4, "Rule", 0.8, 4.05, 2.3, theme.secondary, 19050))
    parts.append(text_box(5, "Footer", 0.8, 6.95, 10.2, 0.25, footer, theme, theme.source_size + 1, "FFFFFF"))
    return slide_xml(parts)


def render_generic(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    body = bullet_text(slide.get("bullets", [])) if slide.get("bullets") else slide.get("body", "")
    parts.append(shape_xml(sid, "Body", theme.margin_x + 0.25, theme.body_y, 11.9, 4.9, "FFFFFF", theme.rule, body, theme.body_size, theme.ink, False, "l", theme.font))
    if slide.get("source"):
        parts.append(text_box(sid + 1, "Source", theme.margin_x + 0.13, 6.72, 11.6, 0.24, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def slide_xml(parts: list[str]) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    {''.join(parts)}
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"""


def render_slide(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    slide_type = slide.get("type", "content")
    if slide_type == "title":
        return render_title(slide, theme, footer, slide_num)
    if slide_type == "summary":
        return render_summary(slide, theme, footer, slide_num)
    if slide_type == "two_column":
        return render_two_column(slide, theme, footer, slide_num)
    if slide_type in {"table", "matrix"}:
        return render_table(slide, theme, footer, slide_num)
    if slide_type == "chart":
        return render_chart(slide, theme, footer, slide_num)
    if slide_type == "roadmap":
        return render_roadmap(slide, theme, footer, slide_num)
    if slide_type == "section":
        return render_section(slide, theme, footer, slide_num)
    return render_generic(slide, theme, footer, slide_num)


def content_types(slide_count: int) -> str:
    overrides = "\n".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, slide_count + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  {overrides}
</Types>
"""


def rels_root() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def presentation_xml(slide_count: int) -> str:
    sld_ids = "\n".join(f'<p:sldId id="{255 + i}" r:id="rId{i}"/>' for i in range(1, slide_count + 1))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId{slide_count + 1}"/></p:sldMasterIdLst>
  <p:sldIdLst>{sld_ids}</p:sldIdLst>
  <p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="wide"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle/>
</p:presentation>
"""


def presentation_rels(slide_count: int) -> str:
    rels = [
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        for i in range(1, slide_count + 1)
    ]
    rels.append(f'<Relationship Id="rId{slide_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>')
    rels.append(f'<Relationship Id="rId{slide_count + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {''.join(rels)}
</Relationships>
"""


def slide_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>
"""


def slide_master_xml() -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>
"""


def slide_master_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
"""


def slide_layout_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
"""


def slide_layout_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
"""


def theme_xml(theme: Theme) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="BankerDeck">
  <a:themeElements>
    <a:clrScheme name="BankerDeck">
      <a:dk1><a:srgbClr val="{theme.ink}"/></a:dk1><a:lt1><a:srgbClr val="{theme.bg}"/></a:lt1>
      <a:dk2><a:srgbClr val="{theme.primary}"/></a:dk2><a:lt2><a:srgbClr val="{theme.light}"/></a:lt2>
      <a:accent1><a:srgbClr val="{theme.primary}"/></a:accent1><a:accent2><a:srgbClr val="{theme.secondary}"/></a:accent2>
      <a:accent3><a:srgbClr val="{theme.accent}"/></a:accent3><a:accent4><a:srgbClr val="{theme.muted}"/></a:accent4>
      <a:accent5><a:srgbClr val="{theme.rule}"/></a:accent5><a:accent6><a:srgbClr val="{theme.light}"/></a:accent6>
      <a:hlink><a:srgbClr val="{theme.secondary}"/></a:hlink><a:folHlink><a:srgbClr val="{theme.muted}"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="BankerDeck"><a:majorFont><a:latin typeface="{safe_text(theme.font)}"/></a:majorFont><a:minorFont><a:latin typeface="{safe_text(theme.font)}"/></a:minorFont></a:fontScheme>
    <a:fmtScheme name="BankerDeck"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
  </a:themeElements>
</a:theme>
"""


def core_props(title: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{safe_text(title)}</dc:title>
  <dc:creator>BankerDeck</dc:creator>
  <cp:lastModifiedBy>BankerDeck</cp:lastModifiedBy>
</cp:coreProperties>
"""


def app_props(slide_count: int) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>BankerDeck</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>{slide_count}</Slides>
</Properties>
"""


def write_pptx(spec: dict[str, Any], output: Path) -> None:
    slides = spec.get("slides") or [{"type": "title", "title": spec.get("title", "BankerDeck"), "subtitle": spec.get("subtitle", "")}]
    theme = load_theme(spec.get("theme", "ib-classic"))
    footer = spec.get("footer", "BankerDeck | Draft")
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types(len(slides)))
        zf.writestr("_rels/.rels", rels_root())
        zf.writestr("ppt/presentation.xml", presentation_xml(len(slides)))
        zf.writestr("ppt/_rels/presentation.xml.rels", presentation_rels(len(slides)))
        zf.writestr("ppt/slideMasters/slideMaster1.xml", slide_master_xml())
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", slide_master_rels())
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout_xml())
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slide_layout_rels())
        zf.writestr("ppt/theme/theme1.xml", theme_xml(theme))
        zf.writestr("docProps/core.xml", core_props(spec.get("title", "BankerDeck")))
        zf.writestr("docProps/app.xml", app_props(len(slides)))
        for idx, slide in enumerate(slides, start=1):
            zf.writestr(f"ppt/slides/slide{idx}.xml", render_slide(slide, theme, footer, idx))
            zf.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", slide_rels())


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate an editable BankerDeck PPTX from a JSON spec.")
    parser.add_argument("--spec", required=True, help="Path to BankerDeck JSON spec.")
    parser.add_argument("--output", required=True, help="Output .pptx path.")
    args = parser.parse_args()

    spec_path = Path(args.spec)
    with spec_path.open("r", encoding="utf-8") as f:
        spec = json.load(f)
    output = Path(args.output)
    write_pptx(spec, output)
    print(f"Wrote editable PPTX: {output}")


if __name__ == "__main__":
    main()
