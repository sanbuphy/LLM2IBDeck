#!/usr/bin/env python3
"""Generate editable investment-banking-style PPTX files from an BankerDeck JSON spec."""

from __future__ import annotations

import argparse
import copy
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
        "house": "classic",
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
        "house": "mckinsey",
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
        "house": "goldman",
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
        "house": "cicc",
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
        "house": "boutique",
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
        "house": "board",
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
    house: str


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
    header_fill = None
    title_color = theme.primary
    if theme.house in {"goldman", "cicc"}:
        header_fill = theme.primary
        title_color = "FFFFFF"
    elif theme.house == "boutique":
        title_color = theme.ink
    parts = [
        shape_xml(2, "Background", 0, 0, 13.333, 7.5, theme.bg, None),
    ]
    sid = 3
    if header_fill:
        parts.append(shape_xml(sid, "Institutional Header Band", 0, 0, 13.333, 0.96, header_fill, None))
        sid += 1
    if theme.house == "mckinsey":
        parts.append(shape_xml(sid, "Consulting Left Rule", mx, theme.title_y + 0.04, 0.08, 0.55, theme.secondary, None))
        sid += 1
        title_x = mx + 0.18
        title_w = 12.25 - title_x
    else:
        title_x = mx
        title_w = 12.45 - mx
    parts.extend([
        text_box(sid, "Action Title", title_x, theme.title_y, title_w, 0.62, title, theme, theme.title_size, title_color, True, font_face=theme.title_font),
        line_xml(sid + 1, "Title Rule", mx, theme.rule_y, 12.6 - mx, theme.secondary if header_fill else theme.rule, theme.line_width),
        text_box(sid + 2, "Footer", mx, theme.footer_y, 10.2, 0.22, footer, theme, theme.source_size, theme.muted),
        text_box(sid + 3, "Page Number", 12.15, theme.footer_y, 0.5, 0.22, str(slide_num), theme, theme.source_size, theme.muted, False, "r"),
        text_box(sid + 4, "Theme Profile Marker", 10.65, 7.25, 2.3, 0.13, f"{theme.house} | {theme.profile} | {theme.density}", theme, 1, theme.bg),
    ])
    return parts, sid + 5


def render_title(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    title = slide.get("title", "Untitled Deck")
    subtitle = slide.get("subtitle", "")
    parts = [shape_xml(2, "Background", 0, 0, 13.333, 7.5, theme.bg, None)]
    if theme.house == "goldman":
        parts.extend([
            shape_xml(3, "Goldman Cover Band", 0, 0, 13.333, 1.15, theme.primary, None),
            shape_xml(4, "Goldman Gold Rule", 0, 1.15, 13.333, 0.05, theme.secondary, None),
            text_box(5, "Deck Title", 0.72, 2.32, 10.8, 0.92, title, theme, theme.title_size + 10, theme.primary, True, font_face=theme.title_font),
            text_box(6, "Deck Subtitle", 0.75, 3.22, 9.8, 0.42, subtitle, theme, theme.body_size + 2, theme.muted),
            text_box(7, "Footer", 0.75, 6.96, 10.2, 0.25, footer, theme, theme.source_size + 1, theme.muted),
            text_box(8, "Theme Profile Marker", 10.2, 7.25, 2.8, 0.13, f"{theme.house} | {theme.profile} | {theme.tone} | {theme.chart_style}", theme, 1, theme.bg),
        ])
        return slide_xml(parts)
    if theme.house == "cicc":
        parts.extend([
            shape_xml(3, "CICC Cover Header", 0, 0, 13.333, 0.9, theme.primary, None),
            shape_xml(4, "CICC Gold Rule", 0, 0.9, 13.333, 0.08, theme.secondary, None),
            shape_xml(5, "CICC Warm Panel", 0.55, 1.55, 12.1, 4.55, theme.light, theme.rule),
            text_box(6, "Deck Title", 0.82, 2.45, 10.8, 0.95, title, theme, theme.title_size + 10, theme.primary, True, font_face=theme.title_font),
            text_box(7, "Deck Subtitle", 0.86, 3.35, 10.0, 0.45, subtitle, theme, theme.body_size + 2, theme.ink),
            line_xml(8, "CICC Title Rule", 0.86, 4.02, 3.2, theme.secondary, 19050),
            text_box(9, "Footer", 0.78, 6.95, 10.2, 0.25, footer, theme, theme.source_size + 1, theme.muted),
            text_box(10, "Theme Profile Marker", 10.2, 7.25, 2.8, 0.13, f"{theme.house} | {theme.profile} | {theme.tone} | {theme.chart_style}", theme, 1, theme.bg),
        ])
        return slide_xml(parts)
    if theme.house == "mckinsey":
        parts.extend([
            shape_xml(3, "McKinsey Blue Anchor", 0.72, 1.55, 0.1, 3.1, theme.secondary, None),
            text_box(4, "Deck Title", 1.02, 2.08, 10.8, 1.05, title, theme, theme.title_size + 12, theme.primary, True, font_face=theme.title_font),
            text_box(5, "Deck Subtitle", 1.04, 3.12, 9.4, 0.45, subtitle, theme, theme.body_size + 3, theme.muted),
            line_xml(6, "Title Rule", 1.04, 3.86, 2.4, theme.secondary, 19050),
            text_box(7, "Footer", 1.04, 6.95, 10.2, 0.25, footer, theme, theme.source_size + 1, theme.muted),
            text_box(8, "Theme Profile Marker", 10.2, 7.25, 2.8, 0.13, f"{theme.house} | {theme.profile} | {theme.tone} | {theme.chart_style}", theme, 1, theme.bg),
        ])
        return slide_xml(parts)
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
    if theme.house == "mckinsey":
        return render_summary_mckinsey(slide, theme, footer, slide_num)
    if theme.house == "goldman":
        return render_summary_goldman(slide, theme, footer, slide_num)
    if theme.house == "cicc":
        return render_summary_cicc(slide, theme, footer, slide_num)
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    bullets = slide.get("bullets", [])
    cols = 1 if theme.density == "board-readable" else min(2, max(1, math.ceil(len(bullets) / 3)))
    if theme.house == "mckinsey":
        cols = 2
    if theme.house in {"goldman", "cicc"} and len(bullets) >= 4:
        cols = 2
    card_w = (11.95 - theme.margin_x) if cols == 1 else 5.85
    card_h = 1.02 if theme.density in {"board-readable", "consulting-medium"} else 0.84
    if theme.house == "mckinsey":
        card_h = 1.18
    if theme.house in {"goldman", "cicc"}:
        card_h = 0.76
    row_gap = 1.24 if card_h > 0.9 else 1.02
    for idx, item in enumerate(bullets):
        col = idx % cols
        row = idx // cols
        x = theme.margin_x + 0.13 + col * 6.05
        y = theme.body_y + row * row_gap
        fill = "FFFFFF" if theme.house == "mckinsey" else theme.light
        line = theme.secondary if theme.house in {"goldman", "cicc"} else theme.rule
        parts.append(shape_xml(sid, f"Message {idx + 1}", x, y, card_w, card_h, fill, line, str(item), theme.body_size, theme.ink, False, "l", theme.font, theme.house not in {"goldman", "cicc"}))
        parts.append(shape_xml(sid + 1, f"Message Accent {idx + 1}", x, y, 0.06, card_h, theme.secondary, None))
        sid += 2
    if slide.get("source"):
        parts.append(text_box(sid, "Source", theme.margin_x + 0.13, 6.72, 11.6, 0.24, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_summary_mckinsey(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    bullets = slide.get("bullets", [])
    parts.append(shape_xml(sid, "Big Insight Panel", 0.78, 1.35, 4.0, 4.7, "FFFFFF", theme.secondary))
    sid += 1
    if bullets:
        parts.append(text_box(sid, "Lead Insight", 1.02, 1.62, 3.45, 1.35, bullets[0], theme, 17, theme.primary, True, font_face=theme.title_font))
        sid += 1
    for idx, item in enumerate(bullets[1:5], start=1):
        y = 1.55 + (idx - 1) * 1.1
        parts.append(shape_xml(sid, f"McKinsey Number {idx}", 5.25, y, 0.34, 0.34, theme.secondary, None, f"{idx}", 10, "FFFFFF", True, "ctr", theme.font))
        parts.append(text_box(sid + 1, f"McKinsey Proof {idx}", 5.78, y - 0.08, 6.2, 0.62, item, theme, 12, theme.ink))
        sid += 2
    if slide.get("source"):
        parts.append(text_box(sid, "Source", 0.78, 6.72, 11.6, 0.24, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_summary_goldman(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    bullets = slide.get("bullets", [])
    parts.append(shape_xml(sid, "Executive Summary Header", 0.42, 1.08, 12.48, 0.36, theme.secondary, None, "Executive summary", 10, "FFFFFF", True, "l", theme.font))
    sid += 1
    for idx, item in enumerate(bullets[:5]):
        y = 1.55 + idx * 0.82
        parts.append(shape_xml(sid, f"Goldman Row Rule {idx + 1}", 0.48, y + 0.68, 12.0, 0.01, theme.rule, None))
        parts.append(text_box(sid + 1, f"Goldman Index {idx + 1}", 0.58, y, 0.42, 0.32, f"{idx + 1}", theme, 10, theme.secondary, True, "ctr"))
        parts.append(text_box(sid + 2, f"Goldman Message {idx + 1}", 1.08, y - 0.04, 10.8, 0.48, item, theme, 11, theme.ink))
        sid += 3
    if slide.get("source"):
        parts.append(text_box(sid, "Source", 0.48, 6.78, 11.6, 0.2, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_summary_cicc(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    bullets = slide.get("bullets", [])
    parts.append(shape_xml(sid, "Core View Label", 0.42, 1.08, 1.45, 0.36, theme.primary, None, "核心观点", 11, "FFFFFF", True, "ctr", theme.font))
    parts.append(shape_xml(sid + 1, "Research Summary Panel", 0.42, 1.52, 12.1, 4.65, theme.light, theme.rule))
    sid += 2
    for idx, item in enumerate(bullets[:5]):
        y = 1.72 + idx * 0.78
        parts.append(shape_xml(sid, f"CICC Bullet {idx + 1}", 0.72, y + 0.03, 0.16, 0.16, theme.secondary, None))
        parts.append(text_box(sid + 1, f"CICC Message {idx + 1}", 1.02, y - 0.06, 10.9, 0.42, item, theme, 11, theme.ink))
        sid += 2
    if slide.get("source"):
        parts.append(text_box(sid, "Source", 0.48, 6.78, 11.6, 0.2, slide["source"], theme, theme.source_size, theme.muted))
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
    if theme.house == "mckinsey":
        return render_table_mckinsey(slide, theme, footer, slide_num)
    if theme.house == "goldman":
        return render_table_goldman(slide, theme, footer, slide_num)
    if theme.house == "cicc":
        return render_table_cicc(slide, theme, footer, slide_num)
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    headers = slide.get("headers", [])
    rows = slide.get("rows", [])
    total_rows = max(1, len(rows) + 1)
    cols = max(1, len(headers))
    x0, y0, w, h = theme.margin_x + 0.13, theme.body_y, 12.0 - max(0, theme.margin_x - 0.45), min(5.35, theme.table_row_h * total_rows)
    col_w = w / cols
    row_h = h / total_rows
    for c, header in enumerate(headers):
        header_fill = theme.primary
        if theme.house == "mckinsey":
            header_fill = "FFFFFF"
        header_color = theme.primary if theme.house == "mckinsey" else "FFFFFF"
        parts.append(shape_xml(sid, f"Header {c + 1}", x0 + c * col_w, y0, col_w, row_h, header_fill, theme.secondary, header, max(8, theme.body_size - 1), header_color, True, "l", theme.font))
        sid += 1
    for r, row in enumerate(rows):
        for c in range(cols):
            fill = "FFFFFF" if r % 2 == 0 else theme.light
            if theme.house == "goldman":
                fill = "FFFFFF" if r % 2 == 0 else "F4F7FB"
            if theme.house == "cicc":
                fill = "FFFFFF" if r % 2 == 0 else "F8F2E6"
            value = row[c] if c < len(row) else ""
            parts.append(shape_xml(sid, f"Cell {r + 1}-{c + 1}", x0 + c * col_w, y0 + (r + 1) * row_h, col_w, row_h, fill, theme.rule, value, max(7, theme.body_size - 2), theme.ink, False, "l", theme.font))
            sid += 1
    if slide.get("source"):
        parts.append(text_box(sid, "Source", theme.margin_x + 0.13, 6.72, 11.6, 0.24, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_table_mckinsey(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    headers = slide.get("headers", [])
    rows = slide.get("rows", [])
    y = 1.32
    for idx, row in enumerate(rows[:4]):
        x = 0.72 + (idx % 2) * 6.0
        y = 1.42 + (idx // 2) * 2.25
        heading = row[0] if row else f"Item {idx + 1}"
        body = " | ".join(str(v) for v in row[1:]) if len(row) > 1 else ""
        parts.append(shape_xml(sid, f"McKinsey Insight Card {idx + 1}", x, y, 5.35, 1.72, "FFFFFF", theme.rule, None, radius=True))
        parts.append(shape_xml(sid + 1, f"McKinsey Card Rule {idx + 1}", x, y, 0.07, 1.72, theme.secondary, None))
        parts.append(text_box(sid + 2, f"McKinsey Card Heading {idx + 1}", x + 0.22, y + 0.18, 4.8, 0.3, str(heading), theme, 13, theme.primary, True))
        parts.append(text_box(sid + 3, f"McKinsey Card Body {idx + 1}", x + 0.22, y + 0.62, 4.8, 0.72, body, theme, 10, theme.ink))
        sid += 4
    if headers:
        parts.append(text_box(sid, "McKinsey Exhibit Note", 0.72, 6.18, 11.4, 0.24, "Exhibit structured as insight cards rather than a dense table.", theme, 8, theme.muted))
        sid += 1
    if slide.get("source"):
        parts.append(text_box(sid, "Source", 0.72, 6.72, 11.6, 0.24, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_table_goldman(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    headers = slide.get("headers", [])
    rows = slide.get("rows", [])
    cols = max(1, len(headers))
    x0, y0, w = 0.42, 1.12, 12.48
    col_w = w / cols
    row_h = 0.38
    parts.append(shape_xml(sid, "Goldman Table Top Rule", x0, y0 - 0.06, w, 0.03, theme.secondary, None))
    sid += 1
    for c, header in enumerate(headers):
        parts.append(shape_xml(sid, f"Goldman Header {c + 1}", x0 + c * col_w, y0, col_w, row_h, theme.primary, "FFFFFF", header, 8, "FFFFFF", True, "l", theme.font))
        sid += 1
    for r, row in enumerate(rows):
        for c in range(cols):
            fill = "FFFFFF" if r % 2 == 0 else "F4F7FB"
            value = row[c] if c < len(row) else ""
            parts.append(shape_xml(sid, f"Goldman Cell {r + 1}-{c + 1}", x0 + c * col_w, y0 + (r + 1) * row_h, col_w, row_h, fill, theme.rule, value, 7, theme.ink, False, "l", theme.font))
            sid += 1
    if slide.get("source"):
        parts.append(text_box(sid, "Source", 0.42, 6.78, 11.6, 0.2, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_table_cicc(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    headers = slide.get("headers", [])
    rows = slide.get("rows", [])
    cols = max(1, len(headers))
    x0, y0, w = 0.42, 1.1, 12.48
    col_w = w / cols
    row_h = 0.38
    parts.append(shape_xml(sid, "CICC Table Label", x0, y0 - 0.46, 1.55, 0.3, theme.primary, None, "图表", 9, "FFFFFF", True, "ctr", theme.font))
    sid += 1
    for c, header in enumerate(headers):
        parts.append(shape_xml(sid, f"CICC Header {c + 1}", x0 + c * col_w, y0, col_w, row_h, theme.primary, theme.secondary, header, 8, "FFFFFF", True, "l", theme.font))
        sid += 1
    for r, row in enumerate(rows):
        for c in range(cols):
            fill = "FFFFFF" if r % 2 == 0 else "F8F2E6"
            value = row[c] if c < len(row) else ""
            parts.append(shape_xml(sid, f"CICC Cell {r + 1}-{c + 1}", x0 + c * col_w, y0 + (r + 1) * row_h, col_w, row_h, fill, theme.rule, value, 7, theme.ink, False, "l", theme.font))
            sid += 1
    if slide.get("source"):
        parts.append(text_box(sid, "Source", 0.42, 6.78, 11.6, 0.2, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_chart(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    if theme.house == "mckinsey":
        return render_chart_mckinsey(slide, theme, footer, slide_num)
    if theme.house == "goldman":
        return render_chart_goldman(slide, theme, footer, slide_num)
    if theme.house == "cicc":
        return render_chart_cicc(slide, theme, footer, slide_num)
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
    if theme.house == "mckinsey":
        colors = [theme.primary, theme.secondary, "7AA6C2"]
    if theme.house == "goldman":
        colors = [theme.primary, theme.secondary, theme.muted]
    if theme.house == "cicc":
        colors = [theme.primary, theme.secondary, "D7B56D"]
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


def render_chart_mckinsey(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    cats = slide.get("categories", [])
    series = slide.get("series", [])
    values = [float(v) for s in series for v in s.get("values", []) if isinstance(v, (int, float))]
    max_v = max(values) if values else 1
    lead_value = int(max_v)
    parts.append(shape_xml(sid, "McKinsey Key Number Panel", 0.78, 1.34, 3.0, 4.95, "FFFFFF", theme.rule))
    parts.append(text_box(sid + 1, "McKinsey Key Number", 1.08, 1.78, 2.4, 0.82, f"{lead_value}", theme, 34, theme.primary, True, "ctr", theme.title_font))
    parts.append(text_box(sid + 2, "McKinsey Key Number Label", 1.02, 2.58, 2.5, 0.54, "indexed peak value", theme, 10, theme.muted, False, "ctr"))
    sid += 3
    x0, y0, chart_w, chart_h = 4.45, 1.55, 7.45, 4.35
    parts.append(line_xml(sid, "McKinsey Axis", x0, y0 + chart_h, chart_w, theme.rule, 9525))
    sid += 1
    group_w = chart_w / max(1, len(cats))
    colors = [theme.primary, theme.secondary, "7AA6C2"]
    for ci, cat in enumerate(cats):
        for si, serie in enumerate(series[:2]):
            vals = serie.get("values", [])
            val = float(vals[ci]) if ci < len(vals) and isinstance(vals[ci], (int, float)) else 0
            bw = (group_w - 0.3) / max(1, len(series[:2]))
            bh = (val / max_v) * (chart_h - 0.28)
            bx = x0 + ci * group_w + 0.16 + si * bw
            by = y0 + chart_h - bh
            parts.append(shape_xml(sid, f"McKinsey Bar {ci + 1}-{si + 1}", bx, by, bw - 0.04, bh, colors[si], None))
            sid += 1
        parts.append(text_box(sid, f"McKinsey Category {ci + 1}", x0 + ci * group_w, y0 + chart_h + 0.12, group_w, 0.25, cat, theme, 8, theme.muted, False, "ctr"))
        sid += 1
    parts.append(text_box(sid, "McKinsey Exhibit Caption", 4.45, 1.18, 6.8, 0.24, "Exhibit uses a large synthesis callout plus simplified supporting chart.", theme, 8, theme.muted))
    sid += 1
    if slide.get("source"):
        parts.append(text_box(sid, "Source", 0.78, 6.72, 11.6, 0.24, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_chart_goldman(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    cats = slide.get("categories", [])
    series = slide.get("series", [])
    values = [float(v) for s in series for v in s.get("values", []) if isinstance(v, (int, float))]
    max_v = max(values) if values else 1
    parts.append(shape_xml(sid, "Goldman Metrics Strip", 0.42, 1.08, 12.48, 0.46, theme.light, theme.secondary))
    sid += 1
    for idx, serie in enumerate(series[:3]):
        vals = [v for v in serie.get("values", []) if isinstance(v, (int, float))]
        metric = max(vals) if vals else 0
        x = 0.62 + idx * 3.8
        parts.append(text_box(sid, f"Goldman Metric {idx + 1}", x, 1.16, 0.85, 0.22, f"{metric:g}", theme, 12, theme.primary, True, "r"))
        parts.append(text_box(sid + 1, f"Goldman Metric Label {idx + 1}", x + 0.95, 1.17, 2.35, 0.22, serie.get("name", f"Series {idx + 1}"), theme, 8, theme.muted))
        sid += 2
    x0, y0, row_h = 0.72, 1.88, 0.62
    colors = [theme.primary, theme.secondary, theme.muted]
    for ci, cat in enumerate(cats):
        y = y0 + ci * row_h
        parts.append(text_box(sid, f"Goldman Category {ci + 1}", x0, y, 1.18, 0.24, cat, theme, 8, theme.muted, False, "r"))
        sid += 1
        for si, serie in enumerate(series[:2]):
            vals = serie.get("values", [])
            val = float(vals[ci]) if ci < len(vals) and isinstance(vals[ci], (int, float)) else 0
            bw = (val / max_v) * 8.4
            by = y + 0.03 + si * 0.22
            parts.append(shape_xml(sid, f"Goldman Horizontal Bar {ci + 1}-{si + 1}", 2.1, by, bw, 0.16, colors[si], None))
            sid += 1
    parts.append(line_xml(sid, "Goldman Chart Base Rule", 2.1, 6.05, 8.8, theme.rule, 9525))
    sid += 1
    if slide.get("source"):
        parts.append(text_box(sid, "Source", 0.42, 6.78, 11.6, 0.2, slide["source"], theme, theme.source_size, theme.muted))
    return slide_xml(parts)


def render_chart_cicc(slide: dict[str, Any], theme: Theme, footer: str, slide_num: int) -> str:
    parts, sid = base_decor(slide_num, slide.get("title", ""), footer, theme)
    cats = slide.get("categories", [])
    series = slide.get("series", [])
    values = [float(v) for s in series for v in s.get("values", []) if isinstance(v, (int, float))]
    max_v = max(values) if values else 1
    parts.append(shape_xml(sid, "CICC Chart Label", 0.42, 1.08, 1.55, 0.3, theme.primary, None, "图表", 9, "FFFFFF", True, "ctr", theme.font))
    sid += 1
    x0, y0, chart_w, chart_h = 0.72, 1.65, 7.45, 3.65
    parts.append(shape_xml(sid, "CICC Chart Panel", 0.42, 1.45, 8.2, 4.3, "FFFFFF", theme.rule))
    parts.append(line_xml(sid + 1, "CICC Axis", x0, y0 + chart_h, chart_w, theme.rule, 9525))
    sid += 2
    group_w = chart_w / max(1, len(cats))
    colors = [theme.primary, theme.secondary, "D7B56D"]
    for ci, cat in enumerate(cats):
        for si, serie in enumerate(series[:2]):
            vals = serie.get("values", [])
            val = float(vals[ci]) if ci < len(vals) and isinstance(vals[ci], (int, float)) else 0
            bw = (group_w - 0.22) / max(1, len(series[:2]))
            bh = (val / max_v) * (chart_h - 0.26)
            bx = x0 + ci * group_w + 0.1 + si * bw
            by = y0 + chart_h - bh
            parts.append(shape_xml(sid, f"CICC Bar {ci + 1}-{si + 1}", bx, by, bw - 0.03, bh, colors[si], None))
            sid += 1
        parts.append(text_box(sid, f"CICC Category {ci + 1}", x0 + ci * group_w, y0 + chart_h + 0.08, group_w, 0.2, cat, theme, 7, theme.muted, False, "ctr"))
        sid += 1
    parts.append(shape_xml(sid, "CICC Side Table Header", 8.9, 1.45, 3.35, 0.36, theme.primary, theme.secondary, "核心假设", 8, "FFFFFF", True, "l", theme.font))
    sid += 1
    for idx, serie in enumerate(series[:3]):
        vals = [v for v in serie.get("values", []) if isinstance(v, (int, float))]
        max_metric = max(vals) if vals else 0
        fill = "FFFFFF" if idx % 2 == 0 else "F8F2E6"
        parts.append(shape_xml(sid, f"CICC Side Table Row {idx + 1}", 8.9, 1.81 + idx * 0.42, 3.35, 0.42, fill, theme.rule, f"{serie.get('name', '')}: {max_metric:g}", 7, theme.ink, False, "l", theme.font))
        sid += 1
    if slide.get("source"):
        parts.append(text_box(sid, "Source", 0.42, 6.78, 11.6, 0.2, slide["source"], theme, theme.source_size, theme.muted))
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


STYLE_CONTENT_PROFILES = {
    "mckinsey": {
        "deck_title": "Digital infrastructure investment priorities",
        "subtitle_suffix": "Executive fact base and implications",
        "footer": "BankerDeck | public-reference calibrated | McKinsey-style discussion draft",
        "title_prefix": "",
        "summary_title": "Power-secure digital infrastructure platforms are best positioned to capture AI-driven demand",
        "summary_bullets": [
            "Power availability is emerging as the primary gating constraint for AI-ready capacity growth.",
            "Investors should prioritize platforms with contracted capacity, credible utility relationships, and clear execution paths.",
            "Higher capital intensity can be underwritten when demand visibility and site control are both strong.",
            "The near-term agenda is to separate structural winners from project-level development risk.",
        ],
        "source": "Source: Public industry sources; BankerDeck synthesis.",
    },
    "goldman": {
        "deck_title": "Digital Infrastructure Sector View",
        "subtitle_suffix": "Capital markets-style sector view",
        "footer": "BankerDeck | public-reference calibrated | Goldman-style market discussion",
        "title_prefix": "",
        "summary_title": "Digital infrastructure risk / reward increasingly depends on power access, tenant quality and capex discipline",
        "summary_bullets": [
            "AI demand remains a constructive secular tailwind, but investable capacity is constrained by power and permitting.",
            "Contracted backlog, counterparty quality and escalation structure should drive valuation dispersion.",
            "Rising capex requirements increase the importance of development track record and funding flexibility.",
            "Key diligence focus: power queue position, pre-leasing visibility, unit economics and exit comparables.",
        ],
        "source": "Source: Public company and industry materials; BankerDeck analysis.",
    },
    "cicc": {
        "deck_title": "算力基础设施行业研究",
        "subtitle_suffix": "行业研究框架与投资含义",
        "footer": "BankerDeck | 公开资料校准 | 中金风格行业研究讨论稿",
        "title_prefix": "",
        "summary_title": "算力基础设施投资主线正从需求弹性转向电力约束、客户质量与资本开支纪律",
        "summary_bullets": [
            "AI 负载增长仍是行业中长期需求主线，但电力获取和审批节奏成为短期供给瓶颈。",
            "具备电力资源、核心区位和高质量客户锁定的平台型资产更可能获得估值溢价。",
            "资本开支上行要求投资人更重视项目回报、融资能力和建设交付确定性。",
            "建议重点验证电力排队位置、预租约质量、单位经济性和可比交易估值区间。",
        ],
        "source": "资料来源：公开公司及行业资料；BankerDeck 分析。",
    },
}


def transform_spec_for_theme(spec: dict[str, Any], theme: Theme) -> dict[str, Any]:
    transformed = copy.deepcopy(spec)
    profile = STYLE_CONTENT_PROFILES.get(theme.house)
    if not profile:
        return transformed
    transformed["title"] = profile["deck_title"]
    transformed["footer"] = profile["footer"]
    transformed["subtitle"] = profile["subtitle_suffix"]
    for slide in transformed.get("slides", []):
        slide_type = slide.get("type")
        if slide_type == "title":
            slide["title"] = profile["deck_title"]
            slide["subtitle"] = profile["subtitle_suffix"]
        elif slide_type == "summary":
            slide["title"] = profile["summary_title"]
            slide["bullets"] = profile["summary_bullets"]
            slide["source"] = profile["source"]
        elif theme.house == "cicc":
            if slide_type in {"table", "matrix"}:
                slide["title"] = "投资判断应回到电力、需求、资本开支与退出假设的交叉验证"
                slide["headers"] = ["研究维度", "核心验证问题", "投资含义"]
                slide["rows"] = [
                    ["需求", "客户预租约、负载类型、上架节奏", "决定收入兑现确定性"],
                    ["电力", "并网排队、PPA 成本、冗余能力", "决定项目供给约束"],
                    ["资本开支", "冷却、电气、土建成本曲线", "决定 IRR 和资金需求"],
                    ["退出", "平台交易、资产交易、REITs 退出窗口", "决定估值区间"],
                ]
                slide["source"] = profile["source"]
            elif slide_type == "chart":
                slide["title"] = "示意测算显示 AI-ready 规格提升将推高资本开支和电力密度"
                slide["source"] = "注：指数化示意测算。资料来源：BankerDeck 分析。"
        elif theme.house == "goldman":
            if slide_type in {"table", "matrix"}:
                slide["title"] = "Underwriting should triangulate demand visibility, power access, capex and exit valuation"
                slide["headers"] = ["Diligence area", "Key underwriting test", "Valuation implication"]
                slide["rows"] = [
                    ["Demand", "Pre-lease probability and tenant credit", "Supports multiple premium"],
                    ["Power", "Queue position, PPA cost and reliability", "Frames development risk"],
                    ["Capex", "Cooling, electrical and construction benchmarks", "Determines funding need"],
                    ["Exit", "Platform and asset transaction comparables", "Defines valuation range"],
                ]
                slide["source"] = profile["source"]
        elif theme.house == "mckinsey":
            if slide_type in {"table", "matrix"}:
                slide["title"] = "Four tests separate scalable platforms from one-off development risk"
                slide["headers"] = ["Test", "What to prove", "Why it matters"]
                slide["rows"] = [
                    ["Power-secure", "Verified grid capacity and utility path", "Constrains near-term supply"],
                    ["Customer-backed", "Anchor demand and pre-lease evidence", "Reduces absorption risk"],
                    ["Capex-disciplined", "Repeatable design and cost controls", "Protects return profile"],
                    ["Exit-ready", "Platform scarcity and buyer universe", "Supports terminal value"],
                ]
                slide["source"] = profile["source"]
    return transformed


def build_slide_xmls(spec: dict[str, Any]) -> tuple[list[str], Theme, str, dict[str, Any]]:
    theme = load_theme(spec.get("theme", "ib-classic"))
    transformed = transform_spec_for_theme(spec, theme)
    slides = transformed.get("slides") or [{"type": "title", "title": transformed.get("title", "BankerDeck"), "subtitle": transformed.get("subtitle", "")}]
    footer = transformed.get("footer", "BankerDeck | Draft")
    slide_xmls = [render_slide(slide, theme, footer, idx) for idx, slide in enumerate(slides, start=1)]
    for idx, xml in enumerate(slide_xmls, start=1):
        if "<p:sld" not in xml or "<p:spTree>" not in xml:
            raise ValueError(f"Invalid slide XML generated for slide {idx}")
    return slide_xmls, theme, footer, transformed


def write_xml_bundle(spec: dict[str, Any], output_dir: Path) -> dict[str, Any]:
    slide_xmls, theme, _footer, transformed = build_slide_xmls(spec)
    output_dir.mkdir(parents=True, exist_ok=True)
    for idx, xml in enumerate(slide_xmls, start=1):
        (output_dir / f"slide{idx}.xml").write_text(xml, encoding="utf-8")
    (output_dir / "theme.json").write_text(
        json.dumps(THEMES.get(spec.get("theme", "ib-classic"), THEMES["ib-classic"]), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (output_dir / "transformed_spec.json").write_text(
        json.dumps(transformed, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return {"slide_count": len(slide_xmls), "theme": theme.house, "output_dir": str(output_dir)}


def write_pptx(spec: dict[str, Any], output: Path) -> None:
    slide_xmls, theme, _footer, transformed = build_slide_xmls(spec)
    slides = transformed.get("slides") or [{"type": "title", "title": transformed.get("title", "BankerDeck"), "subtitle": transformed.get("subtitle", "")}]
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
        zf.writestr("docProps/core.xml", core_props(transformed.get("title", "BankerDeck")))
        zf.writestr("docProps/app.xml", app_props(len(slides)))
        for idx, xml in enumerate(slide_xmls, start=1):
            zf.writestr(f"ppt/slides/slide{idx}.xml", xml)
            zf.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", slide_rels())


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate an editable BankerDeck PPTX from a JSON spec.")
    parser.add_argument("--spec", required=True, help="Path to BankerDeck JSON spec.")
    parser.add_argument("--output", required=True, help="Output .pptx path.")
    parser.add_argument("--xml-dir", help="Optional directory for pre-PPTX slide XML and transformed spec.")
    parser.add_argument("--theme", help="Override the spec theme, e.g. mckinsey-inspired, goldman-inspired, cicc-inspired.")
    args = parser.parse_args()

    spec_path = Path(args.spec)
    with spec_path.open("r", encoding="utf-8") as f:
        spec = json.load(f)
    if args.theme:
        spec["theme"] = args.theme
    output = Path(args.output)
    if args.xml_dir:
        info = write_xml_bundle(spec, Path(args.xml_dir))
        print(f"Wrote pre-PPTX XML bundle: {info['output_dir']} ({info['slide_count']} slides)")
    write_pptx(spec, output)
    print(f"Wrote editable PPTX: {output}")


if __name__ == "__main__":
    main()
