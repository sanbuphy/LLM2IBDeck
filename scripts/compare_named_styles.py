#!/usr/bin/env python3
"""Generate same-input named style decks and write a comparison report."""

from __future__ import annotations

import copy
import json
import sys
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
TEMP_OUTPUT = ROOT / "temp" / "style-comparison"
BASE_SPEC = ROOT / "banker-deck" / "assets" / "finance_industry_research.json"

sys.path.insert(0, str(ROOT / "banker-deck" / "scripts"))
from generate_banker_deck import THEMES, build_slide_xmls, write_pptx  # noqa: E402


STYLE_CHECKS = {
    "mckinsey-inspired": {
        "expected": [
            "large white-space title architecture",
            "answer-first title tone",
            "blue anchor/rule",
            "consulting-medium density",
            "structured message boxes",
        ],
        "tokens": [
            "mckinsey",
            "consulting-medium",
            "003A70",
            "00A3E0",
            "Power-secure digital infrastructure platforms are best positioned",
            "McKinsey Blue Anchor",
            "McKinsey Insight Card",
            "McKinsey Key Number Panel",
        ],
        "gap": "Still needs richer native chart annotation patterns and more visual hierarchy variants for strategy pages.",
    },
    "goldman-inspired": {
        "expected": [
            "institutional navy header band",
            "compact capital-markets density",
            "formal valuation-aware tone",
            "small source notes",
            "premium restrained blue/gold accents",
        ],
        "tokens": [
            "goldman",
            "capital-markets-compact",
            "0B1F3A",
            "C9A227",
            "risk / reward increasingly depends on power access",
            "Goldman Cover Band",
            "Goldman Table Top Rule",
            "Goldman Metrics Strip",
        ],
        "gap": "Still needs more football-field, valuation range, and comparable-company page archetypes.",
    },
    "cicc-inspired": {
        "expected": [
            "red/gold institutional header rhythm",
            "research-very-dense tables",
            "formal Chinese source-forward tone",
            "table-forward exhibits",
            "policy/industry framing language",
        ],
        "tokens": [
            "cicc",
            "research-very-dense",
            "B20D1E",
            "B88900",
            "算力基础设施投资主线",
            "资料来源",
            "CICC Cover Header",
            "CICC Table Label",
            "CICC Side Table Header",
        ],
        "gap": "Still needs more Chinese brokerage-style appendix tables and risk-disclosure pages.",
    },
}


def write_xml_bundle(spec: dict, theme_name: str) -> tuple[str, int]:
    slide_xmls, _theme, _footer, transformed = build_slide_xmls(spec)
    xml_dir = TEMP_OUTPUT / "xml" / theme_name
    xml_dir.mkdir(parents=True, exist_ok=True)
    for idx, xml in enumerate(slide_xmls, start=1):
        ET.fromstring(xml)
        (xml_dir / f"slide{idx}.xml").write_text(xml, encoding="utf-8")
    (xml_dir / "transformed_spec.json").write_text(
        json.dumps(transformed, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return "\n".join(slide_xmls), len(slide_xmls)


def main() -> None:
    TEMP_OUTPUT.mkdir(parents=True, exist_ok=True)
    base = json.loads(BASE_SPEC.read_text(encoding="utf-8"))
    report = [
        "# Named Style Comparison",
        "",
        f"Base input: `{BASE_SPEC.relative_to(ROOT)}`",
        "",
        "All outputs start from the same base input and slide count, then pass through style-specific content transformation before XML is generated.",
        "",
    ]

    for theme_name, checks in STYLE_CHECKS.items():
        spec = copy.deepcopy(base)
        spec["theme"] = theme_name
        spec["title"] = f"{base['title']} - {theme_name}"
        spec["footer"] = f"BankerDeck | public-reference calibrated | {theme_name}"
        blob, slide_count = write_xml_bundle(spec, theme_name)
        missing = [token for token in checks["tokens"] if token not in blob]
        status = "PASS" if not missing else f"FAIL missing {', '.join(missing)}"
        if missing:
            raise AssertionError(f"{theme_name}: {status}")
        output = TEMP_OUTPUT / f"{theme_name}.pptx"
        write_pptx(spec, output)
        report.extend([
            f"## {theme_name}",
            "",
            f"Output: `{output.relative_to(ROOT)}`",
            f"Pre-PPTX XML: `{(TEMP_OUTPUT / 'xml' / theme_name).relative_to(ROOT)}`",
            f"Slide count: {slide_count}",
            f"Status: {status}",
            "",
            "Expected mapping signals:",
            "",
        ])
        report.extend(f"- {item}" for item in checks["expected"])
        report.extend([
            "",
            f"Known remaining gap: {checks['gap']}",
            "",
        ])

    report_path = TEMP_OUTPUT / "comparison_report.md"
    report_path.write_text("\n".join(report), encoding="utf-8")
    print(f"wrote {report_path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
