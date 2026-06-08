---
name: banker-deck
description: Generate editable investment banking-style PowerPoint decks and supporting outlines for financial analysis, industry research, investor presentations, pitch books, management presentations, market maps, company profiles, transaction overviews, and board-style materials. Use when Codex needs to create a PPT/PPTX, deck outline, page-by-page content plan, investment banking slide narrative, source/reference plan, or themed consulting/investment-bank deliverable with editable XML PowerPoint output.
---

# BankerDeck

## Overview

Use BankerDeck to turn a business, finance, market, industry, company, or transaction topic into a standard investment banking deliverable. Produce either a structured outline/content plan or an editable `.pptx` built through a professional presentation-rendering workflow with preview QA.

## Core Workflow

1. Clarify the deck use case: pitch book, investor presentation, industry research, company profile, board memo, transaction overview, valuation, or market scan.
2. Build the story first: situation, key question, answer, supporting analyses, implications, and recommendation.
3. Select slide modules from `references/slide-modules.md`; prefer modules that fit the evidence available.
4. Select a theme from `references/themes.md`; default to `ib-classic` unless the user asks for another style.
5. For named institution styles or "near 1:1 mapping" requests, read `references/style-calibration.md`, search public references, extract a style map, and post-process layout/tone choices before generating.
6. Create a page-by-page plan with page title, main message, visual module, content bullets, required data, and source notes.
7. If generating a production PPTX, use the artifact-tool path in `scripts/generate_real_deck.mjs` or an equivalent presentation JSX workflow; render previews, layout JSON, and a contact sheet before delivery.
8. Use `scripts/generate_banker_deck.py` only for XML-first smoke tests, fallback generation, or regression checks.
9. Review the output for editability, hierarchy, consistent titles, source notes, and slide-to-slide narrative flow.

## Reference Map

- `references/workflow.md`: detailed topic understanding, storyline, outline, and quality checks.
- `references/slide-modules.md`: reusable IB slide patterns and when to use each.
- `references/themes.md`: default and optional house-style-inspired themes.
- `references/production-architecture.md`: market-informed architecture for generating real professional PPT outputs.
- `references/style-calibration.md`: near 1:1 public-reference style mapping workflow for named themes.
- `references/reference-library.md`: recommended sources and citation/source-note patterns.
- `references/cases.md`: example user prompts, interpretation, outline, and generation approach.
- `references/spec-schema.md`: JSON spec fields accepted by the generator.

Load only the reference files needed for the current request.

## Generate Editable PPTX

Use the artifact-tool generator when the user asks for a concrete production-quality deck file.

```bash
node banker-deck/scripts/generate_real_deck.mjs
```

By default it generates `mckinsey-inspired`, `goldman-inspired`, and `cicc-inspired` demo decks with preview PNGs, layout JSON, and contact sheets under `temp/real-deck/`. For a single theme:

```bash
node banker-deck/scripts/generate_real_deck.mjs --theme cicc-inspired
```

Use the Python XML generator only for smoke tests or fallback output.

```bash
mkdir -p temp/generated
python3 banker-deck/scripts/generate_banker_deck.py \
  --spec banker-deck/assets/sample_market_scan.json \
  --output temp/generated/banker-deck_market_scan.pptx
```

The generator creates a real `.pptx` package using editable PowerPoint XML objects. Do not rasterize slides into images unless the user explicitly asks for image-only output.

## Output Standards

- Use action titles: each slide title should state the conclusion, not just the topic.
- Keep one main message per slide.
- Include sources in small footer notes when claims depend on external facts.
- Prefer tables, waterfall/bridge logic, market maps, timelines, and bar/line charts over decorative graphics.
- Keep all generated elements editable: text boxes, shapes, tables, and charts.
- For real PPT quality or named themes, inspect rendered previews, contact sheets, and layout JSON before delivery.
- Use neutral professional language; avoid exaggerated claims unless sourced.
- When information is missing, create clearly labeled placeholders and a data request list.

## Theme Selection

Default to `ib-classic`. Offer optional style directions from `themes.md`: `mckinsey-inspired`, `goldman-inspired`, `cicc-inspired`, `boutique-dark`, and `board-clean`. These are high-fidelity inspired visual systems, not official templates or brand assets. For McKinsey, Goldman, CICC, or any named institution style, always calibrate from public references first.

## Typical Requests

- "Use BankerDeck to create a 12-page market entry deck for AI data centers in Southeast Asia."
- "Generate a Goldman-style editable PPTX about a potential acquisition of a SaaS company."
- "Create the outline and page content references for an investor presentation on GLP-1 supply chain opportunities."
- "Build a CICC-style industry research deck in Chinese for humanoid robotics."
