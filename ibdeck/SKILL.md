---
name: ibdeck
description: Generate editable investment banking-style PowerPoint decks and supporting outlines for financial analysis, industry research, investor presentations, pitch books, management presentations, market maps, company profiles, transaction overviews, and board-style materials. Use when Codex needs to create a PPT/PPTX, deck outline, page-by-page content plan, investment banking slide narrative, source/reference plan, or themed consulting/investment-bank deliverable with editable XML PowerPoint output.
---

# IBDeck

## Overview

Use IBDeck to turn a business, finance, market, industry, company, or transaction topic into a standard investment banking deliverable. Produce either a structured outline/content plan or an editable `.pptx` generated from a JSON spec using PowerPoint-native text boxes, shapes, tables, and charts.

## Core Workflow

1. Clarify the deck use case: pitch book, investor presentation, industry research, company profile, board memo, transaction overview, valuation, or market scan.
2. Build the story first: situation, key question, answer, supporting analyses, implications, and recommendation.
3. Select slide modules from `references/slide-modules.md`; prefer modules that fit the evidence available.
4. Select a theme from `references/themes.md`; default to `ib-classic` unless the user asks for another style.
5. Create a page-by-page plan with page title, main message, visual module, content bullets, required data, and source notes.
6. If generating a PPTX, write a JSON spec and run `scripts/generate_ibdeck.py`.
7. Review the output for editability, hierarchy, consistent titles, source notes, and slide-to-slide narrative flow.

## Reference Map

- `references/workflow.md`: detailed topic understanding, storyline, outline, and quality checks.
- `references/slide-modules.md`: reusable IB slide patterns and when to use each.
- `references/themes.md`: default and optional house-style-inspired themes.
- `references/reference-library.md`: recommended sources and citation/source-note patterns.
- `references/cases.md`: example user prompts, interpretation, outline, and generation approach.
- `references/spec-schema.md`: JSON spec fields accepted by the generator.

Load only the reference files needed for the current request.

## Generate Editable PPTX

Use the generator when the user asks for a concrete deck file.

```bash
python3 ibdeck/scripts/generate_ibdeck.py \
  --spec ibdeck/assets/sample_market_scan.json \
  --output output/ibdeck_market_scan.pptx
```

The generator creates a real `.pptx` package using editable PowerPoint XML objects. Do not rasterize slides into images unless the user explicitly asks for image-only output.

## Output Standards

- Use action titles: each slide title should state the conclusion, not just the topic.
- Keep one main message per slide.
- Include sources in small footer notes when claims depend on external facts.
- Prefer tables, waterfall/bridge logic, market maps, timelines, and bar/line charts over decorative graphics.
- Keep all generated elements editable: text boxes, shapes, tables, and charts.
- Use neutral professional language; avoid exaggerated claims unless sourced.
- When information is missing, create clearly labeled placeholders and a data request list.

## Theme Selection

Default to `ib-classic`. Offer optional style directions from `themes.md`: `mckinsey-inspired`, `goldman-inspired`, `cicc-inspired`, `boutique-dark`, and `board-clean`. These are inspired visual systems, not official templates or brand assets.

## Typical Requests

- "Use IBDeck to create a 12-page market entry deck for AI data centers in Southeast Asia."
- "Generate a Goldman-style editable PPTX about a potential acquisition of a SaaS company."
- "Create the outline and page content references for an investor presentation on GLP-1 supply chain opportunities."
- "Build a CICC-style industry research deck in Chinese for humanoid robotics."
