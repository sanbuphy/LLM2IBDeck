# Style Calibration

Use this when the user asks for a named institution style such as McKinsey, Goldman Sachs, or CICC, especially when they ask for a near 1:1 mapping feel.

## Boundary

Do not copy official logos, proprietary templates, exact slide masters, watermarks, or confidential layouts. Do not imply the output was produced by the named institution. Build an original editable PPTX that maps closely to public visual grammar, tone, density, and page rhythm.

If the user provides a template they own, use that template as the stronger source of truth.

## Calibration Workflow

1. Search the web for public reference PDFs, reports, investor presentations, or official pages for the requested style.
2. Prefer official public sources first. Use third-party commentary only to fill gaps.
3. Inspect 5-10 representative pages, covering title, executive summary, chart, table, matrix, roadmap, and appendix pages when available.
4. Extract a style map:
   - Page margins, title position, title height, title rule, footer/source location.
   - Typography hierarchy: title, section header, chart labels, table cells, source notes.
   - Layout grammar: number of columns, card shape, table density, chart placement, whitespace.
   - Exhibit grammar: table header fill, alternating rows, bar colors, legend position, annotation style.
   - Tone: title phrasing, claim strength, source language, caveat language.
   - Deck rhythm: cover, key messages, context, analysis pages, recommendation, appendix.
5. Map the style to a BankerDeck theme profile:
   - `mckinsey-inspired`: consulting-medium density, larger answer-first titles, structured messages, more whitespace.
   - `goldman-inspired`: capital-markets-compact density, formal action titles, valuation/comps sensitivity language.
   - `cicc-inspired`: research-very-dense, source-forward Chinese footers, table-heavy pages, formal securities-research tone.
6. Generate the deck spec and then post-process page choices to match the style map:
   - McKinsey-like: fewer crowded tables, stronger synthesis, more structured boxes and clean charts.
   - Goldman-like: compact tables, ranges, valuation language, small source notes, premium restraint.
   - CICC-like: denser tables, Chinese research phrasing, explicit `资料来源` and `注`, policy/industry framing.
7. Run `python3 scripts/test_examples.py` and inspect at least one generated PPTX for the requested style.

## Required Output Note

When delivering a named style, mention that the deck is an original editable BankerDeck output calibrated from public references, not an official template.

## Reference Search Patterns

Use searches like:

- `site:mckinsey.com filetype:pdf McKinsey report chart layout`
- `site:goldmansachs.com filetype:pdf Goldman Sachs investor presentation`
- `site:cicc.com filetype:pdf CICC annual results presentation`
- `麦肯锡 研究报告 PDF 图表`
- `高盛 投资者 演示 PDF`
- `中金公司 业绩 发布 演示 PDF`

