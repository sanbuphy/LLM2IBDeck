# Production Architecture

Use this when the user expects a deck that looks like a real professional PPT, not only a valid editable `.pptx` file.

## Market Findings

Public tool behavior points to a clear pattern:

- Beautiful.ai maintains a design/editor system and then exports either editable PowerPoint or image-based PowerPoint. Its public support page notes that editable PowerPoint export keeps elements editable, while missing fonts can change formatting; image exports are static. Reference: https://support.beautiful.ai/hc/en-us/articles/30629528652685-Exporting-your-slides-and-presentations
- Gamma and similar AI slide tools optimize fast polished drafts and support PPTX export, but the web/card authoring model still requires export review when the final deliverable must be a PowerPoint file. Reference: https://help.gamma.app/fr/articles/8022861
- Marp is strong for Markdown/CSS-driven HTML, PDF, and PowerPoint conversion, but it is a presentation converter rather than an investment-banking slide design system. Reference: https://marp.app/
- PptxGenJS can create major editable PowerPoint objects, including text, tables, shapes, images, and charts, but it is a low-level generation library rather than a professional deck judgment layer. Reference: https://gitbrent.github.io/PptxGenJS/

Conclusion: do not treat raw OOXML as the primary authoring layer. Use a higher-level deck IR, professional layout components, rendered preview QA, and only then export PPTX.

## Recommended Stack

1. Story IR: topic, audience, decision, claim spine, source plan, slide list.
2. Style Map: public-reference calibrated layout grammar, tone, density, chart/table behavior.
3. Slide IR: one JSON object per slide with claim, proof object, data, source, module, and style role.
4. Professional Renderer: artifact-tool presentation JSX as the main editable deck builder.
5. Preview QA: export PNG previews, layout JSON, and a contact sheet.
6. PPTX Export: package only after preview/layout QA passes.
7. XML/OOXML Smoke: use direct XML generation only for low-level fallback and regression tests.

## Why XML-First Is Not Enough

XML-first catches broken structure, missing text, and non-editable output. It does not guarantee a real deck feel. A real deck requires:

- distinct slide rhythm across the contact sheet
- visual hierarchy that survives thumbnail view
- non-generic charts and tables
- source note discipline
- business language tuned to the institution style
- slide layouts that look authored, not assembled from boxes

## BankerDeck Rule

For production-quality output, prefer `scripts/generate_real_deck.mjs`. Use `scripts/generate_banker_deck.py` only as a fallback smoke generator or for XML regression checks.
