# BankerDeck

AI tool to generate standard investment banking editable PPT outputs.

This repository contains the `banker-deck` Codex skill:

- `banker-deck/SKILL.md`: skill trigger and operating workflow.
- `banker-deck/references/`: investment banking deck workflow, themes, slide modules, reference library, cases, and generator schema.
- `banker-deck/scripts/generate_real_deck.mjs`: artifact-tool production path with preview and layout QA.
- `banker-deck/scripts/generate_banker_deck.py`: XML-first fallback/smoke generator.
- `banker-deck/assets/*.json`: bundled sample deck specs.

Naming note: the product/display name is `BankerDeck`, while the Codex skill folder and skill name are lowercase `banker-deck`. Do not use `ibdack`; that is a typo and will not match the skill metadata.

## Install

For local development in this repository, use the skill directly from `banker-deck/`.

To install it into Codex's discovered skills directory:

```bash
mkdir -p ~/.codex/skills
cp -R banker-deck ~/.codex/skills/banker-deck
```

Then start a new Codex session and invoke it with prompts like:

```text
Use $banker-deck to create an editable investment banking PPTX on a healthcare SaaS acquisition target.
```

## Generate

Generate a production-path demo with editable PPTX, slide previews, layout JSON, and contact sheet:

```bash
node banker-deck/scripts/generate_real_deck.mjs
```

By default this generates all three public-reference calibrated named profiles under `temp/real-deck/`:

- `mckinsey-inspired`
- `goldman-inspired`
- `cicc-inspired`

Generate a single named production-path theme:

```bash
node banker-deck/scripts/generate_real_deck.mjs --theme goldman-inspired
```

Generate a single fallback XML-smoke sample deck into `temp/generated/`:

```bash
mkdir -p temp/generated
python3 banker-deck/scripts/generate_banker_deck.py \
  --spec banker-deck/assets/sample_market_scan.json \
  --output temp/generated/sample_market_scan.pptx
```

Generate and inspect XML before PPTX packaging:

```bash
python3 banker-deck/scripts/generate_banker_deck.py \
  --spec banker-deck/assets/finance_industry_research.json \
  --theme cicc-inspired \
  --xml-dir temp/generated/xml-first-cicc \
  --output temp/generated/cicc_xml_first_check.pptx
```

The `--xml-dir` output contains `slide*.xml`, `theme.json`, and `transformed_spec.json`. The PPTX is packaged only after the style-specific spec has been transformed and the slide XML has been generated.

Default finance examples:

- `banker-deck/assets/finance_cicc_china_industry.json`
- `banker-deck/assets/finance_mna_target_profile.json`
- `banker-deck/assets/finance_ipo_investor_update.json`
- `banker-deck/assets/finance_industry_research.json`
- `banker-deck/assets/sample_market_scan.json`

## Test

Run all bundled examples and validate the generated PPTX packages:

```bash
python3 scripts/test_examples.py
```

Run the production renderer test for McKinsey, Goldman, and CICC profiles. This generates editable PPTX files, PNG previews, contact sheets, and layout JSON under `temp/real-deck/`, then checks package XML and style-specific text/layout signals:

```bash
python3 scripts/test_real_decks.py
```

Generate same-input comparison decks for McKinsey, Goldman, and CICC profiles. This transforms content and writes pre-PPTX XML before packaging each deck:

```bash
python3 scripts/compare_named_styles.py
```

This writes:

- `temp/style-comparison/mckinsey-inspired.pptx`
- `temp/style-comparison/goldman-inspired.pptx`
- `temp/style-comparison/cicc-inspired.pptx`
- `temp/style-comparison/comparison_report.md`

The test writes generated decks only under `temp/generated/`. It validates:

- every bundled example spec in `banker-deck/assets/`
- every built-in theme, including `mckinsey-inspired`, `goldman-inspired`, and `cicc-inspired`
- each theme profile's palette, font, density, tone, and layout marker inside the generated PPTX XML
- PPTX package structure, slide count, and XML parseability

`temp/` and `output/` are ignored by Git so runtime artifacts do not get committed.

The named institution profiles are high-fidelity inspired profiles calibrated from public materials. They do not copy official logos, proprietary templates, or exact official slide masters.
