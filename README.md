# IBDeck

AI tool to generate standard investment banking editable PPT outputs.

This repository contains the `ibdeck` Codex skill:

- `ibdeck/SKILL.md`: skill trigger and operating workflow.
- `ibdeck/references/`: investment banking deck workflow, themes, slide modules, reference library, cases, and generator schema.
- `ibdeck/scripts/generate_ibdeck.py`: JSON spec to editable `.pptx` generator.
- `ibdeck/assets/*.json`: bundled sample deck specs.

Naming note: the product/display name is `IBDeck`, while the Codex skill folder and skill name are lowercase `ibdeck`. Do not use `ibdack`; that is a typo and will not match the skill metadata.

## Install

For local development in this repository, use the skill directly from `ibdeck/`.

To install it into Codex's discovered skills directory:

```bash
mkdir -p ~/.codex/skills
cp -R ibdeck ~/.codex/skills/ibdeck
```

Then start a new Codex session and invoke it with prompts like:

```text
Use $ibdeck to create an editable investment banking PPTX on a healthcare SaaS acquisition target.
```

## Generate

Generate a single sample deck into `temp/generated/`:

```bash
mkdir -p temp/generated
python3 ibdeck/scripts/generate_ibdeck.py \
  --spec ibdeck/assets/sample_market_scan.json \
  --output temp/generated/sample_market_scan.pptx
```

Default finance examples:

- `ibdeck/assets/finance_cicc_china_industry.json`
- `ibdeck/assets/finance_mna_target_profile.json`
- `ibdeck/assets/finance_ipo_investor_update.json`
- `ibdeck/assets/finance_industry_research.json`
- `ibdeck/assets/sample_market_scan.json`

## Test

Run all bundled examples and validate the generated PPTX packages:

```bash
python3 scripts/test_examples.py
```

The test writes generated decks only under `temp/generated/`. It validates:

- every bundled example spec in `ibdeck/assets/`
- every built-in theme, including `mckinsey-inspired`, `goldman-inspired`, and `cicc-inspired`
- PPTX package structure, slide count, and XML parseability

`temp/` and `output/` are ignored by Git so runtime artifacts do not get committed.
