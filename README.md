# BankerDeck

AI tool to generate standard investment banking editable PPT outputs.

This repository contains the `banker-deck` Codex skill:

- `banker-deck/SKILL.md`: skill trigger and operating workflow.
- `banker-deck/references/`: investment banking deck workflow, themes, slide modules, reference library, cases, and generator schema.
- `banker-deck/scripts/generate_banker_deck.py`: JSON spec to editable `.pptx` generator.
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

Generate a single sample deck into `temp/generated/`:

```bash
mkdir -p temp/generated
python3 banker-deck/scripts/generate_banker_deck.py \
  --spec banker-deck/assets/sample_market_scan.json \
  --output temp/generated/sample_market_scan.pptx
```

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

The test writes generated decks only under `temp/generated/`. It validates:

- every bundled example spec in `banker-deck/assets/`
- every built-in theme, including `mckinsey-inspired`, `goldman-inspired`, and `cicc-inspired`
- each theme profile's palette, font, density, tone, and layout marker inside the generated PPTX XML
- PPTX package structure, slide count, and XML parseability

`temp/` and `output/` are ignored by Git so runtime artifacts do not get committed.

The named institution profiles are high-fidelity inspired profiles calibrated from public materials. They do not copy official logos, proprietary templates, or exact official slide masters.
