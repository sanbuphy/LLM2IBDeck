# IBDeck

AI tool to generate standard investment banking editable PPT outputs.

This repository contains the `ibdeck` Codex skill:

- `ibdeck/SKILL.md`: skill trigger and operating workflow.
- `ibdeck/references/`: investment banking deck workflow, themes, slide modules, reference library, cases, and generator schema.
- `ibdeck/scripts/generate_ibdeck.py`: JSON spec to editable `.pptx` generator.
- `ibdeck/assets/sample_market_scan.json`: sample deck spec.

Generate the sample deck:

```bash
python3 ibdeck/scripts/generate_ibdeck.py \
  --spec ibdeck/assets/sample_market_scan.json \
  --output output/sample_market_scan.pptx
```
