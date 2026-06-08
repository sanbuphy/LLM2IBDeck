#!/usr/bin/env python3
"""Generate and validate all bundled IBDeck example specs under temp/."""

from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "ibdeck" / "assets"
TEMP_OUTPUT = ROOT / "temp" / "generated"

sys.path.insert(0, str(ROOT / "ibdeck" / "scripts"))
from generate_ibdeck import write_pptx  # noqa: E402


def validate_pptx(path: Path, expected_slides: int) -> None:
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        slide_names = sorted(
            name
            for name in names
            if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        )
        if len(slide_names) != expected_slides:
            raise AssertionError(f"{path.name}: expected {expected_slides} slides, found {len(slide_names)}")
        required = {
            "[Content_Types].xml",
            "_rels/.rels",
            "ppt/presentation.xml",
            "ppt/_rels/presentation.xml.rels",
            "ppt/theme/theme1.xml",
        }
        missing = required.difference(names)
        if missing:
            raise AssertionError(f"{path.name}: missing package files {sorted(missing)}")
        for name in names:
            if name.endswith(".xml"):
                ET.fromstring(zf.read(name))


def main() -> None:
    TEMP_OUTPUT.mkdir(parents=True, exist_ok=True)
    specs = sorted(ASSETS.glob("*.json"))
    if not specs:
        raise SystemExit("No example specs found.")

    for spec_path in specs:
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
        output = TEMP_OUTPUT / f"{spec_path.stem}.pptx"
        write_pptx(spec, output)
        validate_pptx(output, len(spec.get("slides", [])))
        print(f"ok {spec_path.name} -> {output.relative_to(ROOT)}")

    print(f"generated {len(specs)} deck(s) under {TEMP_OUTPUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
