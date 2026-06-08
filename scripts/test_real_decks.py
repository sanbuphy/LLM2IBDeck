#!/usr/bin/env python3
"""Generate and validate production-path BankerDeck demos."""

from __future__ import annotations

import json
import subprocess
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
GENERATOR = ROOT / "banker-deck" / "scripts" / "generate_real_deck.mjs"
TEMP_OUTPUT = ROOT / "temp" / "real-deck"
BUNDLED_NODE = Path(
    "/Users/sanbu/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node"
)

THEME_CHECKS = {
    "mckinsey-inspired": [
        "Digital infrastructure investment priorities",
        "Power-secure platforms are best positioned",
        "Executive fact base and implications",
        "BankerDeck | public-reference calibrated | McKinsey-inspired",
    ],
    "goldman-inspired": [
        "BANKERDECK CAPITAL MARKETS",
        "Risk / reward increasingly depends",
        "Capital markets diligence materials",
        "Goldman-inspired",
    ],
    "cicc-inspired": [
        "算力基础设施投资主线",
        "核心观点",
        "资料来源",
        "CICC-inspired",
    ],
}


def node_command() -> list[str]:
    if BUNDLED_NODE.exists():
        return [str(BUNDLED_NODE)]
    return ["node"]


def run_generator() -> dict:
    result = subprocess.run(
        [*node_command(), str(GENERATOR)],
        cwd=ROOT,
        check=True,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    return json.loads(result.stdout)


def pptx_xml_blob(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        required = {
            "[Content_Types].xml",
            "_rels/.rels",
            "ppt/presentation.xml",
            "ppt/_rels/presentation.xml.rels",
        }
        missing = required.difference(names)
        if missing:
            raise AssertionError(f"{path.name}: missing package files {sorted(missing)}")
        for name in names:
            if name.endswith(".xml"):
                ET.fromstring(zf.read(name))
        return "\n".join(
            zf.read(name).decode("utf-8", errors="ignore")
            for name in names
            if name.endswith(".xml")
        )


def validate_theme(theme_name: str, manifest: dict) -> None:
    output = Path(manifest["output"])
    if not output.exists() or output.stat().st_size <= 0:
        raise AssertionError(f"{theme_name}: missing PPTX output")
    if manifest["slideCount"] != 4:
        raise AssertionError(f"{theme_name}: expected 4 slides, got {manifest['slideCount']}")

    contact_sheet = Path(manifest["contactSheet"])
    if not contact_sheet.exists() or contact_sheet.stat().st_size <= 0:
        raise AssertionError(f"{theme_name}: missing contact sheet")

    preview_paths = [Path(item) for item in manifest["previewPaths"]]
    if len(preview_paths) != 4 or any(not item.exists() or item.stat().st_size <= 0 for item in preview_paths):
        raise AssertionError(f"{theme_name}: expected four non-empty preview PNGs")

    layout_dir = Path(manifest["layoutDir"])
    layout_paths = sorted(layout_dir.glob("*.layout.json"))
    if len(layout_paths) != 4:
        raise AssertionError(f"{theme_name}: expected four layout JSON files")
    for layout_path in layout_paths:
        layout = json.loads(layout_path.read_text(encoding="utf-8"))
        if not layout.get("elements"):
            raise AssertionError(f"{theme_name}: empty layout elements in {layout_path.name}")

    blob = pptx_xml_blob(output)
    for token in THEME_CHECKS[theme_name]:
        if token not in blob:
            raise AssertionError(f"{theme_name}: missing theme signal {token!r}")


def main() -> None:
    manifest = run_generator()
    generated = {Path(item["output"]).parts[-3]: item for item in manifest["generated"]}
    for theme_name in THEME_CHECKS:
        if theme_name not in generated:
            raise AssertionError(f"missing generated manifest for {theme_name}")
        validate_theme(theme_name, generated[theme_name])
        print(f"ok production theme {theme_name}")
    print(f"validated production decks under {TEMP_OUTPUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
