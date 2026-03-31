from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT / "src") not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT / "src"))

from pptx import Presentation  # noqa: E402

from ppt_template_sdk import RenderContext, TextReplacer  # noqa: E402


ASSETS_DIR = Path(__file__).resolve().parent / "assets"
OUTPUT_DIR = Path(__file__).resolve().parent / "output"


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    prs = Presentation(str(ASSETS_DIR / "text_replace_template.pptx"))
    result = TextReplacer().replace_presentation_text(
        prs,
        context=RenderContext(
            data={
                "project": {"name": "Aurora"},
                "owner": {"name": "Bob"},
                "report_date": "2026-04-01",
            }
        ),
    )
    output_path = OUTPUT_DIR / "text_replace_output.pptx"
    prs.save(str(output_path))
    print("replaced_count:", result.replaced_count)
    print("warnings:", result.warnings)
    print("output:", output_path)


if __name__ == "__main__":
    main()
