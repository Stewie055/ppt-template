from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT / "singlefile") not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT / "singlefile"))

from ppt_template_sdk import PptOperations  # noqa: E402


ASSETS_DIR = Path(__file__).resolve().parent / "assets"
OUTPUT_DIR = Path(__file__).resolve().parent / "output"


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    ops = PptOperations.load(template_path=str(ASSETS_DIR / "operations_template.pptx"))
    ops.insert_slide(target_index=1, layout_index=6)
    ops.add_section(name="正文", start_slide_index=1)
    ops.delete_table_row(slide_index=0, shape_locator="ops-table", row_index=1)
    ops.delete_table_column(slide_index=0, shape_locator="ops-table", column_index=1)
    ops.merge_table_cells(
        slide_index=0,
        shape_locator="ops-table",
        first_row=0,
        first_col=0,
        last_row=1,
        last_col=1,
    )
    output_path = OUTPUT_DIR / "operations_output.pptx"
    ops.save_to_path(str(output_path))
    print("output:", output_path)


if __name__ == "__main__":
    main()
