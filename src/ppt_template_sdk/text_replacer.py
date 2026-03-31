from __future__ import annotations

import re
from typing import Optional

from .context import RenderContext
from .exceptions import FieldReplaceError
from .models.text import TextReplaceResult
from .parser.template_parser import iter_shapes


class TextReplacer:
    def __init__(self, pattern: Optional[str] = None):
        self.pattern = pattern

    def replace_presentation_text(
        self,
        presentation,
        context: RenderContext,
        rendered_shape_ids: Optional[set[int]] = None,
        pattern: Optional[str] = None,
    ) -> TextReplaceResult:
        field_re = re.compile(pattern or self.pattern or r"\{\{([\w\.]+)\}\}")
        rendered_shape_ids = rendered_shape_ids or set()
        warnings: list[str] = []
        replaced_count = 0

        try:
            for slide in presentation.slides:
                for shape in iter_shapes(slide.shapes):
                    if getattr(shape, "shape_id", None) in rendered_shape_ids:
                        continue
                    if getattr(shape, "has_text_frame", False):
                        new_text, count = self._replace_text(shape.text_frame.text, field_re, context, warnings)
                        shape.text = new_text
                        replaced_count += count
                    if getattr(shape, "has_table", False):
                        for row in shape.table.rows:
                            for cell in row.cells:
                                new_text, count = self._replace_text(cell.text, field_re, context, warnings)
                                cell.text = new_text
                                replaced_count += count
        except Exception as exc:
            raise FieldReplaceError(str(exc)) from exc

        return TextReplaceResult(replaced_count=replaced_count, warnings=warnings)

    @staticmethod
    def _replace_text(text: str, field_re, context: RenderContext, warnings: list[str]) -> tuple[str, int]:
        replacements = 0

        def repl(match):
            nonlocal replacements
            path = match.group(1)
            value = context.get_value(path)
            replacements += 1
            if value is None:
                warnings.append(f"missing text field '{path}'")
                return ""
            return str(value)

        return field_re.sub(repl, text), replacements
