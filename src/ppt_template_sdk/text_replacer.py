"""TextReplacer 提供独立于模板渲染主链路的文本替换能力。

适合对已有 PPT 做字段替换，或在 ``PptTemplateEngine`` 渲染后复用同一套
字段解析逻辑处理剩余文本框和表格单元格。
"""

from __future__ import annotations

import re
from typing import Optional

from .context import RenderContext
from .exceptions import FieldReplaceError
from .models.text import TextReplaceResult
from .parser.template_parser import iter_shapes


class TextReplacer:
    """执行普通文本框和表格单元格中的字段替换。

    示例：
        ```python
        prs = Presentation("examples/assets/text_replace_template.pptx")
        result = TextReplacer().replace_presentation_text(
            prs,
            context=RenderContext(data={"project": {"name": "Aurora"}}),
        )
        ```
    """

    def __init__(self, pattern: Optional[str] = None):
        """创建文本替换器。

        参数：
            pattern: 可选自定义字段正则；为空时使用默认 ``{{path.to.value}}`` 语法。
        """

        self.pattern = pattern

    def replace_presentation_text(
        self,
        presentation,
        context: RenderContext,
        rendered_shape_ids: Optional[set[int]] = None,
        pattern: Optional[str] = None,
    ) -> TextReplaceResult:
        """对整个 Presentation 执行字段替换。

        参数：
            presentation: ``python-pptx`` 的 ``Presentation`` 实例。
            context: 字段取值上下文。
            rendered_shape_ids: 可选 shape id 集合；这些 shape 会被跳过。
            pattern: 本次调用临时覆盖的匹配正则。

        返回：
            ``TextReplaceResult``，包含替换次数与 warning 列表。

        异常：
            FieldReplaceError: 替换过程发生异常时抛出。
        """

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
