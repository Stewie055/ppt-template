"""PptTemplateEngine 是 SDK 的模板渲染主入口。

该模块负责加载模板、解析占位块、调度 renderer、执行文本替换，并输出
最终的 PPT 文件或内存字节流。
"""

from __future__ import annotations

from collections import defaultdict
from typing import Optional

from .adapter.pptx_adapter import PptxAdapter
from .context import RenderContext
from .exceptions import (
    ContentTypeMismatchError,
    DuplicatePlaceholderError,
    PlaceholderFormatError,
    RendererNotFoundError,
)
from .models.content import ChartContent, ImageContent, TableContent, TextContent
from .models.result import RenderResult
from .options import EngineOptions
from .parser.template_parser import parse_presentation
from .text_replacer import TextReplacer
from .validator import validate_presentation


CONTENT_TYPE_MAP = {
    "text": TextContent,
    "image": ImageContent,
    "table": TableContent,
    "chart": ChartContent,
}


class PptTemplateEngine:
    """PPT 模板渲染主引擎。

    示例：
        ```python
        engine = PptTemplateEngine(registry=registry)
        result = engine.render(
            template_path="examples/assets/report_template.pptx",
            output_path="examples/output/report_output.pptx",
            context=context,
        )
        ```
    """

    def __init__(self, registry, options: Optional[EngineOptions] = None):
        """初始化渲染引擎。

        参数：
            registry: 已注册 renderer 的 ``RendererRegistry``。
            options: 可选的引擎配置；为空时使用默认 ``EngineOptions``。
        """

        self.registry = registry
        self.options = options or EngineOptions()
        self.adapter = PptxAdapter()
        self.text_replacer = TextReplacer(self.options.text_field_pattern)

    def render(
        self,
        template_path: Optional[str] = None,
        template_bytes: Optional[bytes] = None,
        output_path: Optional[str] = None,
        context: Optional[RenderContext] = None,
    ) -> RenderResult:
        """执行模板渲染。

        参数：
            template_path: 模板文件路径，与 ``template_bytes`` 二选一。
            template_bytes: 模板字节流，与 ``template_path`` 二选一。
            output_path: 可选输出路径；传入后会在返回 bytes 的同时落盘。
            context: 渲染上下文；为空时使用空字典上下文。

        返回：
            ``RenderResult``，包含输出 bytes、渲染计数和 warning。

        异常：
            PlaceholderFormatError: 模板中存在非法占位块命名。
            DuplicatePlaceholderError: 重复 key 且策略为 ``error``。
            RendererNotFoundError: 模板存在未注册 renderer 的占位块。
            ContentTypeMismatchError: renderer 返回类型与占位块类型不匹配。
        """

        context = context or RenderContext(data={})
        presentation = self.adapter.load(template_path=template_path, template_bytes=template_bytes)
        parsed = parse_presentation(presentation, self.options.text_field_pattern)
        if parsed.invalid_placeholders:
            raise PlaceholderFormatError(parsed.invalid_placeholders[0])

        placeholder_groups = defaultdict(list)
        for placeholder in parsed.placeholders:
            placeholder_groups[placeholder.key].append(placeholder)

        rendered_shape_ids: set[int] = set()
        rendered_count = 0
        warnings: list[str] = []

        for key, placeholders in placeholder_groups.items():
            if len(placeholders) > 1 and self.options.duplicate_key_policy == "error":
                raise DuplicatePlaceholderError(f"duplicate placeholder key '{key}' appears {len(placeholders)} times")
            renderer = self.registry.get(key)
            if renderer is None:
                raise RendererNotFoundError(f"missing renderer for placeholder key '{key}'")
            content = renderer.render(placeholders[0], context)
            expected_type = CONTENT_TYPE_MAP[placeholders[0].type]
            if not isinstance(content, expected_type):
                raise ContentTypeMismatchError(
                    f"renderer '{key}' returned {type(content).__name__}, expected {expected_type.__name__}"
                )
            for placeholder in placeholders:
                self.adapter.write_content(presentation, placeholder, content)
                rendered_shape_ids.add(placeholder.shape_id)
                rendered_count += 1

        if self.options.enable_text_field_replace:
            replace_result = self.text_replacer.replace_presentation_text(
                presentation,
                context=context,
                rendered_shape_ids=rendered_shape_ids,
                pattern=self.options.text_field_pattern,
            )
            warnings.extend(replace_result.warnings)

        output_bytes = self.adapter.save_to_bytes(presentation)
        if output_path:
            self.adapter.save_to_path(presentation, output_path)

        return RenderResult(
            success=True,
            output_path=output_path,
            output_bytes=output_bytes,
            rendered_count=rendered_count,
            skipped_count=0,
            warnings=warnings,
        )

    def validate(
        self,
        template_path: Optional[str] = None,
        template_bytes: Optional[bytes] = None,
    ):
        """对模板执行静态校验。

        参数：
            template_path: 模板文件路径，与 ``template_bytes`` 二选一。
            template_bytes: 模板字节流，与 ``template_path`` 二选一。

        返回：
            ``ValidationReport``，用于查看错误、warning 和未使用 renderer。
        """

        presentation = self.adapter.load(template_path=template_path, template_bytes=template_bytes)
        return validate_presentation(presentation, self.registry, self.options)
