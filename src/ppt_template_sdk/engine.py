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
from .validator import validate_presentation


CONTENT_TYPE_MAP = {
    "text": TextContent,
    "image": ImageContent,
    "table": TableContent,
    "chart": ChartContent,
}


class PptTemplateEngine:
    def __init__(self, registry, options: Optional[EngineOptions] = None):
        self.registry = registry
        self.options = options or EngineOptions()
        self.adapter = PptxAdapter()

    def render(
        self,
        template_path: Optional[str] = None,
        template_bytes: Optional[bytes] = None,
        output_path: Optional[str] = None,
        context: Optional[RenderContext] = None,
    ) -> RenderResult:
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
            warnings.extend(
                self.adapter.replace_text_fields(
                    presentation,
                    rendered_shape_ids,
                    lambda path: context.get_value(path),
                    self.options.text_field_pattern,
                )
            )

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
        presentation = self.adapter.load(template_path=template_path, template_bytes=template_bytes)
        return validate_presentation(presentation, self.registry, self.options)
