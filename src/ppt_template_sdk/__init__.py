"""`ppt_template_sdk` 对外公开的稳定接口入口。

调用方通常只需要从这个模块导入引擎、内容模型、操作模块和异常类型，而无需
直接了解内部子模块结构。
"""

from .context import RenderContext
from .engine import PptTemplateEngine
from .exceptions import (
    ContentTypeMismatchError,
    DuplicatePlaceholderError,
    FieldReplaceError,
    OperationError,
    PlaceholderFormatError,
    PptTemplateSdkError,
    RendererNotFoundError,
    ShapeOperationError,
    TemplateParseError,
)
from .models.content import ChartContent, Content, ImageContent, TableContent, TextContent
from .models.placeholder import Placeholder
from .models.report import ValidationReport
from .models.result import RenderResult
from .models.text import TextReplaceResult
from .options import EngineOptions
from .operations import PptOperations
from .registry import BaseRenderer, RendererRegistry
from .text_replacer import TextReplacer

__all__ = [
    "BaseRenderer",
    "ChartContent",
    "Content",
    "ContentTypeMismatchError",
    "DuplicatePlaceholderError",
    "EngineOptions",
    "FieldReplaceError",
    "ImageContent",
    "OperationError",
    "Placeholder",
    "PlaceholderFormatError",
    "PptTemplateEngine",
    "PptOperations",
    "PptTemplateSdkError",
    "RenderContext",
    "RenderResult",
    "RendererNotFoundError",
    "RendererRegistry",
    "ShapeOperationError",
    "TableContent",
    "TemplateParseError",
    "TextReplaceResult",
    "TextReplacer",
    "TextContent",
    "ValidationReport",
]
