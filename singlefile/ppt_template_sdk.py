"""单文件版 `ppt_template_sdk`。

该文件可直接复制到用户项目中使用，导入方式保持：

```python
from ppt_template_sdk import PptTemplateEngine, RendererRegistry
```

它覆盖当前包版的全部公开能力，但不依赖仓库内其他本地模块。
"""

from __future__ import annotations

import re
import uuid
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Any, Callable, Iterable, Optional, Union

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


_MISSING = object()
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
DEFAULT_SECTION_NAME = "Section 1"
PLACEHOLDER_RE = re.compile(r"^ph:(text|image|table|chart):([\w\.-]+)$")


class PptTemplateSdkError(Exception):
    """SDK 所有自定义异常的基类。"""


class TemplateParseError(PptTemplateSdkError):
    """模板解析阶段发生错误时抛出。"""


class PlaceholderFormatError(TemplateParseError):
    """占位块命名不符合 ``ph:<type>:<key>`` 规范时抛出。"""


class DuplicatePlaceholderError(PptTemplateSdkError):
    """检测到重复占位块 key 且当前策略不允许时抛出。"""


class RendererNotFoundError(PptTemplateSdkError):
    """模板中的占位块缺少对应 renderer 时抛出。"""


class ContentTypeMismatchError(PptTemplateSdkError):
    """renderer 返回的内容类型与占位块类型不匹配时抛出。"""


class ShapeOperationError(PptTemplateSdkError):
    """对底层 shape 执行写回或定位操作失败时抛出。"""


class FieldReplaceError(PptTemplateSdkError):
    """文本字段替换阶段发生异常时抛出。"""


class OperationError(PptTemplateSdkError):
    """执行 slide、section 或表格结构操作失败时抛出。"""


@dataclass(slots=True)
class RenderContext:
    """承载模板渲染和文本替换所需的业务上下文。"""

    data: Any
    extras: dict[str, Any] = field(default_factory=dict)

    def get_value(self, path: str, default: Any = None) -> Any:
        """按点路径读取 ``data`` 中的值。"""

        current = self.data
        for part in path.split("."):
            current = self._resolve_part(current, part, _MISSING)
            if current is _MISSING:
                return default
        return current

    def has_value(self, path: str) -> bool:
        """判断给定点路径是否存在有效值。"""

        return self.get_value(path, _MISSING) is not _MISSING

    @staticmethod
    def _resolve_part(value: Any, part: str, default: Any) -> Any:
        if value is None:
            return default
        if isinstance(value, dict):
            return value.get(part, default)
        if isinstance(value, (list, tuple)) and part.isdigit():
            index = int(part)
            if 0 <= index < len(value):
                return value[index]
            return default
        if hasattr(value, part):
            return getattr(value, part)
        if hasattr(value, "model_dump"):
            return value.model_dump().get(part, default)
        if hasattr(value, "dict") and callable(value.dict):
            return value.dict().get(part, default)
        return default


@dataclass(slots=True)
class EngineOptions:
    """控制渲染引擎行为的配置对象。"""

    duplicate_key_policy: str = "broadcast"
    enable_text_field_replace: bool = True
    text_field_pattern: str = r"\{\{([\w\.]+)\}\}"
    text_field_replace_mode: str = "plain"
    strict: bool = True


class Content:
    """所有渲染内容类型的基类。"""


@dataclass(slots=True)
class TextContent(Content):
    """文本占位块的渲染结果。"""

    text: str


@dataclass(slots=True)
class ImageContent(Content):
    """图片占位块的渲染结果。"""

    image_path: str


@dataclass(slots=True)
class TableContent(Content):
    """表格占位块的渲染结果。"""

    headers: list[str]
    rows: list[list[str]]


@dataclass(slots=True)
class ChartContent(Content):
    """图表占位块的渲染结果。"""

    image_path: str


@dataclass(slots=True)
class Placeholder:
    """模板占位块的标准化描述对象。"""

    type: str
    key: str
    slide_index: int
    shape_id: int
    shape_name: str
    left: int
    top: int
    width: int
    height: int
    shape: Any


@dataclass(slots=True)
class RenderResult:
    """一次模板渲染的结构化返回结果。"""

    success: bool
    output_path: Optional[str] = None
    output_bytes: Optional[bytes] = None
    rendered_count: int = 0
    skipped_count: int = 0
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ValidationReport:
    """静态模板校验结果。"""

    success: bool
    placeholder_count: int = 0
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    unused_renderers: list[str] = field(default_factory=list)


@dataclass(slots=True)
class TextReplaceResult:
    """独立文本替换操作的返回结果。"""

    replaced_count: int = 0
    warnings: list[str] = field(default_factory=list)


class BaseRenderer:
    """自定义 renderer 的基类。"""

    supported_types: set[str] = set()

    def render(self, placeholder: Placeholder, context: RenderContext):
        """根据占位块和上下文生成渲染结果。"""

        raise NotImplementedError


class _FunctionRenderer(BaseRenderer):
    def __init__(self, func: Callable):
        self._func = func

    def render(self, placeholder: Placeholder, context: RenderContext):
        return self._func(placeholder, context)


class RendererRegistry:
    """管理模板占位块 renderer 的注册与查询。"""

    def __init__(self) -> None:
        self._renderers: dict[str, BaseRenderer] = {}

    def register(self, key: str, renderer: BaseRenderer) -> None:
        """注册类式 renderer。"""

        self._renderers[key] = renderer

    def register_func(self, key: str, func: Callable) -> None:
        """注册函数式 renderer。"""

        self.register(key, _FunctionRenderer(func))

    def renderer(self, key: str):
        """返回装饰器，用于以声明式方式注册 renderer。"""

        def decorator(func: Callable):
            self.register_func(key, func)
            return func

        return decorator

    def get(self, key: str):
        """按 key 获取 renderer；若不存在则返回 ``None``。"""

        return self._renderers.get(key)

    def keys(self) -> list[str]:
        """返回已注册的全部 key，按字典序排序。"""

        return sorted(self._renderers.keys())


class PptxAdapter:
    """单文件版内部使用的 Presentation 读写适配层。"""

    def load(self, template_path: Optional[str] = None, template_bytes: Optional[bytes] = None) -> Presentation:
        if bool(template_path) == bool(template_bytes):
            raise ValueError("exactly one of template_path or template_bytes must be provided")
        if template_path:
            return Presentation(template_path)
        from io import BytesIO

        return Presentation(BytesIO(template_bytes))

    @staticmethod
    def save_to_bytes(presentation) -> bytes:
        from io import BytesIO

        buffer = BytesIO()
        presentation.save(buffer)
        return buffer.getvalue()

    @staticmethod
    def save_to_path(presentation, output_path: str) -> None:
        presentation.save(output_path)

    @staticmethod
    def get_slide(presentation, slide_index: int):
        if slide_index < 0 or slide_index >= len(presentation.slides):
            raise ShapeOperationError(f"slide index {slide_index} out of range")
        return presentation.slides[slide_index]

    @staticmethod
    def find_shape(slide, shape_locator: Union[int, str]):
        if isinstance(shape_locator, int):
            for shape in slide.shapes:
                if getattr(shape, "shape_id", None) == shape_locator:
                    return shape
            raise ShapeOperationError(f"shape id {shape_locator} not found")
        for shape in slide.shapes:
            if getattr(shape, "name", None) == shape_locator:
                return shape
        raise ShapeOperationError(f"shape '{shape_locator}' not found")

    @staticmethod
    def write_content(presentation, placeholder: Placeholder, content: Content) -> None:
        slide = presentation.slides[placeholder.slide_index]
        shape = placeholder.shape
        if isinstance(content, TextContent):
            if not getattr(shape, "has_text_frame", False):
                raise ShapeOperationError(f"shape '{placeholder.shape_name}' cannot accept text content")
            shape.text = content.text
            return
        if isinstance(content, (ImageContent, ChartContent)):
            slide.shapes.add_picture(content.image_path, placeholder.left, placeholder.top, placeholder.width, placeholder.height)
            PptxAdapter._remove_shape(shape)
            return
        if isinstance(content, TableContent):
            if getattr(shape, "has_table", False) and PptxAdapter._rewrite_table(shape.table, content):
                return
            rows, cols, grid = PptxAdapter._build_table_grid(content)
            table_shape = slide.shapes.add_table(rows, cols, placeholder.left, placeholder.top, placeholder.width, placeholder.height)
            table = table_shape.table
            for row_index, row_values in enumerate(grid):
                for col_index, value in enumerate(row_values):
                    table.cell(row_index, col_index).text = value
            PptxAdapter._remove_shape(shape)
            return
        raise ShapeOperationError(f"unsupported content type '{type(content).__name__}'")

    @staticmethod
    def _build_table_grid(content: TableContent) -> tuple[int, int, list[list[str]]]:
        grid: list[list[str]] = []
        if content.headers:
            grid.append([str(value) for value in content.headers])
        for row in content.rows:
            grid.append([str(value) for value in row])
        cols = max((len(row) for row in grid), default=1)
        normalized = [row + [""] * (cols - len(row)) for row in grid] or [[""]]
        return len(normalized), cols, normalized

    @staticmethod
    def _rewrite_table(table, content: TableContent) -> bool:
        rows, cols, grid = PptxAdapter._build_table_grid(content)
        if len(table.rows) != rows or len(table.columns) != cols:
            return False
        for row_index, row_values in enumerate(grid):
            for col_index, value in enumerate(row_values):
                table.cell(row_index, col_index).text = value
        return True

    @staticmethod
    def _remove_shape(shape) -> None:
        parent = shape.element.getparent()
        if parent is None:
            raise ShapeOperationError("unable to remove placeholder shape after overlay insertion")
        parent.remove(shape.element)


@dataclass(slots=True)
class ParsedTemplate:
    placeholders: list[Placeholder] = field(default_factory=list)
    invalid_placeholders: list[str] = field(default_factory=list)
    text_field_paths: set[str] = field(default_factory=set)


def iter_shapes(shape_collection) -> Iterable:
    """遍历 slide 中的 shape，并递归展开 group shape。"""

    for shape in shape_collection:
        yield shape
        if getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)


def parse_presentation(presentation, text_field_pattern: str) -> ParsedTemplate:
    """扫描模板中的占位块和文本字段。"""

    parsed = ParsedTemplate()
    field_re = re.compile(text_field_pattern)
    for slide_index, slide in enumerate(presentation.slides):
        for shape in iter_shapes(slide.shapes):
            name = getattr(shape, "name", "") or ""
            if name.startswith("ph:"):
                match = PLACEHOLDER_RE.match(name)
                if not match:
                    parsed.invalid_placeholders.append(
                        f"slide {slide_index + 1} shape {getattr(shape, 'shape_id', '?')}: invalid placeholder name '{name}'"
                    )
                else:
                    parsed.placeholders.append(
                        Placeholder(
                            type=match.group(1),
                            key=match.group(2),
                            slide_index=slide_index,
                            shape_id=getattr(shape, "shape_id", -1),
                            shape_name=name,
                            left=getattr(shape, "left", 0),
                            top=getattr(shape, "top", 0),
                            width=getattr(shape, "width", 0),
                            height=getattr(shape, "height", 0),
                            shape=shape,
                        )
                    )
            if getattr(shape, "has_text_frame", False):
                parsed.text_field_paths.update(field_re.findall(shape.text_frame.text))
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        parsed.text_field_paths.update(field_re.findall(cell.text))
    return parsed


def validate_presentation(presentation, registry, options) -> ValidationReport:
    """对模板执行静态校验。"""

    parsed = parse_presentation(presentation, options.text_field_pattern)
    errors = list(parsed.invalid_placeholders)
    warnings: list[str] = []

    key_counter = Counter(placeholder.key for placeholder in parsed.placeholders)
    for key, count in sorted(key_counter.items()):
        if count < 2:
            continue
        message = f"duplicate placeholder key '{key}' appears {count} times"
        if options.duplicate_key_policy == "error":
            errors.append(message)
        else:
            warnings.append(message)

    used_keys = set()
    for placeholder in parsed.placeholders:
        renderer = registry.get(placeholder.key)
        if renderer is None:
            errors.append(
                f"missing renderer for placeholder '{placeholder.shape_name}' on slide {placeholder.slide_index + 1}"
            )
            continue
        used_keys.add(placeholder.key)
        supported_types = getattr(renderer, "supported_types", set()) or set()
        if supported_types and placeholder.type not in supported_types:
            errors.append(f"renderer '{placeholder.key}' does not support placeholder type '{placeholder.type}'")

    unused_renderers = sorted(set(registry.keys()) - used_keys)
    if unused_renderers:
        warnings.append(f"unused renderers: {', '.join(unused_renderers)}")

    return ValidationReport(
        success=not errors,
        placeholder_count=len(parsed.placeholders),
        errors=errors,
        warnings=warnings,
        unused_renderers=unused_renderers,
    )


class TextReplacer:
    """执行普通文本框和表格单元格中的字段替换。"""

    def __init__(self, pattern: Optional[str] = None):
        self.pattern = pattern

    def replace_presentation_text(
        self,
        presentation,
        context: RenderContext,
        rendered_shape_ids: Optional[set[int]] = None,
        pattern: Optional[str] = None,
    ) -> TextReplaceResult:
        """对整个 Presentation 执行字段替换。"""

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


CONTENT_TYPE_MAP = {
    "text": TextContent,
    "image": ImageContent,
    "table": TableContent,
    "chart": ChartContent,
}


class PptTemplateEngine:
    """PPT 模板渲染主引擎。"""

    def __init__(self, registry, options: Optional[EngineOptions] = None):
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
        """执行模板渲染。"""

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
    ) -> ValidationReport:
        """对模板执行静态校验。"""

        presentation = self.adapter.load(template_path=template_path, template_bytes=template_bytes)
        return validate_presentation(presentation, self.registry, self.options)


def _tag(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


class PptOperations:
    """封装常见的 PPT 结构与表格操作。"""

    def __init__(self, presentation, adapter: Optional[PptxAdapter] = None):
        self.presentation = presentation
        self.adapter = adapter or PptxAdapter()

    @classmethod
    def load(cls, template_path: Optional[str] = None, template_bytes: Optional[bytes] = None):
        """从路径或字节流加载 PPT 并创建操作对象。"""

        adapter = PptxAdapter()
        return cls(adapter.load(template_path=template_path, template_bytes=template_bytes), adapter=adapter)

    def save_to_bytes(self) -> bytes:
        """将当前 Presentation 保存为内存字节流。"""

        return self.adapter.save_to_bytes(self.presentation)

    def save_to_path(self, output_path: str) -> None:
        """将当前 Presentation 保存到指定路径。"""

        self.adapter.save_to_path(self.presentation, output_path)

    def delete_slide(self, slide_index: int) -> int:
        """删除指定索引的 slide。"""

        slide_id_el = self._get_slide_id_element(slide_index)
        slide_id = int(slide_id_el.get("id"))
        rel_id = slide_id_el.get(_tag(R_NS, "id"))
        self.presentation.part.drop_rel(rel_id)
        self.presentation.slides._sldIdLst.remove(slide_id_el)
        if self._has_sections():
            groups = self._read_sections()
            for group in groups:
                group["slide_ids"] = [item for item in group["slide_ids"] if item != slide_id]
            self._write_sections([group for group in groups if group["slide_ids"]])
        return slide_id

    def insert_slide(self, target_index: int, layout_index: int):
        """使用模板现有 ``layout_index`` 新建并插入 slide。"""

        if target_index < 0 or target_index > len(self.presentation.slides):
            raise OperationError(f"target_index {target_index} out of range")
        if layout_index < 0 or layout_index >= len(self.presentation.slide_layouts):
            raise OperationError(f"layout_index {layout_index} out of range")

        previous_slide_ids = self._slide_ids_in_order()
        groups = self._read_sections()
        slide = self.presentation.slides.add_slide(self.presentation.slide_layouts[layout_index])
        slide_id = slide.slide_id
        slide_id_el = self.presentation.slides._sldIdLst[-1]
        self.presentation.slides._sldIdLst.remove(slide_id_el)
        self.presentation.slides._sldIdLst.insert(target_index, slide_id_el)

        if groups:
            if target_index >= len(previous_slide_ids):
                target_group = len(groups) - 1
                groups[target_group]["slide_ids"].append(slide_id)
            else:
                anchor_slide_id = previous_slide_ids[target_index]
                target_group = self._group_index_for_slide(groups, anchor_slide_id)
                anchor_pos = groups[target_group]["slide_ids"].index(anchor_slide_id)
                groups[target_group]["slide_ids"].insert(anchor_pos, slide_id)
            self._write_sections(groups)
        return slide

    def add_section(self, name: str, start_slide_index: int) -> None:
        """在指定 slide 位置开始一个新的 section。"""

        slide_ids = self._slide_ids_in_order()
        if not slide_ids:
            raise OperationError("cannot add a section to an empty presentation")
        if start_slide_index < 0 or start_slide_index >= len(slide_ids):
            raise OperationError(f"start_slide_index {start_slide_index} out of range")

        start_slide_id = slide_ids[start_slide_index]
        groups = self._read_sections()
        if not groups:
            before = slide_ids[:start_slide_index]
            after = slide_ids[start_slide_index:]
            groups = []
            if before:
                groups.append(self._make_section(DEFAULT_SECTION_NAME, before))
            groups.append(self._make_section(name, after))
            self._write_sections(groups)
            return

        group_index = self._group_index_for_slide(groups, start_slide_id)
        group = groups[group_index]
        first_slide_id = group["slide_ids"][0]
        if first_slide_id == start_slide_id:
            group["name"] = name
            self._write_sections(groups)
            return

        split_at = group["slide_ids"].index(start_slide_id)
        before = group["slide_ids"][:split_at]
        after = group["slide_ids"][split_at:]
        groups[group_index] = self._make_section(group["name"], before, guid=group["id"])
        groups.insert(group_index + 1, self._make_section(name, after))
        self._write_sections(groups)

    def delete_section(self, section_index: int) -> None:
        """删除指定 section，但保留其中 slides。"""

        groups = self._read_sections()
        if not groups:
            raise OperationError("presentation does not contain sections")
        if section_index < 0 or section_index >= len(groups):
            raise OperationError(f"section_index {section_index} out of range")
        if len(groups) == 1:
            self._write_sections([])
            return

        removed = groups.pop(section_index)
        if section_index == 0:
            groups[0]["slide_ids"] = removed["slide_ids"] + groups[0]["slide_ids"]
        else:
            groups[section_index - 1]["slide_ids"].extend(removed["slide_ids"])
        self._write_sections(groups)

    def delete_table_row(self, slide_index: int, shape_locator: Union[int, str], row_index: int) -> None:
        """删除指定表格中的一行。"""

        table = self._resolve_table(slide_index, shape_locator)
        self._ensure_unmerged_table(table)
        if row_index < 0 or row_index >= len(table.rows):
            raise OperationError(f"row_index {row_index} out of range")
        table._tbl.remove(table._tbl.tr_lst[row_index])

    def delete_table_column(self, slide_index: int, shape_locator: Union[int, str], column_index: int) -> None:
        """删除指定表格中的一列。"""

        table = self._resolve_table(slide_index, shape_locator)
        self._ensure_unmerged_table(table)
        if column_index < 0 or column_index >= len(table.columns):
            raise OperationError(f"column_index {column_index} out of range")
        table._tbl.tblGrid.remove(table._tbl.tblGrid.gridCol_lst[column_index])
        for tr in table._tbl.tr_lst:
            tr.remove(tr.tc_lst[column_index])

    def merge_table_cells(
        self,
        slide_index: int,
        shape_locator: Union[int, str],
        first_row: int,
        first_col: int,
        last_row: int,
        last_col: int,
    ) -> None:
        """合并指定矩形区域内的表格单元格。"""

        table = self._resolve_table(slide_index, shape_locator)
        self._validate_merge_bounds(table, first_row, first_col, last_row, last_col)
        table.cell(first_row, first_col).merge(table.cell(last_row, last_col))

    def _resolve_table(self, slide_index: int, shape_locator: Union[int, str]):
        slide = self.adapter.get_slide(self.presentation, slide_index)
        shape = self.adapter.find_shape(slide, shape_locator)
        if not getattr(shape, "has_table", False):
            raise ShapeOperationError("target shape is not a table")
        return shape.table

    @staticmethod
    def _validate_merge_bounds(table, first_row: int, first_col: int, last_row: int, last_col: int) -> None:
        if first_row > last_row or first_col > last_col:
            raise OperationError("merge bounds must define a top-left to bottom-right rectangle")
        if first_row < 0 or first_col < 0 or last_row >= len(table.rows) or last_col >= len(table.columns):
            raise OperationError("merge bounds out of range")

    @staticmethod
    def _ensure_unmerged_table(table) -> None:
        for tr in table._tbl.tr_lst:
            for tc in tr.tc_lst:
                if tc.get("rowSpan") or tc.get("gridSpan") or tc.get("hMerge") or tc.get("vMerge"):
                    raise OperationError("row/column deletion is not supported on merged tables")

    def _get_slide_id_element(self, slide_index: int):
        slide_ids = list(self.presentation.slides._sldIdLst)
        if slide_index < 0 or slide_index >= len(slide_ids):
            raise OperationError(f"slide_index {slide_index} out of range")
        return slide_ids[slide_index]

    def _slide_ids_in_order(self) -> list[int]:
        return [slide.slide_id for slide in self.presentation.slides]

    def _group_index_for_slide(self, groups: list[dict], slide_id: int) -> int:
        for index, group in enumerate(groups):
            if slide_id in group["slide_ids"]:
                return index
        raise OperationError(f"slide id {slide_id} is not assigned to any section")

    @staticmethod
    def _make_section(name: str, slide_ids: list[int], guid: Optional[str] = None) -> dict:
        return {"name": name, "id": guid or f"{{{str(uuid.uuid4()).upper()}}}", "slide_ids": slide_ids}

    def _has_sections(self) -> bool:
        return self._find_section_ext() is not None

    def _read_sections(self) -> list[dict]:
        section_ext = self._find_section_ext()
        if section_ext is None:
            return []
        section_lst = section_ext.find(_tag(P14_NS, "sectionLst"))
        if section_lst is None:
            return []
        groups = []
        for section in section_lst.findall(_tag(P14_NS, "section")):
            slide_ids = [int(sld_id.get("id")) for sld_id in section.find(_tag(P14_NS, "sldIdLst")).findall(_tag(P14_NS, "sldId"))]
            groups.append(
                {
                    "name": section.get("name") or DEFAULT_SECTION_NAME,
                    "id": section.get("id") or f"{{{str(uuid.uuid4()).upper()}}}",
                    "slide_ids": slide_ids,
                }
            )
        return groups

    def _write_sections(self, groups: list[dict]) -> None:
        presentation_el = self.presentation.part._element
        ext_lst = presentation_el.find(_tag(P_NS, "extLst"))
        section_ext = self._find_section_ext()
        if not groups:
            if section_ext is not None:
                ext_lst.remove(section_ext)
            if ext_lst is not None and len(ext_lst) == 0:
                presentation_el.remove(ext_lst)
            return

        ordered_ids = self._slide_ids_in_order()
        order_map = {slide_id: index for index, slide_id in enumerate(ordered_ids)}
        normalized_groups = []
        for group in groups:
            slide_ids = sorted([slide_id for slide_id in group["slide_ids"] if slide_id in order_map], key=order_map.__getitem__)
            if slide_ids:
                normalized_groups.append({**group, "slide_ids": slide_ids})

        if ext_lst is None:
            ext_lst = etree.SubElement(presentation_el, _tag(P_NS, "extLst"))
        if section_ext is None:
            section_ext = etree.SubElement(ext_lst, _tag(P_NS, "ext"), uri=SECTION_EXT_URI)
        else:
            for child in list(section_ext):
                section_ext.remove(child)

        section_lst = etree.SubElement(section_ext, _tag(P14_NS, "sectionLst"), nsmap={"p14": P14_NS})
        for group in normalized_groups:
            section_el = etree.SubElement(section_lst, _tag(P14_NS, "section"), name=group["name"], id=group["id"])
            slide_id_lst = etree.SubElement(section_el, _tag(P14_NS, "sldIdLst"))
            for slide_id in group["slide_ids"]:
                etree.SubElement(slide_id_lst, _tag(P14_NS, "sldId"), id=str(slide_id))

    def _find_section_ext(self):
        presentation_el = self.presentation.part._element
        ext_lst = presentation_el.find(_tag(P_NS, "extLst"))
        if ext_lst is None:
            return None
        for ext in ext_lst.findall(_tag(P_NS, "ext")):
            if ext.get("uri") == SECTION_EXT_URI:
                return ext
        return None


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
