from __future__ import annotations

import base64
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import pytest
from pptx import Presentation
from pptx.util import Inches

from ppt_template_sdk import (
    ChartContent,
    ContentTypeMismatchError,
    DuplicatePlaceholderError,
    EngineOptions,
    ImageContent,
    OperationError,
    PptTemplateEngine,
    PptOperations,
    RenderContext,
    RendererRegistry,
    TableContent,
    TextReplacer,
    TextContent,
)
from ppt_template_sdk.registry import BaseRenderer


PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9n6n8AAAAASUVORK5CYII="
)


def _write_png(path: Path) -> None:
    path.write_bytes(PNG_1X1)


def _build_template(path: Path) -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    title.name = "ph:text:title"
    title.text = "placeholder"

    logo = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(2), Inches(1))
    logo.name = "ph:image:logo"
    logo.text = "img"

    chart = slide.shapes.add_textbox(Inches(4), Inches(2), Inches(2), Inches(1))
    chart.name = "ph:chart:sales"
    chart.text = "chart"

    table = slide.shapes.add_textbox(Inches(1), Inches(3.5), Inches(5), Inches(1.5))
    table.name = "ph:table:risk_table"
    table.text = "table"

    field_box = slide.shapes.add_textbox(Inches(1), Inches(5.3), Inches(4), Inches(1))
    field_box.text = "项目：{{project.name}} 日期：{{report_date}} 缺失：{{missing.value}}"

    native_table = slide.shapes.add_table(2, 2, Inches(1), Inches(6.1), Inches(5), Inches(1.5)).table
    native_table.cell(0, 0).text = "负责人"
    native_table.cell(0, 1).text = "{{owner.name}}"
    native_table.cell(1, 0).text = "项目"
    native_table.cell(1, 1).text = "{{project.name}}"

    prs.save(path)


@dataclass
class Report:
    title: str


class TitleRenderer(BaseRenderer):
    supported_types = {"text"}

    def render(self, placeholder, context):
        report = context.extras["report"]
        return TextContent(text=report.title)


def test_render_full_flow(tmp_path: Path):
    template_path = tmp_path / "template.pptx"
    output_path = tmp_path / "output.pptx"
    logo_path = tmp_path / "logo.png"
    chart_path = tmp_path / "chart.png"
    _write_png(logo_path)
    _write_png(chart_path)
    _build_template(template_path)

    registry = RendererRegistry()
    registry.register("title", TitleRenderer())
    registry.register_func("logo", lambda placeholder, context: ImageContent(image_path=str(logo_path)))
    registry.register_func("sales", lambda placeholder, context: ChartContent(image_path=str(chart_path)))
    registry.register_func(
        "risk_table",
        lambda placeholder, context: TableContent(headers=["风险", "等级"], rows=[["现金流", "高"], ["履约", "中"]]),
    )

    engine = PptTemplateEngine(registry, EngineOptions(duplicate_key_policy="broadcast"))
    result = engine.render(
        template_path=str(template_path),
        output_path=str(output_path),
        context=RenderContext(
            data={"project": {"name": "北极星"}, "report_date": "2026-03-31", "owner": {"name": "Alice"}},
            extras={"report": Report(title="Q1 经营分析")},
        ),
    )

    assert result.success is True
    assert result.output_bytes
    assert output_path.exists()
    assert result.rendered_count == 4
    assert any("missing text field 'missing.value'" in warning for warning in result.warnings)

    rendered = Presentation(str(output_path))
    texts = [shape.text for shape in rendered.slides[0].shapes if getattr(shape, "has_text_frame", False)]
    assert "Q1 经营分析" in texts
    assert "项目：北极星 日期：2026-03-31 缺失：" in texts
    assert any(getattr(shape, "has_table", False) for shape in rendered.slides[0].shapes)


def test_render_from_bytes(tmp_path: Path):
    template_path = tmp_path / "template.pptx"
    _build_template(template_path)
    logo_path = tmp_path / "logo.png"
    chart_path = tmp_path / "chart.png"
    _write_png(logo_path)
    _write_png(chart_path)

    template_bytes = template_path.read_bytes()
    registry = RendererRegistry()
    registry.register_func("title", lambda placeholder, context: TextContent(text="Bytes"))
    registry.register_func("logo", lambda placeholder, context: ImageContent(image_path=str(logo_path)))
    registry.register_func("sales", lambda placeholder, context: ChartContent(image_path=str(chart_path)))
    registry.register_func("risk_table", lambda placeholder, context: TableContent(headers=["A"], rows=[["B"]]))

    engine = PptTemplateEngine(registry)
    result = engine.render(template_bytes=template_bytes, context=RenderContext(data={}))

    assert result.success is True
    assert isinstance(result.output_bytes, bytes)


def test_validate_reports_static_issues(tmp_path: Path):
    template_path = tmp_path / "template.pptx"
    _build_template(template_path)
    registry = RendererRegistry()
    registry.register_func("unused", lambda placeholder, context: TextContent(text="x"))

    engine = PptTemplateEngine(registry)
    report = engine.validate(template_path=str(template_path))

    assert report.success is False
    assert report.placeholder_count == 4
    assert any("missing renderer" in error for error in report.errors)
    assert report.unused_renderers == ["unused"]


def test_duplicate_key_error_policy(tmp_path: Path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    first = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    first.name = "ph:text:title"
    second = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(3), Inches(1))
    second.name = "ph:text:title"
    path = tmp_path / "dup.pptx"
    prs.save(path)

    registry = RendererRegistry()
    registry.register_func("title", lambda placeholder, context: TextContent(text="dup"))
    engine = PptTemplateEngine(registry, EngineOptions(duplicate_key_policy="error"))

    with pytest.raises(DuplicatePlaceholderError):
        engine.render(template_path=str(path), context=RenderContext(data={}))


def test_content_type_mismatch(tmp_path: Path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    shape.name = "ph:image:cover"
    path = tmp_path / "mismatch.pptx"
    prs.save(path)

    registry = RendererRegistry()
    registry.register_func("cover", lambda placeholder, context: TextContent(text="wrong"))
    engine = PptTemplateEngine(registry)

    with pytest.raises(ContentTypeMismatchError):
        engine.render(template_path=str(path), context=RenderContext(data={}))


def test_text_replacer_public_api(tmp_path: Path):
    template_path = tmp_path / "template.pptx"
    _build_template(template_path)
    prs = Presentation(str(template_path))
    replacer = TextReplacer()
    result = replacer.replace_presentation_text(
        prs,
        context=RenderContext(data={"project": {"name": "Aurora"}, "report_date": "2026-04-01", "owner": {"name": "Bob"}}),
    )

    texts = [shape.text for shape in prs.slides[0].shapes if getattr(shape, "has_text_frame", False)]
    assert "项目：Aurora 日期：2026-04-01 缺失：" in texts
    assert result.replaced_count >= 4
    assert any("missing text field 'missing.value'" in warning for warning in result.warnings)


def test_operations_slide_and_section_flow(tmp_path: Path):
    prs = Presentation()
    for idx in range(3):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1)).text = f"slide-{idx}"
    path = tmp_path / "ops.pptx"
    prs.save(path)

    ops = PptOperations.load(template_path=str(path))
    ops.add_section("Content", 1)
    ops.insert_slide(target_index=1, layout_index=6)
    ops.delete_slide(3)
    ops.delete_section(0)

    output = tmp_path / "ops-out.pptx"
    ops.save_to_path(str(output))
    rendered = Presentation(str(output))
    assert len(rendered.slides) == 3
    slide0_texts = [shape.text for shape in rendered.slides[0].shapes if getattr(shape, "has_text_frame", False)]
    slide1_texts = [shape.text for shape in rendered.slides[1].shapes if getattr(shape, "has_text_frame", False)]
    assert slide0_texts == ["slide-0"]
    assert slide1_texts == []

    with ZipFile(output) as zf:
        presentation_xml = zf.read("ppt/presentation.xml").decode("utf-8")
    assert 'name="Content"' in presentation_xml
    assert "sectionLst" in presentation_xml
    assert presentation_xml.count("<p:sldId ") == 3


def test_operations_table_row_column_and_merge(tmp_path: Path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_table(3, 3, Inches(1), Inches(1), Inches(4), Inches(2))
    shape.name = "ops-table"
    table = shape.table
    for row_index in range(3):
        for col_index in range(3):
            table.cell(row_index, col_index).text = f"{row_index},{col_index}"
    path = tmp_path / "table-ops.pptx"
    prs.save(path)

    ops = PptOperations.load(template_path=str(path))
    ops.delete_table_row(0, "ops-table", 1)
    ops.delete_table_column(0, "ops-table", 1)
    ops.merge_table_cells(0, "ops-table", 0, 0, 1, 1)

    output_bytes = ops.save_to_bytes()
    rendered = Presentation(BytesIO(output_bytes))
    shape = rendered.slides[0].shapes[0]
    table = shape.table
    assert len(table.rows) == 2
    assert len(table.columns) == 2
    assert table.cell(0, 0).is_merge_origin is True


def test_row_column_delete_rejects_merged_tables(tmp_path: Path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(4), Inches(2))
    shape.name = "merged-table"
    shape.table.cell(0, 0).merge(shape.table.cell(1, 1))
    path = tmp_path / "merged.pptx"
    prs.save(path)

    ops = PptOperations.load(template_path=str(path))
    with pytest.raises(OperationError):
        ops.delete_table_row(0, "merged-table", 0)
