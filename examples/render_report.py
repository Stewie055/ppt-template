from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT / "singlefile") not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT / "singlefile"))

from ppt_template_sdk import (  # noqa: E402
    ChartContent,
    EngineOptions,
    ImageContent,
    PptTemplateEngine,
    RenderContext,
    RendererRegistry,
    TableContent,
    TextContent,
)


ASSETS_DIR = Path(__file__).resolve().parent / "assets"
OUTPUT_DIR = Path(__file__).resolve().parent / "output"


@dataclass
class Report:
    title: str
    logo_path: str
    chart_path: str
    risks: list[list[str]]


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)

    report = Report(
        title="2026 Q1 经营分析",
        logo_path=str(ASSETS_DIR / "sample_logo.png"),
        chart_path=str(ASSETS_DIR / "sample_chart.png"),
        risks=[["现金流", "高"], ["履约", "中"], ["供应链", "低"]],
    )

    registry = RendererRegistry()

    @registry.renderer("title")
    def render_title(placeholder, context):
        report_obj = context.extras["report"]
        return TextContent(text=report_obj.title)

    @registry.renderer("logo")
    def render_logo(placeholder, context):
        report_obj = context.extras["report"]
        return ImageContent(image_path=report_obj.logo_path)

    @registry.renderer("sales_chart")
    def render_chart(placeholder, context):
        report_obj = context.extras["report"]
        return ChartContent(image_path=report_obj.chart_path)

    @registry.renderer("risk_table")
    def render_table(placeholder, context):
        report_obj = context.extras["report"]
        return TableContent(headers=["风险", "等级"], rows=report_obj.risks)

    engine = PptTemplateEngine(
        registry=registry,
        options=EngineOptions(
            duplicate_key_policy="broadcast",
            enable_text_field_replace=True,
            strict=True,
        ),
    )

    result = engine.render(
        template_path=str(ASSETS_DIR / "report_template.pptx"),
        output_path=str(OUTPUT_DIR / "report_output.pptx"),
        context=RenderContext(
            data={
                "project": {"name": "北极星项目"},
                "owner": {"name": "Alice"},
                "report_date": "2026-03-31",
            },
            extras={"report": report},
        ),
    )

    print("rendered_count:", result.rendered_count)
    print("warnings:", result.warnings)
    print("output:", OUTPUT_DIR / "report_output.pptx")


if __name__ == "__main__":
    main()
