# ppt-template-sdk

`ppt_template_sdk` 是一个可复用的 Python SDK，用于读取 `.pptx` 模板、识别占位块、调用接入方 renderer，并输出渲染后的 PPT。

## 安装

```bash
pip install -e .[dev]
```

要求：

- Python `>= 3.10`
- 依赖 `python-pptx`

## 能力概览

- `PptTemplateEngine`：模板渲染主入口
- `RendererRegistry` / `BaseRenderer`：占位块渲染注册与分发
- `RenderContext`：统一业务数据与扩展对象访问
- `TextReplacer`：独立文本替换能力
- `PptOperations`：PPT 结构与表格操作能力
- `validate()`：模板静态校验

## 模板协议

Shape 级占位块通过 `shape.name` 声明，格式为：

```text
ph:<type>:<key>
```

示例：

```text
ph:text:title
ph:image:logo
ph:table:risk_table
ph:chart:sales_chart
```

普通文本替换字段格式为：

```text
{{field_name}}
{{project.name}}
```

默认执行顺序：

1. 先执行 shape 级占位块渲染
2. 再执行普通文本与表格单元格字段替换

## 5 分钟上手

```python
from dataclasses import dataclass

from ppt_template_sdk import (
    EngineOptions,
    PptTemplateEngine,
    RenderContext,
    RendererRegistry,
    TableContent,
    TextContent,
)

@dataclass
class Report:
    title: str
    risks: list[list[str]]

registry = RendererRegistry()

@registry.renderer("title")
def render_title(placeholder, context):
    report = context.extras["report"]
    return TextContent(text=report.title)

@registry.renderer("risk_table")
def render_risk_table(placeholder, context):
    report = context.extras["report"]
    return TableContent(headers=["风险", "等级"], rows=report.risks)

engine = PptTemplateEngine(
    registry=registry,
    options=EngineOptions(duplicate_key_policy="broadcast"),
)

report = Report(
    title="2026 Q1 经营分析",
    risks=[["现金流", "高"], ["履约", "中"]],
)

result = engine.render(
    template_path="examples/assets/report_template.pptx",
    output_path="examples/output/report_output.pptx",
    context=RenderContext(
        data={
            "report_date": "2026-03-31",
            "project": {"name": "北极星项目"},
            "owner": {"name": "Alice"},
        },
        extras={"report": report},
    ),
)
```

`result.output_bytes` 可直接用于 Web/API 返回；若传了 `output_path`，也会同步落盘。

## 核心接口

### `PptTemplateEngine`

- `render(template_path|template_bytes, output_path=None, context=None) -> RenderResult`
- `validate(template_path|template_bytes) -> ValidationReport`

### `TextReplacer`

适合在不走模板占位块渲染时，单独对现有 PPT 做字段替换：

```python
from pptx import Presentation
from ppt_template_sdk import RenderContext, TextReplacer

prs = Presentation("examples/assets/text_replace_template.pptx")
result = TextReplacer().replace_presentation_text(
    prs,
    context=RenderContext(data={"project": {"name": "Aurora"}}),
)
prs.save("examples/output/text_replaced.pptx")
```

### `PptOperations`

适合渲染后或独立脚本中做结构调整：

```python
from ppt_template_sdk import PptOperations

ops = PptOperations.load(template_path="examples/assets/operations_template.pptx")
ops.insert_slide(target_index=1, layout_index=6)
ops.add_section("正文", start_slide_index=1)
ops.delete_table_row(slide_index=0, shape_locator="ops-table", row_index=1)
ops.save_to_path("examples/output/operations_output.pptx")
```

## 常见约束

- `duplicate_key_policy` 默认是 `broadcast`
- 文本缺失字段会替换为空串，并记录 warning
- `slide_index`、`section_index`、表格行列索引全部是 `0-based`
- `shape_locator` 优先建议传 `shape_id`，也支持 `shape_name`
- 已合并表格不支持删行删列
- `section` 相关能力通过 Open XML 实现，不是 `python-pptx` 官方高层 API

## 文档与示例

- [完整使用指南](docs/usage.md)
- [操作模块与文本替换说明](docs/operations.md)
- [全链路渲染示例](examples/render_report.py)
- [独立文本替换示例](examples/text_replace.py)
- [操作模块示例](examples/operations_demo.py)

## 示例素材

`examples/assets/` 包含可直接运行的模板与图片：

- `report_template.pptx`
- `text_replace_template.pptx`
- `operations_template.pptx`
- `sample_logo.png`
- `sample_chart.png`
