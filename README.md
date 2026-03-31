# ppt-template-sdk

`ppt_template_sdk` 是一个可复用的 Python SDK，用于读取 `.pptx` 模板、识别占位块、调用接入方 renderer，并输出渲染后的 PPT。

## 安装

```bash
pip install -e .[dev]
```

## 最小示例

```python
from ppt_template_sdk import (
    EngineOptions,
    PptTemplateEngine,
    PptOperations,
    RenderContext,
    RendererRegistry,
    TextContent,
    TextReplacer,
)

registry = RendererRegistry()

@registry.renderer("title")
def render_title(placeholder, context):
    report = context.extras["report"]
    return TextContent(text=report.title)

engine = PptTemplateEngine(
    registry=registry,
    options=EngineOptions(duplicate_key_policy="broadcast"),
)

result = engine.render(
    template_path="template.pptx",
    output_path="output.pptx",
    context=RenderContext(
        data={"report_date": "2026-03-31"},
        extras={"report": type("Report", (), {"title": "Q1"} )()},
    ),
)

ops = PptOperations.load(template_path="output.pptx")
ops.insert_slide(target_index=1, layout_index=6)
ops.save_to_path("output_with_slide.pptx")
```

## 当前能力

- `ph:<type>:<key>` 占位块识别
- `text` / `image` / `table` / `chart(图片方式)` 渲染
- `TextReplacer` 公共文本替换模块
- `PptOperations` 公共操作模块
- 文本框与表格单元格字段替换
- `validate()` 静态模板校验
- 输出为文件或内存 bytes
