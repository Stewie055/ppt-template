# 使用指南

本文面向接入 `ppt_template_sdk` 的调用方，说明如何设计模板、注册 renderer、执行渲染与校验。

## 1. 基本流程

1. 准备 `.pptx` 模板
2. 在模板中给 shape 设置 `ph:<type>:<key>`
3. 注册 renderer
4. 组装 `RenderContext`
5. 调用 `render()`
6. 读取 `RenderResult` 或输出文件

## 2. 模板设计

### Shape 级占位块

形如：

```text
ph:text:title
ph:image:logo
ph:table:risk_table
ph:chart:sales_chart
```

规则：

- `type` 目前支持 `text`、`image`、`table`、`chart`
- `key` 需要和 `RendererRegistry` 中注册的 key 对齐
- SDK 依赖 `shape.name`，不依赖 shape 原始文本

### 文本字段

文本框和表格单元格内支持：

```text
{{report_date}}
{{project.name}}
{{owner.name}}
```

缺失字段默认替换为空串，并在 `RenderResult.warnings` 中记录 warning。

## 3. 注册 Renderer

### 函数式

```python
from ppt_template_sdk import RendererRegistry, TextContent

registry = RendererRegistry()

@registry.renderer("title")
def render_title(placeholder, context):
    return TextContent(text=context.get_value("project.name", "未命名项目"))
```

同一个函数可以按不同 key 重复注册，并在注册时绑定固定参数：

```python
def render_label(placeholder, context, prefix):
    return TextContent(text=f"{prefix}{context.get_value('project.name')}")

registry.register_func("title", render_label, prefix="标题：")
registry.register_func("subtitle", render_label, prefix="副标题：")
```

### 类式

```python
from ppt_template_sdk import BaseRenderer, TableContent

class RiskTableRenderer(BaseRenderer):
    supported_types = {"table"}

    def render(self, placeholder, context, **kwargs):
        report = context.extras["report"]
        return TableContent(
            headers=["风险", "等级"],
            rows=report.risks,
        )
```

## 4. RenderContext

`RenderContext` 统一承载业务输入：

```python
from ppt_template_sdk import RenderContext

context = RenderContext(
    data={
        "project": {"name": "北极星"},
        "report_date": "2026-03-31",
    },
    extras={
        "report": report_object,
    },
)
```

约定：

- `data` 适合直接做字段替换的简单数据
- `extras` 适合复杂对象、服务对象、聚合模型
- `context.get_value("project.name")` 与模板字段取值规则一致

## 5. 执行渲染

```python
from ppt_template_sdk import EngineOptions, PptTemplateEngine

engine = PptTemplateEngine(
    registry=registry,
    options=EngineOptions(
        duplicate_key_policy="broadcast",
        enable_text_field_replace=True,
        strict=True,
    ),
)

result = engine.render(
    template_path="examples/assets/report_template.pptx",
    output_path="examples/output/report_output.pptx",
    context=context,
)
```

`RenderResult` 重点字段：

- `success`
- `output_path`
- `output_bytes`
- `rendered_count`
- `warnings`

## 6. 模板校验

```python
report = engine.validate(template_path="examples/assets/report_template.pptx")
```

`ValidationReport` 会检查：

- 占位块命名是否合法
- 是否缺少 renderer
- renderer 类型声明是否与占位块匹配
- 是否存在未使用 renderer
- 重复 key 在当前策略下是 error 还是 warning

说明：

- `validate()` 不依赖业务上下文
- 文本字段缺失不在 `validate()` 阶段检查

## 7. 建议的接入方式

- 业务系统负责准备模板、业务数据、图片、表格数据
- SDK 只负责解析模板、调度 renderer、写回 PPT
- `text` placeholder 会保留原文本框格式，并继承原 placeholder 首段首 run 的主样式
- 原生 `table` placeholder 会原位写回，保留列宽、行高和单元格样式
- 原生 `table` placeholder 与 `TableContent` 尺寸不一致时会直接报错
- 若渲染后还需要结构调整，使用 `PptOperations`
- 若只是做字段替换，不必走完整 `PptTemplateEngine`，可直接用 `TextReplacer`

## 8. 可直接运行的示例

- 渲染主链路：[`examples/render_report.py`](../examples/render_report.py)
- 文本替换：[`examples/text_replace.py`](../examples/text_replace.py)
- 操作模块：[`examples/operations_demo.py`](../examples/operations_demo.py)

## 9. 单文件版本

如果你的接入方式更偏向“复制一个文件到项目中”，仓库提供了单文件版：

- [`singlefile/ppt_template_sdk.py`](../singlefile/ppt_template_sdk.py)

使用建议：

1. 将该文件复制到你的业务项目
2. 文件名保持 `ppt_template_sdk.py`
3. 安装 `python-pptx`
4. 继续使用与包版相同的导入方式

说明：

- 单文件版覆盖当前全部公开能力
- 更适合 vendoring，不建议和包版同时装在同一解释器里混用
