# PPT 模板渲染 SDK 需求说明

请实现一个可复用的 Python 功能模块 / SDK，用于给其他应用集成 PPT 模板渲染能力。

## 1. 模块定位

这是一个 PPT 模板渲染引擎 SDK，不是独立业务系统，而是供其他 Python 应用调用的基础能力组件。

模块负责：
- 读取 `.pptx` 模板
- 扫描模板中的占位块
- 调用接入方实现的渲染逻辑
- 将内容替换回 PPT
- 输出 `.pptx` 文件或内存字节流

模块不负责：
- 业务数据获取
- 数据分析
- 图表业务逻辑生成
- 文案生成
- 页面业务编排决策

这些都由接入方实现。

---

## 2. 核心使用模式

接入方在使用本模块时：
1. 提供一个 PPT 模板
2. 在模板中定义内容占位块
3. 实现自己的 renderer / handler
4. 提供业务上下文数据
5. 调用 SDK 生成最终 PPT

---

## 3. 模板协议

### 3.1 Shape 级占位块命名规则

使用 `shape.name` 作为占位块标识，命名格式为：

```text
ph:<type>:<key>
```

例如：

```text
ph:text:title
ph:text:summary
ph:image:cover
ph:chart:sales_trend
ph:table:risk_table
```

### 3.2 语义说明

- `type`：占位块类型，由 SDK 识别
- `key`：业务唯一标识，由接入方定义
- shape 的位置和大小即最终渲染区域
- SDK 不依赖 shape 中原始文本内容，只依赖 `shape.name`

### 3.3 推荐支持的 type

第一版至少支持：
- `text`
- `image`
- `table`
- `chart`

说明：
- `chart` 可以先按“图表图片插入”实现，不要求原生 PPT 图表对象编辑

---

## 4. 文本字段替换能力

除了 shape 级占位块替换外，还需要支持文本中的字段替换。

### 4.1 字段语法

推荐语法：

```text
{{field_name}}
```

也建议支持点路径：

```text
{{project.name}}
{{report.date}}
{{owner.name}}
```

### 4.2 应用范围

第一版至少支持：
- 普通文本框中的字段替换
- 表格单元格中的字段替换

例如模板文本：

```text
项目名称：{{project_name}}
报告日期：{{report_date}}
负责人：{{owner}}
```

### 4.3 与 shape 占位块的执行顺序

推荐执行顺序：
1. 先执行 shape 级占位块替换
2. 再执行文本字段替换

规则：
- 若一个 shape 已按 `ph:*:*` 被整体渲染，则默认不再做文本字段替换
- 普通文本框、表格单元格可执行字段替换

### 4.4 第一版实现建议

第一版先实现 plain 模式：
- 直接按文本整体做字符串替换
- 不强求完整保留局部 run 富文本样式

后续可扩展 `preserve_style` 模式。

---

## 5. 重名占位块策略

需要考虑多个 shape 使用相同 `key` 的情况。

### 5.1 默认策略

默认不允许同名 key，对重复 key 报错。

原因：
- 行为更明确
- 适合做 SDK 的安全默认行为
- 避免模板隐式错误

### 5.2 可选策略

通过配置支持：
- `error`：发现重复 key 直接报错
- `broadcast`：同 key 的多个 shape 使用同一个渲染结果

例如：

```python
EngineOptions(duplicate_key_policy="broadcast")
```

`broadcast` 适合：
- 页眉页脚
- 日期
- 公司名
- 文档编号
- 多页重复标签

第一版不要求实现更复杂的 group/component 模式。

---

## 6. 对外公开 API 要求

SDK 对外尽量暴露少量稳定接口。

建议至少包含：
- `PptTemplateEngine`
- `RendererRegistry`
- `RenderContext`
- `EngineOptions`
- `BaseRenderer`
- `TextContent`
- `ImageContent`
- `TableContent`
- `ChartContent`
- `RenderResult`
- `ValidationReport`

---

## 7. 核心接口设计建议

### 7.1 Engine

```python
class PptTemplateEngine:
    def __init__(
        self,
        registry: RendererRegistry,
        options: EngineOptions | None = None,
    ): ...

    def render(
        self,
        template_path: str | None = None,
        template_bytes: bytes | None = None,
        output_path: str | None = None,
        context: RenderContext | None = None,
    ) -> RenderResult: ...

    def validate(
        self,
        template_path: str | None = None,
        template_bytes: bytes | None = None,
    ) -> ValidationReport: ...
```

要求：
- 支持模板文件路径输入
- 支持模板 bytes 输入，方便 Web/API 场景
- 支持输出到文件
- 支持输出为内存 bytes

### 7.2 Registry

```python
class RendererRegistry:
    def register(self, key: str, renderer: BaseRenderer) -> None: ...
    def register_func(self, key: str, func) -> None: ...
    def get(self, key: str): ...
    def keys(self) -> list[str]: ...
```

要求：
- 支持注册类式 renderer
- 支持注册函数式 renderer

### 7.3 Renderer Contract

```python
class BaseRenderer:
    supported_types: set[str] = set()

    def render(
        self,
        placeholder: Placeholder,
        context: RenderContext,
    ) -> Content:
        raise NotImplementedError
```

要求：
- renderer 不直接操作 `slide` / `shape`
- renderer 只接收占位块描述和业务上下文
- renderer 返回标准化内容对象
- PPT 写入逻辑由 SDK 内部统一处理

### 7.4 Context

```python
from dataclasses import dataclass, field
from typing import Any

@dataclass
class RenderContext:
    data: Any
    extras: dict = field(default_factory=dict)

    def get_value(self, path: str, default=None):
        ...
```

说明：
- `data` 支持 `dict`、dataclass、普通对象、pydantic model 等
- `extras`：日志、临时目录、请求 ID、环境参数、复杂业务对象、服务对象等扩展信息
- SDK 内部应提供统一路径取值能力，例如 `get_value("report.title")`
- 模板字段替换和 renderer 取值都应复用同一套路径解析逻辑

### 7.5 Renderer 的数据传入方式

SDK 不应由 engine 预先把单个字段拆出来传给 renderer，而应将完整 `RenderContext` 传给 renderer：

```python
renderer.render(placeholder, context)
```

推荐的数据使用方式：
- 简单字段放在 `context.data` 中，供模板字段替换和简单 renderer 直接使用
- 复杂对象、聚合对象、服务对象放在 `context.extras` 中，供 renderer 调用其属性和方法

例如：

```python
context = RenderContext(
    data={
        "title": "2026年Q1经营分析报告",
        "report_date": "2026-03-31",
    },
    extras={
        "report": report_data,
    }
)
```

renderer 中可以这样使用：

```python
class SummaryRenderer(BaseRenderer):
    supported_types = {"text"}

    def render(self, placeholder, context):
        report = context.extras["report"]
        return TextContent(text=report.get_summary())
```

如果主数据本身就是一个对象，也允许直接传入：

```python
context = RenderContext(data=report_data)
```

renderer 中：

```python
class TitleRenderer(BaseRenderer):
    supported_types = {"text"}

    def render(self, placeholder, context):
        report = context.data
        return TextContent(text=report.title)
```

---

## 8. 内容对象模型

建议定义统一基类 `Content`，并至少实现：

```python
class Content: ...
```

```python
from dataclasses import dataclass

@dataclass
class TextContent(Content):
    text: str

@dataclass
class ImageContent(Content):
    image_path: str

@dataclass
class TableContent(Content):
    headers: list[str]
    rows: list[list[str]]

@dataclass
class ChartContent(Content):
    image_path: str
```

说明：
- `chart` 第一版按图片处理
- 表格内容先支持基础二维数据写入

---

## 9. 引擎执行流程

推荐实现流程：

```text
1. 读取模板
2. 扫描所有 slide 和 shape
3. 识别 ph:<type>:<key> 占位块
4. 按 key 聚合占位块
5. 根据 duplicate_key_policy 做检查或广播处理
6. 根据 key 从 registry 中获取 renderer
7. 调用 renderer.render(placeholder, context)
8. 校验 content 类型是否和 placeholder.type 匹配
9. 将内容写入对应 shape
10. 对剩余普通文本框 / 表格单元格执行字段替换
11. 输出最终 PPT
```

---

## 10. 校验能力要求

SDK 需要提供独立的模板校验能力 `validate()`。

至少检查以下内容：

### 10.1 占位块命名是否合法
例如非法格式：
- `ph:title`
- `summary`
- `placeholder:text:title`

### 10.2 模板中的占位块是否缺少 renderer
例如模板中有：
- `ph:text:summary`
- `ph:chart:sales`

但未注册对应 renderer

### 10.3 registry 中是否存在未使用 renderer
便于清理无效配置

### 10.4 类型是否匹配
例如：
- 模板是 `ph:image:cover`
- renderer 返回 `TextContent`

### 10.5 重复 key
根据策略进行报错或警告

### 10.6 字段占位符缺失
例如模板中存在：

```text
{{report_date}}
```

但上下文中没有对应值

这一项可作为 warning，而不一定强制 error。

---

## 11. 配置项设计建议

```python
from dataclasses import dataclass

@dataclass
class EngineOptions:
    duplicate_key_policy: str = "boradcast"   # error | broadcast
    enable_text_field_replace: bool = True
    text_field_pattern: str = r"\{\{([\w\.]+)\}\}"
    text_field_replace_mode: str = "plain"   # plain | preserve_style
    strict: bool = True
```

说明：
- `duplicate_key_policy`
  - `error`：重复 key 报错
  - `broadcast`：同 key 广播替换

- `enable_text_field_replace`
  - 是否启用文本字段替换

- `text_field_replace_mode`
  - `plain`：整体文本替换
  - `preserve_style`：后续增强

- `strict`
  - 严格模式下，未注册 renderer、类型不匹配等直接报错

---

## 12. 返回结果设计

不要只返回是否成功，建议提供结构化结果对象。

```python
from dataclasses import dataclass, field

@dataclass
class RenderResult:
    success: bool
    output_path: str | None = None
    output_bytes: bytes | None = None
    rendered_count: int = 0
    skipped_count: int = 0
    warnings: list[str] = field(default_factory=list)
```

可选再补充：
- `placeholder_reports`
- `field_replace_reports`

---

## 13. 异常体系要求

不要只抛标准异常，建议定义模块自己的异常体系。

例如：

```python
class PptTemplateSdkError(Exception): ...
class TemplateParseError(PptTemplateSdkError): ...
class PlaceholderFormatError(PptTemplateSdkError): ...
class DuplicatePlaceholderError(PptTemplateSdkError): ...
class RendererNotFoundError(PptTemplateSdkError): ...
class ContentTypeMismatchError(PptTemplateSdkError): ...
class ShapeOperationError(PptTemplateSdkError): ...
class FieldReplaceError(PptTemplateSdkError): ...
```

要求：
- 上层应用可统一捕获 SDK 异常
- 错误信息清晰可定位

---

## 14. 项目结构建议

建议做成标准 Python package，类似：

```text
ppt_template_sdk/
├── pyproject.toml
├── README.md
├── src/
│   └── ppt_template_sdk/
│       ├── __init__.py
│       ├── engine.py
│       ├── registry.py
│       ├── context.py
│       ├── options.py
│       ├── exceptions.py
│       ├── validator.py
│       │
│       ├── contracts/
│       │   ├── renderer.py
│       │   ├── provider.py
│       │   └── hook.py
│       │
│       ├── models/
│       │   ├── placeholder.py
│       │   ├── content.py
│       │   ├── result.py
│       │   └── report.py
│       │
│       ├── parser/
│       │   └── template_parser.py
│       │
│       ├── renderer/
│       │   ├── dispatcher.py
│       │   ├── text_renderer.py
│       │   ├── image_renderer.py
│       │   ├── table_renderer.py
│       │   └── chart_renderer.py
│       │
│       ├── adapter/
│       │   ├── pptx_adapter.py
│       │   └── shape_operator.py
│       │
│       └── utils/
│           ├── naming.py
│           ├── logging.py
│           └── tempfiles.py
└── tests/
```

---

## 15. 技术实现要求

### 15.1 基础库
- 使用 `python-pptx`

### 15.2 兼容性
- Python 版本建议 3.10+
- 代码尽量使用类型注解

### 15.3 代码要求
- 模块化清晰
- 可测试
- 易扩展
- 公共接口稳定
- 对业务代码低侵入

---

## 16. 第一版必须实现的最小能力

第一版 MVP 至少实现：

1. 读取 `.pptx`
2. 扫描 `shape.name`
3. 识别 `ph:<type>:<key>`
4. 注册 renderer 并执行
5. 支持：
   - `text`
   - `image`
   - `table`
   - `chart(图片方式)`
6. 支持文本框字段替换
7. 支持表格单元格字段替换
8. 支持 `duplicate_key_policy=error`
9. 支持 `duplicate_key_policy=broadcast`
10. 支持 `validate()`
11. 支持输出到文件和 bytes
12. 提供清晰异常体系
13. `RenderContext.data` 支持 dict 与对象
14. renderer 可通过 `context.data` 和 `context.extras` 使用复杂对象及其方法

---

## 17. 暂不要求或后续增强项

这些可以作为后续版本增强，不要求第一版完成：
- 富文本 run 级字段替换
- 保留局部文本样式的替换
- 复杂组件 group 渲染
- 自动分页表格
- 模板页复制
- 条件渲染 / 循环语法
- 原生 PPT 图表对象深度编辑
- 高级样式继承与主题映射

---

## 18. 推荐示例调用方式

```python
from ppt_template_sdk import (
    PptTemplateEngine,
    RendererRegistry,
    RenderContext,
    EngineOptions,
    TextContent,
)

registry = RendererRegistry()

@registry.renderer("title")
def render_title(placeholder, context):
    report = context.extras["report"]
    return TextContent(text=report.title)

engine = PptTemplateEngine(
    registry=registry,
    options=EngineOptions(
        duplicate_key_policy="error",
        enable_text_field_replace=True,
        text_field_replace_mode="plain",
        strict=True,
    ),
)

report_data = ReportData(
    title="2026年Q1经营分析报告",
    sales=[120, 150, 210],
)

result = engine.render(
    template_path="report_template.pptx",
    output_path="report_output.pptx",
    context=RenderContext(
        data={
            "title": report_data.title,
            "report_date": "2026-03-31",
        },
        extras={
            "report": report_data,
        }
    ),
)
```

---

## 19. 最终设计原则

请严格按以下原则设计：

1. 模块只负责模板解析、占位块调度、内容写回
2. 业务生成逻辑全部由接入方实现
3. renderer 不直接操作底层 PPT 对象
4. shape 占位块替换与文本字段替换分开设计
5. 默认行为保守、安全、可预测
6. 公开 API 尽量少，但扩展点清晰
7. 适合作为 SDK 被其他项目长期引用

