# 文本替换与操作模块

本文说明 `TextReplacer` 和 `PptOperations` 的用途、方法和限制。

## 1. TextReplacer

适用场景：

- 只想替换字段，不需要 shape 级 renderer
- 渲染主链路之外，单独处理已有 PPT

### 用法

```python
from pptx import Presentation
from ppt_template_sdk import RenderContext, TextReplacer

prs = Presentation("examples/assets/text_replace_template.pptx")

result = TextReplacer().replace_presentation_text(
    prs,
    context=RenderContext(
        data={
            "project": {"name": "Aurora"},
            "owner": {"name": "Bob"},
            "report_date": "2026-04-01",
        }
    ),
)

prs.save("examples/output/text_replaced.pptx")
print(result.replaced_count)
print(result.warnings)
```

### 行为

- 遍历所有 slide 的文本框与表格单元格
- 支持组内 shape
- 可通过 `rendered_shape_ids` 跳过已整体渲染的 shape
- 缺失字段替换为空串，并在 `warnings` 中记录

## 2. PptOperations

适用场景：

- 渲染后插入/删除页
- 按 section 组织页面
- 删除表格行列或合并单元格

### 加载与保存

```python
from ppt_template_sdk import PptOperations

ops = PptOperations.load(template_path="examples/assets/operations_template.pptx")
ops.save_to_path("examples/output/ops_output.pptx")
```

### 索引规则

- 所有 slide、section、row、column 索引都是 `0-based`
- `shape_locator` 支持 `shape_id` 或 `shape_name`
- 推荐优先使用 `shape_id`，避免同名 shape

## 3. Slide 操作

### 插入 slide

```python
ops.insert_slide(target_index=1, layout_index=6)
```

说明：

- `layout_index` 来自模板现有 `slide_layouts`
- 会先创建新 slide，再移动到目标位置

### 删除 slide

```python
ops.delete_slide(slide_index=2)
```

## 4. Section 操作

### 新增 section

```python
ops.add_section(name="正文", start_slide_index=1)
```

行为：

- 在 `start_slide_index` 对应 slide 处建立 section 边界
- 若模板当前没有 section，会自动初始化 section 列表
- 若目标 slide 已经是某 section 的起始页，则该 section 会被重命名

### 删除 section

```python
ops.delete_section(section_index=0)
```

行为：

- 只删除 section 元数据
- slides 会并入相邻 section，不会删除页面

## 5. 表格操作

### 删除表格行

```python
ops.delete_table_row(slide_index=0, shape_locator="ops-table", row_index=1)
```

### 删除表格列

```python
ops.delete_table_column(slide_index=0, shape_locator="ops-table", column_index=1)
```

### 合并单元格

```python
ops.merge_table_cells(
    slide_index=0,
    shape_locator="ops-table",
    first_row=0,
    first_col=0,
    last_row=1,
    last_col=1,
)
```

限制：

- 已合并表格不支持删行删列
- merge 区域必须是左上到右下的矩形
- 找不到 slide / shape / table 时会抛异常

## 6. 异常建议

- `OperationError`：索引越界、非法 merge 区域、对已合并表格删行删列
- `ShapeOperationError`：shape 不存在或目标不是表格
- `FieldReplaceError`：文本替换过程失败

## 7. 对应示例

- 文本替换：[`examples/text_replace.py`](../examples/text_replace.py)
- 操作模块：[`examples/operations_demo.py`](../examples/operations_demo.py)
