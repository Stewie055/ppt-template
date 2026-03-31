"""模板校验报告模型定义。"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(slots=True)
class ValidationReport:
    """静态模板校验结果。

    字段说明：
        success: 是否通过静态校验。
        placeholder_count: 模板中识别到的合法占位块数量。
        errors: 阻断型问题列表。
        warnings: 非阻断型问题列表。
        unused_renderers: 已注册但未在模板中使用的 renderer key。
    """

    success: bool
    placeholder_count: int = 0
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    unused_renderers: list[str] = field(default_factory=list)
