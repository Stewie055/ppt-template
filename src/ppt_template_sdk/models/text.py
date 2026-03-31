"""文本替换结果模型定义。"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(slots=True)
class TextReplaceResult:
    """独立文本替换操作的返回结果。

    字段说明：
        replaced_count: 本次命中的字段总数。
        warnings: 替换过程中产生的 warning，例如缺失字段。
    """

    replaced_count: int = 0
    warnings: list[str] = field(default_factory=list)
