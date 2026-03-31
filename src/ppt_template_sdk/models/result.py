"""渲染结果模型定义。"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass(slots=True)
class RenderResult:
    """一次模板渲染的结构化返回结果。

    字段说明：
        success: 是否渲染成功。
        output_path: 若调用时指定了输出路径，则为最终落盘路径。
        output_bytes: 渲染后的 PPT 字节流。
        rendered_count: 实际完成 shape 级渲染的占位块数量。
        skipped_count: 预留字段，用于记录未处理占位块数量。
        warnings: 渲染阶段产生的 warning，例如缺失文本字段。
    """

    success: bool
    output_path: Optional[str] = None
    output_bytes: Optional[bytes] = None
    rendered_count: int = 0
    skipped_count: int = 0
    warnings: list[str] = field(default_factory=list)
