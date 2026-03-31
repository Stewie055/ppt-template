"""Placeholder 描述模板中一个已识别的占位块。"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(slots=True)
class Placeholder:
    """模板占位块的标准化描述对象。

    字段说明：
        type: 占位块类型，例如 ``text``、``image``、``table``、``chart``。
        key: 业务侧注册的唯一 key。
        slide_index: 占位块所在 slide 的 ``0-based`` 索引。
        shape_id: 原始 shape id，常用于日志或后续操作定位。
        shape_name: 原始 ``shape.name``。
        left/top/width/height: 占位区域位置与尺寸。
        shape: 原始底层 shape 对象，供内部写回时使用。
    """

    type: str
    key: str
    slide_index: int
    shape_id: int
    shape_name: str
    left: int
    top: int
    width: int
    height: int
    shape: Any
