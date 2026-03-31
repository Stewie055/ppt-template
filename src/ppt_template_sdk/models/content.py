"""内容模型定义了 renderer 可返回的标准化结果类型。

接入方无需直接操作底层 slide/shape，只需返回这些内容对象，具体写回逻辑由
SDK 内部统一处理。
"""

from __future__ import annotations

from dataclasses import dataclass


class Content:
    """所有渲染内容类型的基类。"""

    pass


@dataclass(slots=True)
class TextContent(Content):
    """文本占位块的渲染结果。

    字段说明：
        text: 将写回到目标文本 shape 的完整字符串。
    """

    text: str


@dataclass(slots=True)
class ImageContent(Content):
    """图片占位块的渲染结果。

    字段说明：
        image_path: 本地图片路径，SDK 会将其插入到目标区域。
    """

    image_path: str


@dataclass(slots=True)
class TableContent(Content):
    """表格占位块的渲染结果。

    字段说明：
        headers: 表头行；为空时仅写入数据行。
        rows: 二维字符串数组，每个子列表代表一行。
    """

    headers: list[str]
    rows: list[list[str]]


@dataclass(slots=True)
class ChartContent(Content):
    """图表占位块的渲染结果。

    当前版本按图片方式写入，因此结构与 ``ImageContent`` 类似。
    """

    image_path: str
