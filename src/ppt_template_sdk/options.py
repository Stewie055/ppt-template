"""EngineOptions 定义渲染引擎的主要运行参数。

该文件面向接入方公开模板渲染的核心开关，例如重复 key 策略和文本替换开关。
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class EngineOptions:
    """控制渲染引擎行为的配置对象。

    字段说明：
        duplicate_key_policy: 重复占位块 key 的处理策略，支持 ``error`` 和 ``broadcast``。
        enable_text_field_replace: 是否启用普通文本与表格单元格字段替换。
        text_field_pattern: 文本字段匹配正则，默认支持 ``{{path.to.value}}``。
        text_field_replace_mode: 当前仅支持 ``plain``。
        strict: 是否在校验失败时尽早报错。
    """

    duplicate_key_policy: str = "broadcast"
    enable_text_field_replace: bool = True
    text_field_pattern: str = r"\{\{([\w\.]+)\}\}"
    text_field_replace_mode: str = "plain"
    strict: bool = True
