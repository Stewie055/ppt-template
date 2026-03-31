"""RenderContext 定义了模板渲染与文本替换共享的数据访问入口。

对接方通常把可直接用于模板字段替换的数据放在 ``data`` 中，把复杂对象、
服务对象或聚合模型放在 ``extras`` 中。公开方法主要用于按点路径读取数据。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


_MISSING = object()


@dataclass(slots=True)
class RenderContext:
    """承载模板渲染和文本替换所需的业务上下文。

    参数：
        data: 业务主数据，可为 ``dict``、对象、dataclass 或 pydantic 风格对象。
        extras: 额外上下文对象，例如服务实例、聚合模型、日志上下文等。

    示例：
        ```python
        context = RenderContext(
            data={"project": {"name": "北极星"}},
            extras={"report": report},
        )
        ```
    """

    data: Any
    extras: dict[str, Any] = field(default_factory=dict)

    def get_value(self, path: str, default: Any = None) -> Any:
        """按点路径读取 ``data`` 中的值。

        参数：
            path: 例如 ``"project.name"`` 或 ``"items.0.title"``。
            default: 路径不存在时返回的默认值。

        返回：
            路径对应的值；若路径不存在，则返回 ``default``。

        示例：
            ```python
            value = context.get_value("project.name", "未命名项目")
            ```
        """

        current = self.data
        for part in path.split("."):
            current = self._resolve_part(current, part, _MISSING)
            if current is _MISSING:
                return default
        return current

    def has_value(self, path: str) -> bool:
        """判断给定点路径是否存在有效值。"""

        return self.get_value(path, _MISSING) is not _MISSING

    @staticmethod
    def _resolve_part(value: Any, part: str, default: Any) -> Any:
        if value is None:
            return default
        if isinstance(value, dict):
            return value.get(part, default)
        if isinstance(value, (list, tuple)) and part.isdigit():
            index = int(part)
            if 0 <= index < len(value):
                return value[index]
            return default
        if hasattr(value, part):
            return getattr(value, part)
        if hasattr(value, "model_dump"):
            return value.model_dump().get(part, default)
        if hasattr(value, "dict") and callable(value.dict):
            return value.dict().get(part, default)
        return default
