from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


_MISSING = object()


@dataclass(slots=True)
class RenderContext:
    data: Any
    extras: dict[str, Any] = field(default_factory=dict)

    def get_value(self, path: str, default: Any = None) -> Any:
        current = self.data
        for part in path.split("."):
            current = self._resolve_part(current, part, _MISSING)
            if current is _MISSING:
                return default
        return current

    def has_value(self, path: str) -> bool:
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
