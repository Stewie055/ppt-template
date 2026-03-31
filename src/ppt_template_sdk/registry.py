from __future__ import annotations

from typing import Callable

from .models.placeholder import Placeholder


class BaseRenderer:
    supported_types: set[str] = set()

    def render(self, placeholder: Placeholder, context):
        raise NotImplementedError


class _FunctionRenderer(BaseRenderer):
    def __init__(self, func: Callable):
        self._func = func

    def render(self, placeholder: Placeholder, context):
        return self._func(placeholder, context)


class RendererRegistry:
    def __init__(self) -> None:
        self._renderers: dict[str, BaseRenderer] = {}

    def register(self, key: str, renderer: BaseRenderer) -> None:
        self._renderers[key] = renderer

    def register_func(self, key: str, func: Callable) -> None:
        self.register(key, _FunctionRenderer(func))

    def renderer(self, key: str):
        def decorator(func: Callable):
            self.register_func(key, func)
            return func

        return decorator

    def get(self, key: str):
        return self._renderers.get(key)

    def keys(self) -> list[str]:
        return sorted(self._renderers.keys())
