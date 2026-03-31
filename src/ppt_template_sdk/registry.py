"""RendererRegistry 管理占位块 key 与 renderer 的映射关系。

该模块面向接入方暴露 renderer 注册与查询接口，既支持类式 renderer，也支持
函数式 renderer 和装饰器注册。
"""

from __future__ import annotations

from typing import Callable

from .models.placeholder import Placeholder


class BaseRenderer:
    """自定义 renderer 的基类。

    子类只需要实现 ``render()``，并返回标准化的 ``Content`` 对象。若希望
    提前参与类型校验，可声明 ``supported_types``。
    """

    supported_types: set[str] = set()

    def render(self, placeholder: Placeholder, context):
        """根据占位块和上下文生成渲染结果。

        参数：
            placeholder: 当前占位块描述。
            context: 当前渲染上下文。

        返回：
            ``TextContent``、``ImageContent``、``TableContent`` 或 ``ChartContent``。
        """

        raise NotImplementedError


class _FunctionRenderer(BaseRenderer):
    def __init__(self, func: Callable):
        self._func = func

    def render(self, placeholder: Placeholder, context):
        return self._func(placeholder, context)


class RendererRegistry:
    """管理模板占位块 renderer 的注册与查询。

    示例：
        ```python
        registry = RendererRegistry()
        
        @registry.renderer("title")
        def render_title(placeholder, context):
            return TextContent(text=context.get_value("project.name", "未命名项目"))
        ```
    """

    def __init__(self) -> None:
        self._renderers: dict[str, BaseRenderer] = {}

    def register(self, key: str, renderer: BaseRenderer) -> None:
        """注册类式 renderer。

        参数：
            key: 模板中的占位块 key，例如 ``"title"``。
            renderer: ``BaseRenderer`` 实例。
        """

        self._renderers[key] = renderer

    def register_func(self, key: str, func: Callable) -> None:
        """注册函数式 renderer。"""

        self.register(key, _FunctionRenderer(func))

    def renderer(self, key: str):
        """返回装饰器，用于以声明式方式注册 renderer。

        示例：
            ```python
            @registry.renderer("title")
            def render_title(placeholder, context):
                return TextContent(text="经营分析")
            ```
        """

        def decorator(func: Callable):
            self.register_func(key, func)
            return func

        return decorator

    def get(self, key: str):
        """按 key 获取 renderer；若不存在则返回 ``None``。"""

        return self._renderers.get(key)

    def keys(self) -> list[str]:
        """返回已注册的全部 key，按字典序排序。"""

        return sorted(self._renderers.keys())
