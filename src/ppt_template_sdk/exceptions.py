"""SDK 对外暴露的异常体系。

调用方可以统一捕获 ``PptTemplateSdkError``，或按细分异常做更精确的错误处理。
"""

class PptTemplateSdkError(Exception):
    """SDK 所有自定义异常的基类。"""


class TemplateParseError(PptTemplateSdkError):
    """模板解析阶段发生错误时抛出。"""

    pass


class PlaceholderFormatError(TemplateParseError):
    """占位块命名不符合 ``ph:<type>:<key>`` 规范时抛出。"""

    pass


class DuplicatePlaceholderError(PptTemplateSdkError):
    """检测到重复占位块 key 且当前策略不允许时抛出。"""

    pass


class RendererNotFoundError(PptTemplateSdkError):
    """模板中的占位块缺少对应 renderer 时抛出。"""

    pass


class ContentTypeMismatchError(PptTemplateSdkError):
    """renderer 返回的内容类型与占位块类型不匹配时抛出。"""

    pass


class ShapeOperationError(PptTemplateSdkError):
    """对底层 shape 执行写回或定位操作失败时抛出。"""

    pass


class FieldReplaceError(PptTemplateSdkError):
    """文本字段替换阶段发生异常时抛出。"""

    pass


class OperationError(PptTemplateSdkError):
    """执行 slide、section 或表格结构操作失败时抛出。"""

    pass
