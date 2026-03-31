class PptTemplateSdkError(Exception):
    """Base exception for all SDK errors."""


class TemplateParseError(PptTemplateSdkError):
    pass


class PlaceholderFormatError(TemplateParseError):
    pass


class DuplicatePlaceholderError(PptTemplateSdkError):
    pass


class RendererNotFoundError(PptTemplateSdkError):
    pass


class ContentTypeMismatchError(PptTemplateSdkError):
    pass


class ShapeOperationError(PptTemplateSdkError):
    pass


class FieldReplaceError(PptTemplateSdkError):
    pass
