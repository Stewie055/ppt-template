from __future__ import annotations

from dataclasses import dataclass


class Content:
    pass


@dataclass(slots=True)
class TextContent(Content):
    text: str


@dataclass(slots=True)
class ImageContent(Content):
    image_path: str


@dataclass(slots=True)
class TableContent(Content):
    headers: list[str]
    rows: list[list[str]]


@dataclass(slots=True)
class ChartContent(Content):
    image_path: str
