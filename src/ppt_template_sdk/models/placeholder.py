from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(slots=True)
class Placeholder:
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
