from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass(slots=True)
class RenderResult:
    success: bool
    output_path: Optional[str] = None
    output_bytes: Optional[bytes] = None
    rendered_count: int = 0
    skipped_count: int = 0
    warnings: list[str] = field(default_factory=list)
