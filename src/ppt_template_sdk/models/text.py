from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(slots=True)
class TextReplaceResult:
    replaced_count: int = 0
    warnings: list[str] = field(default_factory=list)
