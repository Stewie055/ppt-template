from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(slots=True)
class ValidationReport:
    success: bool
    placeholder_count: int = 0
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    unused_renderers: list[str] = field(default_factory=list)
