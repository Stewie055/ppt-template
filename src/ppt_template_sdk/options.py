from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class EngineOptions:
    duplicate_key_policy: str = "broadcast"
    enable_text_field_replace: bool = True
    text_field_pattern: str = r"\{\{([\w\.]+)\}\}"
    text_field_replace_mode: str = "plain"
    strict: bool = True
