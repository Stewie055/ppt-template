from __future__ import annotations

from collections import Counter

from .models.report import ValidationReport
from .parser.template_parser import parse_presentation


def validate_presentation(presentation, registry, options) -> ValidationReport:
    parsed = parse_presentation(presentation, options.text_field_pattern)
    errors = list(parsed.invalid_placeholders)
    warnings: list[str] = []

    key_counter = Counter(placeholder.key for placeholder in parsed.placeholders)
    for key, count in sorted(key_counter.items()):
        if count < 2:
            continue
        message = f"duplicate placeholder key '{key}' appears {count} times"
        if options.duplicate_key_policy == "error":
            errors.append(message)
        else:
            warnings.append(message)

    used_keys = set()
    for placeholder in parsed.placeholders:
        renderer = registry.get(placeholder.key)
        if renderer is None:
            errors.append(
                f"missing renderer for placeholder '{placeholder.shape_name}' on slide {placeholder.slide_index + 1}"
            )
            continue
        used_keys.add(placeholder.key)
        supported_types = getattr(renderer, "supported_types", set()) or set()
        if supported_types and placeholder.type not in supported_types:
            errors.append(
                f"renderer '{placeholder.key}' does not support placeholder type '{placeholder.type}'"
            )

    unused_renderers = sorted(set(registry.keys()) - used_keys)
    if unused_renderers:
        warnings.append(f"unused renderers: {', '.join(unused_renderers)}")

    return ValidationReport(
        success=not errors,
        placeholder_count=len(parsed.placeholders),
        errors=errors,
        warnings=warnings,
        unused_renderers=unused_renderers,
    )
