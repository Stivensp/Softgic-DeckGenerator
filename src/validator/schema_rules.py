from __future__ import annotations

import logging

from src.exceptions import SchemaValidationError

logger = logging.getLogger("softgic.validator.schema")

REGISTERED_TYPES: frozenset[str] = frozenset(
    {
        "cover",
        "section_divider",
        "content_two_col",
        "pricing_table",
        "profile_card",
        "stat_callout",
        "closing",
    }
)

REQUIRED_FIELDS: dict[str, list[str]] = {
    "cover": ["title"],
    "section_divider": ["section_number", "section_title"],
    "content_two_col": ["title", "left", "right"],
    "pricing_table": ["title", "currency", "rows", "totals"],
    "profile_card": ["name", "role", "years_experience", "seniority", "skills"],
    "stat_callout": ["big_number", "label"],
    "closing": ["headline", "contact_name", "contact_email"],
}

FIELD_TYPES: dict[str, dict[str, type | tuple[type, ...]]] = {
    "cover": {"title": str},
    "section_divider": {"section_number": str, "section_title": str},
    "content_two_col": {"title": str, "left": dict, "right": dict},
    "pricing_table": {"title": str, "currency": str, "rows": list, "totals": (int, float)},
    "profile_card": {
        "name": str,
        "role": str,
        "years_experience": int,
        "seniority": str,
        "skills": list,
    },
    "stat_callout": {"big_number": str, "label": str},
    "closing": {"headline": str, "contact_name": str, "contact_email": str},
}


class SchemaRulesValidator:
    def validate(self, data: dict) -> None:  # type: ignore[type-arg]
        self._check_top_level(data)
        for idx, slide in enumerate(data["slides"]):
            self._check_slide(slide, idx)

    def _check_top_level(self, data: dict) -> None:  # type: ignore[type-arg]
        if "slides" not in data:
            raise SchemaValidationError("Missing required top-level key 'slides'")
        if not isinstance(data["slides"], list):
            raise SchemaValidationError("'slides' must be a list")
        if len(data["slides"]) == 0:
            raise SchemaValidationError("'slides' list must not be empty")

    def _check_slide(self, slide: dict, idx: int) -> None:  # type: ignore[type-arg]
        if not isinstance(slide, dict):
            raise SchemaValidationError(
                f"Each slide must be a mapping, got {type(slide).__name__}", slide_index=idx
            )
        if "type" not in slide:
            raise SchemaValidationError("Missing required field 'type'", field="type", slide_index=idx)

        slide_type: str = slide["type"]
        if slide_type not in REGISTERED_TYPES:
            raise SchemaValidationError(
                f"Unknown slide type '{slide_type}'. Registered types: {sorted(REGISTERED_TYPES)}",
                field="type",
                slide_index=idx,
            )

        for field in REQUIRED_FIELDS.get(slide_type, []):
            if field not in slide or slide[field] is None:
                raise SchemaValidationError(
                    f"Required field '{field}' is missing or null",
                    field=field,
                    slide_index=idx,
                )

        for field, expected_type in FIELD_TYPES.get(slide_type, {}).items():
            if field in slide and slide[field] is not None:
                if not isinstance(slide[field], expected_type):
                    actual = type(slide[field]).__name__
                    expected_name = (
                        expected_type.__name__
                        if isinstance(expected_type, type)
                        else " or ".join(t.__name__ for t in expected_type)
                    )
                    raise SchemaValidationError(
                        f"Field '{field}' must be {expected_name}, got {actual}",
                        field=field,
                        slide_index=idx,
                    )

        if slide_type in ("content_two_col",):
            self._check_bullet_column(slide.get("left", {}), "left", idx)
            self._check_bullet_column(slide.get("right", {}), "right", idx)

        if slide_type == "pricing_table":
            rows = slide.get("rows", [])
            if not rows:
                raise SchemaValidationError(
                    "'rows' must not be empty", field="rows", slide_index=idx
                )
            for row_idx, row in enumerate(rows):
                if not isinstance(row, dict):
                    raise SchemaValidationError(
                        f"Pricing row {row_idx} must be a mapping", field="rows", slide_index=idx
                    )
                for req in ("description", "quantity", "unit_price", "total"):
                    if req not in row:
                        raise SchemaValidationError(
                            f"Pricing row {row_idx} missing field '{req}'",
                            field=f"rows[{row_idx}].{req}",
                            slide_index=idx,
                        )

    def _check_bullet_column(self, col: dict, col_name: str, idx: int) -> None:  # type: ignore[type-arg]
        if "bullets" not in col or not isinstance(col.get("bullets"), list):
            raise SchemaValidationError(
                f"Column '{col_name}' must have a 'bullets' list",
                field=f"{col_name}.bullets",
                slide_index=idx,
            )
        if len(col["bullets"]) == 0:
            raise SchemaValidationError(
                f"Column '{col_name}.bullets' must not be empty",
                field=f"{col_name}.bullets",
                slide_index=idx,
            )
