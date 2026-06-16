"""
Targeted tests for schema validator branches not yet covered.
"""
from __future__ import annotations

import pytest

from src.exceptions import SchemaValidationError
from src.validator.schema_rules import SchemaRulesValidator


def test_slides_key_not_a_list() -> None:
    with pytest.raises(SchemaValidationError, match="list"):
        SchemaRulesValidator().validate({"slides": "not_a_list"})


def test_field_wrong_type_cover_title_is_int() -> None:
    with pytest.raises(SchemaValidationError, match="str"):
        SchemaRulesValidator().validate({"slides": [{"type": "cover", "title": 123}]})


def test_field_wrong_type_pricing_totals_is_string() -> None:
    with pytest.raises(SchemaValidationError, match="int or float"):
        SchemaRulesValidator().validate(
            {
                "slides": [
                    {
                        "type": "pricing_table",
                        "title": "T",
                        "currency": "USD",
                        "rows": [
                            {"description": "X", "quantity": 1, "unit_price": 10.0, "total": 10.0}
                        ],
                        "totals": "wrong",
                    }
                ]
            }
        )


def test_pricing_row_missing_field() -> None:
    with pytest.raises(SchemaValidationError, match="unit_price"):
        SchemaRulesValidator().validate(
            {
                "slides": [
                    {
                        "type": "pricing_table",
                        "title": "T",
                        "currency": "USD",
                        "rows": [{"description": "X", "quantity": 1, "total": 10.0}],
                        "totals": 10.0,
                    }
                ]
            }
        )


def test_pricing_row_not_a_dict() -> None:
    with pytest.raises(SchemaValidationError, match="mapping"):
        SchemaRulesValidator().validate(
            {
                "slides": [
                    {
                        "type": "pricing_table",
                        "title": "T",
                        "currency": "USD",
                        "rows": ["not_a_dict"],
                        "totals": 0,
                    }
                ]
            }
        )


def test_content_two_col_empty_bullets_list() -> None:
    with pytest.raises(SchemaValidationError, match="empty"):
        SchemaRulesValidator().validate(
            {
                "slides": [
                    {
                        "type": "content_two_col",
                        "title": "T",
                        "left": {"bullets": []},
                        "right": {"bullets": ["x"]},
                    }
                ]
            }
        )


def test_field_wrong_type_years_experience_is_string() -> None:
    with pytest.raises(SchemaValidationError, match="int"):
        SchemaRulesValidator().validate(
            {
                "slides": [
                    {
                        "type": "profile_card",
                        "name": "A",
                        "role": "B",
                        "years_experience": "five",
                        "seniority": "Mid",
                        "skills": ["x"],
                    }
                ]
            }
        )


def test_required_null_field_treated_as_missing() -> None:
    with pytest.raises(SchemaValidationError, match="'title'"):
        SchemaRulesValidator().validate({"slides": [{"type": "cover", "title": None}]})
