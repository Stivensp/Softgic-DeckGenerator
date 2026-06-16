from __future__ import annotations

import pytest

from src.exceptions import (
    AssetValidationError,
    BusinessRuleError,
    SchemaValidationError,
)
from src.validator.asset_validator import AssetValidator
from src.validator.business_rules import BusinessRulesValidator
from src.validator.deck_validator import DeckValidator
from src.validator.schema_rules import SchemaRulesValidator


# ── Schema Rules ─────────────────────────────────────────────────────────────

def test_schema_valid_minimal() -> None:
    SchemaRulesValidator().validate({"slides": [{"type": "cover", "title": "T"}]})


def test_schema_missing_slides_key() -> None:
    with pytest.raises(SchemaValidationError, match="'slides'"):
        SchemaRulesValidator().validate({})


def test_schema_empty_slides_list() -> None:
    with pytest.raises(SchemaValidationError, match="empty"):
        SchemaRulesValidator().validate({"slides": []})


def test_schema_slide_not_dict() -> None:
    with pytest.raises(SchemaValidationError):
        SchemaRulesValidator().validate({"slides": ["not_a_dict"]})


def test_schema_missing_type_field() -> None:
    with pytest.raises(SchemaValidationError, match="type"):
        SchemaRulesValidator().validate({"slides": [{"title": "Hi"}]})


def test_schema_unknown_type() -> None:
    with pytest.raises(SchemaValidationError, match="Unknown slide type"):
        SchemaRulesValidator().validate({"slides": [{"type": "wizard_slide"}]})


def test_schema_missing_required_field_cover() -> None:
    with pytest.raises(SchemaValidationError, match="'title'"):
        SchemaRulesValidator().validate({"slides": [{"type": "cover"}]})


def test_schema_missing_required_field_profile_card() -> None:
    with pytest.raises(SchemaValidationError):
        SchemaRulesValidator().validate(
            {"slides": [{"type": "profile_card", "name": "Ana"}]}
        )


def test_schema_pricing_table_empty_rows() -> None:
    with pytest.raises(SchemaValidationError, match="'rows'"):
        SchemaRulesValidator().validate(
            {
                "slides": [
                    {
                        "type": "pricing_table",
                        "title": "T",
                        "currency": "USD",
                        "rows": [],
                        "totals": 0,
                    }
                ]
            }
        )


def test_schema_content_two_col_missing_bullets() -> None:
    with pytest.raises(SchemaValidationError):
        SchemaRulesValidator().validate(
            {
                "slides": [
                    {
                        "type": "content_two_col",
                        "title": "T",
                        "left": {"title": "L"},  # no bullets
                        "right": {"title": "R", "bullets": ["x"]},
                    }
                ]
            }
        )


# ── Business Rules ────────────────────────────────────────────────────────────

def test_business_title_too_long() -> None:
    with pytest.raises(BusinessRuleError, match="title"):
        BusinessRulesValidator().validate(
            {"slides": [{"type": "cover", "title": "A" * 81}]},
            limits={"title_max_chars": 80},
        )


def test_business_too_many_bullets() -> None:
    with pytest.raises(BusinessRuleError, match="left"):
        BusinessRulesValidator().validate(
            {
                "slides": [
                    {
                        "type": "content_two_col",
                        "title": "T",
                        "left": {"bullets": ["b"] * 7},
                        "right": {"bullets": ["x"]},
                    }
                ]
            },
            limits={"bullets_per_column_max": 6},
        )


def test_business_invalid_email() -> None:
    with pytest.raises(BusinessRuleError, match="email"):
        BusinessRulesValidator().validate(
            {
                "slides": [
                    {
                        "type": "closing",
                        "headline": "H",
                        "contact_name": "N",
                        "contact_email": "not-an-email",
                    }
                ]
            }
        )


def test_business_valid_email_passes() -> None:
    BusinessRulesValidator().validate(
        {
            "slides": [
                {
                    "type": "closing",
                    "headline": "H",
                    "contact_name": "N",
                    "contact_email": "user@domain.com",
                }
            ]
        }
    )


def test_business_negative_years_experience() -> None:
    with pytest.raises(BusinessRuleError, match="years_experience"):
        BusinessRulesValidator().validate(
            {
                "slides": [
                    {
                        "type": "profile_card",
                        "name": "A",
                        "role": "B",
                        "years_experience": -1,
                        "seniority": "Senior",
                        "skills": ["Python"],
                    }
                ]
            }
        )


def test_business_too_many_skills() -> None:
    with pytest.raises(BusinessRuleError, match="skills"):
        BusinessRulesValidator().validate(
            {
                "slides": [
                    {
                        "type": "profile_card",
                        "name": "A",
                        "role": "B",
                        "years_experience": 3,
                        "seniority": "Mid",
                        "skills": ["s"] * 9,
                    }
                ]
            },
            limits={"skills_max": 8},
        )


# ── Asset Validator ───────────────────────────────────────────────────────────

def test_asset_missing_photo_raises_error(tmp_path) -> None:
    with pytest.raises(AssetValidationError, match="not found"):
        AssetValidator().validate(
            {
                "slides": [
                    {
                        "type": "profile_card",
                        "name": "A",
                        "role": "B",
                        "years_experience": 3,
                        "seniority": "Mid",
                        "skills": ["x"],
                        "photo": str(tmp_path / "ghost.jpg"),
                    }
                ]
            }
        )


def test_asset_no_photo_passes() -> None:
    AssetValidator().validate(
        {
            "slides": [
                {
                    "type": "profile_card",
                    "name": "A",
                    "role": "B",
                    "years_experience": 3,
                    "seniority": "Mid",
                    "skills": ["x"],
                }
            ]
        }
    )


# ── DeckValidator orchestration ───────────────────────────────────────────────

def test_deck_validator_all_passes(full_raw_deck) -> None:
    DeckValidator().validate(full_raw_deck)


def test_deck_validator_stops_on_schema_error() -> None:
    with pytest.raises(SchemaValidationError):
        DeckValidator().validate({"slides": [{"type": "unknown_type"}]})
