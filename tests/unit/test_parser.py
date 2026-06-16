from __future__ import annotations

import pytest

from src.exceptions import ParseError
from src.models.slides import (
    ClosingSlide,
    ContentTwoColSlide,
    CoverSlide,
    PricingTableSlide,
    ProfileCardSlide,
    SectionDividerSlide,
    StatCalloutSlide,
)
from src.parser.deck_parser import DeckParser


@pytest.fixture
def parser() -> DeckParser:
    return DeckParser()


def test_parse_cover(parser: DeckParser) -> None:
    deck = parser.parse({"slides": [{"type": "cover", "title": "Hello"}]})
    assert len(deck.slides) == 1
    assert isinstance(deck.slides[0], CoverSlide)
    assert deck.slides[0].title == "Hello"


def test_parse_section_divider(parser: DeckParser) -> None:
    deck = parser.parse(
        {
            "slides": [
                {"type": "section_divider", "section_number": "02", "section_title": "Context"}
            ]
        }
    )
    slide = deck.slides[0]
    assert isinstance(slide, SectionDividerSlide)
    assert slide.section_number == "02"


def test_parse_content_two_col(parser: DeckParser) -> None:
    deck = parser.parse(
        {
            "slides": [
                {
                    "type": "content_two_col",
                    "title": "T",
                    "left": {"bullets": ["a"]},
                    "right": {"bullets": ["b"]},
                }
            ]
        }
    )
    assert isinstance(deck.slides[0], ContentTwoColSlide)


def test_parse_pricing_table(parser: DeckParser) -> None:
    deck = parser.parse(
        {
            "slides": [
                {
                    "type": "pricing_table",
                    "title": "P",
                    "currency": "COP",
                    "rows": [
                        {"description": "X", "quantity": 1, "unit_price": 10.0, "total": 10.0}
                    ],
                    "totals": 10.0,
                }
            ]
        }
    )
    assert isinstance(deck.slides[0], PricingTableSlide)
    assert deck.slides[0].currency == "COP"


def test_parse_profile_card(parser: DeckParser) -> None:
    deck = parser.parse(
        {
            "slides": [
                {
                    "type": "profile_card",
                    "name": "Ana",
                    "role": "Dev",
                    "years_experience": 5,
                    "seniority": "Mid",
                    "skills": ["Python"],
                }
            ]
        }
    )
    assert isinstance(deck.slides[0], ProfileCardSlide)
    assert deck.slides[0].skills == ["Python"]


def test_parse_stat_callout(parser: DeckParser) -> None:
    deck = parser.parse(
        {"slides": [{"type": "stat_callout", "big_number": "99%", "label": "Uptime"}]}
    )
    assert isinstance(deck.slides[0], StatCalloutSlide)


def test_parse_closing(parser: DeckParser) -> None:
    deck = parser.parse(
        {
            "slides": [
                {
                    "type": "closing",
                    "headline": "H",
                    "contact_name": "N",
                    "contact_email": "n@s.com",
                }
            ]
        }
    )
    assert isinstance(deck.slides[0], ClosingSlide)


def test_parse_optional_fields_default_none(parser: DeckParser) -> None:
    deck = parser.parse({"slides": [{"type": "cover", "title": "Min"}]})
    slide = deck.slides[0]
    assert isinstance(slide, CoverSlide)
    assert slide.subtitle is None
    assert slide.client is None
    assert slide.author is None


def test_parse_all_seven_types(parser: DeckParser, full_raw_deck: dict) -> None:
    deck = parser.parse(full_raw_deck)
    assert len(deck.slides) == 7


def test_parse_invalid_data_raises_parse_error(parser: DeckParser) -> None:
    with pytest.raises(ParseError):
        parser.parse({"slides": [{"type": "cover"}]})  # missing required title
