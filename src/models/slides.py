from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, Field


class CoverSlide(BaseModel):
    type: Literal["cover"]
    title: str
    subtitle: str | None = None
    client: str | None = None
    date: str | None = None
    author: str | None = None


class SectionDividerSlide(BaseModel):
    type: Literal["section_divider"]
    section_number: str
    section_title: str
    tagline: str | None = None


class BulletColumn(BaseModel):
    title: str | None = None
    bullets: list[str]


class ContentTwoColSlide(BaseModel):
    type: Literal["content_two_col"]
    title: str
    left: BulletColumn
    right: BulletColumn


class ContentOneColSlide(BaseModel):
    type: Literal["content_one_col"]
    title: str
    body_title: str | None = None
    bullets: list[str]


class PricingRow(BaseModel):
    description: str
    quantity: int
    unit: str | None = None  # unidad (horas, dias, licencias, etc.)
    unit_price: float
    total: float


class PricingTableSlide(BaseModel):
    type: Literal["pricing_table"]
    title: str
    currency: str
    rows: list[PricingRow]
    totals: float
    notes: str | None = None


class ProfileCardSlide(BaseModel):
    type: Literal["profile_card"]
    name: str
    role: str
    years_experience: int
    seniority: str
    skills: list[str]
    highlights: list[str] | None = None
    photo: str | None = None


class StatCalloutSlide(BaseModel):
    type: Literal["stat_callout"]
    big_number: str
    label: str
    context: str | None = None
    source: str | None = None


class ClosingSlide(BaseModel):
    type: Literal["closing"]
    headline: str
    contact_name: str
    contact_email: str
    contact_phone: str | None = None
    cta: str | None = None


class ChartSeries(BaseModel):
    name: str
    values: list[float]


class GraficoSlide(BaseModel):
    type: Literal["grafico"]
    title: str
    chart_title: str | None = None
    chart_type: str = "column"  # column, bar, line, pie
    categories: list[str]
    series: list[ChartSeries]
    source: str | None = None


SlideModel = Annotated[
    Union[
        CoverSlide,
        SectionDividerSlide,
        ContentTwoColSlide,
        ContentOneColSlide,
        PricingTableSlide,
        ProfileCardSlide,
        StatCalloutSlide,
        ClosingSlide,
        GraficoSlide,
    ],
    Field(discriminator="type"),
]
