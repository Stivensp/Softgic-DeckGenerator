from __future__ import annotations

import logging
import re

from src.exceptions import BusinessRuleError

logger = logging.getLogger("softgic.validator.business")

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


class BusinessRulesValidator:
    def validate(self, data: dict, limits: dict | None = None) -> None:  # type: ignore[type-arg]
        lim = limits or {}
        for idx, slide in enumerate(data.get("slides", [])):
            slide_type = slide.get("type", "")
            getattr(self, f"_check_{slide_type}", self._noop)(slide, idx, lim)

    def _noop(self, *args: object) -> None:
        pass

    def _check_cover(self, slide: dict, idx: int, lim: dict) -> None:  # type: ignore[type-arg]
        self._max_len(slide, "title", lim.get("title_max_chars", 80), idx)
        self._max_len(slide, "subtitle", lim.get("subtitle_max_chars", 120), idx)

    def _check_section_divider(self, slide: dict, idx: int, lim: dict) -> None:  # type: ignore[type-arg]
        self._max_len(slide, "section_title", lim.get("title_max_chars", 80), idx)

    def _check_content_two_col(self, slide: dict, idx: int, lim: dict) -> None:  # type: ignore[type-arg]
        self._max_len(slide, "title", lim.get("title_max_chars", 80), idx)
        max_bullets = lim.get("bullets_per_column_max", 6)
        bullet_max_chars = lim.get("bullet_max_chars", 100)
        for col_name in ("left", "right"):
            col = slide.get(col_name, {})
            bullets = col.get("bullets", [])
            if len(bullets) > max_bullets:
                raise BusinessRuleError(
                    f"Column '{col_name}' has {len(bullets)} bullets; maximum is {max_bullets}",
                    field=f"{col_name}.bullets",
                    slide_index=idx,
                )
            self._max_len_list(bullets, bullet_max_chars, f"{col_name}.bullets", idx)

    def _check_content_one_col(self, slide: dict, idx: int, lim: dict) -> None:  # type: ignore[type-arg]
        self._max_len(slide, "title", lim.get("title_max_chars", 80), idx)
        max_bullets = lim.get("bullets_per_column_max", 6) * 2
        bullets = slide.get("bullets", [])
        if len(bullets) > max_bullets:
            raise BusinessRuleError(
                f"content_one_col has {len(bullets)} bullets; maximum is {max_bullets}",
                field="bullets",
                slide_index=idx,
            )
        self._max_len_list(bullets, lim.get("bullet_max_chars", 100), "bullets", idx)

    def _check_grafico(self, slide: dict, idx: int, lim: dict) -> None:  # type: ignore[type-arg]
        self._max_len(slide, "title", lim.get("title_max_chars", 80), idx)
        categories = slide.get("categories", [])
        if not categories:
            raise BusinessRuleError(
                "grafico: 'categories' must have at least one entry",
                field="categories",
                slide_index=idx,
            )
        series = slide.get("series", [])
        if not series:
            raise BusinessRuleError(
                "grafico: 'series' must have at least one entry",
                field="series",
                slide_index=idx,
            )
        for i, s in enumerate(series):
            if isinstance(s, dict):
                vals = s.get("values", [])
                if len(vals) != len(categories):
                    raise BusinessRuleError(
                        f"grafico: series[{i}] has {len(vals)} values but there are {len(categories)} categories",
                        field=f"series[{i}].values",
                        slide_index=idx,
                    )

    def _check_pricing_table(self, slide: dict, idx: int, lim: dict) -> None:  # type: ignore[type-arg]
        self._max_len(slide, "title", lim.get("title_max_chars", 80), idx)
        max_rows = lim.get("pricing_rows_max", 10)
        rows = slide.get("rows", [])
        if len(rows) > max_rows:
            raise BusinessRuleError(
                f"Pricing table has {len(rows)} rows; maximum is {max_rows}",
                field="rows",
                slide_index=idx,
            )
        computed_total = 0.0
        for i, row in enumerate(rows):
            if isinstance(row, dict):
                qty = row.get("quantity", 0)
                if isinstance(qty, int) and qty < 0:
                    raise BusinessRuleError(
                        f"Row {i}: 'quantity' must be >= 0", field=f"rows[{i}].quantity", slide_index=idx
                    )
                computed_total += float(row.get("total", 0.0))
                self._max_len(row, "description", lim.get("bullet_max_chars", 100), idx, field_label=f"rows[{i}].description")
        declared = slide.get("totals")
        if declared is not None and abs(computed_total - float(declared)) > 0.01:
            raise BusinessRuleError(
                f"pricing_table: suma de row.total ({computed_total:.2f}) != totals declarado ({float(declared):.2f})",
                field="totals",
                slide_index=idx,
            )

    def _check_profile_card(self, slide: dict, idx: int, lim: dict) -> None:  # type: ignore[type-arg]
        self._max_len(slide, "name", lim.get("title_max_chars", 80), idx)
        max_skills = lim.get("skills_max", 8)
        skills = slide.get("skills", [])
        if len(skills) > max_skills:
            raise BusinessRuleError(
                f"Profile card has {len(skills)} skills; maximum is {max_skills}",
                field="skills",
                slide_index=idx,
            )
        yoe = slide.get("years_experience")
        if yoe is not None and isinstance(yoe, int) and yoe < 0:
            raise BusinessRuleError(
                "'years_experience' must be >= 0", field="years_experience", slide_index=idx
            )
        self._max_len_list(slide.get("highlights") or [], lim.get("bullet_max_chars", 100), "highlights", idx)

    def _check_stat_callout(self, slide: dict, idx: int, lim: dict) -> None:  # type: ignore[type-arg]
        self._max_len(slide, "label", lim.get("title_max_chars", 80), idx)

    def _check_closing(self, slide: dict, idx: int, lim: dict) -> None:  # type: ignore[type-arg]
        self._max_len(slide, "headline", lim.get("title_max_chars", 80), idx)
        email = slide.get("contact_email", "")
        if email and not _EMAIL_RE.match(email):
            raise BusinessRuleError(
                f"'contact_email' is not a valid email address: '{email}'",
                field="contact_email",
                slide_index=idx,
            )

    def _max_len(self, slide: dict, field: str, max_chars: int, idx: int, field_label: str | None = None) -> None:  # type: ignore[type-arg]
        value = slide.get(field)
        if value and isinstance(value, str) and len(value) > max_chars:
            raise BusinessRuleError(
                f"Field '{field_label or field}' is {len(value)} chars; maximum is {max_chars}",
                field=field_label or field,
                slide_index=idx,
            )

    def _max_len_list(self, items: list, max_chars: int, field_label: str, idx: int) -> None:  # type: ignore[type-arg]
        """Valida la longitud de cada item de texto en una lista (bullets, highlights, etc.).

        Consistente con _max_len: si un item individual excede el limite, falla
        duro ANTES del render — evita que el renderer trunque silenciosamente
        con '...' sin que el usuario se entere (asimetria que existia antes:
        solo se validaba la CANTIDAD de items, nunca el largo de cada uno).
        """
        for i, item in enumerate(items):
            if isinstance(item, str) and len(item) > max_chars:
                raise BusinessRuleError(
                    f"Item {i} in '{field_label}' is {len(item)} chars; maximum is {max_chars}",
                    field=f"{field_label}[{i}]",
                    slide_index=idx,
                )
