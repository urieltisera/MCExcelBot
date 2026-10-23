from __future__ import annotations

import calendar
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator


class SalesGcDayReading(BaseModel):
    model_config = ConfigDict(extra="forbid")

    day: int = Field(ge=1, le=31)
    real_sales: int | None = Field(default=None, ge=0)
    real_gc: int | None = Field(default=None, ge=0)
    sales_confidence: float | None = Field(default=None, ge=0, le=1)
    gc_confidence: float | None = Field(default=None, ge=0, le=1)
    notes: list[str] = Field(default_factory=list)


class SalesGcPhotoReading(BaseModel):
    model_config = ConfigDict(extra="forbid")

    template_type: Literal["ventas_gc"]
    template_code: str | None = None
    year: int | None = Field(default=None, ge=2000, le=2100)
    month: int | None = Field(default=None, ge=1, le=12)
    rows: list[SalesGcDayReading]
    image_quality: Literal["good", "acceptable", "poor"]
    warnings: list[str] = Field(default_factory=list)

    @model_validator(mode="after")
    def reject_duplicate_days(self) -> "SalesGcPhotoReading":
        days = [row.day for row in self.rows]
        if len(days) != len(set(days)):
            raise ValueError("La lectura contiene días duplicados.")
        return self

    def validate_period(
        self,
        *,
        expected_code: str,
        expected_year: int,
        expected_month: int,
    ) -> None:
        if self.template_code is not None and self.template_code != expected_code:
            raise ValueError("El código leído no coincide con el QR.")
        if self.year is not None and self.year != expected_year:
            raise ValueError("El año leído no coincide con el QR.")
        if self.month is not None and self.month != expected_month:
            raise ValueError("El mes leído no coincide con el QR.")

        days_in_month = calendar.monthrange(expected_year, expected_month)[1]
        invalid_days = sorted(
            row.day for row in self.rows if row.day > days_in_month
        )
        if invalid_days:
            invalid = ", ".join(str(day) for day in invalid_days)
            raise ValueError(
                f"La lectura contiene días inexistentes para el período: {invalid}."
            )

    def populated_field_count(self) -> int:
        return sum(
            value is not None
            for row in self.rows
            for value in (row.real_sales, row.real_gc)
        )
