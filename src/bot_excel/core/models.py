from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Sequence


@dataclass(frozen=True)
class MonthlyTemplateRequest:
    year: int
    month: int
    output_dir: Path
    projected_sales: Sequence[int | float] = field(default_factory=tuple)
    projected_gc: Sequence[int | float] = field(default_factory=tuple)
    real_gc: Sequence[int | float] = field(default_factory=tuple)
    projected_hours: Sequence[int | float] = field(default_factory=tuple)
    real_hours: Sequence[int | float] = field(default_factory=tuple)
    ip_target: float = 6.5
    template_id: str | None = None


@dataclass(frozen=True)
class AnnualSalesMonth:
    month: int
    projected_sales: int | float | None = None
    real_sales: int | float | None = None
    projected_gc: int | float | None = None
    real_gc: int | float | None = None
    average_ticket: int | float | None = None


@dataclass(frozen=True)
class AnnualSalesRequest:
    year: int
    output_dir: Path
    months: Sequence[AnnualSalesMonth] = field(default_factory=tuple)
    template_id: str | None = None


@dataclass(frozen=True)
class GeneratedTemplate:
    template_id: str
    workbook_path: Path
    qr_path: Path
    period: str
    template_type: str
