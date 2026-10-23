from __future__ import annotations

from collections.abc import Sequence

from bot_excel.core.models import AnnualSalesMonth


def normalize_annual_months(
    months: Sequence[AnnualSalesMonth],
) -> tuple[AnnualSalesMonth, ...]:
    """Return one normalized record per calendar month.

    This boundary is intentionally independent from workbook generation so a
    future importer can read monthly workbooks and feed the same annual model.
    """
    by_month: dict[int, AnnualSalesMonth] = {}
    for item in months:
        if not 1 <= item.month <= 12:
            raise ValueError("El mes anual debe estar entre 1 y 12.")
        if item.month in by_month:
            raise ValueError(f"El mes {item.month} está repetido.")
        for field_name, value in (
            ("venta proyectada", item.projected_sales),
            ("venta real", item.real_sales),
            ("GC proyectados", item.projected_gc),
            ("GC reales", item.real_gc),
            ("ticket promedio", item.average_ticket),
        ):
            if value is not None and value < 0:
                raise ValueError(
                    f"{field_name}: no se permiten valores negativos."
                )
        by_month[item.month] = item

    return tuple(
        by_month.get(month, AnnualSalesMonth(month=month))
        for month in range(1, 13)
    )
