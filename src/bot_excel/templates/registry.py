from __future__ import annotations

from bot_excel.templates.base import AnnualTemplateGenerator, MonthlyTemplateGenerator
from bot_excel.templates.ip.generator import IpTemplateGenerator
from bot_excel.templates.productividad_controlables.generator import (
    ProductivityControllablesTemplateGenerator,
)
from bot_excel.templates.ventas_anuales.generator import AnnualSalesTemplateGenerator
from bot_excel.templates.ventas_gc.generator import SalesGcTemplateGenerator


class TemplateRegistry:
    def __init__(self) -> None:
        generators: list[MonthlyTemplateGenerator | AnnualTemplateGenerator] = [
            SalesGcTemplateGenerator(),
            ProductivityControllablesTemplateGenerator(),
            IpTemplateGenerator(),
            AnnualSalesTemplateGenerator(),
        ]
        self._generators = {generator.template_key: generator for generator in generators}

    def get(
        self,
        key: str,
    ) -> MonthlyTemplateGenerator | AnnualTemplateGenerator:
        try:
            return self._generators[key]
        except KeyError as exc:
            raise KeyError(f"Plantilla no registrada: {key}") from exc

    def available(self) -> tuple[str, ...]:
        return tuple(self._generators.keys())
