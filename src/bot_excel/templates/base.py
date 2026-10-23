from __future__ import annotations

from abc import ABC, abstractmethod
from bot_excel.core.models import (
    AnnualSalesRequest,
    GeneratedTemplate,
    MonthlyTemplateRequest,
)


class MonthlyTemplateGenerator(ABC):
    template_key: str
    display_name: str

    @abstractmethod
    def generate(self, request: MonthlyTemplateRequest) -> GeneratedTemplate:
        raise NotImplementedError


class AnnualTemplateGenerator(ABC):
    template_key: str
    display_name: str

    @abstractmethod
    def generate(self, request: AnnualSalesRequest) -> GeneratedTemplate:
        raise NotImplementedError
