import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from bot_excel.core.ids import build_annual_template_id, build_template_id
from bot_excel.core.models import MonthlyTemplateRequest
from bot_excel.templates.registry import TemplateRegistry
from bot_excel.templates.ventas_gc.generator import SalesGcTemplateGenerator


class CoreTests(unittest.TestCase):
    def test_template_id(self):
        self.assertEqual(build_template_id("VT", 2026, 9, "A7K2"), "VT-2026-09-A7K2")

    def test_annual_template_id(self):
        self.assertEqual(
            build_annual_template_id("VA", 2026, "A7K2"),
            "VA-2026-A7K2",
        )

    def test_registry(self):
        self.assertEqual(
            TemplateRegistry().available(),
            (
                "ventas_gc",
                "productividad_controlables",
                "ip_semanal_mensual",
                "ventas_consolidadas_anuales",
            ),
        )

    def test_invalid_month(self):
        generator = SalesGcTemplateGenerator()
        request = MonthlyTemplateRequest(2026, 13, Path("/tmp"))
        with self.assertRaises(ValueError):
            generator._validate_request(request)

    def test_negative_projection(self):
        generator = SalesGcTemplateGenerator()
        request = MonthlyTemplateRequest(2026, 9, Path("/tmp"), projected_sales=(-1,))
        with self.assertRaises(ValueError):
            generator._validate_request(request)


if __name__ == "__main__":
    unittest.main()
