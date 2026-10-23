import re
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from openpyxl import load_workbook

from bot_excel.core.models import MonthlyTemplateRequest
from bot_excel.templates.ventas_gc.generator import SalesGcTemplateGenerator


DAILY_ROWS = [*range(4, 19), *range(20, 36)]


class SalesAccumulatorTests(unittest.TestCase):
    def generate_sheet(self, *, year=2027, month=5):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        result = SalesGcTemplateGenerator().generate(
            MonthlyTemplateRequest(
                year=year,
                month=month,
                output_dir=Path(temporary.name),
                template_id=f"VT-{year}-{month:02d}-ACC",
            )
        )
        workbook = load_workbook(result.workbook_path, data_only=False)
        self.addCleanup(workbook.close)
        return workbook["CONSOLIDADO"]

    @staticmethod
    def evaluate_accumulation(sheet, real_column, projected_column):
        running_total = 0
        displayed = {}
        for row in DAILY_ROWS:
            if sheet[f"B{row}"].value in (None, ""):
                continue
            real_value = sheet[f"{real_column}{row}"].value
            projected_value = sheet[f"{projected_column}{row}"].value
            if real_value in (None, ""):
                displayed[row] = ""
                continue
            difference = real_value - projected_value
            running_total += difference
            displayed[row] = running_total
        return displayed

    def test_interleaved_values_do_not_chain_through_blank_accumulator(self):
        sheet = self.generate_sheet()
        sheet["D4"] = 345
        sheet["H4"] = 793
        sheet["D5"] = ""
        sheet["H5"] = ""
        sheet["D6"] = 890
        sheet["H6"] = 543

        sales = self.evaluate_accumulation(sheet, "D", "C")
        gc = self.evaluate_accumulation(sheet, "H", "G")

        self.assertEqual(sales[4], 345)
        self.assertEqual(sales[5], "")
        self.assertEqual(sales[6], 1235)
        self.assertEqual(gc[4], 793)
        self.assertEqual(gc[5], "")
        self.assertEqual(gc[6], 1336)
        self.assertEqual(sheet["F6"].value, '=IF(E6="","",SUM($E$4:E6))')
        self.assertEqual(sheet["J6"].value, '=IF(I6="","",SUM($I$4:I6))')

    def test_multiple_blanks_first_midmonth_and_after_first_half(self):
        sheet = self.generate_sheet()
        sheet["D10"] = 100
        sheet["H10"] = 200
        sheet["D20"] = 300
        sheet["H20"] = 400
        sheet["D23"] = 500
        sheet["H23"] = 600

        sales = self.evaluate_accumulation(sheet, "D", "C")
        gc = self.evaluate_accumulation(sheet, "H", "G")

        self.assertEqual(sales[10], 100)
        self.assertEqual(sales[20], 400)
        self.assertEqual(sales[23], 900)
        self.assertEqual(gc[10], 200)
        self.assertEqual(gc[20], 600)
        self.assertEqual(gc[23], 1200)
        self.assertEqual(
            sheet["F20"].value,
            '=IF(E20="","",SUM($E$4:$E$18,$E$20:E20))',
        )
        self.assertEqual(
            sheet["J23"].value,
            '=IF(I23="","",SUM($I$4:$I$18,$I$20:I23))',
        )

    def test_empty_and_full_months(self):
        empty_sheet = self.generate_sheet()
        empty_sales = self.evaluate_accumulation(empty_sheet, "D", "C")
        empty_gc = self.evaluate_accumulation(empty_sheet, "H", "G")
        self.assertTrue(all(value == "" for value in empty_sales.values()))
        self.assertTrue(all(value == "" for value in empty_gc.values()))

        full_sheet = self.generate_sheet()
        month_rows = [
            row for row in DAILY_ROWS if full_sheet[f"B{row}"].value not in (None, "")
        ]
        for row in month_rows:
            full_sheet[f"D{row}"] = 10
            full_sheet[f"H{row}"] = 2

        full_sales = self.evaluate_accumulation(full_sheet, "D", "C")
        full_gc = self.evaluate_accumulation(full_sheet, "H", "G")
        self.assertEqual(full_sales[month_rows[-1]], 10 * len(month_rows))
        self.assertEqual(full_gc[month_rows[-1]], 2 * len(month_rows))

    def test_month_lengths_and_all_accumulators_use_sum_ranges(self):
        cases = (
            (2025, 2, 28),
            (2024, 2, 29),
            (2027, 4, 30),
            (2027, 5, 31),
        )
        chained_pattern = re.compile(r"[FJ]\d+\+[EI]\d+")

        for year, month, expected_days in cases:
            with self.subTest(year=year, month=month):
                sheet = self.generate_sheet(year=year, month=month)
                populated_days = [
                    row
                    for row in DAILY_ROWS
                    if sheet[f"B{row}"].value not in (None, "")
                ]
                self.assertEqual(len(populated_days), expected_days)

                for row in DAILY_ROWS:
                    for column in ("F", "J"):
                        formula = sheet[f"{column}{row}"].value
                        self.assertIsInstance(formula, str)
                        self.assertIn("SUM(", formula)
                        self.assertIsNone(chained_pattern.search(formula))

                self.assertEqual(
                    sheet["F35"].value,
                    '=IF(E35="","",SUM($E$4:$E$18,$E$20:E35))',
                )
                self.assertEqual(
                    sheet["J35"].value,
                    '=IF(I35="","",SUM($I$4:$I$18,$I$20:I35))',
                )


if __name__ == "__main__":
    unittest.main()
