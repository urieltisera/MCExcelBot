import calendar
import unittest

from pydantic import ValidationError

from bot_excel.core.digitization_models import (
    SalesGcDayReading,
    SalesGcPhotoReading,
)


def reading(**overrides):
    values = {
        "template_type": "ventas_gc",
        "template_code": "VT-2026-06-ABCD",
        "year": 2026,
        "month": 6,
        "rows": [
            {
                "day": 1,
                "real_sales": 9_564_195,
                "real_gc": 504,
                "sales_confidence": 0.98,
                "gc_confidence": 0.96,
            }
        ],
        "image_quality": "good",
    }
    values.update(overrides)
    return SalesGcPhotoReading.model_validate(values)


class DigitizationModelTests(unittest.TestCase):
    def test_valid_structured_model(self):
        result = reading()
        self.assertEqual(result.populated_field_count(), 2)

    def test_negative_sales_is_rejected(self):
        with self.assertRaises(ValidationError):
            SalesGcDayReading(day=1, real_sales=-1)

    def test_negative_gc_is_rejected(self):
        with self.assertRaises(ValidationError):
            SalesGcDayReading(day=1, real_gc=-1)

    def test_confidence_outside_range_is_rejected(self):
        with self.assertRaises(ValidationError):
            SalesGcDayReading(day=1, sales_confidence=1.1)

    def test_duplicate_days_are_rejected(self):
        with self.assertRaises(ValidationError):
            reading(rows=[
                {"day": 1, "real_sales": 10},
                {"day": 1, "real_gc": 20},
            ])

    def test_day_outside_month_is_rejected(self):
        result = reading(
            template_code="VT-2026-02-ABCD",
            month=2,
            rows=[{"day": 30, "real_sales": 10}],
        )
        with self.assertRaisesRegex(ValueError, "inexistentes"):
            result.validate_period(
                expected_code="VT-2026-02-ABCD",
                expected_year=2026,
                expected_month=2,
            )

    def test_leap_day_is_accepted(self):
        self.assertEqual(calendar.monthrange(2028, 2)[1], 29)
        result = reading(
            template_code="VT-2028-02-ABCD",
            year=2028,
            month=2,
            rows=[{"day": 29, "real_sales": 10}],
        )
        result.validate_period(
            expected_code="VT-2028-02-ABCD",
            expected_year=2028,
            expected_month=2,
        )

    def test_code_period_mismatch_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "código"):
            reading().validate_period(
                expected_code="VT-2026-06-WXYZ",
                expected_year=2026,
                expected_month=6,
            )

    def test_extra_json_fields_are_rejected(self):
        with self.assertRaises(ValidationError):
            reading(unexpected="not allowed")


if __name__ == "__main__":
    unittest.main()
