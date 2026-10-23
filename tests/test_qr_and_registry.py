import tempfile
import unittest
from pathlib import Path

import qrcode

from bot_excel.core.models import MonthlyTemplateRequest
from bot_excel.services.generated_file_registry import (
    GeneratedFileNotFoundError,
    GeneratedFileRegistry,
    MultipleGeneratedFilesError,
)
from bot_excel.services.qr_reader import (
    InvalidQrPayloadError,
    UnsupportedQrTypeError,
    decode_sales_gc_qr,
    parse_sales_gc_qr,
)
from bot_excel.templates.ventas_gc.generator import SalesGcTemplateGenerator


VALID_PAYLOAD = (
    "MCEXCEL|TIPO=VENTAS_GC|CODIGO=VT-2026-06-ABCD|PERIODO=2026-06"
)


class QrAndRegistryTests(unittest.TestCase):
    def test_parse_valid_qr(self):
        parsed = parse_sales_gc_qr(VALID_PAYLOAD)
        self.assertEqual(parsed.type, "VENTAS_GC")
        self.assertEqual(parsed.code, "VT-2026-06-ABCD")
        self.assertEqual(parsed.period, "2026-06")

    def test_reject_invalid_prefix(self):
        with self.assertRaises(InvalidQrPayloadError):
            parse_sales_gc_qr(VALID_PAYLOAD.replace("MCEXCEL", "OTHER"))

    def test_reject_wrong_template_type(self):
        payload = VALID_PAYLOAD.replace("VENTAS_GC", "PRODUCTIVIDAD_CONTROLABLES")
        with self.assertRaises(UnsupportedQrTypeError):
            parse_sales_gc_qr(payload)

    def test_reject_mismatched_code_period(self):
        payload = VALID_PAYLOAD.replace("VT-2026-06-", "VT-2026-05-")
        with self.assertRaisesRegex(InvalidQrPayloadError, "no coinciden"):
            parse_sales_gc_qr(payload)

    def test_decode_qr_image_locally(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "qr.png"
            qrcode.make(VALID_PAYLOAD).save(path)
            self.assertEqual(decode_sales_gc_qr(path).code, "VT-2026-06-ABCD")

    def test_registry_finds_exact_original(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = SalesGcTemplateGenerator().generate(
                MonthlyTemplateRequest(
                    year=2026,
                    month=6,
                    output_dir=Path(tmp),
                    template_id="VT-2026-06-ABCD",
                )
            )
            found = GeneratedFileRegistry(Path(tmp)).find_sales_gc(
                "VT-2026-06-ABCD"
            )
            self.assertEqual(found, result.workbook_path.resolve())

    def test_registry_ignores_digitized_copy(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = SalesGcTemplateGenerator().generate(
                MonthlyTemplateRequest(
                    year=2026,
                    month=6,
                    output_dir=Path(tmp),
                    template_id="VT-2026-06-ABCD",
                )
            )
            result.workbook_path.with_name(
                f"{result.workbook_path.stem}_DIGITALIZADA.xlsx"
            ).write_bytes(result.workbook_path.read_bytes())
            found = GeneratedFileRegistry(Path(tmp)).find_sales_gc(
                "VT-2026-06-ABCD"
            )
            self.assertEqual(found, result.workbook_path.resolve())

    def test_registry_rejects_zero_matches(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(GeneratedFileNotFoundError):
                GeneratedFileRegistry(Path(tmp)).find_sales_gc(
                    "VT-2026-06-ABCD"
                )

    def test_registry_rejects_multiple_originals(self):
        with tempfile.TemporaryDirectory() as tmp:
            result = SalesGcTemplateGenerator().generate(
                MonthlyTemplateRequest(
                    year=2026,
                    month=6,
                    output_dir=Path(tmp),
                    template_id="VT-2026-06-ABCD",
                )
            )
            duplicate = Path(tmp) / (
                "Ventas_GC_Copia_VT-2026-06-ABCD.xlsx"
            )
            duplicate.write_bytes(result.workbook_path.read_bytes())
            with self.assertRaises(MultipleGeneratedFilesError):
                GeneratedFileRegistry(Path(tmp)).find_sales_gc(
                    "VT-2026-06-ABCD"
                )

    def test_registry_rejects_path_traversal_code(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(GeneratedFileNotFoundError):
                GeneratedFileRegistry(Path(tmp)).find_sales_gc("../secret")


if __name__ == "__main__":
    unittest.main()
