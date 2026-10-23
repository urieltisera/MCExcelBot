import shutil
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock, patch

from bot_excel.bot.digitization_handlers import (
    PENDING_KEY,
    _validate_reading,
    handle_digitization_action,
    receive_sales_gc_photo,
    reject_unsupported_document,
)
from bot_excel.core.digitization_models import SalesGcPhotoReading
from bot_excel.core.models import MonthlyTemplateRequest
from bot_excel.services.gemini_reader import GeminiReaderError
from bot_excel.services.qr_reader import ParsedQr, QrNotDetectedError
from bot_excel.templates.ventas_gc.generator import SalesGcTemplateGenerator


def valid_reading(confidence=0.99):
    return SalesGcPhotoReading.model_validate({
        "template_type": "ventas_gc",
        "template_code": "VT-2026-06-ABCD",
        "year": 2026,
        "month": 6,
        "rows": [{
            "day": 1,
            "real_sales": 345,
            "real_gc": 793,
            "sales_confidence": confidence,
            "gc_confidence": confidence,
        }],
        "image_quality": "good",
    })


class FakeTelegramFile:
    def __init__(self, source: Path):
        self.source = source

    async def download_to_drive(self, custom_path):
        shutil.copy2(self.source, custom_path)
        return Path(custom_path)


class FakePhoto:
    def __init__(self, source: Path):
        self.file_size = source.stat().st_size
        self.file = FakeTelegramFile(source)

    async def get_file(self):
        return self.file


class FakeReader:
    def __init__(self, result=None, error=None):
        self.result = result
        self.error = error

    def read_image(self, image_path, *, mime_type):
        if self.error:
            raise self.error
        return self.result


class DigitizationBotFlowTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.temp = tempfile.TemporaryDirectory()
        root = Path(self.temp.name)
        self.output_dir = root / "output"
        self.temp_dir = root / "runtime" / "temp"
        self.output_dir.mkdir()
        self.temp_dir.mkdir(parents=True)
        self.original = SalesGcTemplateGenerator().generate(
            MonthlyTemplateRequest(
                year=2026,
                month=6,
                output_dir=self.output_dir,
                template_id="VT-2026-06-ABCD",
            )
        ).workbook_path
        self.source_image = root / "source.jpg"
        self.source_image.write_bytes(b"test photo")
        self.message = SimpleNamespace(
            photo=[FakePhoto(self.source_image)],
            document=None,
            reply_text=AsyncMock(),
            reply_document=AsyncMock(),
        )
        self.bot_data = {
            "output_dir": str(self.output_dir),
            "temp_dir": str(self.temp_dir),
            "max_photo_bytes": 10 * 1024 * 1024,
            "digitization_ttl_seconds": 900,
            "max_real_sales": 1_000_000_000,
            "max_real_gc": 100_000,
            "gemini_api_key": "not-used-by-fake",
            "gemini_model": "gemini-3.5-flash-lite",
            "gemini_reader": FakeReader(valid_reading()),
        }
        self.context = SimpleNamespace(
            user_data={},
            application=SimpleNamespace(bot_data=self.bot_data),
        )
        self.update = SimpleNamespace(effective_message=self.message)
        self.qr = ParsedQr(
            type="VENTAS_GC",
            code="VT-2026-06-ABCD",
            period="2026-06",
            year=2026,
            month=6,
            payload=(
                "MCEXCEL|TIPO=VENTAS_GC|CODIGO="
                "VT-2026-06-ABCD|PERIODO=2026-06"
            ),
        )

    async def asyncTearDown(self):
        self.temp.cleanup()

    async def receive(self):
        with patch(
            "bot_excel.bot.digitization_handlers.decode_sales_gc_qr",
            return_value=self.qr,
        ):
            await receive_sales_gc_photo(self.update, self.context)

    async def test_photo_flow_waits_for_confirmation(self):
        await self.receive()
        self.assertIn(PENDING_KEY, self.context.user_data)
        self.assertFalse(any(
            "_DIGITALIZADA" in path.name
            for path in self.output_dir.glob("*.xlsx")
        ))
        summary = self.message.reply_text.await_args_list[-1].args[0]
        self.assertIn("Ventas y GC — Junio 2026", summary)
        self.assertIn("$345", summary)

    async def test_confirmation_creates_and_sends_copy(self):
        await self.receive()
        pending = self.context.user_data[PENDING_KEY]
        query = SimpleNamespace(
            data=f"digitize:confirm:{pending['operation_id']}",
            answer=AsyncMock(),
            edit_message_text=AsyncMock(),
            message=self.message,
        )
        await handle_digitization_action(
            SimpleNamespace(callback_query=query),
            self.context,
        )
        self.message.reply_document.assert_awaited_once()
        self.assertNotIn(PENDING_KEY, self.context.user_data)
        self.assertEqual(list(self.temp_dir.iterdir()), [])

    async def test_cancel_cleans_temp_without_copy(self):
        await self.receive()
        pending = self.context.user_data[PENDING_KEY]
        query = SimpleNamespace(
            data=f"digitize:cancel:{pending['operation_id']}",
            answer=AsyncMock(),
            edit_message_text=AsyncMock(),
            message=self.message,
        )
        await handle_digitization_action(
            SimpleNamespace(callback_query=query),
            self.context,
        )
        self.assertNotIn(PENDING_KEY, self.context.user_data)
        self.assertEqual(list(self.temp_dir.iterdir()), [])
        self.message.reply_document.assert_not_awaited()

    async def test_expired_operation_is_cleaned(self):
        await self.receive()
        pending = self.context.user_data[PENDING_KEY]
        pending["expires_at"] = time.time() - 1
        query = SimpleNamespace(
            data=f"digitize:confirm:{pending['operation_id']}",
            answer=AsyncMock(),
            edit_message_text=AsyncMock(),
            message=self.message,
        )
        await handle_digitization_action(
            SimpleNamespace(callback_query=query),
            self.context,
        )
        self.assertNotIn(PENDING_KEY, self.context.user_data)
        self.assertIn(
            "venció",
            query.edit_message_text.await_args.args[0],
        )

    async def test_gemini_error_cleans_temp(self):
        self.bot_data["gemini_reader"] = FakeReader(
            error=GeminiReaderError("Falla segura.")
        )
        await self.receive()
        self.assertNotIn(PENDING_KEY, self.context.user_data)
        self.assertEqual(list(self.temp_dir.iterdir()), [])
        self.assertIn(
            "original no fue modificado",
            self.message.reply_text.await_args_list[-1].args[0],
        )

    async def test_document_image_is_accepted(self):
        document = SimpleNamespace(
            mime_type="image/png",
            file_size=self.source_image.stat().st_size,
            get_file=AsyncMock(
                return_value=FakeTelegramFile(self.source_image)
            ),
        )
        self.message.photo = []
        self.message.document = document
        await self.receive()
        self.assertIn(PENDING_KEY, self.context.user_data)

    async def test_unsupported_document_gets_safe_message(self):
        await reject_unsupported_document(self.update, self.context)
        self.assertIn(
            "JPG, PNG o WEBP",
            self.message.reply_text.await_args.args[0],
        )

    async def test_oversized_photo_is_rejected_before_download(self):
        self.bot_data["max_photo_bytes"] = 1
        await self.receive()
        self.assertNotIn(PENDING_KEY, self.context.user_data)
        self.assertIn(
            "demasiado grande",
            self.message.reply_text.await_args.args[0],
        )

    async def test_photo_without_qr_is_rejected_and_cleaned(self):
        with patch(
            "bot_excel.bot.digitization_handlers.decode_sales_gc_qr",
            side_effect=QrNotDetectedError("not found"),
        ):
            await receive_sales_gc_photo(self.update, self.context)
        self.assertNotIn(PENDING_KEY, self.context.user_data)
        self.assertEqual(list(self.temp_dir.iterdir()), [])
        self.assertIn(
            "No pude identificar el QR",
            self.message.reply_text.await_args_list[-1].args[0],
        )

    def test_doubtful_confidence_is_classified(self):
        validation = _validate_reading(
            valid_reading(0.80),
            max_real_sales=1_000_000_000,
            max_real_gc=100_000,
        )
        self.assertEqual(len(validation["doubtful"]), 2)
        self.assertEqual(validation["illegible"], [])

    def test_low_confidence_is_illegible(self):
        validation = _validate_reading(
            valid_reading(0.50),
            max_real_sales=1_000_000_000,
            max_real_gc=100_000,
        )
        self.assertEqual(len(validation["illegible"]), 2)


if __name__ == "__main__":
    unittest.main()
