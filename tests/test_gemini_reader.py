import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

from bot_excel.core.digitization_models import SalesGcPhotoReading
from bot_excel.services.gemini_reader import (
    GeminiConfigurationError,
    GeminiReaderError,
    GeminiResponseError,
    GeminiSalesGcReader,
)


VALID = SalesGcPhotoReading.model_validate({
    "template_type": "ventas_gc",
    "template_code": "VT-2026-06-ABCD",
    "year": 2026,
    "month": 6,
    "rows": [{
        "day": 1,
        "real_sales": 100,
        "real_gc": 20,
        "sales_confidence": 0.99,
        "gc_confidence": 0.98,
    }],
    "image_quality": "good",
})


class FakeModels:
    def __init__(self, response=None, error=None):
        self.response = response
        self.error = error
        self.last_call = None

    def generate_content(self, **kwargs):
        self.last_call = kwargs
        if self.error:
            raise self.error
        return self.response


class FakeClient:
    def __init__(self, models):
        self.models = models
        self.closed = False

    def close(self):
        self.closed = True


class GeminiReaderTests(unittest.TestCase):
    def image_file(self, directory: str) -> Path:
        path = Path(directory) / "photo.jpg"
        path.write_bytes(b"fake image bytes")
        return path

    def test_missing_key_is_rejected_without_api_call(self):
        with tempfile.TemporaryDirectory() as tmp:
            reader = GeminiSalesGcReader(api_key=None)
            with self.assertRaises(GeminiConfigurationError):
                reader.read_image(
                    self.image_file(tmp),
                    mime_type="image/jpeg",
                )

    def test_structured_pydantic_response_is_returned(self):
        with tempfile.TemporaryDirectory() as tmp:
            models = FakeModels(SimpleNamespace(parsed=VALID, text=None))
            client = FakeClient(models)
            reader = GeminiSalesGcReader(
                api_key="secret",
                client_factory=lambda **_: client,
            )
            result = reader.read_image(
                self.image_file(tmp),
                mime_type="image/jpeg",
            )
            self.assertEqual(result.rows[0].real_sales, 100)
            self.assertEqual(models.last_call["model"], "gemini-3.5-flash-lite")
            self.assertTrue(client.closed)

    def test_invalid_json_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            client = FakeClient(FakeModels(
                SimpleNamespace(parsed=None, text='{"bad": true}')
            ))
            reader = GeminiSalesGcReader(
                api_key="secret",
                client_factory=lambda **_: client,
            )
            with self.assertRaises(GeminiResponseError):
                reader.read_image(
                    self.image_file(tmp),
                    mime_type="image/png",
                )

    def test_quota_error_is_safely_classified(self):
        with tempfile.TemporaryDirectory() as tmp:
            secret = "VERY_SECRET_API_KEY"
            client = FakeClient(FakeModels(
                error=RuntimeError(f"429 RESOURCE_EXHAUSTED {secret}")
            ))
            reader = GeminiSalesGcReader(
                api_key=secret,
                client_factory=lambda **_: client,
            )
            with self.assertRaises(GeminiReaderError) as caught:
                reader.read_image(
                    self.image_file(tmp),
                    mime_type="image/webp",
                )
            self.assertIn("cuota", str(caught.exception).lower())
            self.assertNotIn(secret, str(caught.exception))

    def test_non_image_mime_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            reader = GeminiSalesGcReader(api_key="secret")
            with self.assertRaises(GeminiReaderError):
                reader.read_image(
                    self.image_file(tmp),
                    mime_type="application/pdf",
                )


if __name__ == "__main__":
    unittest.main()
