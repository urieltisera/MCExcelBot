import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from bot_excel.config import Settings


class GeminiConfigurationTests(unittest.TestCase):
    def test_gemini_key_is_loaded_without_being_exposed(self):
        secret = "AIzaSy_TEST_ONLY_NOT_REAL"
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {
                "TELEGRAM_BOT_TOKEN": "123456:TEST_TOKEN",
                "BOT_OUTPUT_DIR": str(Path(tmp) / "output"),
                "GEMINI_API_KEY": secret,
            },
            clear=False,
        ):
            settings = Settings.from_env()
        self.assertEqual(settings.gemini_api_key, secret)
        self.assertNotIn(secret, repr({
            "output_dir": str(settings.output_dir),
            "model": settings.gemini_model,
        }))

    def test_missing_gemini_key_does_not_prevent_bot_startup(self):
        with tempfile.TemporaryDirectory() as tmp, patch.dict(
            os.environ,
            {
                "TELEGRAM_BOT_TOKEN": "123456:TEST_TOKEN",
                "BOT_OUTPUT_DIR": str(Path(tmp) / "output"),
                "GEMINI_API_KEY": "",
            },
            clear=False,
        ):
            settings = Settings.from_env()
        self.assertIsNone(settings.gemini_api_key)


if __name__ == "__main__":
    unittest.main()
