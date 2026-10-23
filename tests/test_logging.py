import logging
import sys
import unittest

from bot_excel.bot.main import _RedactingFormatter


class LoggingSecurityTests(unittest.TestCase):
    def test_token_is_redacted_from_log_message_and_exception(self):
        token = "123456789:TEST_TOKEN_VALUE"
        formatter = _RedactingFormatter("%(message)s", token)
        try:
            raise RuntimeError(f"request failed for bot{token}/getMe")
        except RuntimeError:
            record = logging.LogRecord(
                name="test",
                level=logging.ERROR,
                pathname=__file__,
                lineno=1,
                msg=f"https://api.telegram.org/bot{token}/getMe",
                args=(),
                exc_info=sys.exc_info(),
            )

        rendered = formatter.format(record)
        self.assertNotIn(token, rendered)
        self.assertIn("<redacted>", rendered)

    def test_gemini_key_is_redacted_from_log_message_and_exception(self):
        token = "123456789:TEST_TOKEN_VALUE"
        api_key = "AIzaSy_TEST_GEMINI_SECRET"
        formatter = _RedactingFormatter("%(message)s", token, api_key)
        try:
            raise RuntimeError(f"Gemini request failed with key={api_key}")
        except RuntimeError:
            record = logging.LogRecord(
                name="test",
                level=logging.ERROR,
                pathname=__file__,
                lineno=1,
                msg=f"Gemini key={api_key}",
                args=(),
                exc_info=sys.exc_info(),
            )

        rendered = formatter.format(record)
        self.assertNotIn(api_key, rendered)
        self.assertIn("<redacted>", rendered)


if __name__ == "__main__":
    unittest.main()
