from __future__ import annotations

import logging

from bot_excel.bot.application import build_application
from bot_excel.config import Settings


class _RedactingFormatter(logging.Formatter):
    def __init__(self, format_string: str, *secrets: str | None) -> None:
        super().__init__(format_string)
        self._secrets = tuple(secret for secret in secrets if secret)

    def format(self, record: logging.LogRecord) -> str:
        rendered = super().format(record)
        for secret in self._secrets:
            rendered = rendered.replace(secret, "<redacted>")
        return rendered


def main() -> None:
    settings = Settings.from_env()
    handler = logging.StreamHandler()
    handler.setFormatter(_RedactingFormatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        settings.telegram_bot_token,
        settings.gemini_api_key,
    ))
    logging.basicConfig(
        level=logging.INFO,
        handlers=[handler],
        force=True,
    )
    logging.getLogger("httpx").setLevel(logging.WARNING)
    logging.getLogger("httpcore").setLevel(logging.WARNING)
    application = build_application(settings)
    application.run_polling(
        allowed_updates=["message", "callback_query"],
        drop_pending_updates=False,
    )


if __name__ == "__main__":
    main()
