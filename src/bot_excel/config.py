from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(override=True)


@dataclass(frozen=True)
class Settings:
    telegram_bot_token: str
    output_dir: Path
    gemini_api_key: str | None = None
    gemini_model: str = "gemini-3.5-flash-lite"
    max_photo_bytes: int = 10 * 1024 * 1024
    digitization_ttl_seconds: int = 15 * 60
    max_real_sales: int = 1_000_000_000
    max_real_gc: int = 100_000

    @classmethod
    def from_env(cls) -> "Settings":
        token = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
        if not token:
            raise RuntimeError(
                "Falta TELEGRAM_BOT_TOKEN. Creá el bot con BotFather "
                "y configurá el token como variable de entorno."
            )

        output_dir = Path(os.getenv("BOT_OUTPUT_DIR", "./output")).resolve()
        output_dir.mkdir(parents=True, exist_ok=True)
        gemini_api_key = os.getenv("GEMINI_API_KEY", "").strip() or None
        return cls(
            telegram_bot_token=token,
            output_dir=output_dir,
            gemini_api_key=gemini_api_key,
            gemini_model=os.getenv(
                "GEMINI_MODEL", "gemini-3.5-flash-lite"
            ).strip(),
            max_photo_bytes=int(
                os.getenv("MAX_PHOTO_BYTES", str(10 * 1024 * 1024))
            ),
            digitization_ttl_seconds=int(
                os.getenv("DIGITIZATION_TTL_SECONDS", str(15 * 60))
            ),
            max_real_sales=int(
                os.getenv("MAX_REAL_SALES", "1000000000")
            ),
            max_real_gc=int(os.getenv("MAX_REAL_GC", "100000")),
        )
