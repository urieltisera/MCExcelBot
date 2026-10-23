from __future__ import annotations

import secrets
import string

_ALPHABET = string.ascii_uppercase + string.digits


def generate_short_code(length: int = 4) -> str:
    if length < 4 or length > 12:
        raise ValueError("La longitud del código debe estar entre 4 y 12.")
    return "".join(secrets.choice(_ALPHABET) for _ in range(length))


def build_template_id(prefix: str, year: int, month: int, short_code: str | None = None) -> str:
    suffix = (short_code or generate_short_code()).upper()
    return f"{prefix}-{year}-{month:02d}-{suffix}"


def build_annual_template_id(
    prefix: str,
    year: int,
    short_code: str | None = None,
) -> str:
    suffix = (short_code or generate_short_code()).upper()
    return f"{prefix}-{year}-{suffix}"
