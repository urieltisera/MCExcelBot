from __future__ import annotations

import math


def parse_callback(data: str, prefix: str) -> str:
    expected = f"{prefix}:"
    if not data.startswith(expected):
        raise ValueError(f"Callback inválido: se esperaba el prefijo {expected!r}.")
    value = data[len(expected):]
    if not value:
        raise ValueError("Callback sin valor.")
    return value


def parse_month(data: str) -> int:
    value = int(parse_callback(data, "month"))
    if not 1 <= value <= 12:
        raise ValueError("Mes fuera de rango.")
    return value


def parse_year(data: str) -> int:
    value = int(parse_callback(data, "year"))
    if not 2000 <= value <= 2100:
        raise ValueError("Año fuera de rango.")
    return value


def parse_ip_target(text: str) -> float:
    normalized = text.strip().replace(",", ".")
    try:
        value = float(normalized)
    except ValueError as exc:
        raise ValueError("El objetivo IP debe ser un número, por ejemplo 6,50.") from exc
    if not math.isfinite(value) or not 0 < value <= 100:
        raise ValueError("El objetivo IP debe ser mayor que 0 y menor o igual a 100.")
    return value
