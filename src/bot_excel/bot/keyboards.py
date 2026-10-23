from __future__ import annotations

from datetime import datetime

from telegram import InlineKeyboardButton, InlineKeyboardMarkup


MONTHS = (
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
)


def templates_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("Ventas y GC", callback_data="template:ventas_gc")],
        [InlineKeyboardButton(
            "Productividad y controlables",
            callback_data="template:productividad_controlables",
        )],
        [InlineKeyboardButton(
            "IP semanal y mensual",
            callback_data="template:ip_semanal_mensual",
        )],
        [InlineKeyboardButton(
            "Ventas consolidadas anuales",
            callback_data="template:ventas_consolidadas_anuales",
        )],
        [InlineKeyboardButton("Cancelar", callback_data="cancel")],
    ])


def months_keyboard() -> InlineKeyboardMarkup:
    rows = []
    for start in range(0, 12, 3):
        rows.append([
            InlineKeyboardButton(MONTHS[index], callback_data=f"month:{index + 1}")
            for index in range(start, start + 3)
        ])
    rows.append([InlineKeyboardButton("Cancelar", callback_data="cancel")])
    return InlineKeyboardMarkup(rows)


def years_keyboard(base_year: int | None = None) -> InlineKeyboardMarkup:
    current = base_year or datetime.now().year
    years = [current, current + 1, current + 2]
    return InlineKeyboardMarkup([
        [InlineKeyboardButton(str(year), callback_data=f"year:{year}") for year in years],
        [InlineKeyboardButton("Cancelar", callback_data="cancel")],
    ])
