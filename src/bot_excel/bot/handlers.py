from __future__ import annotations

import logging
from pathlib import Path

from telegram import Update
from telegram.ext import ContextTypes, ConversationHandler

from bot_excel.bot.digitization_handlers import cleanup_pending_digitization
from bot_excel.bot.keyboards import MONTHS, months_keyboard, templates_keyboard, years_keyboard
from bot_excel.bot.parsing import (
    parse_callback,
    parse_ip_target,
    parse_month,
    parse_year,
)
from bot_excel.bot.states import CreateState
from bot_excel.core.models import AnnualSalesRequest, MonthlyTemplateRequest
from bot_excel.templates.registry import TemplateRegistry


LOGGER = logging.getLogger(__name__)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    message = (
        "Hola. Soy el Bot de Excel de Telegram.\n\n"
        "Puedo generar las cuatro planillas y digitalizar fotografías "
        "de Ventas y GC.\n\n"
        "Usá /crear para generar una planilla o enviá una fotografía "
        "completa con el QR visible."
    )
    await update.effective_message.reply_text(message)


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.effective_message.reply_text(
        "/start — presentación\n"
        "/crear — generar una nueva planilla\n"
        "/cancelar — cancelar la operación actual\n"
        "/ayuda — mostrar esta ayuda\n\n"
        "Digitalizar:\n"
        "Enviá una fotografía completa de una planilla de Ventas y GC "
        "que contenga el QR."
    )


async def create_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    cleanup_pending_digitization(context)
    context.user_data.clear()
    await update.effective_message.reply_text(
        "Elegí el tipo de planilla:",
        reply_markup=templates_keyboard(),
    )
    return CreateState.TEMPLATE


async def choose_template(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    if query.data == "cancel":
        await query.edit_message_text("Operación cancelada.")
        return ConversationHandler.END

    if query.data == "disabled":
        await query.answer("Este formato todavía no está implementado.", show_alert=True)
        return CreateState.TEMPLATE

    template_key = parse_callback(query.data, "template")
    context.user_data["template_key"] = template_key

    if template_key == "ventas_consolidadas_anuales":
        await query.edit_message_text(
            "Elegí el año:",
            reply_markup=years_keyboard(),
        )
        return CreateState.YEAR

    await query.edit_message_text(
        "Elegí el mes:",
        reply_markup=months_keyboard(),
    )
    return CreateState.MONTH


async def choose_month(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    if query.data == "cancel":
        await query.edit_message_text("Operación cancelada.")
        return ConversationHandler.END

    month = parse_month(query.data)
    context.user_data["month"] = month

    await query.edit_message_text(
        f"Mes seleccionado: {MONTHS[month - 1]}.\nElegí el año:",
        reply_markup=years_keyboard(),
    )
    return CreateState.YEAR


async def choose_year(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()

    if query.data == "cancel":
        await query.edit_message_text("Operación cancelada.")
        return ConversationHandler.END

    year = parse_year(query.data)
    template_key = str(context.user_data["template_key"])
    context.user_data["year"] = year

    if template_key == "ventas_consolidadas_anuales":
        await query.edit_message_text(
            f"Generando consolidado anual {year}…"
        )
        return await _generate_and_send(query.message, context)

    month = int(context.user_data["month"])
    if template_key == "ip_semanal_mensual":
        await query.edit_message_text(
            f"Período: {MONTHS[month - 1]} {year}.\n"
            "Ingresá el objetivo IP, por ejemplo 6,50.\n"
            "Podés usar coma o punto decimal."
        )
        return CreateState.IP_TARGET

    await query.edit_message_text(f"Generando {MONTHS[month - 1]} {year}…")
    return await _generate_and_send(query.message, context)


async def receive_ip_target(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> int:
    try:
        ip_target = parse_ip_target(update.effective_message.text or "")
    except ValueError as exc:
        await update.effective_message.reply_text(
            f"{exc}\nIntentá nuevamente o usá /cancelar."
        )
        return CreateState.IP_TARGET

    context.user_data["ip_target"] = ip_target
    month = int(context.user_data["month"])
    year = int(context.user_data["year"])
    await update.effective_message.reply_text(
        f"Generando {MONTHS[month - 1]} {year} "
        f"con objetivo IP {ip_target:.2f}…"
    )
    return await _generate_and_send(update.effective_message, context)


async def _generate_and_send(message, context: ContextTypes.DEFAULT_TYPE) -> int:
    year = int(context.user_data["year"])
    template_key = str(context.user_data["template_key"])
    try:
        output_dir = Path(context.application.bot_data["output_dir"])
        generator = TemplateRegistry().get(template_key)
        if template_key == "ventas_consolidadas_anuales":
            result = generator.generate(AnnualSalesRequest(
                year=year,
                output_dir=output_dir,
            ))
            period_caption = str(year)
        else:
            month = int(context.user_data["month"])
            result = generator.generate(MonthlyTemplateRequest(
                year=year,
                month=month,
                output_dir=output_dir,
                ip_target=float(context.user_data.get("ip_target", 6.5)),
            ))
            period_caption = f"{MONTHS[month - 1]} {year}"

        with result.workbook_path.open("rb") as document:
            caption = (
                f"Planilla creada correctamente.\n"
                f"Código: {result.template_id}\n"
                f"Período: {period_caption}"
            )
            if template_key == "ip_semanal_mensual":
                target = float(context.user_data["ip_target"])
                caption += f"\nObjetivo IP: {target:.2f}".replace(".", ",")
            await message.reply_document(
                document=document,
                filename=result.workbook_path.name,
                caption=caption,
            )
    except Exception:
        LOGGER.exception("No se pudo generar la planilla.")
        await message.reply_text(
            "No pude generar la planilla. El error quedó registrado."
        )
    finally:
        context.user_data.clear()

    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    cleanup_pending_digitization(context)
    context.user_data.clear()
    await update.effective_message.reply_text("Operación cancelada.")
    return ConversationHandler.END


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    LOGGER.exception("Error no controlado procesando una actualización.", exc_info=context.error)
