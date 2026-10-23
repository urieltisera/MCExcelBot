from __future__ import annotations

import asyncio
import logging
import time
import uuid
from pathlib import Path
from typing import Any

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import ContextTypes

from bot_excel.core.digitization_models import SalesGcPhotoReading
from bot_excel.services.gemini_reader import (
    GeminiReaderError,
    GeminiSalesGcReader,
)
from bot_excel.services.generated_file_registry import (
    GeneratedFileRegistry,
    GeneratedFileRegistryError,
)
from bot_excel.services.qr_reader import (
    InvalidQrPayloadError,
    QrNotDetectedError,
    UnsupportedQrTypeError,
    decode_sales_gc_qr,
)
from bot_excel.services.sales_gc_digitizer import (
    SalesGcDigitizer,
    SalesGcDigitizerError,
)


LOGGER = logging.getLogger(__name__)
PENDING_KEY = "sales_gc_digitization"
ALLOWED_MIME_TYPES = {
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
}
MONTHS_ES = (
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
)


def cleanup_pending_digitization(context: ContextTypes.DEFAULT_TYPE) -> None:
    pending = context.user_data.pop(PENDING_KEY, None)
    if not isinstance(pending, dict):
        return
    temp_path = pending.get("temp_path")
    if temp_path:
        _safe_unlink_temp(
            Path(temp_path),
            Path(context.application.bot_data["temp_dir"]),
        )


async def receive_sales_gc_photo(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    message = update.effective_message
    cleanup_pending_digitization(context)
    operation_id = uuid.uuid4().hex[:12]
    phase = "recepción"
    template_code: str | None = None
    temp_path: Path | None = None
    awaiting_confirmation = False

    try:
        telegram_file, mime_type, suffix, file_size = await _select_image(message)
        max_bytes = int(context.application.bot_data["max_photo_bytes"])
        if file_size is not None and file_size > max_bytes:
            await message.reply_text(
                "La imagen es demasiado grande. Enviá una foto de hasta "
                f"{max_bytes // (1024 * 1024)} MB."
            )
            return

        await message.reply_text(
            "Recibí la fotografía. Estoy identificando la planilla…"
        )
        temp_dir = Path(context.application.bot_data["temp_dir"])
        temp_dir.mkdir(parents=True, exist_ok=True)
        temp_path = temp_dir / f"{uuid.uuid4().hex}{suffix}"
        await telegram_file.download_to_drive(custom_path=temp_path)
        if temp_path.stat().st_size > max_bytes:
            raise ValueError("La imagen descargada supera el tamaño permitido.")

        phase = "qr"
        parsed_qr = await asyncio.to_thread(decode_sales_gc_qr, temp_path)
        template_code = parsed_qr.code

        phase = "registro"
        output_dir = Path(context.application.bot_data["output_dir"])
        original_path = await asyncio.to_thread(
            GeneratedFileRegistry(output_dir).find_sales_gc,
            parsed_qr.code,
        )

        phase = "gemini"
        reader = context.application.bot_data.get("gemini_reader")
        if reader is None:
            reader = GeminiSalesGcReader(
                api_key=context.application.bot_data.get("gemini_api_key"),
                model=str(context.application.bot_data["gemini_model"]),
            )
        reading = await asyncio.to_thread(
            reader.read_image,
            temp_path,
            mime_type=mime_type,
        )

        phase = "validación"
        reading.validate_period(
            expected_code=parsed_qr.code,
            expected_year=parsed_qr.year,
            expected_month=parsed_qr.month,
        )
        validation = _validate_reading(
            reading,
            max_real_sales=int(
                context.application.bot_data["max_real_sales"]
            ),
            max_real_gc=int(context.application.bot_data["max_real_gc"]),
        )
        if reading.image_quality == "poor" or validation["illegible"]:
            await message.reply_text(
                "No pude leer la fotografía con suficiente seguridad. "
                "Probá nuevamente con una foto completa, nítida y bien iluminada."
            )
            return

        expires_at = (
            time.time()
            + int(context.application.bot_data["digitization_ttl_seconds"])
        )
        context.user_data[PENDING_KEY] = {
            "operation_id": operation_id,
            "temp_path": str(temp_path),
            "template_code": parsed_qr.code,
            "year": parsed_qr.year,
            "month": parsed_qr.month,
            "reading": reading,
            "original_path": str(original_path),
            "expires_at": expires_at,
        }
        temp_path = None

        buttons = [[
            InlineKeyboardButton(
                "Confirmar",
                callback_data=f"digitize:confirm:{operation_id}",
            ),
            InlineKeyboardButton(
                "Cancelar",
                callback_data=f"digitize:cancel:{operation_id}",
            ),
        ]]
        if validation["doubtful"]:
            buttons.append([
                InlineKeyboardButton(
                    "Revisar dudosos",
                    callback_data=f"digitize:review:{operation_id}",
                )
            ])
        await message.reply_text(
            _build_summary(
                reading=reading,
                template_code=parsed_qr.code,
                year=parsed_qr.year,
                month=parsed_qr.month,
                doubtful=validation["doubtful"],
                warnings=validation["warnings"],
            ),
            reply_markup=InlineKeyboardMarkup(buttons),
        )
        awaiting_confirmation = True
        LOGGER.info(
            "digitization_ready correlation_id=%s code=%s fields=%s",
            operation_id,
            parsed_qr.code,
            reading.populated_field_count(),
        )
    except QrNotDetectedError:
        await message.reply_text(
            "No pude identificar el QR de la planilla.\n"
            "Probá nuevamente con una foto completa, nítida y bien iluminada."
        )
    except UnsupportedQrTypeError as exc:
        await message.reply_text(str(exc))
    except InvalidQrPayloadError:
        await message.reply_text(
            "El QR no pertenece a una planilla válida de McExcel."
        )
    except GeneratedFileRegistryError as exc:
        await message.reply_text(str(exc))
    except GeminiReaderError as exc:
        await message.reply_text(
            f"No pude completar la lectura. {exc}\n"
            "El archivo original no fue modificado."
        )
    except ValueError as exc:
        LOGGER.warning(
            "digitization_validation_failed correlation_id=%s phase=%s error=%s",
            operation_id,
            phase,
            type(exc).__name__,
        )
        await message.reply_text(
            f"No pude validar la fotografía. {exc}\n"
            "El archivo original no fue modificado."
        )
    except Exception as exc:
        LOGGER.error(
            "digitization_failed correlation_id=%s code=%s phase=%s error=%s",
            operation_id,
            template_code or "unknown",
            phase,
            type(exc).__name__,
        )
        await message.reply_text(
            "No pude completar la digitalización. "
            "El archivo original no fue modificado."
        )
    finally:
        pending = context.user_data.get(PENDING_KEY)
        if (
            not awaiting_confirmation
            and isinstance(pending, dict)
            and pending.get("operation_id") == operation_id
        ):
            cleanup_pending_digitization(context)
        if temp_path is not None:
            _safe_unlink_temp(
                temp_path,
                Path(context.application.bot_data["temp_dir"]),
            )


async def reject_unsupported_document(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    await update.effective_message.reply_text(
        "En esta etapa solo acepto fotografías JPG, PNG o WEBP. "
        "No envíes PDF, ZIP ni otros archivos."
    )


async def handle_digitization_action(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
) -> None:
    query = update.callback_query
    await query.answer()
    try:
        _, action, operation_id = (query.data or "").split(":", 2)
    except ValueError:
        await query.edit_message_text("Esta operación ya no es válida.")
        return

    pending = context.user_data.get(PENDING_KEY)
    if (
        not isinstance(pending, dict)
        or pending.get("operation_id") != operation_id
    ):
        await query.edit_message_text(
            "Esta operación venció o ya fue procesada."
        )
        return

    if float(pending["expires_at"]) < time.time():
        cleanup_pending_digitization(context)
        await query.edit_message_text(
            "La operación venció. Enviá nuevamente la fotografía."
        )
        return

    if action == "cancel":
        cleanup_pending_digitization(context)
        await query.edit_message_text(
            "Digitalización cancelada. El archivo original no fue modificado."
        )
        return
    if action == "review":
        cleanup_pending_digitization(context)
        await query.edit_message_text(
            "La primera versión no modifica campos dudosos. "
            "Enviá otra foto más nítida para revisarlos."
        )
        return
    if action != "confirm":
        await query.edit_message_text("Acción no válida.")
        return

    operation_id = str(pending["operation_id"])
    template_code = str(pending["template_code"])
    try:
        result = await asyncio.to_thread(
            SalesGcDigitizer().create_digitized_copy,
            original_path=Path(pending["original_path"]),
            reading=pending["reading"],
            year=int(pending["year"]),
            month=int(pending["month"]),
        )
        with result.workbook_path.open("rb") as document:
            await query.message.reply_document(
                document=document,
                filename=result.workbook_path.name,
                caption=(
                    "Planilla digitalizada correctamente.\n\n"
                    f"Código: {template_code}\n"
                    f"Valores cargados: {result.values_written}\n"
                    f"Campos omitidos: {result.fields_omitted}"
                ),
            )
        await query.edit_message_text(
            "Confirmación recibida. Te envié la copia digitalizada."
        )
        LOGGER.info(
            "digitization_completed correlation_id=%s code=%s fields=%s",
            operation_id,
            template_code,
            result.values_written,
        )
    except (SalesGcDigitizerError, OSError) as exc:
        LOGGER.error(
            "digitization_write_failed correlation_id=%s code=%s error=%s",
            operation_id,
            template_code,
            type(exc).__name__,
        )
        await query.edit_message_text(
            "No pude completar la digitalización. "
            "El archivo original no fue modificado."
        )
    finally:
        cleanup_pending_digitization(context)


async def _select_image(message) -> tuple[Any, str, str, int | None]:
    if message.photo:
        photo = message.photo[-1]
        telegram_file = await photo.get_file()
        return telegram_file, "image/jpeg", ".jpg", photo.file_size

    document = message.document
    if document is None or document.mime_type not in ALLOWED_MIME_TYPES:
        raise ValueError("El archivo no es una imagen permitida.")
    telegram_file = await document.get_file()
    return (
        telegram_file,
        document.mime_type,
        ALLOWED_MIME_TYPES[document.mime_type],
        document.file_size,
    )


def _validate_reading(
    reading: SalesGcPhotoReading,
    *,
    max_real_sales: int,
    max_real_gc: int,
) -> dict[str, Any]:
    if not reading.rows or reading.populated_field_count() == 0:
        raise ValueError("No se detectaron valores de Venta real o GC real.")

    doubtful: list[str] = []
    illegible: list[str] = []
    warnings = list(reading.warnings)
    for row in reading.rows:
        for label, value, confidence, maximum in (
            ("Venta real", row.real_sales, row.sales_confidence, max_real_sales),
            ("GC real", row.real_gc, row.gc_confidence, max_real_gc),
        ):
            if value is None:
                continue
            if value > maximum:
                warnings.append(
                    f"Día {row.day}: {label} fuera del rango habitual."
                )
            if confidence is None or confidence < 0.75:
                illegible.append(f"Día {row.day} — {label}")
            elif confidence < 0.90:
                doubtful.append(f"Día {row.day} — {label}: {_format_value(label, value)}")
    return {
        "doubtful": doubtful,
        "illegible": illegible,
        "warnings": warnings,
    }


def _build_summary(
    *,
    reading: SalesGcPhotoReading,
    template_code: str,
    year: int,
    month: int,
    doubtful: list[str],
    warnings: list[str],
) -> str:
    lines = [
        "Planilla detectada:",
        f"Ventas y GC — {MONTHS_ES[month - 1]} {year}",
        f"Código: {template_code}",
        "",
        "Valores leídos:",
    ]
    for row in sorted(reading.rows, key=lambda item: item.day):
        if row.real_sales is None and row.real_gc is None:
            continue
        lines.extend([
            "",
            f"Día {row.day}",
            f"Venta real: {_format_value('Venta real', row.real_sales)}",
            f"GC real: {_format_value('GC real', row.real_gc)}",
        ])
    if doubtful:
        lines.extend(["", "Campos dudosos:", *doubtful])
    if warnings:
        lines.extend(["", "Advertencias:", *warnings])
    lines.extend(["", "¿Confirmás los valores?"])
    return "\n".join(lines)


def _format_value(label: str, value: int | None) -> str:
    if value is None:
        return "omitido"
    rendered = f"{value:,}".replace(",", ".")
    return f"${rendered}" if label == "Venta real" else rendered


def _safe_unlink_temp(path: Path, temp_dir: Path) -> None:
    resolved_path = Path(path).resolve()
    resolved_temp = Path(temp_dir).resolve()
    try:
        resolved_path.relative_to(resolved_temp)
    except ValueError:
        LOGGER.error("temp_cleanup_rejected path_outside_temp")
        return
    resolved_path.unlink(missing_ok=True)
