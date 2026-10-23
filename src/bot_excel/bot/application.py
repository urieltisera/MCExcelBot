from __future__ import annotations

from pathlib import Path

from telegram.ext import (
    Application,
    ApplicationBuilder,
    CallbackQueryHandler,
    CommandHandler,
    ConversationHandler,
    MessageHandler,
    filters,
)

from bot_excel.bot.digitization_handlers import (
    handle_digitization_action,
    receive_sales_gc_photo,
    reject_unsupported_document,
)
from bot_excel.bot.handlers import (
    cancel,
    choose_month,
    choose_template,
    choose_year,
    create_start,
    error_handler,
    help_command,
    receive_ip_target,
    start,
)
from bot_excel.bot.states import CreateState
from bot_excel.config import Settings


def build_application(settings: Settings) -> Application:
    application = (
        ApplicationBuilder()
        .token(settings.telegram_bot_token)
        .concurrent_updates(False)
        .build()
    )
    application.bot_data["output_dir"] = str(settings.output_dir)
    temp_dir = settings.output_dir.parent / "runtime" / "temp"
    temp_dir.mkdir(parents=True, exist_ok=True)
    application.bot_data.update({
        "temp_dir": str(temp_dir),
        "gemini_api_key": settings.gemini_api_key,
        "gemini_model": settings.gemini_model,
        "max_photo_bytes": settings.max_photo_bytes,
        "digitization_ttl_seconds": settings.digitization_ttl_seconds,
        "max_real_sales": settings.max_real_sales,
        "max_real_gc": settings.max_real_gc,
    })

    create_conversation = ConversationHandler(
        entry_points=[CommandHandler("crear", create_start)],
        states={
            CreateState.TEMPLATE: [
                CallbackQueryHandler(choose_template, pattern=r"^(template:|disabled$|cancel$)")
            ],
            CreateState.MONTH: [
                CallbackQueryHandler(choose_month, pattern=r"^(month:|cancel$)")
            ],
            CreateState.YEAR: [
                CallbackQueryHandler(choose_year, pattern=r"^(year:|cancel$)")
            ],
            CreateState.IP_TARGET: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_ip_target)
            ],
        },
        fallbacks=[CommandHandler("cancelar", cancel)],
        allow_reentry=True,
    )

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("ayuda", help_command))
    application.add_handler(create_conversation)
    allowed_documents = (
        filters.Document.MimeType("image/jpeg")
        | filters.Document.MimeType("image/png")
        | filters.Document.MimeType("image/webp")
    )
    application.add_handler(
        MessageHandler(
            filters.PHOTO | allowed_documents,
            receive_sales_gc_photo,
        )
    )
    application.add_handler(
        CallbackQueryHandler(
            handle_digitization_action,
            pattern=r"^digitize:(confirm|cancel|review):[a-f0-9]{12}$",
        )
    )
    application.add_handler(
        MessageHandler(filters.Document.ALL, reject_unsupported_document)
    )
    application.add_handler(CommandHandler("cancelar", cancel))
    application.add_error_handler(error_handler)

    return application
