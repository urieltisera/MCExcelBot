import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import AsyncMock

from telegram.ext import ConversationHandler

from bot_excel.bot.handlers import (
    _generate_and_send,
    choose_template,
)
from bot_excel.bot.states import CreateState


class AnnualConversationFlowTests(unittest.IsolatedAsyncioTestCase):
    async def test_annual_template_requests_year_without_month(self):
        query = SimpleNamespace(
            data="template:ventas_consolidadas_anuales",
            answer=AsyncMock(),
            edit_message_text=AsyncMock(),
        )
        update = SimpleNamespace(callback_query=query)
        context = SimpleNamespace(user_data={})

        state = await choose_template(update, context)

        self.assertEqual(state, CreateState.YEAR)
        self.assertEqual(
            context.user_data["template_key"],
            "ventas_consolidadas_anuales",
        )
        self.assertNotIn("month", context.user_data)
        query.edit_message_text.assert_awaited_once()
        self.assertIn(
            "año",
            query.edit_message_text.await_args.args[0].lower(),
        )

    async def test_bot_generates_and_sends_annual_workbook(self):
        with tempfile.TemporaryDirectory() as tmp:
            message = SimpleNamespace(
                reply_document=AsyncMock(),
                reply_text=AsyncMock(),
            )
            context = SimpleNamespace(
                user_data={
                    "template_key": "ventas_consolidadas_anuales",
                    "year": 2026,
                },
                application=SimpleNamespace(
                    bot_data={"output_dir": str(Path(tmp))}
                ),
            )

            state = await _generate_and_send(message, context)

            self.assertEqual(state, ConversationHandler.END)
            message.reply_text.assert_not_awaited()
            message.reply_document.assert_awaited_once()
            call = message.reply_document.await_args.kwargs
            self.assertTrue(
                call["filename"].startswith(
                    "Ventas_Consolidadas_Anuales_2026_VA-2026-"
                )
            )
            self.assertIn("Período: 2026", call["caption"])
            self.assertEqual(context.user_data, {})
            self.assertTrue((Path(tmp) / call["filename"]).exists())


if __name__ == "__main__":
    unittest.main()
