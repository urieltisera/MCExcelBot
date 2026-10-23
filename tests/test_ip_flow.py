import unittest
from types import SimpleNamespace
from unittest.mock import AsyncMock

from bot_excel.bot.handlers import choose_year
from bot_excel.bot.states import CreateState


class IpConversationFlowTests(unittest.IsolatedAsyncioTestCase):
    async def test_ip_year_selection_requests_target(self):
        query = SimpleNamespace(
            data="year:2026",
            answer=AsyncMock(),
            edit_message_text=AsyncMock(),
        )
        update = SimpleNamespace(callback_query=query)
        context = SimpleNamespace(user_data={
            "month": 7,
            "template_key": "ip_semanal_mensual",
        })

        state = await choose_year(update, context)

        self.assertEqual(state, CreateState.IP_TARGET)
        self.assertEqual(context.user_data["year"], 2026)
        query.edit_message_text.assert_awaited_once()
        self.assertIn(
            "objetivo IP",
            query.edit_message_text.await_args.args[0],
        )


if __name__ == "__main__":
    unittest.main()
