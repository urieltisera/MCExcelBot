import unittest

from bot_excel.bot.keyboards import templates_keyboard


class TemplateKeyboardTests(unittest.TestCase):
    def test_productivity_template_is_enabled(self):
        callbacks = [
            button.callback_data
            for row in templates_keyboard().inline_keyboard
            for button in row
        ]
        self.assertIn("template:ventas_gc", callbacks)
        self.assertIn("template:productividad_controlables", callbacks)
        self.assertIn("template:ip_semanal_mensual", callbacks)
        self.assertIn("template:ventas_consolidadas_anuales", callbacks)
        self.assertEqual(callbacks.count("disabled"), 0)


if __name__ == "__main__":
    unittest.main()
