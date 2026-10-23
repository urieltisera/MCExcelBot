import unittest

from bot_excel.bot.parsing import (
    parse_callback,
    parse_ip_target,
    parse_month,
    parse_year,
)


class BotParsingTests(unittest.TestCase):
    def test_parse_callback(self):
        self.assertEqual(parse_callback("template:ventas_gc", "template"), "ventas_gc")

    def test_parse_month(self):
        self.assertEqual(parse_month("month:9"), 9)

    def test_invalid_month(self):
        with self.assertRaises(ValueError):
            parse_month("month:13")

    def test_parse_year(self):
        self.assertEqual(parse_year("year:2026"), 2026)

    def test_invalid_prefix(self):
        with self.assertRaises(ValueError):
            parse_callback("year:2026", "month")

    def test_parse_ip_target_with_comma_or_dot(self):
        self.assertEqual(parse_ip_target("6,50"), 6.5)
        self.assertEqual(parse_ip_target("7.25"), 7.25)

    def test_invalid_ip_target(self):
        for value in ("", "abc", "0", "-1", "nan", "101"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                parse_ip_target(value)
