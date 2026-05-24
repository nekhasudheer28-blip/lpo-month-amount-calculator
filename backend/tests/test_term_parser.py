import unittest
from datetime import datetime

from services.term_parser import parse_term


class TermParserTests(unittest.TestCase):
    def labels(self, term):
        result = parse_term(term, datetime(2026, 2, 9))
        return result.month_labels if result else None

    def test_ex_stock_uses_lpo_month(self):
        self.assertEqual(self.labels("EX-STOCK"), ["Feb-2026"])

    def test_days_are_added_to_lpo_date(self):
        self.assertEqual(self.labels("120 Days PDC"), ["Jun-2026"])

    def test_weeks_are_added_to_lpo_date(self):
        self.assertEqual(self.labels("2 weeks"), ["Feb-2026"])

    def test_month_ranges_return_multiple_months(self):
        self.assertEqual(self.labels("Material by July/Aug-2026"), ["Jul-2026", "Aug-2026"])

    def test_unclear_term_returns_none(self):
        self.assertIsNone(parse_term("After approval", datetime(2026, 2, 9)))


if __name__ == "__main__":
    unittest.main()