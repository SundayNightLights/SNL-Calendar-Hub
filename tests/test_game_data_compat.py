"""Backward-compatible SNL Excel column and diamond-label tests."""

import pathlib
import sys
import unittest

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1] / "scripts"))

from generate_calendars import normalize_game_columns
from generate_website_data import date_times


class WorkbookCompatibilityTests(unittest.TestCase):
    def setUp(self):
        self.old_headers = {
            "A": "Game ID", "B": "Week", "C": "Away Team",
            "D": "Home Team", "E": "Date", "F": "Time",
            "G": "Slot", "H": "Diamond", "I": "Away Score",
            "J": "Home Score", "K": "Away Points", "L": "Home Points",
            "M": "Game Status",
        }
        self.reordered_headers = {
            "A": "Game ID", "B": "Week", "C": "Date",
            "D": "Time", "E": "Slot", "F": "Diamond",
            "G": "Away Team", "H": "Home Team", "I": "Game Status",
            "J": "Away Score", "K": "Home Score",
            "L": "Away Points", "M": "Home Points",
        }

    def test_previous_workbook_layout_unchanged(self):
        data = {4: self.old_headers, 5: {
            "A": "F26-001", "C": "Buffaloes", "D": "IHT",
            "E": "46327", "F": "10:00–10:45 PM", "G": "1",
            "H": "1 (Near)", "I": "3", "J": "2",
            "K": "2", "L": "0", "M": "Final",
        }}
        self.assertIs(normalize_game_columns(data, "Regular Season Data"), data)

    def test_new_workbook_layout_is_normalized(self):
        data = {4: self.reordered_headers, 5: {
            "A": "F26-001", "B": "1", "C": "46327",
            "D": "10:00–10:45 PM", "E": "1",
            "F": "Diamond 1 • Near", "G": "Buffaloes", "H": "IHT",
            "I": "Final", "J": "3", "K": "2",
            "L": "2", "M": "0", "P": "Buffaloes_1",
        }}
        fixed = normalize_game_columns(data, "Regular Season Data")
        self.assertEqual(fixed[4]["C"], "Away Team")
        self.assertEqual(fixed[4]["I"], "Away Score")
        self.assertEqual(fixed[5]["C"], "Buffaloes")
        self.assertEqual(fixed[5]["D"], "IHT")
        self.assertEqual(fixed[5]["E"], "46327")
        self.assertEqual(fixed[5]["H"], "Diamond 1 • Near")
        self.assertEqual(fixed[5]["I"], "3")
        self.assertEqual(fixed[5]["J"], "2")
        self.assertEqual(fixed[5]["M"], "Final")
        self.assertEqual(fixed[5]["P"], "Buffaloes_1")

    def test_playoff_data_uses_same_column_adapter(self):
        rows = {4: self.reordered_headers, 5: {
            "A": "F26-P01", "G": "Seed 6", "H": "Seed 3",
            "I": "Unplayed", "J": "", "K": "",
        }}
        normalized = normalize_game_columns(rows, "Playoff Data")
        self.assertEqual(normalized[5]["C"], "Seed 6")
        self.assertEqual(normalized[5]["D"], "Seed 3")
        self.assertEqual(normalized[5]["M"], "Unplayed")

    def test_diamond_labels_old_and_new(self):
        for label, expected in (
            ("1 (Near)", 1), ("2 (Far)", 2),
            ("Diamond 1 • Near", 1), ("Diamond 2 • Far", 2),
        ):
            with self.subTest(label=label):
                row = {
                    "E": "46327", "F": "10:00–10:45 PM",
                    "G": "1", "H": label,
                }
                self.assertEqual(date_times(row, "F26-001")["diamond"], expected)

    def test_unrelated_sheets_not_modified(self):
        rows = {4: {"A": "Team"}}
        self.assertIs(normalize_game_columns(rows, "Playoff Seeding"), rows)


if __name__ == "__main__":
    unittest.main()
