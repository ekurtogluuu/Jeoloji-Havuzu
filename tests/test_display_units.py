import json
import unittest
from pathlib import Path

import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from display_units import find_unit_by_name, format_unit_for_display


class DisplayUnitsTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.units = json.loads((ROOT / "output" / "geology_units.json").read_text(encoding="utf-8"))

    def test_dataset_record_count_is_stable(self):
        self.assertEqual(len(self.units), 43)

    def test_find_unit_by_name_handles_turkish_names(self):
        unit = find_unit_by_name("Kapaklı Formasyonu", units=self.units)
        self.assertIsNotNone(unit)
        self.assertEqual(unit["code"], "PTRk")

    def test_repeated_contact_and_thickness_are_merged(self):
        unit = find_unit_by_name("Yayalar Formasyonu", units=self.units)
        output = format_unit_for_display(unit, mode="short")
        self.assertIn("## Dokanak ve Kalınlık", output)
        self.assertNotIn("## Dokanak İlişkileri", output)
        self.assertNotIn("## Kalınlık", output)

    def test_missing_fields_are_quality_notes_not_empty_sections(self):
        unit = find_unit_by_name("Sultanbeyli Formasyonu", units=self.units)
        output = format_unit_for_display(unit, mode="short")
        self.assertNotIn("## Dokanak İlişkileri\n##", output)
        self.assertNotIn("## Fosil Kapsamı\n##", output)
        self.assertIn("Eksik / Kalite Notları", output)
        self.assertIn("fosil kapsamı", output)

    def test_display_text_is_cleaned(self):
        unit = find_unit_by_name("Yayalar Formasyonu", units=self.units)
        output = format_unit_for_display(unit, mode="short")
        self.assertNotIn("şşeyl", output)
        self.assertNotIn("İİstanbul", output)
        self.assertNotIn("�", output)


if __name__ == "__main__":
    unittest.main()
