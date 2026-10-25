import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook
import procurement as app


def section(number, heading, *lines):
    return dict(section=number, heading=heading, page=1, lines=list(lines))


class ExtractionTests(unittest.TestCase):
    def test_scope_and_financial_semantics(self):
        sections = [section('1.1.', 'Beschaffer', 'Offizielle Bezeichnung: Buyer'),
                    section('8.1.', 'ORG-1', 'Offizielle Bezeichnung: Supplier'),
                    section('5.1.3.', 'Geschätzte Dauer', 'Laufzeit: 24 Monate'),
                    section('5.1.5.', 'Wert', 'Geschätzter Wert ohne MwSt.: 1 000,00 EUR'),
                    section('5.1.10.', 'Zuschlagskriterien', 'Art: Preis', 'Zuschlagskriterium — Zahl: 100')]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'source.pdf'
            path.write_bytes(b'test')
            with patch.object(app, 'read_sections', return_value=sections):
                result = app.extract(path)
        self.assertEqual(result['fields']['Auftraggeber'], 'Buyer')
        self.assertIsNone(result['fields']['Erwarteter Umsatz'])
        self.assertEqual(len(result['evidence']['financial_candidates']), 1)
        self.assertEqual(result['evidence']['financial_candidates'][0]['value'], '1 000,00 EUR')

    def test_multiple_lots_not_silently_selected(self):
        sections = [section('5.1.3.', 'Geschätzte Dauer', 'Laufzeit: 12 Monate')] * 2
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'a.pdf'
            path.write_bytes(b'test')
            with patch.object(app, 'read_sections', return_value=sections):
                result = app.extract(path)
        self.assertIsNone(result['fields']['Laufzeit'])
        self.assertTrue(any('Multiple duration' in x for x in result['review_issues']))

    def test_mapping_unique_conflicting_missing_and_blank(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'mapping.xlsx'
            book = Workbook()
            sheet = book.active
            sheet.append(['Auftraggeber', 'KAM'])
            sheet.append(['Buyer', 'Alex'])
            sheet.append([' BUYER ', 'Alex'])
            book.save(path)
            self.assertEqual(app.lookup_kam(path, 'buyer'), ('Alex', None))
            self.assertIsNone(app.lookup_kam(path, 'Other')[0])
            sheet.append(['Buyer', 'Sam'])
            book.save(path)
            self.assertIsNone(app.lookup_kam(path, 'Buyer')[0])
            sheet.delete_rows(4)
            sheet.append(['Buyer', None])
            book.save(path)
            self.assertIsNone(app.lookup_kam(path, 'Buyer')[0])


if __name__ == '__main__':
    unittest.main()
