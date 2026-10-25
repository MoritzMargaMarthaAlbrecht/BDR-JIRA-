import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook
import procurement as app


def section(number, heading, *lines):
    return dict(section=number, heading=heading, page=1, lines=list(lines))


class ExtractionTests(unittest.TestCase):
    def extract_sections(self, sections, basis):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "source.pdf"
            path.write_bytes(b"test")
            with patch.object(app, "read_sections", return_value=sections):
                return app.extract(path, financial_basis=basis)

    def test_scope_and_financial_semantics(self):
        sections = [
            section("1.1.", "Beschaffer", "Offizielle Bezeichnung: Buyer"),
            section("8.1.", "ORG-1", "Offizielle Bezeichnung: Supplier"),
            section("5.1.3.", "Geschätzte Dauer", "Laufzeit: 24 Monate"),
            section("5.1.5.", "Wert", "Geschätzter Wert ohne MwSt.: 1 000,00 EUR"),
            section(
                "5.1.10.",
                "Zuschlagskriterien",
                "Art: Preis",
                "Zuschlagskriterium — Zahl: 100",
            ),
        ]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "source.pdf"
            path.write_bytes(b"test")
            with patch.object(app, "read_sections", return_value=sections):
                result = app.extract(path)
        self.assertEqual(result["fields"]["Auftraggeber"], "Buyer")
        self.assertIsNone(result["fields"]["Erwarteter Umsatz"])
        self.assertEqual(len(result["evidence"]["financial_candidates"]), 1)
        self.assertEqual(
            result["evidence"]["financial_candidates"][0]["value"], "1 000,00 EUR"
        )

    def test_multiple_lots_not_silently_selected(self):
        sections = [section("5.1.3.", "Geschätzte Dauer", "Laufzeit: 12 Monate")] * 2
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "a.pdf"
            path.write_bytes(b"test")
            with patch.object(app, "read_sections", return_value=sections):
                result = app.extract(path)
        self.assertIsNone(result["fields"]["Laufzeit"])
        self.assertTrue(any("Multiple duration" in x for x in result["review_issues"]))

    def test_mapping_unique_conflicting_missing_and_blank(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "mapping.xlsx"
            book = Workbook()
            sheet = book.active
            sheet.append(["Auftraggeber", "KAM"])
            sheet.append(["Buyer", "Alex"])
            sheet.append([" BUYER ", "Alex"])
            book.save(path)
            self.assertEqual(app.lookup_kam(path, "buyer"), ("Alex", None))
            self.assertIsNone(app.lookup_kam(path, "Other")[0])
            sheet.append(["Buyer", "Sam"])
            book.save(path)
            self.assertIsNone(app.lookup_kam(path, "Buyer")[0])
            sheet.delete_rows(4)
            sheet.append(["Buyer", None])
            book.save(path)
            self.assertIsNone(app.lookup_kam(path, "Buyer")[0])

    def test_explicit_reference_basis_excludes_price_weightings(self):
        sections = [
            section("5.1.", "Los: LOT-0001"),
            section("5.1.5.", "Wert", "Geschätzter Wert ohne MwSt.: 1 000,00 EUR"),
            section(
                "5.1.10.",
                "Zuschlagskriterien",
                "Art: Preis",
                "Zuschlagskriterium — Zahl: 100",
            ),
            section(
                "6.",
                "Ergebnisse",
                "Wert aller in dieser Bekanntmachung vergebenen Verträge: 900,00 EUR",
            ),
            section("6.1.2.", "Informationen", "Wert des Angebots: 800,00 EUR"),
        ]
        for basis, expected in [
            ("estimated", "1 000,00 EUR"),
            ("awarded", "900,00 EUR"),
        ]:
            with self.subTest(basis=basis):
                result = self.extract_sections(sections, basis)
                self.assertEqual(result["fields"]["Erwarteter Umsatz"], expected)
                self.assertEqual(result["financial_selection"]["basis"], basis)
                self.assertEqual(len(result["financial_selection"]["sources"]), 1)
                self.assertTrue(
                    any(
                        "Draft reference value only" in x
                        for x in result["review_issues"]
                    )
                )

    def test_conflicting_estimates_and_multiple_lots_need_review(self):
        estimates = [
            section("2.1.3.", "Wert", "Geschätzter Wert ohne MwSt.: 1 000,00 EUR"),
            section("5.1.5.", "Wert", "Geschätzter Wert ohne MwSt.: 2 000,00 EUR"),
        ]
        result = self.extract_sections(estimates, "estimated")
        self.assertIsNone(result["fields"]["Erwarteter Umsatz"])
        self.assertTrue(
            any("Conflicting estimated" in x for x in result["review_issues"])
        )
        for kind in ["Los", "Teil"]:
            result = self.extract_sections(
                [
                    estimates[0],
                    section("5.1.", f"{kind}: A"),
                    section("5.1.", f"{kind}: B"),
                ],
                "estimated",
            )
            self.assertIsNone(result["fields"]["Erwarteter Umsatz"])
            self.assertTrue(
                any("Multiple lots or parts" in x for x in result["review_issues"])
            )

    def test_identical_ceiling_values_are_one_reference(self):
        sections = [
            section(
                "2.1.3.", "Wert", "Höchstwert der Rahmenvereinbarung: 1 000,00 EUR"
            ),
            section("5.1.5.", "Wert", "Höchstwert der Rahmenvereinbarung: 1000,00 EUR"),
        ]
        result = self.extract_sections(sections, "ceiling")
        self.assertEqual(result["fields"]["Erwarteter Umsatz"], "1 000,00 EUR")
        self.assertEqual(len(result["financial_selection"]["sources"]), 2)

    def test_missing_nonpublished_and_offer_only_values_not_selected(self):
        for lines in [
            [],
            [
                "Wert aller in dieser Bekanntmachung vergebenen Verträge: Nicht veröffentlicht"
            ],
            ["Wert des Angebots: 900,00 EUR"],
            ["Wert aller in dieser Bekanntmachung vergebenen Verträge: 100 %"],
        ]:
            with self.subTest(lines=lines):
                result = self.extract_sections(
                    [section("6.", "Ergebnisse", *lines)], "awarded"
                )
                self.assertIsNone(result["fields"]["Erwarteter Umsatz"])
                self.assertEqual(result["financial_selection"]["sources"], [])


if __name__ == "__main__":
    unittest.main()
