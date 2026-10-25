"""Prepare reviewable Jira field drafts from German TED PDF notices."""

from __future__ import annotations
import argparse
import hashlib
import json
import re
import unicodedata
from decimal import Decimal
from pathlib import Path
from pypdf import PdfReader
from openpyxl import load_workbook

HEADING = re.compile(r"^(\d+(?:\.\d+)*\.)\s+(.+)$")
MONEY = re.compile(
    r"^(Geschätzter Wert ohne MwSt\.|Höchstwert der Rahmenvereinbarung|Wert aller in dieser Bekanntmachung vergebenen Verträge|Wert des Angebots):\s*(.+)$"
)
FINANCIAL_KINDS = {
    "estimated": "Geschätzter Wert ohne MwSt.",
    "awarded": "Wert aller in dieser Bekanntmachung vergebenen Verträge",
    "ceiling": "Höchstwert der Rahmenvereinbarung",
}
AMOUNT = re.compile(r"^((?:\d{1,3}(?:[ .]\d{3})+|\d+)(?:,\d{2})?)\s+([A-Z]{3})$")


def normalize(value):
    return " ".join(unicodedata.normalize("NFKC", str(value)).casefold().split())


def read_sections(path):
    sections, current = [], None
    reader = PdfReader(path)
    for page_no, page in enumerate(reader.pages, 1):
        for raw in (page.extract_text(extraction_mode="layout") or "").splitlines():
            line = raw.strip()
            if (
                not line
                or re.match(r"^\d+-\d{4}\s+Page \d+/\d+$", line)
                or line == "See the notice on TED website"
            ):
                continue
            match = HEADING.match(line)
            if match:
                current = {
                    "section": match[1],
                    "heading": match[2],
                    "page": page_no,
                    "lines": [],
                }
                sections.append(current)
            elif current:
                current["lines"].append(line)
    return sections


def evidence(section, value):
    return {
        "value": value,
        "section": section["section"],
        "heading": section["heading"],
        "page": section["page"],
    }


def lookup_kam(path, client):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = iter(workbook.active.values)
        headers = [normalize(x or "") for x in next(rows, ())]
        if headers.count("auftraggeber") != 1 or headers.count("kam") != 1:
            raise ValueError(
                "Excel needs unique Auftraggeber and KAM column headers in row 1."
            )
        ci, ki = headers.index("auftraggeber"), headers.index("kam")
        matches = [
            row[ki]
            for row in rows
            if row[ci] is not None and normalize(row[ci]) == normalize(client)
        ]
        names = {str(x).strip() for x in matches if x is not None and str(x).strip()}
        if not matches:
            return None, "No exact customer match in Excel."
        if any(x is None or not str(x).strip() for x in matches) or len(names) != 1:
            return None, "Blank or conflicting KAM assignments in Excel."
        return names.pop(), None
    finally:
        workbook.close()


def select_financial_reference(amounts, sections, basis):
    """Use only the explicitly requested category; never infer sales revenue."""
    selection = {"basis": basis, "value": None, "sources": []}
    if basis is None:
        return selection, None
    if basis not in FINANCIAL_KINDS:
        raise ValueError("Unknown financial basis.")
    lots = {
        m[1]
        for s in sections
        if (m := re.match(r"^(?:Los|Teil):\s*(\S+)", s["heading"]))
    }
    if len(lots) > 1:
        return (
            selection,
            "Multiple lots or parts; select the financial reference manually.",
        )
    values = {}
    for candidate in amounts:
        if candidate["kind"] != FINANCIAL_KINDS[basis]:
            continue
        text = unicodedata.normalize("NFKC", candidate["value"])
        match = AMOUNT.fullmatch(text)
        if match:
            number = Decimal(
                match[1].replace(" ", "").replace(".", "").replace(",", ".")
            )
            values.setdefault((number, match[2]), []).append(candidate)
    if not values:
        return (
            selection,
            f"No published numeric {basis} value with currency; check manually.",
        )
    if len(values) != 1:
        return (
            selection,
            f"Conflicting {basis} values; select the financial reference manually.",
        )
    sources = next(iter(values.values()))
    selection.update(value=sources[0]["value"], sources=sources)
    return (
        selection,
        "Draft reference value only; confirm expected revenue before Jira transfer.",
    )


def extract(path, mapping=None, financial_basis=None):
    sections = read_sections(path)
    issues, clients, amounts, durations, criteria, extensions = [], [], [], [], [], []
    for section in sections:
        lines = section["lines"]
        if section["section"] == "1.1." and section["heading"] == "Beschaffer":
            for i, line in enumerate(lines):
                if line.startswith("Offizielle Bezeichnung:"):
                    parts = [line.split(":", 1)[1].strip()]
                    for following in lines[i + 1 :]:
                        if ":" in following:
                            break
                        parts.append(following)
                    clients.append(evidence(section, " ".join(parts)))
        for line in lines:
            match = MONEY.match(line)
            if match:
                amounts.append({**evidence(section, match[2]), "kind": match[1]})
        if section["section"].startswith(("3.", "5.")) and section["heading"] in (
            "Dauer",
            "Geschätzte Dauer",
            "Laufzeit",
        ):
            durations.append(evidence(section, "\n".join(lines)))
        if section["heading"] == "Verlängerung":
            extensions.append(evidence(section, "\n".join(lines)))
        if section["heading"] == "Zuschlagskriterien":
            criteria.append(evidence(section, "\n".join(lines)))
    client = clients[0]["value"] if len(clients) == 1 else None
    if not client:
        issues.append("Missing or repeated buyer section; check Auftraggeber.")
    # A notice value is evidence, never an automatic sales revenue forecast.
    for key, values in [
        ("financial figure", amounts),
        ("duration", durations),
        ("award criteria", criteria),
    ]:
        if not values or any(not x["value"] for x in values):
            issues.append(f"Missing {key}.")
        elif key != "financial figure" and len(values) != 1:
            issues.append(f"Multiple {key} sections; select the intended lot manually.")
    if extensions:
        issues.append(
            "Contract extension options exist; review extension evidence separately from base duration."
        )
    selection, financial_issue = select_financial_reference(
        amounts, sections, financial_basis
    )
    if financial_issue:
        issues.append(financial_issue)
    elif amounts:
        issues.append(
            "Confirm financial basis manually: notice estimates, ceilings and awarded values are not expected revenue."
        )
    if any("Unbekannt" in x["value"] for x in durations):
        issues.append("Contract duration is explicitly unknown.")
    if any("Vergabeunterlagen" in x["value"] for x in criteria):
        issues.append("Award criteria refer to procurement documents; consult them.")
    kam = None
    if mapping and client:
        kam, problem = lookup_kam(mapping, client)
        if problem:
            issues.append(problem)
    elif not mapping:
        issues.append("No Excel KAM mapping supplied.")
    if not sections:
        issues.append(
            "No TED sections found; scanned PDFs require OCR or a different parser."
        )
    return {
        "source": str(path.resolve()),
        "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        "status": "manual_review_required",
        "fields": {
            "Auftraggeber": client,
            "Erwarteter Umsatz": selection["value"],
            "Zuschlagskriterien": criteria[0]["value"] if len(criteria) == 1 else None,
            "Laufzeit": durations[0]["value"] if len(durations) == 1 else None,
            "KAM": kam,
        },
        "evidence": {
            "clients": clients,
            "financial_candidates": amounts,
            "durations": durations,
            "extensions": extensions,
            "award_criteria": criteria,
        },
        "financial_selection": selection,
        "review_issues": issues,
        "reviewed": False,
    }


def fill_jira(report_path, config_path):
    """Explicit opt-in; only text inputs, never submit or select an assignee."""
    from playwright.sync_api import sync_playwright

    report = json.loads(report_path.read_text(encoding="utf-8"))
    config = json.loads(config_path.read_text(encoding="utf-8"))
    fields = report["fields"]
    if (
        report.get("reviewed") is not True
        or report.get("review_issues")
        or any(
            not str(fields.get(k) or "").strip()
            for k in (
                "Auftraggeber",
                "Erwarteter Umsatz",
                "Zuschlagskriterien",
                "Laufzeit",
                "KAM",
            )
        )
    ):
        raise ValueError(
            "Review all fields, resolve review_issues and set reviewed=true before filling Jira."
        )
    from urllib.parse import urlsplit

    if urlsplit(config["url"]).scheme != "https":
        raise ValueError("Jira URL must use HTTPS.")
    selectors = config["selectors"]
    if set(selectors) != set(fields):
        raise ValueError("Configure exactly one text-field selector per report field.")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        try:
            page = browser.new_page()
            page.goto(config["url"])
            input(
                "Sign in and open the create form, then press Enter here to fill it: "
            )
            locators = {
                key: page.locator(selector) for key, selector in selectors.items()
            }
            for locator in locators.values():
                if locator.count() != 1 or not locator.is_editable():
                    raise ValueError(
                        "Every selector must identify one editable text field."
                    )
            for key, locator in locators.items():
                locator.fill(str(fields[key]))
            input(
                "Review entries, confirm KAM, attach PDF and submit manually if appropriate. Enter closes browser: "
            )
        finally:
            browser.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    prepare = sub.add_parser("prepare")
    prepare.add_argument("pdfs", nargs="+", type=Path)
    prepare.add_argument("--mapping", type=Path)
    prepare.add_argument("--output", type=Path, default=Path("output"))
    prepare.add_argument(
        "--financial-basis",
        choices=tuple(FINANCIAL_KINDS),
        help="Explicit draft reference category; always requires revenue confirmation.",
    )
    fill = sub.add_parser("fill-jira")
    fill.add_argument("report", type=Path)
    fill.add_argument("--config", required=True, type=Path)
    args = parser.parse_args()
    if args.command == "fill-jira":
        fill_jira(args.report, args.config)
        return 0
    args.output.mkdir(parents=True, exist_ok=True)
    failed = False
    for path in args.pdfs:
        try:
            report = extract(path, args.mapping, args.financial_basis)
            target = args.output / f'{path.stem}-{report["sha256"][:12]}.json'
            # Never overwrite a colleague's reviewed report.
            with target.open("x", encoding="utf-8") as stream:
                json.dump(report, stream, ensure_ascii=False, indent=2)
                stream.write("\n")
            print(
                f'{path.name}: {target} ({len(report["review_issues"])} review issues)'
            )
            for field, value in report["fields"].items():
                print(
                    f'  {field}: {value if value is not None else "[not available / needs review]"}'
                )
            for issue in report["review_issues"]:
                print(f"  REVIEW: {issue}")
        except Exception as error:
            print(f"{path}: ERROR: {error}")
            failed = True
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
