# BDR-JIRA-
JIRA ticket automation

# Procurement notice → Jira draft

Python 3.10+ tool for German TED procurement notices with numbered sections.
Extracts Auftraggeber, duration and award criteria; lists labeled financial
candidates with section and page references. Matches Auftraggeber to KAM in Excel.
No AI service, API key or Jira login is needed for extraction.

## Install (VS Code terminal or any terminal)

```sh
python -m venv .venv
# macOS/Linux:
source .venv/bin/activate
# Windows PowerShell instead: .venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Prepare documents

```sh
python procurement.py prepare "notice.pdf" --mapping "kam_mapping.xlsx" --output output
# Multiple files are supported:
python procurement.py prepare "notice1.pdf" "notice2.pdf" --output output
python -m unittest discover -s tests -v
```

Excel's active worksheet needs row-one columns `Auftraggeber` and `KAM`.
Matches ignore case and whitespace; there is no fuzzy assignment. Duplicate rows
with the same KAM are allowed; conflicting or blank assignments require review.
Without a mapping, extraction still runs and flags the missing KAM.

Each PDF produces a UTF-8 JSON report containing five target fields, evidence,
review issues, source path and SHA-256. Existing reports are never overwritten;
use a new output directory to repeat a run. Exit 0 means reports were generated,
not that their contents were approved. Exit 1 means at least one file failed;
other files still process. Keep these reports and spreadsheets private.

Compare the report against the PDF. `Erwarteter Umsatz` intentionally starts
empty: a procurement estimate, framework ceiling or awarded contract value is
not automatically your expected revenue. Select and enter the business-approved
figure manually. Missing values, unknown durations, multiple lots and references
to external award documents require review. Base duration and extension options
are recorded separately. Page references identify the start of each section.
The parser supports text PDFs of the supplied TED format; scans need OCR first.
It does not interpret arbitrary PDF layouts, forecast revenue or aggregate lots.

## Optional Jira form filling

Document processing works without this extension. Live Jira integration has not
been tested. Use it only where your organisation permits browser automation.
A desktop/visible browser is required; headless cloud extraction still works.

```sh
python -m pip install playwright==1.56.0
python -m playwright install chromium
```

Copy `jira.example.json` to `jira.local.json`. Set the actual HTTPS Jira URL and
selectors for your form. Every selector must match a single editable text input
or textarea. Jira assignee/KAM autocomplete controls may need a site-specific
adapter; typing a name does not confirm an account selection.

After checking a generated report, enter/correct every field, resolve and remove
the entries in `review_issues`, then set `reviewed` to `true`. Keep source evidence
in the report for comparison.

```sh
python procurement.py fill-jira output/your-report.json --config jira.local.json
```

Sign in normally in the opened browser and navigate to the create form. Press
Enter in the terminal to fill the configured fields. Review all fields, confirm
the KAM, attach the PDF, and submit manually if appropriate. The script contains
no submit action and creates no Jira ticket via an API. Browser sessions are
not saved. Report text is inserted as data, never executed as instructions.

## Included validation

The five supplied notices were processed locally. The Fraport planning notice
has no financial figure/award criteria and an unknown duration. The gematik
notice has no structured financial figure/award criteria and includes extension
options. Munich contains different procedure and lot estimates and a framework
ceiling; these remain distinct candidates. LKA and Bruchsal contain awarded
values, which remain labeled as such. Automated tests cover buyer versus supplier
scope, financial versus weighting values, multiple durations, and Excel matching.
Actual KAM assignments and live Jira remain unverified without your mapping and
project-specific configuration. No documents, customer spreadsheet or credentials
are bundled in the repository.

## Dummy KAM unit mapping

`examples/kam_mapping_dummy.xlsx` maps the five supplied sample buyers to
KAM BD, KAM SI, KAM DG and KAM FFO. All assignments are fictional test data.
The active sheet has `Auftraggeber`, `KAM` and an informational `Hinweis` column.
The KAM column includes a dropdown for the four units.

```sh
python procurement.py prepare "your-document.pdf" --mapping examples/kam_mapping_dummy.xlsx --output output/dummy-demo
```

Here `KAM` is a unit label, not a person's name or a Jira account ID.
It can populate a suitable Jira text field; assignee/group selection requires
configuration for your actual Jira form. Replace dummy assignments with verified
ones in a separate private workbook before operational use.
