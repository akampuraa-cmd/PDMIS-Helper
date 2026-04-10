# PDMIS-Helper

Playwright-based Python automation script for the **PDMIS Farmer Institutional Support (FIS)** portal at `https://fis.pdmis.go.ug/`.

## Features

| # | Feature |
|---|---------|
| 1 | Automatic login detection — prompts for credentials in the terminal |
| 2 | Math CAPTCHA auto-solve ("Are you human? X + Y = ?") |
| 3 | Navigation: Loan Management → Approved Loan |
| 4 | Single-select dropdown filters: Region, District, Subcounty |
| 5 | Multi-select dropdown filters: Parish, Payment Status |
| 6 | Search bar with configurable category and query |
| 7 | Primary table extraction (Applicant/ID, Loan Amount, Payment Status, Subsector, Date of Creation) |
| 8 | Deep-dive "View All" per beneficiary: Owner Name(s), NIN, Tel. Contact(s) |
| 9 | Export combined data to `.xlsx` (Excel) or `.csv` |

---

## Quick-Start

### 1 — Install dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

### 2 — Customise the configuration

Open `pdmis_helper.py` and edit the **CONFIGURATION** block near the bottom of the file:

```python
# Single-select dropdowns (set to None to skip)
REGION     = "Central"
DISTRICT   = "Kampala"
SUBCOUNTY  = "Nakawa"

# Multi-select dropdowns (list of option texts, or None to skip)
PARISHES         = ["Banda", "Nakawa"]
PAYMENT_STATUSES = ["Approved"]

# Search bar
SEARCH_CATEGORY = "Applicant Name"
SEARCH_QUERY    = "John"

# Output file (.xlsx or .csv)
OUTPUT_FILE = "pdmis_approved_loans.xlsx"

# Limit deep-dive to first N rows (None = all rows)
MAX_ROWS = None

# Show browser window for debugging
HEADLESS = True
```

### 3 — Run

```bash
python pdmis_helper.py
```

The script will:
1. Open the portal and detect whether login is required.
2. If so, prompt you for your **Phone number / Email** and **Password** in the terminal (credentials are never stored).
3. Automatically read and solve the math CAPTCHA.
4. Apply your configured filters and search terms.
5. Print Table 1 and Table 2 to the console.
6. Save the combined data to the file specified in `OUTPUT_FILE`.

---

## Programmatic Usage

You can also import and drive `PDMISHelper` from your own script:

```python
from pdmis_helper import PDMISHelper

helper = PDMISHelper(headless=False, slow_mo=100)  # slow_mo useful for debugging
df = helper.run(
    region="Central",
    district="Kampala",
    payment_statuses=["Approved", "Pending"],
    search_category="Applicant Name",
    search_query="John",
    output_file="results.xlsx",
    max_rows=10,   # only process first 10 rows for a quick test
)
print(df)
```

---

## Requirements

- Python 3.9+
- [Playwright for Python](https://playwright.dev/python/) — `pip install playwright && playwright install chromium`
- [pandas](https://pandas.pydata.org/) — `pip install pandas`
- [openpyxl](https://openpyxl.readthedocs.io/) — `pip install openpyxl` (needed for `.xlsx` export)

---

## Notes

- The selectors used in the script are best-effort based on common portal patterns. If the portal's HTML structure differs, edit the relevant method's selector lists (each method contains a comment indicating which selectors to update).
- The script handles paginated results tables automatically (keeps clicking "Next" until all pages are scraped).
- Set `HEADLESS = False` to watch the browser in real time — useful when first testing or updating selectors.

