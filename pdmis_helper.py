"""
PDMIS Helper — Playwright automation script for https://fis.pdmis.go.ug/
=========================================================================
Automates:
  1. Login (with math-CAPTCHA solving)
  2. Navigation  → Loan Management → Approved Loan
  3. Dropdown filtering (Region / District / Subcounty / Parish / Payment Status)
  4. Search bar with category dropdown
  5. Primary table extraction  (Table 1)
  6. "View All" deep-dive per beneficiary row
  7. Secondary data extraction (Table 2)
  8. Combined export to Excel (.xlsx) or CSV (.csv)

Usage
-----
  pip install -r requirements.txt
  playwright install chromium
  python pdmis_helper.py

Customisation
-------------
Edit the CONFIGURATION block near the bottom of this file (or call
PDMISHelper directly from your own script) to choose:
  • region / district / subcounty          — single-select
  • parishes / payment_statuses            — multi-select (list of strings)
  • search_category / search_query         — search bar
  • output_file                            — path for the exported file
"""

from __future__ import annotations

import re
import sys
import getpass
from typing import Optional

import pandas as pd
from playwright.sync_api import sync_playwright, Page, Locator


# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------

def _safe_text(loc: Locator) -> str:
    """Return stripped inner-text or empty string if the element is missing."""
    try:
        return loc.inner_text().strip()
    except Exception:
        return ""


def _wait_and_click(page: Page, selector: str, timeout: int = 15_000) -> None:
    """Wait for an element to be visible and then click it."""
    page.wait_for_selector(selector, state="visible", timeout=timeout)
    page.click(selector)


def _select_dropdown_option(page: Page, dropdown_selector: str, option_text: str) -> None:
    """
    Open a generic <select> or custom dropdown and choose an option by its
    visible text.  Tries <select> first, then falls back to looking for a
    list item / anchor whose text matches.
    """
    element = page.query_selector(dropdown_selector)
    if element is None:
        raise RuntimeError(f"Dropdown not found: {dropdown_selector}")

    tag = element.evaluate("el => el.tagName").upper()
    if tag == "SELECT":
        page.select_option(dropdown_selector, label=option_text)
    else:
        # Custom dropdown — click to open, then click matching option
        element.click()
        page.wait_for_timeout(500)
        option_loc = page.locator(
            f"li:has-text('{option_text}'), a:has-text('{option_text}')"
        ).first
        option_loc.wait_for(state="visible", timeout=10_000)
        option_loc.click()


# ---------------------------------------------------------------------------
# Main automation class
# ---------------------------------------------------------------------------

class PDMISHelper:
    """
    Orchestrates the full PDMIS scraping workflow.

    Parameters
    ----------
    headless : bool
        Run the browser without a visible window (default True).
    slow_mo : int
        Milliseconds added between Playwright actions — useful for debugging.
    timeout : int
        Default navigation / element timeout in milliseconds.
    """

    BASE_URL = "https://fis.pdmis.go.ug/"

    def __init__(
        self,
        headless: bool = True,
        slow_mo: int = 0,
        timeout: int = 30_000,
    ) -> None:
        self.headless = headless
        self.slow_mo = slow_mo
        self.timeout = timeout

        self._playwright = None
        self._browser = None
        self._context = None
        self.page: Optional[Page] = None

    # ------------------------------------------------------------------
    # Browser lifecycle
    # ------------------------------------------------------------------

    def start(self) -> None:
        """Launch the browser and open a blank page."""
        self._playwright = sync_playwright().start()
        self._browser = self._playwright.chromium.launch(
            headless=self.headless,
            slow_mo=self.slow_mo,
        )
        self._context = self._browser.new_context(
            viewport={"width": 1280, "height": 900},
            # Accept all cookies / permissions that the portal may request
            permissions=["geolocation"],
        )
        self.page = self._context.new_page()
        self.page.set_default_timeout(self.timeout)

    def stop(self) -> None:
        """Close browser and Playwright engine."""
        if self._browser:
            self._browser.close()
        if self._playwright:
            self._playwright.stop()

    # ------------------------------------------------------------------
    # Step 1 — Login detection + CAPTCHA solving
    # ------------------------------------------------------------------

    def login_if_required(self) -> None:
        """
        Navigate to the base URL.  If a login page is detected, prompt
        the operator for credentials, solve the math CAPTCHA, and submit.
        """
        print(f"[*] Opening {self.BASE_URL} …")
        self.page.goto(self.BASE_URL, wait_until="domcontentloaded")
        self.page.wait_for_load_state("networkidle", timeout=self.timeout)

        if not self._is_login_page():
            print("[*] Already authenticated — skipping login.")
            return

        print("[*] Login page detected.")

        # ---- Prompt for credentials (never hardcoded) ----
        username = input("    Enter your Phone number / Email: ").strip()
        password = getpass.getpass("    Enter your Password: ")

        self._fill_credentials(username, password)
        self._solve_math_captcha()
        self._submit_login()

        # Wait for post-login navigation
        self.page.wait_for_load_state("networkidle", timeout=self.timeout)

        if self._is_login_page():
            raise RuntimeError(
                "Login failed — still on the login page.  "
                "Please check your credentials."
            )

        print("[*] Login successful.")

    def _is_login_page(self) -> bool:
        """Return True if the current page contains a login / credential form."""
        indicators = [
            "input[type='password']",
            "input[name='password']",
            "#loginForm",
            ".login-form",
            "form[action*='login']",
        ]
        for selector in indicators:
            if self.page.query_selector(selector):
                return True
        # Also check the URL
        return "login" in self.page.url.lower()

    def _fill_credentials(self, username: str, password: str) -> None:
        """Locate the username and password fields and fill them in."""
        # Username / email / phone field — try common selectors in order
        username_selectors = [
            "input[name='username']",
            "input[name='email']",
            "input[name='phone']",
            "input[type='email']",
            "input[type='text'][placeholder*='mail' i]",
            "input[type='text'][placeholder*='phone' i]",
            "input[type='text']",  # fallback: first text input
        ]
        for sel in username_selectors:
            el = self.page.query_selector(sel)
            if el:
                el.fill(username)
                break
        else:
            raise RuntimeError("Could not find a username/email/phone input field.")

        # Password field
        self.page.fill("input[type='password']", password)

    def _solve_math_captcha(self) -> None:
        """
        Detect the math CAPTCHA question ("X + Y = ?"), compute the answer,
        and fill it into the verification input.

        The portal typically renders something like:
            <span class="captcha-question">5 + 7 = ?</span>
            <input id="captcha" …>
        """
        # Broad selector — covers most common patterns on this portal
        captcha_selectors = [
            ".captcha-question",
            ".captcha_question",
            "[class*='captcha']",
            "label[for*='captcha' i]",
            "p:has-text('Are you human')",
            "div:has-text('Are you human')",
            "span:has-text('+'):has-text('=')",
            "p:has-text('+'):has-text('=')",
        ]

        # NOTE: This script only supports addition CAPTCHAs ("X + Y = ?").
        # If the portal uses subtraction or multiplication, update the
        # pattern and the arithmetic in this method accordingly.
        _CAPTCHA_PATTERN = re.compile(r"\d+\s*\+\s*\d+")

        question_text = ""
        for sel in captcha_selectors:
            el = self.page.query_selector(sel)
            if el:
                candidate = el.inner_text().strip()
                if _CAPTCHA_PATTERN.search(candidate):
                    question_text = candidate
                    break

        if not question_text:
            # Try to grab any visible text that looks like "N + M = ?"
            body_text = self.page.inner_text("body")
            match = re.search(r"(\d+)\s*\+\s*(\d+)\s*=", body_text)
            if match:
                question_text = match.group(0)
            else:
                print("[!] Math CAPTCHA not found on this page — skipping.")
                return

        # Parse the two operands
        numbers = re.findall(r"\d+", question_text)
        if len(numbers) < 2:
            raise RuntimeError(
                f"Failed to parse numbers from CAPTCHA question: '{question_text}'"
            )
        answer = int(numbers[0]) + int(numbers[1])
        print(f"    CAPTCHA detected: '{question_text.strip()}' → answer = {answer}")

        # Fill in the answer
        captcha_input_selectors = [
            "input#captcha",
            "input[name='captcha']",
            "input[id*='captcha' i]",
            "input[name*='captcha' i]",
            "input[placeholder*='answer' i]",
            "input[placeholder*='sum' i]",
            "input[placeholder*='human' i]",
        ]
        for sel in captcha_input_selectors:
            el = self.page.query_selector(sel)
            if el:
                el.fill(str(answer))
                print(f"    CAPTCHA answer '{answer}' filled successfully.")
                return

        raise RuntimeError(
            "CAPTCHA answer input field not found.  "
            "Please inspect the portal and update the selectors in _solve_math_captcha()."
        )

    def _submit_login(self) -> None:
        """Click the login / verify / submit button."""
        submit_selectors = [
            "button[type='submit']",
            "input[type='submit']",
            "button:has-text('Login')",
            "button:has-text('Sign in')",
            "button:has-text('Verify')",
            "a:has-text('Login')",
            "#loginBtn",
            ".login-btn",
        ]
        for sel in submit_selectors:
            el = self.page.query_selector(sel)
            if el:
                el.click()
                return
        raise RuntimeError(
            "Login submit button not found.  Update the selectors in _submit_login()."
        )

    # ------------------------------------------------------------------
    # Step 2 — Navigate to Loan Management → Approved Loan
    # ------------------------------------------------------------------

    def navigate_to_approved_loans(self) -> None:
        """Click 'Loan Management' in the main navigation, then 'Approved Loan'."""
        print("[*] Navigating to Loan Management → Approved Loan …")

        # Click top-level "Loan Management" menu / tab
        loan_mgmt_selectors = [
            "a:has-text('Loan Management')",
            "span:has-text('Loan Management')",
            "li:has-text('Loan Management') > a",
            "nav a:has-text('Loan')",
        ]
        for sel in loan_mgmt_selectors:
            el = self.page.query_selector(sel)
            if el:
                el.click()
                self.page.wait_for_load_state("networkidle", timeout=self.timeout)
                break
        else:
            raise RuntimeError(
                "Could not find 'Loan Management' navigation element.  "
                "Update the selectors in navigate_to_approved_loans()."
            )

        # Click "Approved Loan" sub-menu item
        approved_selectors = [
            "a:has-text('Approved Loan')",
            "li:has-text('Approved Loan') > a",
            "span:has-text('Approved Loan')",
        ]
        for sel in approved_selectors:
            el = self.page.query_selector(sel)
            if el:
                el.click()
                self.page.wait_for_load_state("networkidle", timeout=self.timeout)
                break
        else:
            raise RuntimeError(
                "Could not find 'Approved Loan' navigation element.  "
                "Update the selectors in navigate_to_approved_loans()."
            )

        print("[*] Navigated to Approved Loan page.")

    # ------------------------------------------------------------------
    # Step 3 — Dropdown filters
    # ------------------------------------------------------------------

    def apply_filters(
        self,
        region: Optional[str] = None,
        district: Optional[str] = None,
        subcounty: Optional[str] = None,
        parishes: Optional[list[str]] = None,
        payment_statuses: Optional[list[str]] = None,
    ) -> None:
        """
        Apply dropdown filters on the Approved Loan page.

        Parameters
        ----------
        region : str, optional
            Exact text of the Region option to select.
        district : str, optional
            Exact text of the District option to select.
        subcounty : str, optional
            Exact text of the Subcounty option to select.
        parishes : list of str, optional
            One or more Parish option texts to select (multi-select).
        payment_statuses : list of str, optional
            One or more Payment Status option texts to select (multi-select).
        """
        # --- Single-select dropdowns ---
        # Selectors below are best-effort and may need adjusting after
        # inspecting the live portal DOM.
        if region:
            print(f"    Selecting Region: {region}")
            self._select_single(
                ["select[name*='region' i]", "#region", "[id*='region' i]"],
                region,
            )

        if district:
            print(f"    Selecting District: {district}")
            self._select_single(
                ["select[name*='district' i]", "#district", "[id*='district' i]"],
                district,
            )

        if subcounty:
            print(f"    Selecting Subcounty: {subcounty}")
            self._select_single(
                ["select[name*='subcounty' i]", "#subcounty", "[id*='subcounty' i]"],
                subcounty,
            )

        # --- Multi-select dropdowns ---
        if parishes:
            print(f"    Selecting Parish(es): {parishes}")
            self._select_multi(
                ["select[name*='parish' i]", "#parish", "[id*='parish' i]"],
                parishes,
            )

        if payment_statuses:
            print(f"    Selecting Payment Status(es): {payment_statuses}")
            self._select_multi(
                [
                    "select[name*='payment_status' i]",
                    "select[name*='paymentstatus' i]",
                    "#payment_status",
                    "[id*='payment' i][id*='status' i]",
                ],
                payment_statuses,
            )

        self.page.wait_for_timeout(1_000)

    def _select_single(self, selectors: list[str], option_text: str) -> None:
        """Try each selector in order and perform a single option selection."""
        for sel in selectors:
            el = self.page.query_selector(sel)
            if el:
                tag = el.evaluate("e => e.tagName").upper()
                if tag == "SELECT":
                    self.page.select_option(sel, label=option_text)
                else:
                    _select_dropdown_option(self.page, sel, option_text)
                self.page.wait_for_load_state("networkidle", timeout=self.timeout)
                return
        print(
            f"    [!] Could not locate dropdown for option '{option_text}'.  "
            "Inspect the portal and update the selectors in _select_single()."
        )

    def _select_multi(self, selectors: list[str], option_texts: list[str]) -> None:
        """
        Select multiple options from a multi-select element.
        Works for both native <select multiple> and custom multi-select widgets.
        """
        for sel in selectors:
            el = self.page.query_selector(sel)
            if el:
                tag = el.evaluate("e => e.tagName").upper()
                if tag == "SELECT":
                    # Native multi-select — pass all labels at once
                    self.page.select_option(sel, label=option_texts)
                else:
                    # Custom widget: open and click each option individually
                    for text in option_texts:
                        _select_dropdown_option(self.page, sel, text)
                self.page.wait_for_load_state("networkidle", timeout=self.timeout)
                return
        print(
            f"    [!] Could not locate multi-select for options {option_texts}.  "
            "Inspect the portal and update the selectors in _select_multi()."
        )

    # ------------------------------------------------------------------
    # Step 4 — Search bar
    # ------------------------------------------------------------------

    def search(
        self,
        search_category: Optional[str] = None,
        search_query: Optional[str] = None,
    ) -> None:
        """
        Select a category from the "Search Categories" dropdown and enter a
        search query, then execute the search.

        Parameters
        ----------
        search_category : str, optional
            Visible text of the search-category option (e.g. 'Applicant Name').
        search_query : str, optional
            Text to type into the search input box.
        """
        if not search_category and not search_query:
            return

        print(f"[*] Searching — category='{search_category}', query='{search_query}'")

        # Select the search category dropdown
        if search_category:
            cat_selectors = [
                "select[name*='search_category' i]",
                "select[id*='search_category' i]",
                "select[name*='searchcategory' i]",
                "#searchCategory",
                ".search-category select",
            ]
            for sel in cat_selectors:
                el = self.page.query_selector(sel)
                if el:
                    self.page.select_option(sel, label=search_category)
                    break
            else:
                print(
                    "    [!] Search-category dropdown not found.  "
                    "Skipping category selection."
                )

        # Enter the search query
        if search_query:
            search_input_selectors = [
                "input[name*='search' i]",
                "input[type='search']",
                "input[placeholder*='search' i]",
                "#searchInput",
                ".search-input",
            ]
            for sel in search_input_selectors:
                el = self.page.query_selector(sel)
                if el:
                    el.fill(search_query)
                    break
            else:
                print("    [!] Search input not found.  Skipping search query input.")

        # Click the Search / Go button
        search_btn_selectors = [
            "button:has-text('Search')",
            "input[type='submit'][value*='Search' i]",
            "button[type='submit']",
            ".search-btn",
            "#searchBtn",
        ]
        for sel in search_btn_selectors:
            el = self.page.query_selector(sel)
            if el:
                el.click()
                self.page.wait_for_load_state("networkidle", timeout=self.timeout)
                break
        else:
            # Fall back to pressing Enter in the search field
            self.page.keyboard.press("Enter")
            self.page.wait_for_load_state("networkidle", timeout=self.timeout)

    # ------------------------------------------------------------------
    # Step 5 — Primary table extraction (Table 1)
    # ------------------------------------------------------------------

    def extract_primary_table(self) -> pd.DataFrame:
        """
        Scrape the beneficiary results table and return a DataFrame with:
          Applicant/ID, Loan Amount, Payment Status, Subsector, Date of Creation.

        The method handles simple pagination by clicking "Next" until all
        pages have been scraped.
        """
        print("[*] Extracting primary table …")

        column_map = {
            # portal header text  →  standardised column name
            # Note: generic keys like "id" and "created" are included as
            # fallbacks; more specific keys take precedence because dict
            # look-ups are exact-match on lowercased header text.
            "applicant":          "Applicant/ID",
            "applicant id":       "Applicant/ID",
            "applicant/id":       "Applicant/ID",
            "loan amount":        "Loan Amount",
            "amount":             "Loan Amount",
            "payment status":     "Payment Status",
            "status":             "Payment Status",
            "subsector":          "Subsector",
            "sub sector":         "Subsector",
            "date of creation":   "Date of Creation",
            "date created":       "Date of Creation",
            "creation date":      "Date of Creation",
        }

        all_rows: list[dict] = []
        page_num = 1

        while True:
            print(f"    Scraping page {page_num} …")
            rows = self._scrape_table_page(column_map)
            all_rows.extend(rows)

            # Check for a "Next" pagination button
            next_btn = self.page.query_selector(
                "a:has-text('Next'), button:has-text('Next'), "
                ".pagination .next:not(.disabled)"
            )
            if not next_btn:
                break
            next_btn.click()
            self.page.wait_for_load_state("networkidle", timeout=self.timeout)
            page_num += 1

        df = pd.DataFrame(all_rows)
        print(f"[*] Primary table: {len(df)} rows extracted.")
        return df

    def _scrape_table_page(self, column_map: dict) -> list[dict]:
        """Return a list of row dicts for the visible table page."""
        # Locate the results table — try common selectors
        table_selectors = ["table.table", "table#resultsTable", "table", ".data-table"]
        table = None
        for sel in table_selectors:
            table = self.page.query_selector(sel)
            if table:
                break
        if not table:
            print("    [!] Results table not found on this page.")
            return []

        # Map header positions → column names we care about
        header_cells = table.query_selector_all("thead th, thead td")
        if not header_cells:
            header_cells = table.query_selector_all("tr:first-child th, tr:first-child td")

        col_indices: dict[str, int] = {}
        for i, cell in enumerate(header_cells):
            text = cell.inner_text().strip().lower()
            if text in column_map:
                col_indices[column_map[text]] = i

        target_cols = {
            "Applicant/ID", "Loan Amount", "Payment Status",
            "Subsector", "Date of Creation",
        }

        rows: list[dict] = []
        body_rows = table.query_selector_all("tbody tr")
        for tr in body_rows:
            cells = tr.query_selector_all("td")
            row: dict[str, str] = {}
            for col_name in target_cols:
                idx = col_indices.get(col_name)
                if idx is not None and idx < len(cells):
                    row[col_name] = cells[idx].inner_text().strip()
                else:
                    row[col_name] = ""
            if any(row.values()):
                rows.append(row)

        return rows

    # ------------------------------------------------------------------
    # Step 6 + 7 — "View All" deep-dive and secondary data extraction
    # ------------------------------------------------------------------

    def extract_secondary_data(self, primary_df: pd.DataFrame) -> pd.DataFrame:
        """
        For each row in primary_df, click the "Tools" button and then
        "View All", extract secondary data (owner names, NIN, phone), then
        navigate back.

        Returns a DataFrame aligned with primary_df (same row order).
        """
        print("[*] Extracting secondary data (View All pages) …")

        secondary_rows: list[dict] = []

        for idx in range(len(primary_df)):
            print(f"    Processing row {idx + 1} / {len(primary_df)} …")

            # Locate the Tools button for this row (1-based index)
            row_data = self._view_all_for_row(idx)
            secondary_rows.append(row_data)

            # Navigate back to the results table
            self.page.go_back()
            self.page.wait_for_load_state("networkidle", timeout=self.timeout)

        secondary_df = pd.DataFrame(secondary_rows)
        return secondary_df

    def _view_all_for_row(self, row_index: int) -> dict:
        """
        Click the Tools button on row `row_index` (0-based), choose "View All",
        then extract owner name, NIN, and phone contact.
        """
        # Re-query the table rows each time (the page may have reloaded)
        table_selectors = ["table.table tbody tr", "table tbody tr", ".data-table tbody tr"]
        rows = []
        for sel in table_selectors:
            rows = self.page.query_selector_all(sel)
            if rows:
                break

        if row_index >= len(rows):
            print(
                f"    [!] Row {row_index} not found (only {len(rows)} rows visible).  "
                "Returning empty secondary record."
            )
            return {"Owner Name(s)": "", "NIN": "", "Tel. Contact(s)": ""}

        row = rows[row_index]

        # Click the "Tools" button / dropdown trigger in this row
        tools_selectors = [
            "button:has-text('Tools')",
            "a:has-text('Tools')",
            ".tools-btn",
            ".dropdown-toggle",
            "button.btn-tools",
        ]
        tools_btn = None
        for sel in tools_selectors:
            tools_btn = row.query_selector(sel)
            if tools_btn:
                break

        if not tools_btn:
            print(f"    [!] Tools button not found for row {row_index}.  Skipping.")
            return {"Owner Name(s)": "", "NIN": "", "Tel. Contact(s)": ""}

        tools_btn.click()
        # Wait for the dropdown menu to become visible rather than using a
        # fixed timeout, which can be brittle on slower connections.
        try:
            self.page.wait_for_selector(
                "a:has-text('View All'), .dropdown-menu, .dropdown-item",
                state="visible",
                timeout=5_000,
            )
        except Exception:
            pass  # proceed anyway — the selector loop below will handle missing menus

        # Click the first option: "View All"
        view_all_selectors = [
            "a:has-text('View All')",
            "li:has-text('View All') > a",
            ".dropdown-menu a:first-child",
            ".dropdown-item:first-child",
        ]
        for sel in view_all_selectors:
            el = self.page.query_selector(sel)
            if el:
                el.click()
                self.page.wait_for_load_state("networkidle", timeout=self.timeout)
                break
        else:
            print(
                f"    [!] 'View All' option not found for row {row_index}.  Skipping."
            )
            return {"Owner Name(s)": "", "NIN": "", "Tel. Contact(s)": ""}

        return self._extract_detail_fields()

    def _extract_detail_fields(self) -> dict:
        """
        On a beneficiary detail page, extract:
          • Owner / business owner name(s)
          • NIN (National Identification Number)
          • Phone / Tel. Contact(s)

        The selectors below cover the most common label-value patterns.
        Adjust them once you have inspected the live portal.
        """
        result = {"Owner Name(s)": "", "NIN": "", "Tel. Contact(s)": ""}

        # Strategy: look for label text, then grab the adjacent value element.
        label_targets = {
            "Owner Name(s)": [
                "owner", "business owner", "name of owner", "proprietor",
            ],
            "NIN": ["nin", "national id", "national identification"],
            "Tel. Contact(s)": [
                "tel", "telephone", "phone", "mobile", "contact",
            ],
        }

        # Grab all label-like elements on the page
        label_els = self.page.query_selector_all(
            "th, td.label, .detail-label, label, dt, strong"
        )

        for label_el in label_els:
            label_text = label_el.inner_text().strip().lower()
            for field, keywords in label_targets.items():
                if result[field]:
                    continue  # already found
                if any(kw in label_text for kw in keywords):
                    # Try sibling <td> first, then next sibling element
                    value = ""
                    sibling = label_el.evaluate_handle(
                        "el => el.nextElementSibling"
                    )
                    if sibling:
                        try:
                            value = sibling.as_element().inner_text().strip()
                        except Exception:
                            pass
                    if not value:
                        # Try the parent row's next cell
                        try:
                            value = label_el.evaluate(
                                "el => {"
                                "  const row = el.closest('tr');"
                                "  if (!row) return '';"
                                "  const cells = row.querySelectorAll('td, th');"
                                "  for (let i = 0; i < cells.length - 1; i++) {"
                                "    if (cells[i] === el) return cells[i+1].innerText.trim();"
                                "  }"
                                "  return '';"
                                "}"
                            )
                        except Exception:
                            pass
                    if value:
                        result[field] = value

        return result

    # ------------------------------------------------------------------
    # Step 8 — Data export
    # ------------------------------------------------------------------

    def export_data(
        self,
        primary_df: pd.DataFrame,
        secondary_df: pd.DataFrame,
        output_file: str = "pdmis_approved_loans.xlsx",
    ) -> None:
        """
        Combine primary and secondary DataFrames and save to Excel or CSV.

        Parameters
        ----------
        primary_df : pd.DataFrame
            Table 1 data (applicant, loan amount, etc.).
        secondary_df : pd.DataFrame
            Table 2 data (owner name, NIN, phone).
        output_file : str
            Destination file path.  Use `.xlsx` for Excel or `.csv` for CSV.
        """
        if primary_df.empty and secondary_df.empty:
            print("[!] Both DataFrames are empty — nothing to export.")
            return

        if len(primary_df) != len(secondary_df):
            print(
                f"[!] Row count mismatch: primary has {len(primary_df)} rows, "
                f"secondary has {len(secondary_df)} rows.  "
                "Some rows will contain NaN values in the export."
            )

        combined = pd.concat(
            [primary_df.reset_index(drop=True), secondary_df.reset_index(drop=True)],
            axis=1,
        )

        if output_file.lower().endswith(".csv"):
            combined.to_csv(output_file, index=False, encoding="utf-8-sig")
        else:
            combined.to_excel(output_file, index=False)

        print(f"[*] Data exported → {output_file}  ({len(combined)} rows)")

    # ------------------------------------------------------------------
    # Convenience: pretty-print a DataFrame in the console
    # ------------------------------------------------------------------

    @staticmethod
    def print_table(df: pd.DataFrame, title: str = "") -> None:
        """Print a DataFrame as a formatted console table."""
        if title:
            print(f"\n{'=' * 60}")
            print(f"  {title}")
            print(f"{'=' * 60}")
        if df.empty:
            print("  (no data)")
        else:
            print(df.to_string(index=False))
        print()

    # ------------------------------------------------------------------
    # Full workflow shortcut
    # ------------------------------------------------------------------

    def run(
        self,
        # ── Dropdown filter settings ─────────────────────────────────
        region: Optional[str] = None,
        district: Optional[str] = None,
        subcounty: Optional[str] = None,
        parishes: Optional[list[str]] = None,
        payment_statuses: Optional[list[str]] = None,
        # ── Search settings ──────────────────────────────────────────
        search_category: Optional[str] = None,
        search_query: Optional[str] = None,
        # ── Export settings ──────────────────────────────────────────
        output_file: str = "pdmis_approved_loans.xlsx",
        # ── Deep-dive settings ───────────────────────────────────────
        max_rows: Optional[int] = None,
    ) -> pd.DataFrame:
        """
        Execute the complete workflow and return the combined DataFrame.

        Parameters
        ----------
        region, district, subcounty : str, optional
            Single-select dropdown values.
        parishes, payment_statuses : list of str, optional
            Multi-select dropdown values.
        search_category : str, optional
            Category text for the Search Categories dropdown.
        search_query : str, optional
            Query text to type in the search bar.
        output_file : str
            Output file path (.xlsx or .csv).
        max_rows : int, optional
            Limit the deep-dive to the first N rows (useful for testing).
        """
        try:
            self.start()
            self.login_if_required()
            self.navigate_to_approved_loans()
            self.apply_filters(
                region=region,
                district=district,
                subcounty=subcounty,
                parishes=parishes,
                payment_statuses=payment_statuses,
            )
            self.search(
                search_category=search_category,
                search_query=search_query,
            )

            primary_df = self.extract_primary_table()
            self.print_table(primary_df, title="Table 1 — Approved Loans")

            # Optionally limit to a subset for the deep-dive
            primary_subset = primary_df.iloc[:max_rows] if max_rows else primary_df
            secondary_df = self.extract_secondary_data(primary_subset)
            self.print_table(secondary_df, title="Table 2 — Beneficiary Details")

            self.export_data(primary_subset, secondary_df, output_file=output_file)

            return pd.concat(
                [primary_subset.reset_index(drop=True),
                 secondary_df.reset_index(drop=True)],
                axis=1,
            )
        finally:
            self.stop()


# ===========================================================================
# CONFIGURATION — Edit this section to customise your run
# ===========================================================================

if __name__ == "__main__":

    # ── Single-select dropdowns ──────────────────────────────────────────────
    # Set to None (or remove) to leave the dropdown at its default value.

    REGION     = None          # e.g. "Central"
    DISTRICT   = None          # e.g. "Kampala"
    SUBCOUNTY  = None          # e.g. "Nakawa"

    # ── Multi-select dropdowns ───────────────────────────────────────────────
    # Provide a list of exact option texts.  Use an empty list [] or None to
    # skip the dropdown entirely.

    PARISHES         = None    # e.g. ["Banda", "Nakawa"]
    PAYMENT_STATUSES = None    # e.g. ["Approved", "Pending"]

    # ── Search bar ───────────────────────────────────────────────────────────
    SEARCH_CATEGORY = None     # e.g. "Applicant Name"
    SEARCH_QUERY    = None     # e.g. "John"

    # ── Output file ──────────────────────────────────────────────────────────
    OUTPUT_FILE = "pdmis_approved_loans.xlsx"   # change to .csv if preferred

    # ── Deep-dive limit ──────────────────────────────────────────────────────
    # Set to an integer (e.g. 5) to only process the first N rows of Table 1.
    # Set to None to process all rows.
    MAX_ROWS = None

    # ── Browser settings ─────────────────────────────────────────────────────
    HEADLESS = True    # set False to watch the browser in action
    SLOW_MO  = 0       # milliseconds between actions (increase for debugging)

    # ── Run ──────────────────────────────────────────────────────────────────
    helper = PDMISHelper(headless=HEADLESS, slow_mo=SLOW_MO)
    combined_df = helper.run(
        region=REGION,
        district=DISTRICT,
        subcounty=SUBCOUNTY,
        parishes=PARISHES,
        payment_statuses=PAYMENT_STATUSES,
        search_category=SEARCH_CATEGORY,
        search_query=SEARCH_QUERY,
        output_file=OUTPUT_FILE,
        max_rows=MAX_ROWS,
    )

    sys.exit(0)
