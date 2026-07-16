import re
import pandas as pd
from playwright.sync_api import sync_playwright

DEBUGGING_URL = "http://127.0.0.1:9222"
OUTPUT_FILE = "tnea_cutoff_table.xlsx"


def get_pagination_range(page):
    # Matches text such as: 1–100 of 840
    return page.get_by_text(
        re.compile(r"^\s*\d+\s*[-–]\s*\d+\s+of\s+\d+\s*$")
    ).first


def get_next_button(page):
    # Finds the paginator containing "1–100 of 840",
    # then returns its final icon-only button: the right-arrow.
    container = get_pagination_range(page)

    for _ in range(5):
        buttons = container.locator("button")

        if buttons.count() >= 2:
            for i in range(buttons.count() - 1, -1, -1):
                button = buttons.nth(i)

                # Numeric page buttons have text; arrow buttons do not.
                if button.inner_text().strip() == "":
                    return button

        container = container.locator("xpath=..")

    return None


def is_disabled(button):
    return (
        button.get_attribute("disabled") is not None
        or button.get_attribute("aria-disabled") == "true"
        or "disabled" in (button.get_attribute("class") or "").lower()
    )


def extract_headers(table):
    header_rows = table.locator("thead tr")

    if header_rows.count() < 2:
        raise RuntimeError("Expected the TNEA table's two header rows.")

    top_cells = header_rows.nth(0).locator("th")
    second_cells = header_rows.nth(1).locator("th")

    top_headers = [
        top_cells.nth(i).inner_text().strip()
        for i in range(top_cells.count())
    ]

    second_headers = [
        second_cells.nth(i).inner_text().strip()
        for i in range(second_cells.count())
    ]

    # Produces: CODE, COLLEGE NAME, BRANCH, OC 2025, OC 2024
    return top_headers[:3] + [
        f"{top_headers[3]} {year}"
        for year in second_headers
    ]


def extract_rows(table):
    rows = []
    table_rows = table.locator("tbody tr")

    for i in range(table_rows.count()):
        cells = table_rows.nth(i).locator("td")

        values = [
            cells.nth(j).inner_text().strip()
            for j in range(cells.count())
        ]

        if values:
            rows.append(values)

    return rows


def main():
    with sync_playwright() as p:
        # Attach to the Chrome window you launched manually
        browser = p.chromium.connect_over_cdp(DEBUGGING_URL)

        pages = [
            page
            for context in browser.contexts
            for page in context.pages
        ]

        page = next(
            (
                page
                for page in pages
                if "cutoff.tneaonline.org" in page.url
            ),
            None,
        )

        if page is None:
            raise RuntimeError(
                "TNEA page not found. Open the TNEA filtered results page first."
            )

        print("Connected to:", page.url)
        input(
            "Manually complete verification, set all filters, and show page 1. "
            "Then press ENTER..."
        )

        table = page.locator("table").first

        if table.count() == 0:
            raise RuntimeError("No results table found.")

        # Header extracted exactly once, from page 1.
        headers = extract_headers(table)
        all_rows = []
        page_number = 1

        while True:
            print(f"Extracting page {page_number}...")

            current_rows = extract_rows(table)
            all_rows.extend(current_rows)

            next_button = get_next_button(page)

            # Stops on last page, where the right arrow is disabled.
            if next_button is None:
                print("Next-page arrow was not found. Stopping.")
                break

            if is_disabled(next_button):
                print("Last page reached.")
                break

            old_range = get_pagination_range(page).inner_text().strip()

            next_button.click()
            page.wait_for_timeout(1000)

            # Wait until page range changes, e.g. 1–100 -> 101–200.
            changed = False

            for _ in range(20):
                new_range = get_pagination_range(page).inner_text().strip()

                if new_range != old_range:
                    changed = True
                    break

                page.wait_for_timeout(500)

            if not changed:
                print("Next page did not load. Stopping to prevent duplicate rows.")
                break

            page_number += 1

        if not all_rows:
            raise RuntimeError("No data rows were extracted.")

        # Ensure every row has the same number of cells as the header.
        column_count = len(headers)

        normalized_rows = [
            row[:column_count] + [""] * max(0, column_count - len(row))
            for row in all_rows
        ]

        # Writes one worksheet with one header row.
        df = pd.DataFrame(normalized_rows, columns=headers)

        df.to_excel(
            OUTPUT_FILE,
            index=False,
            sheet_name="TNEA Cutoff",
        )

        print(f"Saved {len(df)} rows to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()