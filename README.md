# Yahoo Finance Earnings Calendar Scraper

This Python script fetches and extracts the earnings calendar data for a given stock ticker from Yahoo Finance. It uses Selenium to control a web browser, allowing it to access dynamically loaded content.

## Prerequisites

*   Python 3.6 or higher.
*   Google Chrome web browser installed.
*   ChromeDriver installed:
    *   Download the ChromeDriver version that matches your Google Chrome browser version from [https://chromedriver.chromium.org/downloads](https://chromedriver.chromium.org/downloads).
    *   The ChromeDriver executable must be in your system's PATH. Alternatively, you would need to specify the path to the executable directly in the script (this script currently assumes it's in PATH).

## Setup

1.  **Install Python dependencies:**
    Open your terminal or command prompt and run:
    ```bash
    pip install selenium
    ```

## Running the Script

Execute the script from your terminal:

```bash
python yahoo_earnings_calendar.py <TICKER_SYMBOL>
```

Replace `<TICKER_SYMBOL>` with the stock ticker you are interested in.

**Example:**

```bash
python yahoo_earnings_calendar.py AAPL
```

## Output

The script will attempt to locate and print the JSON object containing the earnings calendar dates and related information for the specified ticker. If the specific earnings data path within the larger page data isn't found directly, it will print a larger portion of the available JSON data for inspection.

## Troubleshooting

*   **`selenium.common.exceptions.WebDriverException: Message: 'chromedriver' executable needs to be in PATH.`**
    This means the ChromeDriver executable was not found. Ensure you have downloaded ChromeDriver, and the directory containing `chromedriver.exe` (Windows) or `chromedriver` (macOS/Linux) is added to your system's PATH environment variable. You might need to restart your terminal or system for PATH changes to take effect.
