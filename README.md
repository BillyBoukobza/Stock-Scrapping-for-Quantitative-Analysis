# Stock Scraping for Quantitative Analysis (Yahoo Finance Scraper)

This project contains a Node.js script (`yahoo_finance_scraper.js`) designed to fetch various financial data points for a list of stock tickers. It retrieves information such as analyst price targets, recommendations, and key statistics, primarily using the `yahoo-finance2` library and `puppeteer` for web scraping.

## Prerequisites

- Node.js (which includes npm)

## Setup

1.  **Clone the repository (if you haven't already):**
    ```bash
    git clone https://github.com/BillyBoukobza/Stock-Scrapping-for-Quantitative-Analysis.git
    cd Stock-Scrapping-for-Quantitative-Analysis
    ```
2.  **Install dependencies:**
    Open your terminal in the project root directory and run:
    ```bash
    npm install
    ```
    This will install all necessary packages defined in `package.json` (including `yahoo-finance2`, `xlsx`, `puppeteer`, `p-limit`, and `delay`).

## Running the Script

To execute the scraper, run the following command in your terminal from the project root:
```bash
npm start
```
Alternatively, you can run it directly with node:
```bash
node yahoo_finance_scraper.js
```

## Configuration: IMPORTANT!

The script currently uses **hardcoded paths** to locate the input Excel file (containing stock tickers) and to determine where the output Excel files will be saved.

You **MUST** modify these paths within the `yahoo_finance_scraper.js` file to match your local environment before running the script.

Look for these lines near the top of `yahoo_finance_scraper.js` and update them:

```javascript
// Path to the Excel file containing tickers
const excelPath = path.join('C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finanaces', 'Investissements', 'Stocks', 'Long', 'IBKR API Python - Stock Info - v2.xlsx');

// Output folder for the generated Excel files
const outputFolder = path.join('C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long', 'XLSX');
```

Failure to update these paths will likely result in the script not finding your input file or saving output to an unintended location.

## Original Scripts

The repository may also contain older scripts (like `combined_stock_scripts.js` or individual Python scripts) from previous iterations. This README focuses on the `yahoo_finance_scraper.js` script. Refer to previous commits or branches if you need information on those.
