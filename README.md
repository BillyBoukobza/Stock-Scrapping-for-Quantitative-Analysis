# Stock-Scrapping-for-Quantitative-Analysis

## Overview

This project contains scripts designed to fetch various types of stock data from online sources (primarily Yahoo Finance) for quantitative analysis. It includes a consolidated JavaScript application for fetching analyst data and key statistics, and a Python script for retrieving detailed financial statements.

## Structure

The main scripts for use are:

*   `combined_stock_scripts.js`: A Node.js script that consolidates the functionality of three original scripts. It fetches:
    *   Analyst price targets.
    *   Analyst recommendations.
    *   Key statistics (e.g., valuation measures).
    It saves this data into `.xlsx` files in a structured output folder, typically creating or updating files like `{ticker}_valuation_measures.xlsx` and `recommendation_trend_{ticker}.xlsx`.

*   `Financials_for_multiple_stocks.py`: A Python script that fetches detailed financial statements (Income Statement, Balance Sheet, Cash Flow) for a list of stock tickers. It saves each statement into a separate sheet within an Excel file named `{ticker}_financial_statements.xlsx`.

The original individual JavaScript files are also still available in the repository:
*   `Analyst Price Targets generalized to all stocks.js`
*   `Analyst Reccomendation_looping on multiple Stock to create new Excel Sheets.js`
*   `Statistics to XLS.js`

Configuration files:
*   `package.json`: Defines the Node.js project, dependencies, and scripts for the JavaScript part.
*   `requirements.txt`: Lists the Python dependencies.
*   `.gitignore`: Specifies intentionally untracked files that Git should ignore (e.g., `node_modules`).

## Prerequisites

To run these scripts, you'll need the following installed on your system:

*   **For `combined_stock_scripts.js`:**
    *   [Node.js](https://nodejs.org/) (which includes npm, the Node.js package manager).
*   **For `Financials_for_multiple_stocks.py`:**
    *   [Python](https://www.python.org/) (version 3.x recommended).
    *   pip (Python package installer, usually comes with Python).

## Setup

1.  **Clone the repository (if you haven't already):**
    ```bash
    git clone --branch feat/consolidated-scripts --single-branch https://github.com/BillyBoukobza/Stock-Scrapping-for-Quantitative-Analysis
    cd Stock-Scrapping-for-Quantitative-Analysis
    ```

2.  **JavaScript Dependencies:**
    Navigate to the project's root directory in your terminal and run:
    ```bash
    npm install
    ```
    This command will download and install the necessary Node.js packages defined in `package.json` (yahoo-finance2, xlsx, puppeteer).

3.  **Python Dependencies:**
    In the project's root directory, run:
    ```bash
    pip install -r requirements.txt
    ```
    This command will install the required Python libraries defined in `requirements.txt` (pandas, yfinance, openpyxl).

## Running the Scripts

### JavaScript (`combined_stock_scripts.js`)

To run the consolidated JavaScript script, navigate to the project's root directory in your terminal and use:
```bash
npm start
```
Alternatively, you can directly execute the file with Node.js:
```bash
node combined_stock_scripts.js
```
This script will process the tickers found in the input Excel file (see Important Notes below) and generate/update Excel files in the specified output directory.

### Python (`Financials_for_multiple_stocks.py`)

To run the Python script, ensure you are in the project's root directory and execute:
```bash
python Financials_for_multiple_stocks.py
```
This script will also use the tickers from the input Excel file and save financial statements into new Excel files in the output directory.

## Important Notes: Hardcoded Paths

**Crucial:** Both the JavaScript (`combined_stock_scripts.js`) and Python (`Financials_for_multiple_stocks.py`) scripts currently use **hardcoded file paths** for:

1.  **Input Ticker File:** The Excel file from which stock tickers are read.
    *   Example path in scripts: `C:\Users\billy\OneDrive\Bureau\Finances\Investissements\Stocks\Long\IBKR API Python - Stock Info - v2.xlsx`

2.  **Output Directory:** The folder where the generated Excel files are saved.
    *   Example path in scripts: `C:\Users\billy\OneDrive\Bureau\Finances\Investissements\Stocks\Long\XLSX`

**You MUST modify these paths directly within the scripts if your file locations or desired output directories are different.**

*   In `combined_stock_scripts.js`: Look for `excelPath` and `outputFolder` constants near the top of the file.
*   In `Financials_for_multiple_stocks.py`: Look for `excel_file_path` and `output_folder` variables near the top of the script.

Failure to update these paths to match your environment will result in the scripts not finding the input file or saving data to unintended locations.
