import pandas as pd
import yfinance as yf
import os
from openpyxl import load_workbook

# === Path to the Excel file containing tickers ===
tickers_file = r"C:\Users\billy\OneDrive\Bureau\Finances\Investissements\Stocks\Long\IBKR API Python - Stock Info - v2.xlsx"

# === Folder to save the financial data files ===
folder_path = r"C:\Users\billy\OneDrive\Bureau\Finances\Investissements\Stocks\Long\XLSX"

# === Read tickers from the Excel file ===
# Assumption: tickers are in a column named 'Ticker' (adapt if your file differs)
tickers_df = pd.read_excel(tickers_file)
tickers_list = tickers_df['Ticker'].dropna().unique().tolist()

def append_new_columns_only(sheet_name, new_df, writer, file_path):
    new_df = new_df.iloc[::-1]  # Reorder columns chronologically

    try:
        old_df = pd.read_excel(file_path, sheet_name=sheet_name, index_col=0)

        new_cols = [col for col in new_df.columns if col not in old_df.columns]
        new_data_to_add = new_df[new_cols]

        if new_data_to_add.empty:
            print(f"🔁 Nothing to add for '{sheet_name}' (all data already present)")
            combined_df = old_df
        else:
            print(f"➕ Adding new columns for '{sheet_name}': {new_cols}")
            combined_df = pd.concat([old_df, new_data_to_add], axis=1)

    except ValueError:
        # Sheet does not exist yet
        print(f"🆕 Creating sheet '{sheet_name}' with current data")
        combined_df = new_df

    combined_df.to_excel(writer, sheet_name=sheet_name, startrow=0)

for ticker_symbol in tickers_list:
    print(f"\nProcessing ticker: {ticker_symbol}")
    file_path = os.path.join(folder_path, f"{ticker_symbol}_valuation_measures.xlsx")

    # Fetch data for the ticker
    ticker = yf.Ticker(ticker_symbol)
    financial_data = {
        "Annual Income Statement": ticker.financials,
        "Quarterly Income Statement": ticker.quarterly_financials,
        "Annual Balance Sheet": ticker.balance_sheet,
        "Quarterly Balance Sheet": ticker.quarterly_balance_sheet,
        "Annual Cash Flow": ticker.cashflow,
        "Quarterly Cash Flow": ticker.quarterly_cashflow,
    }

    if os.path.exists(file_path):
        book = load_workbook(file_path)
        with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            writer._book = book
            for sheet_name, df in financial_data.items():
                append_new_columns_only(sheet_name, df, writer, file_path)
        print(f"✅ Financial data updated in: {file_path}")
    else:
        # Create a new file from scratch (write all sheets)
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            for sheet_name, df in financial_data.items():
                df.iloc[::-1].to_excel(writer, sheet_name=sheet_name, startrow=0)
        print(f"🆕 New financial data file created: {file_path}")