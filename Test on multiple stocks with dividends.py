import yfinance as yf
import pandas as pd
import os

# === Configuration ===
input_excel_path = r"C:\Users\billy\OneDrive\Bureau\Finances\Investissements\Stocks\Long\IBKR API Python - Stock Info - v2.xlsx"
xlsx_output_folder = r"C:\Users\billy\OneDrive\Bureau\Finances\Investissements\Stocks\Long\XLSX"

intervals = {
    "1d": "Daily",
    "1wk": "Weekly",
    "1mo": "Monthly"
}

# === Utilitaire : vérifier que le fichier Excel est lisible ===
def is_valid_excel_file(file_path):
    try:
        pd.read_excel(file_path, engine="openpyxl", nrows=1)
        return True
    except:
        return False

# === Read tickers ===
try:
    df_tickers = pd.read_excel(input_excel_path)
    tickers = df_tickers["Ticker"].dropna().unique().tolist()
except Exception as e:
    raise RuntimeError(f"❌ Could not read ticker list from Excel: {e}")

# === Process each ticker ===
for ticker in tickers:
    print(f"\n📈 Processing {ticker}...")
    output_file = os.path.join(xlsx_output_folder, f"{ticker}_valuation_measures.xlsx")

    # Create or reset workbook if not exists or corrupted
    if not os.path.exists(output_file) or not is_valid_excel_file(output_file):
        print(f"  ⚠️ {output_file} is missing or corrupted. Re-creating it.")
        with pd.ExcelWriter(output_file, engine="openpyxl", mode="w") as writer:
            pass

    with pd.ExcelWriter(output_file, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        # Historical Prices
        for interval, sheet_name in intervals.items():
            try:
                print(f"  ⏳ Fetching {interval} data...")
                df = yf.download(ticker, period="max", interval=interval, auto_adjust=False)

                if isinstance(df.columns, pd.MultiIndex):
                    df.columns = ['_'.join(col).strip() for col in df.columns.values]

                df.reset_index(inplace=True)
                df.to_excel(writer, sheet_name=sheet_name, index=False)

            except Exception as e:
                print(f"  ⚠️ Failed to fetch or write {interval} data for {ticker}: {e}")

        # Dividends
        try:
            print(f"  ⏳ Fetching dividends...")
            dividends = yf.Ticker(ticker).dividends

            # Make sure index is datetime without timezone even if empty
            if isinstance(dividends.index, pd.DatetimeIndex):
                dividends.index = dividends.index.tz_localize(None)

            if not dividends.empty:
                df_div = dividends.reset_index()
                df_div.columns = ["Date", "Dividend"]
            else:
                print(f"  ℹ️ No dividend data for {ticker}. Creating empty sheet.")
                df_div = pd.DataFrame(columns=["Date", "Dividend"])

            # Force 'Date' column to datetime type even if empty
            if "Date" in df_div.columns:
                df_div["Date"] = pd.to_datetime(df_div["Date"], errors='coerce')

            df_div.to_excel(writer, sheet_name="Dividends", index=False)

        except Exception as e:
            print(f"  ⚠️ Failed to fetch/write dividends for {ticker}: {e}")

print("\n✅ All price and dividend data successfully written.")