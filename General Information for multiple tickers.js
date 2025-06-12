const yahooFinance = require('yahoo-finance2').default;
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const inputExcelPath = path.join(
  'C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long',
  'IBKR API Python - Stock Info - v2.xlsx'
);

const outputDir = path.join(
  'C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long', 'XLSX'
);

// Ensure output directory exists
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
  console.log(`Created output directory: ${outputDir}`);
}

// Read tickers from Excel
function getTickersFromExcel(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet);

  // Extract Ticker column, filter out empty or invalid entries
  return data
    .map(row => row.Ticker)
    .filter(ticker => typeof ticker === 'string' && ticker.trim().length > 0);
}

// Fetch info for one ticker
async function fetchTickerInfo(ticker) {
  try {
    const summary = await yahooFinance.quoteSummary(ticker, {
      modules: ['assetProfile', 'price']
    });

    const profile = summary.assetProfile || {};
    const price = summary.price || {};

    return {
      Ticker: ticker,
      TradingCurrency: price.currency || 'N/A',
      Sector: profile.sector || 'N/A',
      Industry: profile.industry || 'N/A',
      FullTimeEmployees: profile.fullTimeEmployees || 'N/A',
    };
  } catch (error) {
    console.error(`Error fetching data for ${ticker}: ${error.message}`);
    return {
      Ticker: ticker,
      TradingCurrency: 'Error',
      Sector: 'Error',
      Industry: 'Error',
      FullTimeEmployees: 'Error',
    };
  }
}

// Save data to Excel file
function saveToExcel(ticker, dataObj) {
  const fileName = `${ticker}_valuation_measures.xlsx`;
  const filePath = path.join(outputDir, fileName);

  // Convert single data object to sheet
  const newSheet = XLSX.utils.json_to_sheet([dataObj]);

  let workbook;

  if (fs.existsSync(filePath)) {
    workbook = XLSX.readFile(filePath);

    // Remove existing "General information" sheet if present to avoid duplicates
    if (workbook.SheetNames.includes('General information')) {
      const idx = workbook.SheetNames.indexOf('General information');
      workbook.SheetNames.splice(idx, 1);
      delete workbook.Sheets['General information'];
    }
  } else {
    workbook = XLSX.utils.book_new();
  }

  // Append new sheet named "General information"
  XLSX.utils.book_append_sheet(workbook, newSheet, 'General information');

  XLSX.writeFile(workbook, filePath);
  console.log(`✅ Added/updated "General information" sheet in ${filePath}`);
}

// Main runner
async function run() {
  const tickers = getTickersFromExcel(inputExcelPath);
  console.log(`Found ${tickers.length} tickers.`);

  for (const ticker of tickers) {
    console.log(`Fetching data for ${ticker}...`);
    const info = await fetchTickerInfo(ticker);
    saveToExcel(ticker, info);
  }

  console.log('All done!');
}

run();