// Combined Community Script

// 1. Dependencies
const yahooFinance = require('yahoo-finance2').default;
const XLSX = require('xlsx');
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

// 2. Configuration (Paths)
const excelPath = path.join(
  'C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long',
  'IBKR API Python - Stock Info - v2.xlsx'
);

const outputFolder = path.join(
  'C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long', 'XLSX'
);

// 3. Helper Function: Read Tickers
async function getTickers(filePath) {
  try {
    const wb = XLSX.readFile(filePath);
    const sheetName = wb.SheetNames[0];
    const tickers = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]).map(row => row.Ticker).filter(Boolean);
    console.log(`🔍 Found ${tickers.length} tickers in ${filePath}`);
    return tickers;
  } catch (err) {
    console.error(`❌ Error reading tickers from ${filePath}:`, err.message);
    throw err; // Re-throw to be caught by calling function
  }
}

// 4. Helper Function: Ensure Output Directory Exists
function ensureOutputDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
    console.log(`📁 Created output directory: ${dirPath}`);
  }
}

// 5. Helper Function: Parse Period (from Analyst Reccomendation_looping on multiple Stock to create new Excel Sheets.js)
function parsePeriod(offsetStr) {
  const match = offsetStr.match(/(-?\d+)m/);
  if (!match) return offsetStr;

  const monthsAgo = parseInt(match[1], 10);
  const baseDate = new Date();
  const targetDate = new Date(baseDate.getFullYear(), baseDate.getMonth() + monthsAgo, 1);
  return targetDate.toISOString().slice(0, 10);
}

// --- Main Functionality Sections ---

async function fetchAnalystPriceTargets(tickers, outputDir) {
  console.log('\n--- Starting Analyst Price Targets Fetch ---');
  // Logic from "Analyst Price Targets generalized to all stocks.js"
  for (const ticker of tickers) {
    try {
      const result = await yahooFinance.quoteSummary(ticker, { modules: ['financialData'] });
      const fd = result.financialData;

      if (!fd) {
        console.warn(`⚠️ No financial data for ${ticker} in Analyst Price Targets.`);
        continue;
      }

      const data = [
        ['Category', 'Value'], // Header row
        ['Current Price', fd.currentPrice || 'N/A'],
        ['Target High Price', fd.targetHighPrice || 'N/A'],
        ['Target Low Price', fd.targetLowPrice || 'N/A'],
        ['Target Mean Price', fd.targetMeanPrice || 'N/A'],
        ['Analyst Opinions', fd.numberOfAnalystOpinions || 'N/A'],
        ['Recommendation', fd.recommendationKey || 'N/A'],
        ['Recommendation Score', fd.recommendationMean || 'N/A'],
      ];

      const ws = XLSX.utils.aoa_to_sheet(data);
      const fileName = `${ticker}_valuation_measures.xlsx`;
      const filePath = path.join(outputDir, fileName);

      let outWorkbook;
      if (fs.existsSync(filePath)) {
        outWorkbook = XLSX.readFile(filePath);
        // Remove existing sheet if it exists to avoid duplication issues, will be re-added
        if (outWorkbook.SheetNames.includes('Analyst Price Target')) {
          const sheetIndex = outWorkbook.SheetNames.indexOf('Analyst Price Target');
          outWorkbook.SheetNames.splice(sheetIndex, 1);
          delete outWorkbook.Sheets['Analyst Price Target'];
         }
      } else {
        outWorkbook = XLSX.utils.book_new();
      }

      XLSX.utils.book_append_sheet(outWorkbook, ws, 'Analyst Price Target');
      XLSX.writeFile(outWorkbook, filePath);
      console.log(`✅ Analyst Price Target data saved for ${ticker} to ${filePath}`);

    } catch (err) {
      console.error(`❌ Error fetching Analyst Price Targets for ${ticker}:`, err.message);
    }
  }
  console.log('--- Finished Analyst Price Targets Fetch ---');
}

async function fetchAnalystRecommendations(tickers, outputDir) {
  console.log('\n--- Starting Analyst Recommendations Fetch ---');
  // Logic from "Analyst Reccomendation_looping on multiple Stock to create new Excel Sheets.js"
  yahooFinance.suppressNotices(['yahooSurvey']); // Suppress notice

  for (const ticker of tickers) {
    try {
      console.log(`\n⏳ Fetching recommendations for: ${ticker} ...`);
      const result = await yahooFinance.quoteSummary(ticker, { modules: ['recommendationTrend'] });
      const trend = result.recommendationTrend?.trend;

      if (!trend || !Array.isArray(trend) || trend.length === 0) {
        console.log(`⚠️ No recommendation data for ${ticker}. Skipping.`);
        continue;
      }

      const data = trend.map(item => ({
        Period: parsePeriod(item.period),
        StrongBuy: item.strongBuy,
        Buy: item.buy,
        Hold: item.hold,
        Sell: item.sell,
        StrongSell: item.strongSell
      }));

      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'RecommendationTrend');

      const filePath = path.join(outputDir, `recommendation_trend_${ticker}.xlsx`);
      XLSX.writeFile(wb, filePath);

      console.log(`✅ Saved recommendation trend for ${ticker} to: ${filePath}`);
    } catch (innerErr) {
      console.error(`❌ Error processing recommendations for ${ticker}:`, innerErr.message);
    }
  }
  console.log('--- Finished Analyst Recommendations Fetch ---');
}

async function fetchKeyStatistics(tickers, outputDir) {
  console.log('\n--- Starting Key Statistics Fetch ---');
  // Logic from "Statistics to XLS.js"
  const browser = await puppeteer.launch({ headless: true, args: ['--lang=fr-FR'] }); // Consider making headless configurable

  for (const ticker of tickers) {
    const page = await browser.newPage();
    try {
      console.log(`\n🌐 Scraping key statistics for ${ticker}...`);
      const url = `https://finance.yahoo.com/quote/${ticker}/key-statistics?p=${ticker}`;
      await page.goto(url, { waitUntil: 'networkidle2', timeout: 60000 }); // Increased timeout

      // Handle potential pop-ups (consent, scroll buttons)
      try {
        const acceptButtonSelector = 'button.accept-all, button[name="agree"], button[value="agree"]'; // More generic selectors
        await page.waitForSelector(acceptButtonSelector, { timeout: 10000 });
        await page.click(acceptButtonSelector);
        console.log(`Clicked consent button for ${ticker}`);
      } catch (e) {
        // console.log(`No consent button found or needed for ${ticker}`);
      }
      
      try {
        // Sometimes a scroll down button appears
        await page.waitForSelector('#scroll-down-btn', { timeout: 5000 });
        await page.click('#scroll-down-btn');
         console.log(`Clicked scroll-down button for ${ticker}`);
      } catch (e) {
        // console.log(`No scroll-down button found for ${ticker}`);
      }

      // Wait for the main statistics table
      const mainTableSelector = 'table.table.yf-kbx2lo, div[data-test="key-stats-table"] table'; // More generic selector for table
      await page.waitForSelector(mainTableSelector, { timeout: 30000 });

      const { data: newData, columns: newColumns, metrics: newMetrics } = await page.evaluate((tableSelector) => {
        const table = document.querySelector(tableSelector);
        if (!table) return { error: 'Valuation table not found on page.' };

        const headers = Array.from(table.querySelectorAll('thead th'))
          .slice(1) // Skip the first header (usually "Metric" or empty)
          .map(th => th.innerText.trim());

        const rows = Array.from(table.querySelectorAll('tbody tr'));
        const resultData = {};
        const resultMetrics = [];

        rows.forEach(row => {
          const cells = Array.from(row.querySelectorAll('td'));
          if (cells.length < 2) return; // Ensure there's a label and at least one value
          
          const labelElement = cells[0].querySelector('span') || cells[0]; // Sometimes the label is wrapped in a span
          const label = labelElement.innerText.trim();
          
          const values = cells.slice(1).map(td => td.innerText.trim());
          resultData[label] = values;
          resultMetrics.push(label);
        });

        return { data: resultData, columns: headers, metrics: resultMetrics };
      }, mainTableSelector);


      if (!newData || Object.keys(newData).length === 0) {
        console.warn(`⚠️ No key statistics data extracted for ${ticker}. It might be due to page structure changes or no data available.`);
        continue;
      }
      
      const outputPath = path.join(outputDir, `${ticker}_valuation_measures.xlsx`);
      let outWorkbook;
      let sheetData = [];
      const sheetNameValuation = 'Valuation';

      let allColumns = ['Metric', ...newColumns]; // Start with Metric and new columns

      if (fs.existsSync(outputPath)) {
        outWorkbook = XLSX.readFile(outputPath);
        if (outWorkbook.SheetNames.includes(sheetNameValuation)) {
          const existingSheet = outWorkbook.Sheets[sheetNameValuation];
          const existingRows = XLSX.utils.sheet_to_json(existingSheet);
          
          // Get all columns from existing sheet, maintaining their order
          if (existingRows.length > 0) {
            allColumns = Object.keys(existingRows[0]); 
            // Ensure 'Metric' is first if it exists, otherwise add it
            if (!allColumns.includes('Metric')) {
                allColumns.unshift('Metric');
            } else if (allColumns.indexOf('Metric') > 0) {
                allColumns = ['Metric', ...allColumns.filter(c => c !== 'Metric')];
            }
             // Add new columns from scraping if they aren't already there
            newColumns.forEach(col => {
                if (!allColumns.includes(col)) allColumns.push(col);
            });
          }


          const existingDataMap = new Map();
          existingRows.forEach(row => existingDataMap.set(row.Metric, row));

          newMetrics.forEach(metric => {
            const row = existingDataMap.get(metric) || { Metric: metric };
            newColumns.forEach((col, i) => {
              row[col] = newData[metric][i];
            });
            existingDataMap.set(metric, row);
          });
          sheetData = Array.from(existingDataMap.values());

           // Remove the old "Valuation" sheet before adding the updated one
           if (outWorkbook.SheetNames.includes(sheetNameValuation)) {
            const sheetIndex = outWorkbook.SheetNames.indexOf(sheetNameValuation);
            outWorkbook.SheetNames.splice(sheetIndex, 1);
            delete outWorkbook.Sheets[sheetNameValuation];
           }

        } else { // File exists, but "Valuation" sheet doesn't
           newMetrics.forEach(metric => {
            const row = { Metric: metric };
            newColumns.forEach((col, i) => {
              row[col] = newData[metric][i];
            });
            sheetData.push(row);
          });
        }
      } else { // File does not exist
        outWorkbook = XLSX.utils.book_new();
        newMetrics.forEach(metric => {
          const row = { Metric: metric };
          newColumns.forEach((col, i) => {
            row[col] = newData[metric][i];
          });
          sheetData.push(row);
        });
      }
      
      // Sort data by metric name for consistency
      sheetData.sort((a, b) => (a.Metric || "").localeCompare(b.Metric || ""));

      const newSheet = XLSX.utils.json_to_sheet(sheetData, { header: allColumns });
      XLSX.utils.book_append_sheet(outWorkbook, newSheet, sheetNameValuation);
      XLSX.writeFile(outWorkbook, outputPath);

      console.log(`✅ Key statistics saved/updated for ${ticker} in ${outputPath}`);

    } catch (error) {
      console.error(`❌ Error scraping key statistics for ${ticker}:`, error.message, error.stack);
      if (page && !page.isClosed()) {
        // Save a screenshot for debugging if page is available
        const screenshotPath = path.join(outputDir, `${ticker}_error_screenshot.png`);
        try {
            await page.screenshot({ path: screenshotPath, fullPage: true });
            console.log(`📸 Screenshot saved to ${screenshotPath} for ${ticker}`);
        } catch (ssError) {
            console.error(`Could not save screenshot for ${ticker}: ${ssError.message}`);
        }
      }
    } finally {
      if (page && !page.isClosed()) {
        await page.close();
      }
    }
  }
  await browser.close();
  console.log('--- Finished Key Statistics Fetch ---');
}


// Main execution function
async function runAll() {
  try {
    console.log('🚀 Starting combined stock script...');
    ensureOutputDir(outputFolder); // Ensure output directory exists

    const tickers = await getTickers(excelPath);

    if (!tickers || tickers.length === 0) {
      console.log('No tickers found. Exiting script.');
      return;
    }

    // Execute functions sequentially
    await fetchAnalystPriceTargets(tickers, outputFolder);
    await fetchAnalystRecommendations(tickers, outputFolder);
    await fetchKeyStatistics(tickers, outputFolder);

    console.log('\n🎉 All tasks completed successfully!');

  } catch (err) {
    console.error('❌ A fatal error occurred in the main execution:', err.message, err.stack);
  }
}

// Run the script
runAll();
