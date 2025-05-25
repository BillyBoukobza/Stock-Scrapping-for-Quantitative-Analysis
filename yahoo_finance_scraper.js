const yahooFinance = require('yahoo-finance2').default;
const XLSX = require('xlsx');
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const pLimit = require('p-limit');
const delay = require('delay');

const CONCURRENCY_LIMIT = 5;
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 3000;

const excelPath = path.join('C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long', 'IBKR API Python - Stock Info - v2.xlsx');
const outputFolder = path.join('C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long', 'XLSX');

function ensureOutputDir(dirPath) {
  if (!fs.existsSync(dirPath)) fs.mkdirSync(dirPath, { recursive: true });
}

function getTickers(filePath) {
  const wb = XLSX.readFile(filePath);
  const sheetName = wb.SheetNames[0];
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName])
    .map(row => row.Ticker)
    .filter(Boolean);
}

async function withRetry(fn, description = '') {
  for (let i = 0; i < MAX_RETRIES; i++) {
    try {
      return await fn();
    } catch (err) {
      console.warn(`⚠️ ${description} failed (try ${i + 1}/${MAX_RETRIES}): ${err.message}`);
      if (i < MAX_RETRIES - 1) await delay(RETRY_DELAY_MS);
    }
  }
  throw new Error(`❌ ${description} failed after ${MAX_RETRIES} retries`);
}

async function fetchAllPriceTargets(tickers) {
  const limit = pLimit(CONCURRENCY_LIMIT);
  await Promise.all(tickers.map(ticker => limit(() => withRetry(
    () => fetchPriceTarget(ticker),
    `PriceTarget for ${ticker}`
  ))));
}

async function fetchPriceTarget(ticker) {
  const result = await yahooFinance.quoteSummary(ticker, { modules: ['financialData'] });
  const fd = result.financialData;
  if (!fd) return;

  const data = [
    ['Category', 'Value'],
    ['Current Price', fd.currentPrice || 'N/A'],
    ['Target High Price', fd.targetHighPrice || 'N/A'],
    ['Target Low Price', fd.targetLowPrice || 'N/A'],
    ['Target Mean Price', fd.targetMeanPrice || 'N/A'],
    ['Analyst Opinions', fd.numberOfAnalystOpinions || 'N/A'],
    ['Recommendation', fd.recommendationKey || 'N/A'],
    ['Recommendation Score', fd.recommendationMean || 'N/A'],
  ];

  writeToExcel(ticker, 'Analyst Price Target', data);
}

async function fetchAllKeyStats(tickers) {
  const browser = await puppeteer.launch({ headless: true });
  const limit = pLimit(CONCURRENCY_LIMIT);

  await Promise.all(tickers.map(ticker => limit(() => withRetry(
    () => fetchKeyStats(ticker, browser),
    `Key Stats for ${ticker}`
  ))));

  await browser.close();
}

async function fetchKeyStats(ticker, browser) {
  const page = await browser.newPage();
  try {
    const url = `https://finance.yahoo.com/quote/${ticker}/key-statistics?p=${ticker}`;
    await page.goto(url, { waitUntil: 'networkidle2', timeout: 60000 });

    const mainTableSelector = 'table.table.yf-kbx2lo, div[data-test="key-stats-table"] table';
    await page.waitForSelector(mainTableSelector, { timeout: 30000 });

    const tableData = await page.evaluate((sel) => {
      const table = document.querySelector(sel);
      if (!table) return null;

      const rows = Array.from(table.querySelectorAll('tr'));
      return rows.map(row => {
        const tds = row.querySelectorAll('td');
        if (tds.length < 2) return null;
        return [tds[0].innerText.trim(), tds[1].innerText.trim()];
      }).filter(Boolean);
    }, mainTableSelector);

    if (tableData && tableData.length) {
      writeToExcel(ticker, 'Valuation', [['Metric', 'Value'], ...tableData]);
    } else {
      console.warn(`⚠️ No key stats found for ${ticker}`);
    }

  } finally {
    await page.close();
  }
}

function writeToExcel(ticker, sheetName, data) {
  const filePath = path.join(outputFolder, `${ticker}_valuation_measures.xlsx`);
  let wb = fs.existsSync(filePath) ? XLSX.readFile(filePath) : XLSX.utils.book_new();

  if (wb.SheetNames.includes(sheetName)) {
    const idx = wb.SheetNames.indexOf(sheetName);
    wb.SheetNames.splice(idx, 1);
    delete wb.Sheets[sheetName];
  }

  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filePath);
}

(async () => {
  console.log('🚀 Starting enhanced script...');
  ensureOutputDir(outputFolder);

  const tickers = getTickers(excelPath);
  if (tickers.length === 0) return console.error('❌ No tickers found.');

  await fetchAllPriceTargets(tickers);
  await fetchAllKeyStats(tickers);

  console.log('✅ All done.');
})();