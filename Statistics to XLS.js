const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// 📥 Lire le fichier Excel contenant les tickers
const excelPath = path.join(
  'C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long',
  'IBKR API Python - Stock Info - v2.xlsx'
);
const workbook = XLSX.readFile(excelPath);
const sheetName = workbook.SheetNames[0];
const tickers = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]).map(row => row.Ticker);

(async () => {
  const browser = await puppeteer.launch({ headless: true, args: ['--lang=fr-FR'] });

  for (const ticker of tickers) {
    const page = await browser.newPage();
    try {
      console.log(`🌐 Scraping data for ${ticker}...`);
      const url = `https://finance.yahoo.com/quote/${ticker}/key-statistics?p=${ticker}`;
      await page.goto(url, { waitUntil: 'networkidle2', timeout: 30000 });

      try {
        await page.waitForSelector('#scroll-down-btn', { timeout: 5000 });
        await page.click('#scroll-down-btn');
      } catch {}

      try {
        await page.waitForSelector('button.accept-all', { timeout: 10000 });
        await page.click('button.accept-all');
      } catch {}

      await page.waitForSelector('table.table.yf-kbx2lo', { timeout: 15000 });

      const { data: newData, columns: newColumns, metrics: newMetrics } = await page.evaluate(() => {
        const table = document.querySelector('table.table.yf-kbx2lo');
        if (!table) return { error: 'Valuation table not found.' };

        const headers = Array.from(table.querySelectorAll('thead th'))
          .slice(1)
          .map(th => th.innerText.trim());

        const rows = Array.from(table.querySelectorAll('tbody tr'));
        const result = {};
        const metrics = [];

        rows.forEach(row => {
          const cells = Array.from(row.querySelectorAll('td'));
          if (cells.length < 2) return;
          const label = cells[0].innerText.trim();
          const values = cells.slice(1).map(td => td.innerText.trim());
          result[label] = values;
          metrics.push(label);
        });

        return { data: result, columns: headers, metrics };
      });

      if (!newData || !newColumns || !newMetrics) {
        throw new Error(`Failed to extract table data for ${ticker}`);
      }

      const outputDir = path.resolve('C:/Users/billy/OneDrive/Bureau/Finances/Investissements/Stocks/Long/XLSX');
      const outputPath = path.join(outputDir, `${ticker}_valuation_measures.xlsx`);

      if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
      }

      let oldData = {};
      let allColumns = ['Metric'];

      if (fs.existsSync(outputPath)) {
        const existingWorkbook = XLSX.readFile(outputPath);
        const existingSheet = existingWorkbook.Sheets[existingWorkbook.SheetNames[0]];
        const existingRows = XLSX.utils.sheet_to_json(existingSheet);
        existingRows.forEach(row => {
          const metric = row['Metric'];
          if (metric) oldData[metric] = row;
        });
        allColumns = Object.keys(existingRows[0] || {}).filter(col => col !== '');
        if (!allColumns.includes('Metric')) allColumns.unshift('Metric');
      }

      newColumns.forEach(col => {
        if (!allColumns.includes(col)) allColumns.push(col);
      });

      const finalData = {};

      newMetrics.forEach(metric => {
        const existingRow = oldData[metric] || { Metric: metric };
        const mergedRow = { ...existingRow };
        newColumns.forEach((col, i) => {
          mergedRow[col] = newData[metric][i];
        });
        finalData[metric] = mergedRow;
      });

      Object.keys(oldData).forEach(metric => {
        if (!finalData[metric]) {
          finalData[metric] = oldData[metric];
        }
      });

      const finalRows = Object.values(finalData).sort((a, b) => a.Metric.localeCompare(b.Metric));

      const newSheet = XLSX.utils.json_to_sheet(finalRows, { header: allColumns });
      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Valuation');

      XLSX.writeFile(newWorkbook, outputPath);

      console.log(`✅ Saved XLSX for ${ticker}: ${outputPath}`);
    } catch (error) {
      console.error(`❌ Error for ${ticker}:`, error.message);
    } finally {
      await page.close();
    }
  }

  await browser.close();
  console.log('👋 All scraping done.');
})();