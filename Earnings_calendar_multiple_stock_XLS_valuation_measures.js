const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const inputExcelPath = path.join(
  'C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long',
  'IBKR API Python - Stock Info - v2.xlsx'
);

const outputFolder = path.join(
  'C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long', 'XLSX'
);

function getTickersFromExcel(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  return data.map(row => row.Ticker).filter(Boolean);
}

async function scrapeEarningsData(page, ticker) {
  const baseUrl = `https://fr.finance.yahoo.com/calendar/earnings?symbol=${ticker}&size=100`;
  let allData = [];

  for (const offset of [0, 100]) {
    await page.goto(`${baseUrl}&offset=${offset}`, { waitUntil: 'domcontentloaded' });

    try {
      await page.waitForSelector('button[name="agree"]', { timeout: 3000 });
      await page.click('button[name="agree"]');
      await page.waitForTimeout(1000);
    } catch {}

    try {
      await page.waitForSelector('#scroll-down-btn', { timeout: 3000 });
      await page.click('#scroll-down-btn');
      await page.waitForTimeout(2000);
    } catch {}

    try {
      await page.waitForSelector('tr.row', { timeout: 5000 });
    } catch (err) {
      console.log(`⛔ No data found for ${ticker} at offset ${offset}`);
      continue;
    }

    const data = await page.evaluate(() => {
      const months = {
        janvier: '01', février: '02', mars: '03', avril: '04',
        mai: '05', juin: '06', juillet: '07', août: '08',
        septembre: '09', octobre: '10', novembre: '11', décembre: '12'
      };

      function parseFrenchDate(str) {
        const regex = /(\d{1,2}) (\w+) (\d{4}) à (\d{1,2}) h UTC([+\-−]\d+)/;
        const match = str.match(regex);
        if (!match) return { date: null, time: null, UTC: null };

        let [, day, monthFr, year, hour, offset] = match;
        const month = months[monthFr.toLowerCase()];
        const paddedDay = day.padStart(2, '0');
        const paddedHour = hour.padStart(2, '0');

        offset = offset.replace('−', '-'); // Remplace tiret unicode par tiret standard
        const utcOffset = parseInt(offset, 10);

        return {
          date: `${year}-${month}-${paddedDay}`,
          time: `${paddedHour}:00:00`,
          UTC: utcOffset
        };
      }

      return Array.from(document.querySelectorAll('tr.row')).map(row => {
        const cells = row.querySelectorAll('td');
        const rawDate = cells[2]?.innerText.trim() || null;
        const parsedDate = rawDate ? parseFrenchDate(rawDate) : { date: null, time: null, UTC: null };

        return {
          ticker: cells[0]?.innerText.trim() || null,
          company: cells[1]?.innerText.trim() || null,
          date: parsedDate.date,
          time: parsedDate.time,
          UTC: parsedDate.UTC,
          estimateEPS: cells[3]?.innerText.trim() || null,
          reportedEPS: cells[4]?.innerText.trim() || null,
          surprise: cells[5]?.innerText.trim() || null
        };
      });
    });

    allData = allData.concat(data);
  }

  return allData;
}

(async () => {
  const tickers = getTickersFromExcel(inputExcelPath);
  const browser = await puppeteer.launch({ headless: true });

  for (const ticker of tickers) {
    const page = await browser.newPage();
    try {
      console.log(`🔍 Scraping earnings for ${ticker}`);
      const earningsData = await scrapeEarningsData(page, ticker);

      if (earningsData.length === 0) {
        console.log(`⚠️ No earnings data found for ${ticker}`);
        continue;
      }

      const filePath = path.join(outputFolder, `${ticker}_valuation_measures.xlsx`);
      let workbook = fs.existsSync(filePath) ? XLSX.readFile(filePath) : XLSX.utils.book_new();

      const safeSheetName = `earnings_${new Date().toISOString().split('T')[0]}`;
      const worksheet = XLSX.utils.json_to_sheet(earningsData);
      XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName);
      XLSX.writeFile(workbook, filePath);

      console.log(`✅ Added earnings sheet for ${ticker} in ${filePath}`);
    } catch (err) {
      console.error(`❌ Error processing ${ticker}: ${err.message}`);
    } finally {
      await page.close();
    }
  }

  await browser.close();
})();