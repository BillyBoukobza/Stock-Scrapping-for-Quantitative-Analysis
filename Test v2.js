const yahooFinance = require('yahoo-finance2').default;
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Chemins à adapter
const excelPath = path.join(
  'C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long',
  'IBKR API Python - Stock Info - v2.xlsx'
);

const outputDir = path.join(
  'C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long', 'XLSX'
);

// Création du dossier de sortie s'il n'existe pas
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

function parsePeriod(period) {
  // Ici tu peux améliorer la fonction pour parser la période si besoin
  return period;
}

async function main() {
  try {
    const workbook = xlsx.readFile(excelPath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(worksheet);

    const tickers = rows.map(row => row.Ticker).filter(Boolean);
    console.log(`🔍 Found ${tickers.length} tickers.`);

    for (const symbol of tickers) {
      try {
        console.log(`\n⏳ Fetching Earnings Estimate for ${symbol} ...`);

        const result = await yahooFinance.quoteSummary(symbol, { modules: ['earningsTrend'] });
        const earningsTrend = result.earningsTrend;

        if (!earningsTrend || !earningsTrend.trend) {
          console.log(`⚠️ No earningsTrend data for ${symbol}. Skipping.`);
          continue;
        }

        const data = earningsTrend.trend.map(item => ({
          Period: parsePeriod(item.period),
          Growth: item.growth,
          EarningsAvg: item.earningsEstimate?.avg,
          EarningsLow: item.earningsEstimate?.low,
          EarningsHigh: item.earningsEstimate?.high,
          RevenueAvg: item.revenueEstimate?.avg,
          RevenueLow: item.revenueEstimate?.low,
          RevenueHigh: item.revenueEstimate?.high,
        }));

        const ws = xlsx.utils.json_to_sheet(data);
        const wb = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(wb, ws, 'EarningsEstimate');

        const filePath = path.join(outputDir, `earnings_estimate_${symbol}.xlsx`);
        xlsx.writeFile(wb, filePath);

        console.log(`✅ Saved: ${filePath}`);
      } catch (innerErr) {
        console.error(`❌ Error processing ${symbol}:`, innerErr.message);
      }
    }

    console.log('\n🎉 All done.');
  } catch (err) {
    console.error('❌ Fatal error:', err.message);
  }
}

main();