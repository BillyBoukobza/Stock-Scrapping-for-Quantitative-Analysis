const yahooFinance = require('yahoo-finance2').default;
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Suppression du message survey Yahoo
yahooFinance.suppressNotices(['yahooSurvey']);

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

function parsePeriod(offsetStr) {
  const match = offsetStr.match(/(-?\d+)m/);
  if (!match) return offsetStr;

  const monthsAgo = parseInt(match[1], 10);
  const baseDate = new Date();
  const targetDate = new Date(baseDate.getFullYear(), baseDate.getMonth() + monthsAgo, 1);
  return targetDate.toISOString().slice(0, 10);
}

(async () => {
  try {
    // Lecture du fichier Excel source
    const workbook = xlsx.readFile(excelPath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Extraction de toutes les lignes sous forme JSON
    const rows = xlsx.utils.sheet_to_json(worksheet);

    // Récupération des tickers (colonne "Ticker")
    const tickers = rows.map(row => row.Ticker).filter(Boolean);

    console.log(`🔍 Found ${tickers.length} tickers.`);

    for (const symbol of tickers) {
      try {
        console.log(`\n⏳ Fetching data for: ${symbol} ...`);

        const result = await yahooFinance.quoteSummary(symbol, { modules: ['recommendationTrend'] });
        const trend = result.recommendationTrend?.trend;

        if (!trend || !Array.isArray(trend)) {
          console.log(`⚠️ No recommendation data for ${symbol}. Skipping.`);
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

        const ws = xlsx.utils.json_to_sheet(data);
        const wb = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(wb, ws, 'RecommendationTrend');

        const filePath = path.join(outputDir, `recommendation_trend_${symbol}.xlsx`);
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
})();