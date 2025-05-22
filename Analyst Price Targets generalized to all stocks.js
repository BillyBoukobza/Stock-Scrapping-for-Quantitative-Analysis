const yahooFinance = require('yahoo-finance2').default;
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Chemin vers le fichier source des tickers
const excelPath = path.join(
  'C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long',
  'IBKR API Python - Stock Info - v2.xlsx'
);

// Dossier de sortie
const outputFolder = path.join(
  'C:', 'Users', 'billy', 'OneDrive', 'Bureau', 'Finances', 'Investissements', 'Stocks', 'Long', 'XLSX'
);

(async () => {
  try {
    // Lire les tickers depuis la feuille source
    const wb = XLSX.readFile(excelPath);
    const sheetName = wb.SheetNames[0];
    const tickers = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]).map(row => row.Ticker).filter(Boolean);

    for (const ticker of tickers) {
      try {
        const result = await yahooFinance.quoteSummary(ticker, { modules: ['financialData'] });

        const fd = result.financialData;
        if (!fd) {
          console.warn(`Pas de données financières pour ${ticker}`);
          continue;
        }

        // Préparer les données à écrire
        const data = [
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
        const filePath = path.join(outputFolder, fileName);

        let outWorkbook;
        if (fs.existsSync(filePath)) {
          // Charger l'existant
          outWorkbook = XLSX.readFile(filePath);
          // Supprimer l'ancienne feuille si elle existe
          if (outWorkbook.SheetNames.includes('Analyst Price Target')) {
            delete outWorkbook.Sheets['Analyst Price Target'];
            outWorkbook.SheetNames = outWorkbook.SheetNames.filter(name => name !== 'Analyst Price Target');
          }
        } else {
          // Nouveau fichier si inexistant
          outWorkbook = XLSX.utils.book_new();
        }

        // Ajouter la nouvelle feuille
        XLSX.utils.book_append_sheet(outWorkbook, ws, 'Analyst Price Target');

        // Sauvegarder
        XLSX.writeFile(outWorkbook, filePath);
        console.log(`✅ Données analystes enregistrées pour ${ticker}`);
      } catch (err) {
        console.error(`❌ Erreur avec le ticker ${ticker}:`, err.message);
      }
    }
  } catch (err) {
    console.error('❌ Erreur générale :', err.message);
  }
})();