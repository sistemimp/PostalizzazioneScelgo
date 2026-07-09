const XLSX = require('xlsx');

function getWorkbookInfo(inputPath) {
  const workbook = XLSX.readFile(inputPath, { cellDates: true });
  const sheetNames = workbook.SheetNames;

  const sheets = sheetNames.map((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    const headers = rows.length > 0 ? Object.keys(rows[0]) : [];

    return {
      name: sheetName,
      headers,
      rowCount: rows.length
    };
  });

  return {
    sheetNames,
    sheets,
    totalSheets: sheetNames.length
  };
}

module.exports = {
  getWorkbookInfo
};
