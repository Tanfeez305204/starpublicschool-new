const Excel = require('exceljs');
const path = require('path');
const file = path.join(__dirname, 'results.xlsx');
(async () => {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(file);
  console.log(workbook.worksheets.map(ws => ws.name).join(', '));
})();
