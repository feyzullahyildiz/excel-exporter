import * as ExcelJS from 'exceljs';
import * as fs from 'fs'
console.log('HOPPA')
import {
    rotateCell,
    addDefaultDarkBorder,
    setCellFontGray,
    setCellBackgroundGray,
    setCellFontBold,
    demoData,
    drawDataModal,
} from './utils';

const filename = './test.xlsx'
if (fs.existsSync(filename)) {
    fs.unlinkSync(filename)
}


// const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(excelFileOptions);
const workbook = new ExcelJS.Workbook();
const writeToExcel = fs.createWriteStream(filename);

// create a sheet with red tab colour
// var sheet = workbook.addWorksheet('My Sheet', { properties: { tabColor: { argb: 'FFC0000' } } });

// create a sheet where the grid lines are hidden
const worksheet = workbook.addWorksheet('My Sheet', { properties: { showGridLines: false } });
worksheet.getCell('A1').value = (Math.random() * 100).toFixed(0)
// create a sheet with the first row and column frozen
// var sheet = workbook.addWorksheet('My Sheet', { views: [{ xSplit: 1, ySplit: 1 }] });

// const columns = [
//     { header: 'İsim', key: 'name' },
//     { header: 'Soyisim', key: 'surname' },
// ]

// sheet.columns = columns;

// const row = sheet.addRow({name: 'Feyz', surname: 'YILDIZ'})
// row.commit()
worksheet.views = [
    { state: 'frozen', xSplit: 3, ySplit: 7, activeCell: 'A1' }
]
// worksheet.mergeCells('A8:A15');

// const columnA = worksheet.getColumn('A')
// columnA.width = 3
// const cellA8 = worksheet.getCell('A8');
// setCellFontGray(cellA8);
// setCellBackgroundGray(cellA8)
// addDefaultDarkBorder(cellA8)
// rotateCell(cellA8)
// setCellFontBold(cellA8)
// cellA8.value = 'SABAH'

drawDataModal(worksheet, worksheet.getCell('A6'), demoData);
drawDataModal(worksheet, worksheet.getCell('A24'), demoData);

// const table = worksheet.addTable({
//     name: 'MyTable',
//     ref: 'B8',
//     headerRow: true,
//     totalsRow: true,
//     style: {
//         theme: 'TableStyleDark3',
//         showRowStripes: true,
//     },
//     columns: [
//         { name: 'Date', totalsRowLabel: 'Totals:', filterButton: true },
//         { name: 'Amount', totalsRowFunction: 'sum', filterButton: false },
//     ],
//     rows: [
//         [new Date('2019-07-20'), 70.10],
//         [new Date('2019-07-21'), 70.60],
//         [new Date('2019-07-22'), 70.10],
//     ],
// });
// sheet.commit();
// workbook.commit();
workbook.xlsx.write(writeToExcel).then(a => {
    console.log('finitto')
})