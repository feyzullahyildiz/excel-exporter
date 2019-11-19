import * as ExcelJS from 'exceljs';
import * as fs  from 'fs'
const filename = './test.xlsx'
const COLOR_OBJ_GRAY = { argb: 'c1c1c1' }
if (fs.existsSync(filename)) {
    fs.unlinkSync(filename)
}

function rotateCell(cell: ExcelJS.Cell) {

}
function setCellGray(cell: ExcelJS.Cell) {

}
// const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(excelFileOptions);
const workbook = new ExcelJS.Workbook();
const writeToExcel = fs.createWriteStream(filename);

// create a sheet with red tab colour
// var sheet = workbook.addWorksheet('My Sheet', { properties: { tabColor: { argb: 'FFC0000' } } });

// create a sheet where the grid lines are hidden
const worksheet = workbook.addWorksheet('My Sheet', { properties: { showGridLines: false } });

// create a sheet with the first row and column frozen
// var sheet = workbook.addWorksheet('My Sheet', { views: [{ xSplit: 1, ySplit: 1 }] });

// const columns = [
//     { header: 'İsim', key: 'name' },
//     { header: 'Soyisim', key: 'surname' },
// ]

// sheet.columns = columns;

// const row = sheet.addRow({name: 'Feyz', surname: 'YILDIZ'})
// row.commit()
const columnA = worksheet.getColumn('A')
columnA.width = 4
worksheet.mergeCells('A8:A15');
const cellA8 = worksheet.getCell('A8');
cellA8.font = {
    color: COLOR_OBJ_GRAY
}
cellA8.fill = {
    type: 'pattern',
    pattern: 'darkVertical',
    fgColor: { argb: 'FFFF0000' }
}
cellA8.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
}
cellA8.alignment = {
    textRotation: 90,
    // textRotation: 'vertical',
    // vertical: 'justify',
    // indent: 1,
    vertical: 'middle',
    horizontal: 'centerContinuous',
};
cellA8.value = 'SABAH22sasdfasf'
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
// workbook.commit().then(a => {
//     console.log('BITTI')
// })