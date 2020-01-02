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


const workbook = new ExcelJS.Workbook();
const writeToExcel = fs.createWriteStream(filename);

const worksheet = workbook.addWorksheet('1', { properties: { showGridLines: false } });
worksheet.getCell('A1').value = (Math.random() * 100).toFixed(0)
worksheet.views = [
    { state: 'frozen', xSplit: 3, ySplit: 7, activeCell: 'A1' }
]

drawDataModal(worksheet, worksheet.getCell('A6'), demoData);
drawDataModal(worksheet, worksheet.getCell('A24'), Object.assign(demoData, {displayName: 'AKŞAM'}));

workbook.xlsx.write(writeToExcel).then(a => {
    console.log('finitto')
})