import * as ExcelJS from 'exceljs';

export const getNextColumn = (col: string) => {
    const val = col.toUpperCase();
    if (val.length !== 1) {
        throw new Error('next column not found')
    }
    if (val === 'Z') {
        throw new Error('next column not found')
    }
    const num = val.charCodeAt(0)
    return String.fromCharCode(num + 1)
}
console.assert(getNextColumn('A') === 'B', 'HATA OLDU')
console.assert(getNextColumn('a') === 'B', 'HATA OLDU')
console.assert(getNextColumn('B') === 'C', 'HATA OLDU')
// console.assert(getNextColumn('Z') === 'AA', 'HATA OLDU')
enum Color {
    GRAY = 'C0C0C0',
    BLACK = '000000'
}
export function rotateCell(cell: ExcelJS.Cell) {
    cell.alignment = {
        textRotation: 90,
        // textRotation: 'vertical',
        // vertical: 'justify',
        // indent: 1,
        vertical: 'middle',
        horizontal: 'centerContinuous',
    };
}
export function getRefCell(cell: ExcelJS.Cell, x: number, y: number) {
    const sheet = cell.worksheet;
    const [row, col] = getCellPosition(cell)
    return sheet.getCell(row + y, col + x);
}
export function setCellFontCenter(cell: ExcelJS.Cell) {
    if (cell.alignment) {
        cell.alignment.vertical = 'middle';
        cell.alignment.horizontal = 'center';

    } else {
        cell.alignment = {
            vertical: 'middle',
            horizontal: 'center',
        }
    }
}
export function setCellFontGray(cell: ExcelJS.Cell) {
    if(cell.font) {
        cell.font.color = { argb: Color.BLACK }
    } else {
        cell.font = {
            color: { argb: Color.BLACK }
        }
    }
}
export function setCellFontSelectedFont(cell: ExcelJS.Cell) {
    if(cell.font) {
        cell.font.name = 'Times New Roman'
    } else {
        cell.font = {
            name: 'Times New Roman'
        }
    }
}
type BorderSide = 'top' | 'left' | 'bottom' | 'right'
export function addDarkBorder(cell: ExcelJS.Cell, ...sides: BorderSide[]) {
    if (cell.border) {
        for (const side of sides) {
            cell.border[side] = { style: 'medium' }
        }

    } else {
        cell.border = {}
        for (const side of sides) {
            cell.border[side] = { style: 'medium' }

        }
    }
}
export function addDefaultDarkBorder(cell: ExcelJS.Cell) {
    cell.border = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        bottom: { style: 'medium' },
        right: { style: 'medium' }
    }
}
export function setCellBackgroundGray(cell: ExcelJS.Cell) {
    cell.fill = {
        type: 'pattern',
        pattern: 'darkVertical',
        fgColor: { argb: Color.GRAY }
    }
}
export function setCellFontItalic(cell: ExcelJS.Cell) {
    if (cell.font) {
        cell.font.italic = true
    } else {
        cell.font = {
            italic: true,
        }
    }
}
export function setCellFontBold(cell: ExcelJS.Cell) {
    if (cell.font) {
        cell.font.bold = true
    } else {
        cell.font = {
            bold: true
        }
    }
}

interface VehicleRaportItem {
    subHeaderCount: number,
    header: {
        displayName: string,
        key: string
    },
    subHeaders: {
        displayName: string,
        key: string
    }[],
    rows: { [dynamicSubHeaderKey: string]: number }[]
}
interface DataModal {
    times: { start: string, end: string }[],
    displayName: string,
    data: VehicleRaportItem[],
    rowCount: number,
}

export const demoData: DataModal = {
    rowCount: 8,
    displayName: 'SABAH',
    times: [
        { start: '7:00', end: '7:15' },
        { start: '7:15', end: '7:30' },
        { start: '7:30', end: '7:45' },
        { start: '7:45', end: '8:00' },
        { start: '8:00', end: '8:15' },
        { start: '8:15', end: '8:30' },
        { start: '8:30', end: '8:45' },
        { start: '8:45', end: '9:00' },
    ],
    data: [
        {
            header: {
                displayName: 'OTOMOBİL',
                key: 'otomobil',
            },
            subHeaderCount: 4,
            subHeaders: [
                { displayName: 'U', key: 'u' },
                { displayName: '1-2', key: '2' },
                { displayName: '1-3', key: '3' },
                { displayName: '1-4', key: '4' },
            ],
            rows: [
                { 'u': 0, '2': 131, '3': 28, '4': 18 },
                { 'u': 0, '2': 116, '3': 50, '4': 23 },
                { 'u': 0, '2': 156, '3': 60, '4': 29 },
                { 'u': 0, '2': 185, '3': 74, '4': 43 },
                { 'u': 0, '2': 215, '3': 73, '4': 40 },
                { 'u': 0, '2': 214, '3': 87, '4': 43 },
                { 'u': 0, '2': 125, '3': 93, '4': 31 },
                { 'u': 0, '2': 137, '3': 103, '4': 31 },
            ],
        },
        {
            header: {
                displayName: 'KAMYONET',
                key: 'kamyonet',
            },
            subHeaderCount: 4,
            subHeaders: [
                { displayName: 'U', key: 'u' },
                { displayName: '1-2', key: '2' },
                { displayName: '1-3', key: '3' },
                { displayName: '1-4', key: '4' },
            ],
            rows: [
                { 'u': 0, '2': 9, '3': 10, '4': 1 },
                { 'u': 0, '2': 7, '3': 8, '4': 0 },
                { 'u': 0, '2': 10, '3': 14, '4': 2 },
                { 'u': 0, '2': 8, '3': 10, '4': 3 },
                { 'u': 0, '2': 9, '3': 3, '4': 1 },
                { 'u': 0, '2': 11, '3': 5, '4': 3 },
                { 'u': 0, '2': 8, '3': 9, '4': 3 },
                { 'u': 0, '2': 11, '3': 10, '4': 5 },
            ],
        },
        {
            header: {
                displayName: 'TAKSİ',
                key: 'taksi',
            },
            subHeaderCount: 4,
            subHeaders: [
                { displayName: 'U', key: 'u' },
                { displayName: '1-2', key: '2' },
                { displayName: '1-3', key: '3' },
                { displayName: '1-4', key: '4' },
            ],
            rows: [
                { 'u': 0, '2': 4, '3': 7, '4': 1 },
                { 'u': 0, '2': 5, '3': 11, '4': 2 },
                { 'u': 0, '2': 8, '3': 6, '4': 2 },
                { 'u': 0, '2': 25, '3': 13, '4': 2 },
                { 'u': 0, '2': 13, '3': 13, '4': 0 },
                { 'u': 0, '2': 17, '3': 8, '4': 0 },
                { 'u': 0, '2': 10, '3': 3, '4': 2 },
                { 'u': 0, '2': 13, '3': 8, '4': 0 },
            ],
        },
        {
            header: {
                displayName: 'T.MİNİBÜS',
                key: 'tminibus',
            },
            subHeaderCount: 4,
            subHeaders: [
                { displayName: 'U', key: 'u' },
                { displayName: '1-2', key: '2' },
                { displayName: '1-3', key: '3' },
                { displayName: '1-4', key: '4' },
            ],
            rows: [
                { 'u': 0, '2': 4, '3': 7, '4': 1 },
                { 'u': 0, '2': 5, '3': 11, '4': 2 },
                { 'u': 0, '2': 8, '3': 6, '4': 2 },
                { 'u': 0, '2': 25, '3': 13, '4': 2 },
                { 'u': 0, '2': 13, '3': 13, '4': 0 },
                { 'u': 0, '2': 17, '3': 8, '4': 0 },
                { 'u': 0, '2': 10, '3': 3, '4': 2 },
                { 'u': 0, '2': 13, '3': 8, '4': 0 },
            ],
        },
        {
            header: {
                displayName: 'T.MİNİBÜS',
                key: 'tminibus',
            },
            subHeaderCount: 4,
            subHeaders: [
                { displayName: 'U', key: 'u' },
                { displayName: '1-2', key: '2' },
                { displayName: '1-3', key: '3' },
                { displayName: '1-4', key: '4' },
            ],
            rows: [
                { 'u': 0, '2': 4, '3': 7, '4': 1 },
                { 'u': 0, '2': 5, '3': 11, '4': 2 },
                { 'u': 0, '2': 8, '3': 6, '4': 2 },
                { 'u': 0, '2': 25, '3': 13, '4': 2 },
                { 'u': 0, '2': 13, '3': 13, '4': 0 },
                { 'u': 0, '2': 17, '3': 8, '4': 0 },
                { 'u': 0, '2': 10, '3': 3, '4': 2 },
                { 'u': 0, '2': 13, '3': 8, '4': 0 },
            ],
        },
    ]
}
const getCellPosition = (cell: ExcelJS.Cell): number[] => {
    const row: number = cell.fullAddress.row as any;
    const col: number = cell.fullAddress.col as any;
    return [row, col]
}
const setDayPeriodCell = (dayPeriodCell: ExcelJS.Cell, dataModal: DataModal) => {
    const sheet = dayPeriodCell.worksheet;
    const [row, col] = getCellPosition(dayPeriodCell);
    const endOfDayPeriodCell = sheet.getCell(row + dataModal.rowCount - 1, col);
    sheet.mergeCells(dayPeriodCell.address, endOfDayPeriodCell.address)
    setCellFontGray(dayPeriodCell);
    setCellBackgroundGray(dayPeriodCell)
    addDefaultDarkBorder(dayPeriodCell)
    rotateCell(dayPeriodCell)
    setCellFontBold(dayPeriodCell)
    dayPeriodCell.value = dataModal.displayName
}
const setColWidthHeighValues = (startCell: ExcelJS.Cell) => {
    const sheet = startCell.worksheet;
    const [row, col] = getCellPosition(startCell);
    sheet.getColumn(col).width = 3;

    sheet.getColumn(col + 1).width = 5;
    sheet.getColumn(col + 2).width = 5;

    sheet.getRow(row).height = 25
}

const setTimeValues = (startCell: ExcelJS.Cell, dataModal: DataModal) => {
    const sheet = startCell.worksheet
    // const [sRow, sCol] = getCellPosition(startCell);
    // const row = sRow + 2
    // const col = sCol + 1
    const [row, col] = getCellPosition(startCell);
    dataModal.times.forEach((time, index) => {

        const timeStartRow = row + index
        const timeStartCol = col
        const startCell = sheet.getCell(timeStartRow, timeStartCol)
        startCell.value = time.start
        setCellFontCenter(startCell)

        const timeEndRow = row + index
        const timeEndCol = col + 1
        const endCell = sheet.getCell(timeEndRow, timeEndCol);
        setCellFontCenter(endCell)
        endCell.value = time.end
    })

}
const drawVehicleRaportItem = (startCell: ExcelJS.Cell, item: VehicleRaportItem) => {
    const sheet = startCell.worksheet;
    const [row, col] = getCellPosition(startCell);
    const titleEndCell = sheet.getCell(row, col + item.subHeaderCount - 1);
    sheet.mergeCells(startCell.address, titleEndCell.address)
    startCell.value = item.header.displayName
    setCellFontCenter(startCell)
    addDefaultDarkBorder(startCell)
    const getColIndexOfCol = (key: string): number => {
        return item.subHeaders.findIndex(a => a.key === key)
    }
    item.subHeaders.forEach((header, index) => {
        const cell = getRefCell(startCell, index, 1);
        const columnObject = sheet.getColumn(cell.col)
        columnObject.width = 5;
        cell.value = header.displayName
        setCellFontCenter(cell);
        addDefaultDarkBorder(cell);

        const countResultStartCell = getRefCell(startCell, index, 1);
        const countResultEndCell = getRefCell(startCell, index, item.rows.length + 1);
        const totalResultCell = getRefCell(startCell, index, item.rows.length + 2);
        setCellFontCenter(totalResultCell)
        addDefaultDarkBorder(totalResultCell)
        totalResultCell.value = {
            formula: `SUM(${countResultStartCell.address}:${countResultEndCell.address})`,
            date1904: true
        }
        setCellBackgroundGray(totalResultCell);
        setCellFontItalic(totalResultCell)
    })
    // Her Saate karşılık gelen data bunun içinde
    item.rows.forEach((row, rowIndex) => {
        const keys = Object.keys(row)
        for (const key of keys) {
            const index = getColIndexOfCol(key)
            if (index !== -1) {
                const cell = getRefCell(startCell, index, rowIndex + 2)
                setCellFontCenter(cell)
                cell.value = row[key];
                if (index === 0) {
                    addDarkBorder(cell, 'left')
                } else if (index === item.subHeaderCount - 1) {
                    addDarkBorder(cell, 'right')
                }
            }
        }
    });
}
const setTimeCellHeader = (startCell: ExcelJS.Cell) => {
    const sheet = startCell.worksheet;
    const [row, col] = getCellPosition(startCell);
    const titleCell = sheet.getCell(row, col + 1)
    const endOfTitleCell = sheet.getCell(row + 1, col + 2)
    sheet.mergeCells(titleCell.address, endOfTitleCell.address)

    setCellFontGray(titleCell);
    addDefaultDarkBorder(titleCell)
    setCellFontBold(titleCell)
    setCellFontCenter(titleCell)
    titleCell.value = `ÇEKİM\nSAATİ`
}

const setTotalCellFooter = (startCell: ExcelJS.Cell) => {
    const sheet = startCell.worksheet;
    const [row, col] = getCellPosition(startCell);
    const totalCell = sheet.getCell(row, col)
    const endOfTotalCell = getRefCell(totalCell, 1, 0);
    sheet.mergeCells(totalCell.address, endOfTotalCell.address);
    addDefaultDarkBorder(totalCell);
    setCellBackgroundGray(totalCell)
    setCellFontItalic(totalCell)
    setCellFontBold(totalCell)
    setCellFontCenter(totalCell)
    totalCell.value = 'TOPLAM';
}
const setWorkSheetFont = (sheet: ExcelJS.Worksheet) => {
    let i = 0;
    sheet.eachRow(row => {
        row.eachCell(cell => {
            if(cell.value !== undefined) {
                setCellFontSelectedFont(cell)
                i++;
            }
        })
    })
    console.log('setCellFontSelectedFont CELL', i)
}
export const drawDataModal = (sheet: ExcelJS.Worksheet, startCell: ExcelJS.Cell, dataModal: DataModal) => {
    const [row, col] = getCellPosition(startCell);
    setColWidthHeighValues(startCell);
    setDayPeriodCell(getRefCell(startCell, 0, 2), dataModal)
    setTimeCellHeader(startCell)
    setTimeValues(getRefCell(startCell, 1, 2), dataModal)
    setTotalCellFooter(getRefCell(startCell, 1, dataModal.rowCount + 2))
    const drawVehicleRaportItemStartCell = getRefCell(startCell, 3, 0);
    let xOffsetValue = 0;
    dataModal.data.forEach((vehicleModal) => {
        const refCell = getRefCell(drawVehicleRaportItemStartCell, xOffsetValue, 0)
        drawVehicleRaportItem(refCell, vehicleModal)
        xOffsetValue += vehicleModal.subHeaderCount
    })
    setWorkSheetFont(sheet)
    // const mergedTimeCellValue = `${col}${row}`
    // // sheet.mergeCells()

    console.log('Row', row)

}