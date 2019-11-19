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
export function setCellFontGray(cell: ExcelJS.Cell) {
    cell.font = {
        color: { argb: Color.BLACK }
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

export const drawDataModal = (sheet: ExcelJS.Worksheet, startCell: ExcelJS.Cell, dataModal: DataModal) => {
    const { row, col } = startCell.fullAddress;
    const titleCell = sheet.getCell()
    setCellFontGray(titleCell);
    setCellBackgroundGray(titleCell)
    addDefaultDarkBorder(titleCell)
    rotateCell(titleCell)
    setCellFontBold(titleCell)
    const mergedTimeCellValue = `${col}${row}`
    // // sheet.mergeCells()
    // console.log('Row', row)

}