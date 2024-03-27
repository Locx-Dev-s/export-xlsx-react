export const addColumnWidthByList = ({ sheet, columns = [], width }) => {
    if (sheet) {
        columns.forEach(colId => {
            const column = sheet.getColumn(colId);
            column.width = width;
        })
    }
}

export const addRowHeightByRange = ({ sheet, start = 0, end, height }) => {
    if (sheet) {
        for (let rowId = start; rowId <= end; rowId += 1) {
            const row = sheet.getRow(rowId);
            row.height = height;
        }
    }
}

export const addCustomCell = ({ sheet, cellAddress, value, valueFormula, numFmt }) => {
    if (sheet && cellAddress) {
        const cell = sheet.getCell(cellAddress);
        if (value) {
            cell.value = value;
        } else if (valueFormula) {
            cell.value = { formula: valueFormula };
        }
        if (numFmt) {
            cell.numFmt = numFmt;
        }
    }
}
