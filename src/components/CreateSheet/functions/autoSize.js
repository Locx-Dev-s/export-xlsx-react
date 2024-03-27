// Link da soluçao: https://stackoverflow.com/a/64097746
export const autoFitColumn = ({ sheet }) => {
    sheet.columns.forEach((column) => {
        let maxLength = 6;

        column.eachCell({ includeEmpty: true }, (cell) => {
            const columnLength = cell.value ? cell.value.toString().length + 3 : 10;
            if (columnLength > maxLength) {
                maxLength = columnLength + 3;
            }
        });

        column.width = maxLength < 10 ? 10 : maxLength;
    });
};