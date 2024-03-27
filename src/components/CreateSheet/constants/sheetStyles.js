const borderDefaultStyle = {
    style: 'thin',
    color: {
        argb:'FFCCCCCC'
    }
}

export const defaultStyle = {
    alignment: {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: false
    },
    fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF7F4F2' }
    },
    font: {
        bold: true,
        name: 'Calibri',
        size: 10,
        color: { argb: 'FF5B5D74' }
    },
    border: {
        left: borderDefaultStyle,
        top: borderDefaultStyle,
        right: borderDefaultStyle,
        bottom: borderDefaultStyle
    }
}
