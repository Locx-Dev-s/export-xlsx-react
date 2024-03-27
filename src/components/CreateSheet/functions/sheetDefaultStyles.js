import { defaultStyle } from '../constants/sheetStyles'
/**
 * Adiciona modelo de estilo padrao na tabela
 * 
 * Obs: Este metodo deve ser executado após a planilha for completamente preenchida,
 * afim de aplicar os estilos apenas onde possua dados
 */
export const addDefaultStyle = ({ sheet }) => {
    sheet.eachRow((_, rowNumber) => {
        sheet.getRow(rowNumber).eachCell({ includeEmpty: true }, (cell_v, colNumber) => {
            if (colNumber <= sheet.columnCount) {
                // Informações resumidas
                if (rowNumber <= 3 && cell_v.value) {
                    cell_v.alignment = defaultStyle.alignment;
                    cell_v.fill = defaultStyle.fill;
                    cell_v.font = {
                        ...defaultStyle.font,
                        size: 18
                    };
                    cell_v.border = defaultStyle.border;
                }
                // Cabeçalhos da tabela
                if (rowNumber === 4) {
                    cell_v.alignment = defaultStyle.alignment;
                    cell_v.fill = defaultStyle.fill;
                    cell_v.font = defaultStyle.font;
                    cell_v.border = defaultStyle.border;
                }
                // Dados da tabela
                if (rowNumber > 4) {
                    cell_v.alignment = defaultStyle.alignment;
                    cell_v.fill = {
                        ...defaultStyle.fill,
                        fgColor: { argb: 'FFFFFF' }
                    };
                    cell_v.font = {
                        ...defaultStyle.font,
                        bold: false
                    };
                    cell_v.border = defaultStyle.border;
                }
            }
        })
    })
}
