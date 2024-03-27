import * as React from "react";
import ExcelJS from "exceljs";

import { saveAs } from '../../utils/saveAs'
import { fetchFile } from '../../utils/fetchFile'
import {
    addColumnWidthByList,
    addRowHeightByRange,
    addDefaultStyle,
    autoFitColumn,
    addCustomCell
} from './functions'
import { defaultStyle } from './constants/sheetStyles'

import { table01Columns, table01MockData } from './constants/mocks'
import {
    dateFormat,
    decimalFormat,
    moneyFormat,
    percentageFormat
} from './constants/format'

// Colunas de valor monetario
const dateColumns = ['E', 'W', 'X'];
const moneyColumns = ['H', 'J', 'K', 'M', 'O', 'P', 'AA', 'AC', 'AG'];
const decimalColumns = ['I', 'AB', 'AE', 'AF'];

const table01Name = 'Table01';

const CreateSheet = () => {
    const data = [table01MockData]
    const dataListSize = data.length;

    const exportExcelFile = async () => {
        const workbook = new ExcelJS.Workbook();
        
        const sheet = workbook.addWorksheet("Visão Geral", {
            properties: {
                showGridLines: true,
            },
            views: [{ showGridLines: false, zoomScale: 80 }],

        });
        sheet.unprotect();

        /** 
         * Adiciona a logo na planilha
         * Linhas 1, 2 e 3 / Celulas: A1:A3
         * */
        const addLogo = async () => {
            // Get image of "public/assets" paste
            const imageFileUrl = `${window.location.origin}/assets/locX.png`;
            const imageFile = await fetchFile({ fileUrl: imageFileUrl })
            // Add image in workbook
            const reactLogo = workbook.addImage({
                buffer: imageFile,
                extension: 'png',
            });
            
            sheet.mergeCells('A1:B3');
            sheet.getCell('A1').fill = defaultStyle.fill;
            // Add image in "A1:B3" cell
            sheet.addImage(reactLogo, {
                /** Comportamento
                 * `undefined` -  Especifica que a imagem será movida e dimensionada com células
                 * `oneCell` - Esta é a predefinição. A imagem será movida com as células mas não será dimensionada
                 * `absolute` - A imagem não será movida ou dimensionada com as células
                 */
                editAs: 'undefined',
                // Posicionamento
                tl: { col: 0.6, row: 1 },
                // Dimensionamento dinamico
                br: { col: 1.99999999, row: 2 },
                // Dimensionamento fixo
                // ext: { width: 100, height: 50 }
            });
        }
        await addLogo()

        /** 
         * Cria um resumo de informaçoes
         * Linhas 1, 2 e 3 
         * */
        const addSummaryOfInformation = () => {
            const table01HeaderRowNumber = 4;
            const startDataCellPosition = table01HeaderRowNumber + 1;
            const endDataCellPosition = table01HeaderRowNumber + dataListSize;

            sheet.getCell('C1').value = 'Saving Total';
            sheet.getCell('C2').value = 'Remuneração LocX';
            sheet.getCell('C3').value = 'Saving Liquido';

            addCustomCell({ sheet, cellAddress: 'D1', valueFormula: 'M3', numFmt: moneyFormat })
            addCustomCell({ sheet, cellAddress: 'D2', valueFormula: 'AG3', numFmt: moneyFormat })
            addCustomCell({ sheet, cellAddress: 'D3', valueFormula: 'D1-D2', numFmt: moneyFormat })
            addCustomCell({ sheet, cellAddress: 'E2', valueFormula: 'D2/D1', numFmt: percentageFormat })
            addCustomCell({ sheet, cellAddress: 'E3', valueFormula: 'D3/D1', numFmt: percentageFormat })
            addCustomCell({ sheet, cellAddress: 'H3', valueFormula: `H${startDataCellPosition}:H${endDataCellPosition}`, numFmt: moneyFormat })
            addCustomCell({ sheet, cellAddress: 'J3', valueFormula: `J${startDataCellPosition}:J${endDataCellPosition}`, numFmt: moneyFormat })
            addCustomCell({ sheet, cellAddress: 'K3', valueFormula: `K${startDataCellPosition}:K${endDataCellPosition}`, numFmt: moneyFormat })
            addCustomCell({ sheet, cellAddress: 'M3', valueFormula: `M${startDataCellPosition}:M${endDataCellPosition}`, numFmt: moneyFormat })
            addCustomCell({ sheet, cellAddress: 'O3', valueFormula: `O${startDataCellPosition}:O${endDataCellPosition}`, numFmt: moneyFormat })
            addCustomCell({ sheet, cellAddress: 'P3', valueFormula: `P${startDataCellPosition}:P${endDataCellPosition}`, numFmt: moneyFormat })
            sheet.getCell('R3').value = 'Isenção de reajuste';
            sheet.mergeCells('R3:U3');
            sheet.getCell('Z3').value = 36;
            addCustomCell({ sheet, cellAddress: 'AA3', valueFormula: `AA${startDataCellPosition}:AA${endDataCellPosition}`, numFmt: moneyFormat })
            addCustomCell({ sheet, cellAddress: 'AB3', value: 12, numFmt: decimalFormat })
            addCustomCell({ sheet, cellAddress: 'AC3', valueFormula: `AC${startDataCellPosition}:AC${endDataCellPosition}`, numFmt: moneyFormat })
            sheet.getCell('AD3').value = 36;
            addCustomCell({ sheet, cellAddress: 'AE3', valueFormula: `AE${startDataCellPosition}:AE${endDataCellPosition}`, numFmt: moneyFormat })
            addCustomCell({ sheet, cellAddress: 'AF3', value: 0.08, numFmt: percentageFormat })
            addCustomCell({ sheet, cellAddress: 'AG3', valueFormula: `AG${startDataCellPosition}:AG${endDataCellPosition}`, numFmt: moneyFormat })
        }
        addSummaryOfInformation();

        /**
         * Cria uma tabela e adiciona dados abaixo delas
         * Linha 4 para baixo 
         * */
        const addTable01 = () => {
            const table01 = sheet.addTable({
                name: table01Name,
                ref: 'A4',
                headerRow: true,
                style: {
                    theme: null,
                },
                columns: table01Columns.map(value => ({ name: value })),
                rows: []
            })
            // adiciona dados na tabela
            data.forEach(rowData => {
                table01.addRow(rowData)
            })
            table01.commit();

            // DADOS
            const startRowData = 5;
            // Obtem intervalo de linhas 
            const rows = sheet.getRows(startRowData, dataListSize)
            // Itera sobre o intervalo de linhas e aplica as formulas
            rows.forEach((row) => {
                const columnsToAddFormatDate = dateColumns.map(col => col + row.number);
                const columnsToAddFormatMoney = moneyColumns.map(col => col + row.number);
                const columnsToAddFormatDecimal = decimalColumns.map(col => col + row.number);

                row.eachCell({ includeEmpty: true }, (cell) => {
                    if (columnsToAddFormatDate.includes(cell.address)) {
                        addCustomCell({ cellAddress: cell.address, numFmt: dateFormat })
                    }
                    if (columnsToAddFormatMoney.includes(cell.address)) {
                        addCustomCell({ cellAddress: cell.address, numFmt: moneyFormat })
                    }
                    if (columnsToAddFormatDecimal.includes(cell.address)) {
                        addCustomCell({ cellAddress: cell.address, numFmt: decimalFormat })
                    }
                })
            });
        }
        addTable01();

        /** 
         * Altera as dimençoes de linhas e colunas 
         **/
        addRowHeightByRange({ sheet, start: 1, end: 3, height: 35 });
        addRowHeightByRange({ sheet, start: 4, end: 4, height: 30 });
        addRowHeightByRange({ sheet, start: 5, end: sheet.rowCount, height: 25 });
        
        autoFitColumn({ sheet });

        addColumnWidthByList({ sheet, width: 30, columns: ['C'] });

        addDefaultStyle({ sheet })

        workbook.xlsx.writeBuffer().then(function (data) {
            const blob = new Blob([data], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            });

            saveAs({ blob, fileName: "created_template_download.xlsx" });
        })
    }

    return (
        <div style={{ padding: "30px" }}>
            <button
                className="btn btn-primary mt-2 mb-2"
                onClick={exportExcelFile}
            >
                CREATE AND EXPORT XLSX
            </button>
        </div>
    );
};

export default CreateSheet;