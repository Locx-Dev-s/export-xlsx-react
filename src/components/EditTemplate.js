import * as React from "react";
import ExcelJS from "exceljs";

import { saveAs } from '../utils/saveAs'
import { fetchFile } from '../utils/fetchFile'

const mockData = {
    installments: [12, 12, 12, 12]
}

const EXCEL_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8"

const EditTemplate = () => {
    const readSheetFile = ({ sheetFile }) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = function () {
                resolve(reader.result);
            }
            reader.onerror = function (error) {
                reject(error);
            }
            reader.readAsArrayBuffer(sheetFile);
        });
    }

    const exportExcelFile = async () => {
        try {
            const workbook = new ExcelJS.Workbook();

            // Url da planilha
            const sheetFileUrl = `${window.location.origin}/assets/modelo02.xlsx`;
            // busca o template via fetch API 
            const sheetFile = await fetchFile({ fileUrl: sheetFileUrl })
            // Cria buffer da planilha informada por parametro
            const arrayBufferSheetFile = await readSheetFile({ sheetFile })
            // Carrega arquivo na classe de Workbook(pasta de trabalho)
            const workbookWithSheetFile = await workbook.xlsx.load(arrayBufferSheetFile)
            // Metodo criado para adicionar um estilo padrão nas celulas de cabeçalho
            const addStylesToHeaderCell = ({ headerCell }) => {
                headerCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'F2EEEB' }
                }
                headerCell.font = {
                    bold: true,
                    name: 'Calibri',
                    size: 10,
                    color: { argb: '5B5D74' }
                }
                headerCell.alignment = {
                    wrapText: true,
                    vertical: 'middle',
                    horizontal: 'left'
                }

                return headerCell
            }
            // Metodo criado para adicionar um estilo padrão nas celulas de dados
            const addStylesToDataCell = ({ dataCell }) => {
                dataCell.alignment = {
                    vertical: 'middle',
                    horizontal: 'left'
                }
                dataCell.font = {
                    name: 'Calibri',
                    size: 10,
                    color: { argb: '5B5D74' }
                }
                dataCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFFFF' }
                }
            }
            // Seleciona uma panilha especifica do template. Neste caso a planilha "Visão geral" 
            const sheetOverview = workbookWithSheetFile.getWorksheet('Visão geral')

            /** 
             * Adiciona a logo na planilha
             * */
            const addLogo = async () => {
                // Get image of "public/assets" paste
                const imageFileUrl = `${window.location.origin}/assets/locx-logo.png`;
                const imageFile = await fetchFile({ fileUrl: imageFileUrl })
                // Add image in workbook
                const reactLogo = workbook.addImage({
                    buffer: imageFile,
                    extension: 'png',
                });

                const cellToAddImage = sheetOverview.getCell('E4:E6')

                cellToAddImage.alignment = {
                    // horizontal: 'center',
                    vertical: 'middle'
                }
                // Add image in "F1" cell
                sheetOverview.addImage(reactLogo, {
                    editAs: 'absolute',
                    // Posicionamento
                    tl: {
                        // acessar uma celula funciona como um array javascrit
                        // col: 0 igual a coluna 1 da planilha(o mesmo serve para row)
                        col: cellToAddImage.col - 1,
                        row: cellToAddImage.row + 0.5,
                    },
                    // dimensoes da imagem
                    ext: { width: 50, height: 30 }
                });
            }

            addStylesToHeaderCell({ headerCell: sheetOverview.getCell('AL7:AM7') })

            /** 
             * Adiciona dados a coluna AL iniciando na linha 7
             * 
             * Cabeçalho: `nº PARCELAS`
             * */
            const addInstallmentsNumber = () => {
                const collumnLetter = 'AL'
                const rowNumber = 7

                const headerCellIndex = `${collumnLetter}${rowNumber}`;

                const headerCell = sheetOverview.getCell(headerCellIndex)
                headerCell.value = 'nº PARCELAS'

                mockData.installments.forEach((value, index) => {
                    const dataCellRowNumber = rowNumber + (index + 1)
                    const dataCellIndex = `${collumnLetter}${dataCellRowNumber}`;

                    const dataCell = sheetOverview.getCell(dataCellIndex)

                    addStylesToDataCell({ dataCell })
                    dataCell.value = value;
                })
            }
            /** 
             * Adiciona dados a coluna AM iniciando na linha 7
             * 
             * Cabeçalho: `R$ PARCELA`
             * */
            const addInstallmentsValue = () => {
                const collumnLetter = 'AM'
                const rowNumber = 7

                const headerCellIndex = `${collumnLetter}${rowNumber}`;

                const headerCell = sheetOverview.getCell(headerCellIndex)
                headerCell.value = 'R$ PARCELA'

                mockData.installments.forEach((_, index) => {
                    const dataCellRowNumber = rowNumber + (index + 1)
                    const dataCellIndex = `${collumnLetter}${dataCellRowNumber}`;

                    const dataCell = sheetOverview.getCell(dataCellIndex)

                    addStylesToDataCell({ dataCell })

                    dataCell.numFmt = 'R$ #,##0.00;[Red]-R$ #,##0.00'
                    const formula = `AK${dataCellRowNumber}/AL${dataCellRowNumber}`;
                    dataCell.value = { formula };
                })
            }

            await addLogo()
            addInstallmentsNumber()
            addInstallmentsValue()

            // Reescreve a planilha com as modificaçoes em um array buffer
            const arrayBufferModified = await workbookWithSheetFile.xlsx.writeBuffer()
            // Converte o array buffer para um blob file
            const blob = new Blob([arrayBufferModified], { type: EXCEL_TYPE });
            // Executa a funçao de Download informando o arquivo(blob) e seu nome
            saveAs({ blob, fileName: "download.xlsx" });
        } catch (error) {
            console.log(error)
        }
    }

    return (
        <button className="btn btn-primary mt-2 mb-2" onClick={exportExcelFile}>
            Export modified Excel
        </button>
    )
};

export default EditTemplate;
