Converter cores Hexadecimais para ARGB
Site: https://www.myfixguide.com/color-converter/

        // const promise = Promise.all(
        //     data?.products?.map(async (product, index) => {
        //         const rowNumber = index + 1;
        //         sheet.addRow({
        //             id: product?.id,
        //             title: product?.title,
        //             brand: product?.brand,
        //             category: product?.category,
        //             price: product?.price,
        //             rating: product?.rating,
        //         });

        //         // const imageFileUrl = product?.thumbnail
        //         const imageFileUrl = `${window.location.origin}/assets/locx-logo.png`;
        //         const imageFile = await fetchFile({ fileUrl: imageFileUrl })

        //         const splitted = imageFileUrl.split(".");
        //         const extName = splitted[splitted.length - 1];

        //         const imageId2 = workbook.addImage({
        //             buffer: imageFile,
        //             extension: extName, 
        //         });

        //         sheet.addImage(imageId2, {
        //             tl: { col: 6, row: rowNumber },
        //             ext: { width: 100, height: 100 },
        //         });
        //     })
        // );

        // promise.then(() => {
        //     const priceCol = sheet.getColumn(5);

        //     // iterate over all current cells in this column
        //     priceCol.eachCell((cell) => {
        //         const cellValue = sheet.getCell(cell?.address).value;
        //         // add a condition to set styling
        //         if (cellValue > 50 && cellValue < 1000) {
        //             sheet.getCell(cell?.address).fill = {
        //                 type: "pattern",
        //                 pattern: "solid",
        //                 fgColor: { argb: "FF0000" },
        //             };
        //         }
        //     });

        //     workbook.xlsx.writeBuffer().then(function (data) {

        //         const blob = new Blob([data], {
        //             type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        //         });

        //         saveAs({
        //             blob,
        //             fileName: "created_template_download.xlsx"
        //         })
        //     });
        // });