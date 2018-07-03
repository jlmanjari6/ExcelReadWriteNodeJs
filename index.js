
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

workbook.xlsx.readFile('Input.xlsx')
    .then(function () {
        var worksheet = workbook.getWorksheet('Sheet2');
        var row = worksheet.getRow(1);
        for (i = 1; i <= row.cellCount; i++) {
            //set header style(background, bold font)
            row.getCell(i).style = {
                fill:
                {
                    type: 'pattern',
                    pattern: 'mediumGray',
                    bgColor: { argb: 'F1C40F' },
                    fgColor: { argb: 'F1C40F' }
                },
                font:
                {
                    bold: true,
                },
                alignment:
                {
                    horizontal: 'center'
                },
                border: {
                    right: { style: 'thin', color: { argb: '17202A' } }
                }
            };
        }
        row.commit();
        return workbook.xlsx.writeFile('new.xlsx');
    })