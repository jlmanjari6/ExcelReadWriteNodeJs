//Read input file
var XLSX = require('xlsx');
var workbook = XLSX.readFile('Input.xlsx');
var sheet_name_list = workbook.SheetNames;
var data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

//sort by Genre ascending and then by Critic score descending
function predicateBy(prop1, prop2) {
    return function (a, b) {
        if (a[prop1] > b[prop1]) 
            return 1;
        else if (a[prop1] < b[prop1]) 
            return -1;        
        else if (b[prop2] > a[prop2])
            return 1;
        else if (b[prop2] < a[prop2])
            return -1;
        else
            return 0;
    }
}
data.sort(predicateBy("Genre", "Critic Score"));

//write the output to new file called "out"
var book = XLSX.utils.book_new();
var worksheet = XLSX.utils.json_to_sheet(data);
XLSX.utils.book_append_sheet(book, worksheet, 'Sheet2');
XLSX.writeFile(book, 'out.xlsx');

//reading the newly generated "out" file
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

workbook.xlsx.readFile('out.xlsx')
    .then(function () {
        var worksheet = workbook.getWorksheet('Sheet2');

        var row = worksheet.getRow(1); //Fetch header row
        for (i = 1; i <= row.cellCount; i++) {
            //set header style(background, bold font, center aligned)
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

        //to set the column header values
        worksheet.getColumn(1).header = "SNO";
        worksheet.getColumn(2).header = "Album Name";
        worksheet.getColumn(3).header = "Genre";
        worksheet.getColumn(4).header = "Artist";
        worksheet.getColumn(5).header = "Release Date";
        worksheet.getColumn(6).header = "Critic Score";

        //update the "out" file with header styling changes
        return workbook.xlsx.writeFile('out.xlsx');
    });

