//Read input file and to sort data, converting to json 
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

//function that generates the random color
function getRandomColor() {
    var letters = '0123456789ABCDEF';
    var color = '';
    for (var i = 0; i < 6; i++) {
        color += letters[Math.floor(Math.random() * 16)];
    }
    return color;
}

//to assign distinct color to each Genre
function getDistinctValues() {
    var distinctGenres = [];
    for (var i = 0; i < data.length; i++) {
        if (!distinctGenres.hasOwnProperty(data[i]["Genre"])) {
            distinctGenres[data[i]["Genre"]] = getRandomColor();
        }
    }
    return distinctGenres;
}

//write the output to new file called "out"
var book = XLSX.utils.book_new();
var worksheet = XLSX.utils.json_to_sheet(data); //converting the sorted json back to sheet
XLSX.utils.book_append_sheet(book, worksheet, 'Sheet2');
XLSX.writeFile(book, 'Output.xlsx');

//reading the newly generated "out" file
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

workbook.xlsx.readFile('Output.xlsx')
    .then(function () {
        var worksheet = workbook.getWorksheet('Sheet2');

        //to fetch the header row and set the header style
        var row = worksheet.getRow(1); 
        for (i = 1; i <= row.cellCount; i++) {
            row.getCell(i).style = {
                fill:
                {
                    type: 'pattern',
                    pattern: 'mediumGray',
                    bgColor: { argb: 'F1C40F' },
                    fgColor: { argb: 'F1C40F' },
                },
                font:
                {
                    bold: true,
                },
                alignment:
                {
                    horizontal: 'center',
                },
                border: {
                    right: { style: 'thin', color: { argb: '17202A' } }
                },
            };


        }
        row.commit();

        //align right if the cell value is a number
        for (i = 2; i <= worksheet.rowCount; i++) {
            for (j = 1; j <= worksheet.getRow(i).cellCount; j++) {
                var reg = new RegExp('^[0-9]+$');
                if (reg.test(worksheet.getRow(i).getCell(j).value) == true) {
                    worksheet.getRow(i).getCell(j).style = {
                        alignment:
                        {
                            horizontal: 'right'
                        },
                    }
                    //converting text to number to remove green triangles in excel sheet
                    worksheet.getRow(i).getCell(j).value = Number(worksheet.getRow(i).getCell(j).value);
                }
            }
        }

        // to get the distinct color codes for different Genre types and color the corresponding rows 
        var arr = getDistinctValues();
        for (i = 2; i <= worksheet.rowCount; i++) {
            for (j = 1; j <= worksheet.getRow(i).cellCount; j++) {
                var cellValue = worksheet.getRow(i).getCell(j).value;
                if (arr.hasOwnProperty(cellValue)) {
                    var fillColor = arr[cellValue];
                    var row = worksheet.getRow(i);
                    for (k = 1; k <= row.cellCount; k++) {
                        row.getCell(k).style = {
                            fill:
                            {
                                type: 'pattern',
                                pattern: 'mediumGray',
                                bgColor: { argb: fillColor },
                                fgColor: { argb: fillColor },
                            },
                            border: {
                                right: { style: 'thin', color: { argb: '17202A' } },
                                bottom: { style: 'thin', color: { argb: '17202A' } }
                            },
                        }
                    }
                }
            }
        }

        //setting up header width for each column
        worksheet.columns = [
            { header: 'SNO', key: 'sno', width: 10 },
            { header: 'Album Name', key: 'album_name', width: 32 },
            { header: 'Genre', key: 'genre', width: 10 },
            { header: 'Artist', key: 'artist', width: 20 },
            { header: 'Release Date', key: 'release_date', width: 15 },
            { header: 'Critic Score', key: 'critic_score', width: 10 }
        ];

        //update the "out" file with header styling changes
        return workbook.xlsx.writeFile('Output.xlsx');
    });

