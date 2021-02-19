const fs = require("fs");
const pdf = require('pdf-parse');
const Excel = require('exceljs');
 
let dataBuffer = fs.readFileSync("/Users/williamwalker/Developer/SWM/afi36-2903.pdf");

pdf(dataBuffer).then(function (data) {
    result = data.text.replace(/\n[^\d+\.]/g, '')
    result = result.split(/(\d\..*)\s/)

    var filtered = result.filter(function (el) {
        return el != '';
      });

    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");

    worksheet.columns = [
    {header: 'Description', key: 'description', width: 70},
    {header: 'ShallWillMust', key: 'shallwillmust', width: 30}
    ];
    
    worksheet.getColumn(1).values = filtered;
    worksheet.getColumn(1).alignment = { wrapText: true };

    const shallwillmust = worksheet.getColumn(2);
    shallwillmust.eachCell(function(cell, rowNumber) {
        cell.value = { formula:
        '=CONCATENATE(IF(IFERROR(FIND(" must ",A' + rowNumber 
        + '),0)>0,"Must",""),IF(IFERROR(FIND(" will ",A' + rowNumber
        + '),0)>0,"Will",""),IF(IFERROR(FIND(" shall ",A' + rowNumber
        + '),0)>0,"Shall",""))'
        }
    });

    worksheet.insertRow(1, ['Description', 'Shall/Will/Must']);
    worksheet.views = [
        {state: 'frozen', xSplit: 0, ySplit: 1}
      ];
    workbook.xlsx.writeFile('export.xlsx');
});

