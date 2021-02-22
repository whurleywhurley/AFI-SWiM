const pdf = require('pdf-parse')
const excel = require('exceljs')

console.log("Converting AFI to Excel...")
pdf("/Users/williamwalker/Developer/AFISWM/afi36-2903.pdf")
.then(function (data) {
    cleanedText = data.text.replace(/\n[^\d+\.]/g, '')
    segmentedText = cleanedText.split(/\n/)

    var found = false;
    var index = 0;
    while (!found) {
        if (segmentedText[index].includes("1.1.  ")) {
            found = true
        } else {
            index++;
        }
    }

    segmentedText.splice(0, 77)

    return segmentedText
})
.then(result => {
// Create our worksheet and set it up
const workbook = new excel.Workbook();
const worksheet = workbook.addWorksheet("AFI Export");
worksheet.columns = [
{header: 'Description', key: 'description', width: 70},
{header: 'Shall/Will/Must', key: 'shallwillmust', width: 30},
{header: 'Tier Level', key: 'tierlevel', width: 30}
];

// Set up the description column from the PDF export
worksheet.getColumn(1).values = result;
worksheet.getColumn(1).alignment = { wrapText: true };

// Set up the SWM column by iterating over the rows, and then create the filter
const shallwillmust = worksheet.getColumn(2);
shallwillmust.eachCell(function(cell, rowNumber) {
    cell.value = { formula:
    '=CONCATENATE(IF(IFERROR(FIND(" shall ",A' + rowNumber 
    + '),0)>0,"Shall",""),IF(IFERROR(FIND(" will ",A' + rowNumber
    + '),0)>0,"Will",""),IF(IFERROR(FIND(" must ",A' + rowNumber
    + '),0)>0,"Must",""))'
    }
});

// Set up the Tier column by iterating over the rows, and then create the filter
const tier = worksheet.getColumn(3);
tier.eachCell(function(cell, rowNumber) {
    cell.value = { formula:
    '=CONCATENATE(IF(IFERROR(FIND("T-0",A' + rowNumber 
    + '),0)>0,"T-0",""),IF(IFERROR(FIND("T-1",A' + rowNumber
    + '),0)>0,"T-1",""),IF(IFERROR(FIND("T-2",A' + rowNumber
    + '),0)>0,"T-2",""),IF(IFERROR(FIND("T-3",A' + rowNumber
    + '),0)>0,"T-3",""))'
    }
});


// Insert the header row,, style, then freeze it
worksheet.getColumn('description').header = "Description"
worksheet.getColumn('shallwillmust').header = "Shall/Will/Must"
worksheet.getColumn('tierlevel').header = "Tier Level"
worksheet.getRow(1).font = {bold: true};
worksheet.views = [{state: 'frozen', xSplit: 0, ySplit: 1}];
worksheet.autoFilter = 'B1:C1';

// return the workbook
return workbook;
})
.then(workbook =>{
    workbook.xlsx.writeFile('test.xlsx');
});