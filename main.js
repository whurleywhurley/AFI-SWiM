// Modules to control application life and create native browser window
const {app, BrowserWindow, dialog, ipcMain} = require('electron')
const path = require('path')
const pdf = require('pdf-parse')
const excel = require('exceljs')

function createWindow () {
  // Create the main program window.
  const mainWindow = new BrowserWindow({
    width: 600,
    height: 400,
    webPreferences: {
      nodeIntegration: true
    }
  })

  // and load the index.html of the app.
  mainWindow.loadFile('index.html')

  // don't allow the window to be resized
  // and remove the menu
  mainWindow.setResizable(false)
  mainWindow.removeMenu()
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.whenReady().then(createWindow)

// Quit when all windows are closed
app.on('window-all-closed', function () {
    app.quit()
})

// Convert the AFI into plaintext and ask the user where to save
function convertAFI(status, dataBuffer) {
  console.log("Converting AFI to Excel...")
  pdf(dataBuffer)
  .then(function (data) {

    // clean the text by removing all line breaks before paragraph numberings
    cleanedText = data.text.replace(/\n[^\d+\.]/g, '')

    // deliminate the text by newline breaks
    segmentedText = cleanedText.split(/\n/)

    // find the first paragraph of the AFI
    var found = false;
    var index = 0;
    while (!found) {
        if (segmentedText[index].includes("1.1.  ")) {
            found = true
        } else {
            index++;
        }
    }

    // remove everything before paragraph 1.1. of the AFI
    segmentedText.splice(0, index)

    // pass off the result to the exceljs handler
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
    status.reply('upload-reply', 'Done')
    dialog.showSaveDialog({ 
        title: 'Select the File Path to save', 
        defaultPath: path.join(__dirname, '../assets/workbook.xlsx'), 
        buttonLabel: 'Save', 
        filters: [ 
            { 
                name: 'Excel Workbook', 
                extensions: ['xlsx'] 
            }, ], 
        properties: [] 
    }).then(file => { 
        console.log(file.filePath.toString()); 
        workbook.xlsx.writeFile(file.filePath.toString());
      })
      .then(() => {
        status.reply('upload-reply', 'Reset')
      });
    });
}

ipcMain.on('upload', (event, arg) => {
  convertAFI(event, arg)
});