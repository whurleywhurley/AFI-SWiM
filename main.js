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
  mainWindow.setResizable(false)

  // Open the DevTools.
  //mainWindow.webContents.openDevTools()
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
      result = data.text.replace(/\n[^\d+\.]/g, '')
      result = result.split(/(\d\..*)\s/)
      return result.filter(function (el) {
          return el != '';
      });
  })
  .then(result => {
    // Create our worksheet and set it up
    const workbook = new excel.Workbook();
    const worksheet = workbook.addWorksheet("AFI Export");
    worksheet.columns = [
    {header: 'Description', key: 'description', width: 70},
    {header: 'Shall/Will/Must', key: 'shallwillmust', width: 30}
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

    // Insert the header row,, style, then freeze it
    worksheet.getColumn('description').header = "Description"
    worksheet.getColumn('shallwillmust').header = "Shall/Will/Must"
    worksheet.getRow(1).font = {bold: true};
    worksheet.views = [{state: 'frozen', xSplit: 0, ySplit: 1}];
    worksheet.autoFilter = 'B1:B1';

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