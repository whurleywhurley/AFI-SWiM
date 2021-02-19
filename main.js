// Modules to control application life and create native browser window
const {app, BrowserWindow, ipcMain, dialog} = require('electron')
const path = require('path'); 
const fs = require("fs");
const pdf = require('pdf-parse');
const Excel = require('exceljs');

function createWindow () {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true
    }
  })

  // and load the index.html of the app.
  mainWindow.loadFile('index.html')

  // Open the DevTools.
  //mainWindow.webContents.openDevTools()
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.whenReady().then(createWindow)

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on('window-all-closed', function () {
    app.quit()
})

app.on('activate', function () {
  // On OS X it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and require them here.

// Convert the AFI into plaintext and store into the global afiText variable
function convertAFI(dataBuffer) {
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
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");

    worksheet.columns = [
    {header: 'Description', key: 'description', width: 70},
    {header: 'ShallWillMust', key: 'shallwillmust', width: 30}
    ];
    
    worksheet.getColumn(1).values = result;
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
    worksheet.views = [{state: 'frozen', xSplit: 0, ySplit: 1}];

    return workbook;
  })
  .then(workbook =>{
    
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
      });
    });
}

ipcMain.on('upload', (event, arg) => {
  convertAFI(arg)
});