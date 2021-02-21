// This file is required by the index.html file and will
// be executed in the renderer process for that window.
// All of the Node.js APIs are available in this process.

const { ipcRenderer } = require('electron')

document.addEventListener('drop', (event) => { 
    event.preventDefault(); 
    event.stopPropagation(); 
  
    for (const f of event.dataTransfer.files) { 
        // Using the path attribute to get absolute file path 
        console.log('File Path of dragged files: ', f.path)
        ipcRenderer.send('upload', f.path);
        document.getElementById("status").innerHTML = "Converting..."
        ipcRenderer.on('upload-reply', (event, arg) => {
            if (arg === 'Done') document.getElementById("status").innerHTML = "Done"
            if (arg === 'Reset') document.getElementById("status").innerHTML = "Drag AFI PDF Here"
        });
    } 
}); 
  
document.addEventListener('dragover', (e) => { 
    e.preventDefault(); 
    e.stopPropagation(); 
  }); 
  
document.addEventListener('dragenter', (event) => { 
    console.log('File is in the Drop Space'); 
}); 
  
document.addEventListener('dragleave', (event) => { 
    console.log('File has left the Drop Space'); 
}); 
