const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');

if (require('electron-squirrel-startup')) app.quit();

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        }
    });

    mainWindow.loadFile(path.join(__dirname, 'index.html'));

    mainWindow.on('closed', function () {
        mainWindow = null;
    });

    mainWindow.webContents.openDevTools();

}

app.on('ready', createWindow);

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', function () {
    if (mainWindow === null) {
        createWindow();
    }
});

ipcMain.handle('show-save-dialog', async (event) => {
  const window = BrowserWindow.fromWebContents(event.sender);
  const options = {
    title: '합쳐진 데이터 저장',
    defaultPath: app.getPath('downloads') + '/filtered_data.xlsx',
    filters: [
      { name: 'Excel', extensions: ['xlsx'] }
    ]
  };

  const { filePath } = await dialog.showSaveDialog(window, options);
  return filePath;  // This will be undefined if the user cancels the dialog
});

ipcMain.handle('show-save-unique-dialog', async (event) => {
  const window = BrowserWindow.fromWebContents(event.sender);
  const options = {
    title: '하나뿐 이었던 데이터 저장',
    defaultPath: app.getPath('downloads') + '/unique_data.xlsx',
    filters: [
      { name: 'Excel', extensions: ['xlsx'] }
    ]
  };

  const { filePath } = await dialog.showSaveDialog(window, options);
  return filePath;  // This will be undefined if the user cancels the dialog
});

ipcMain.handle('show-save-duplicated-dialog', async (event) => {
  const window = BrowserWindow.fromWebContents(event.sender);
  const options = {
    title: '중복된 데이터 저장',
    defaultPath: app.getPath('downloads') + '/duplicated_data.xlsx',
    filters: [
      { name: 'Excel', extensions: ['xlsx'] }
    ]
  };

  const { filePath } = await dialog.showSaveDialog(window, options);
  return filePath;  // This will be undefined if the user cancels the dialog
});


// For development purposes
require('electron-reload')(__dirname);

