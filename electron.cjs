const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const fs = require("fs");
const { fileURLToPath } = require("url");

function createWindow() {
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(path.dirname(__filename), "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  win.loadFile(path.join(path.dirname(__filename), "dist", "index.html"));
}

ipcMain.on("save-excel-files", async (event, archivos) => {
  const folderPath = dialog.showOpenDialogSync({
    properties: ["openDirectory"],
    title: "Selecciona una carpeta para guardar los archivos Excel",
  });

  if (!folderPath || !folderPath[0]) return;

  archivos.forEach((archivo) => {
    const fullPath = path.join(folderPath[0], archivo.nombre);
    fs.writeFileSync(fullPath, archivo.buffer);
  });

  event.sender.send("save-excel-files-done");
});

ipcMain.on("save-excel-file", async (event, archivos) => {
  const folderPath = dialog.showOpenDialogSync({
    properties: ["openDirectory"],
    title: "Selecciona una carpeta para guardar el archivo Excel",
  });

  if (!folderPath || !folderPath[0]) return;

  archivos.forEach((archivo) => {
    const fullPath = path.join(folderPath[0], archivo.nombre);
    fs.writeFileSync(fullPath, archivo.buffer);
  });

  event.sender.send("save-excel-file-done");
});

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
