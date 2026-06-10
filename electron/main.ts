import { app, BrowserWindow, ipcMain, dialog, Menu, shell } from "electron";
import { join } from "node:path";
import { readFile, writeFile } from "node:fs/promises";

// Bundled to CommonJS for Electron's main process, so __dirname is a builtin.
const DEV_URL = process.env.VITE_DEV_SERVER_URL;

let mainWindow: BrowserWindow | null = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 640,
    minHeight: 400,
    title: "BoringSheets",
    backgroundColor: "#ffffff",
    webPreferences: {
      preload: join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  if (DEV_URL) {
    mainWindow.loadURL(DEV_URL);
  } else {
    mainWindow.loadFile(join(__dirname, "../dist/index.html"));
  }

  // Excel-style: zoom should affect the worksheet only, never the whole UI.
  // Disable Electron's page zoom (pinch + factor) so our grid handles zoom.
  mainWindow.webContents.on("did-finish-load", () => {
    mainWindow?.webContents.setVisualZoomLevelLimits(1, 1); // no trackpad pinch zoom
    mainWindow?.webContents.setZoomFactor(1);
  });
  // If anything still changes the page zoom, snap it back to 100%.
  mainWindow.webContents.on("zoom-changed", () => mainWindow?.webContents.setZoomFactor(1));

  // Prompt to save unsaved changes before closing.
  let forceClose = false;
  mainWindow.on("close", async (e) => {
    if (forceClose || !mainWindow) return;
    e.preventDefault();
    const win = mainWindow;
    let dirty = false;
    try {
      dirty = await win.webContents.executeJavaScript("(window.app && window.app.isDirty && window.app.isDirty()) || false");
    } catch {
      dirty = false;
    }
    if (!dirty) {
      forceClose = true;
      win.close();
      return;
    }
    const { response } = await dialog.showMessageBox(win, {
      type: "warning",
      buttons: ["Save", "Don't Save", "Cancel"],
      defaultId: 0,
      cancelId: 2,
      message: "Do you want to save the changes you made?",
      detail: "Your changes will be lost if you don't save them.",
    });
    if (response === 2) return; // Cancel — stay open
    if (response === 1) {
      forceClose = true; // Don't Save
      win.close();
      return;
    }
    // Save, then close if it actually persisted (the user may cancel Save As).
    let saved = false;
    try {
      saved = await win.webContents.executeJavaScript("window.app.saveForClose()");
    } catch {
      saved = false;
    }
    if (saved) {
      forceClose = true;
      win.close();
    }
  });

  mainWindow.on("closed", () => {
    mainWindow = null;
  });
}

// --- IPC: native file dialogs + filesystem (renderer is sandboxed) ---

ipcMain.handle("dialog:openXlsx", async () => {
  if (!mainWindow) return null;
  const res = await dialog.showOpenDialog(mainWindow, {
    properties: ["openFile"],
    filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
  });
  if (res.canceled || res.filePaths.length === 0) return null;
  const path = res.filePaths[0];
  const data = await readFile(path);
  return { path, data: data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) };
});

ipcMain.handle(
  "dialog:saveXlsx",
  async (_e, args: { data: ArrayBuffer; defaultPath?: string }) => {
    if (!mainWindow) return null;
    const res = await dialog.showSaveDialog(mainWindow, {
      defaultPath: args.defaultPath ?? "Untitled.xlsx",
      filters: [{ name: "Excel Workbook", extensions: ["xlsx"] }],
    });
    if (res.canceled || !res.filePath) return null;
    await writeFile(res.filePath, Buffer.from(args.data));
    return { path: res.filePath };
  }
);

ipcMain.handle("file:save", async (_e, args: { path: string; data: ArrayBuffer }) => {
  await writeFile(args.path, Buffer.from(args.data));
  return { path: args.path };
});

ipcMain.handle("dialog:openCsv", async () => {
  if (!mainWindow) return null;
  const res = await dialog.showOpenDialog(mainWindow, {
    properties: ["openFile"],
    filters: [{ name: "CSV", extensions: ["csv", "txt"] }],
  });
  if (res.canceled || res.filePaths.length === 0) return null;
  const text = await readFile(res.filePaths[0], "utf8");
  return { path: res.filePaths[0], text };
});

ipcMain.handle("dialog:saveCsv", async (_e, args: { text: string; defaultPath?: string }) => {
  if (!mainWindow) return null;
  const res = await dialog.showSaveDialog(mainWindow, {
    defaultPath: args.defaultPath ?? "export.csv",
    filters: [{ name: "CSV", extensions: ["csv"] }],
  });
  if (res.canceled || !res.filePath) return null;
  await writeFile(res.filePath, args.text, "utf8");
  return { path: res.filePath };
});

ipcMain.handle("print:pdf", async (_e, args: { defaultPath?: string }) => {
  if (!mainWindow) return null;
  const res = await dialog.showSaveDialog(mainWindow, {
    defaultPath: args.defaultPath ?? "sheet.pdf",
    filters: [{ name: "PDF", extensions: ["pdf"] }],
  });
  if (res.canceled || !res.filePath) return null;
  const data = await mainWindow.webContents.printToPDF({ landscape: true, printBackground: true });
  await writeFile(res.filePath, data);
  return { path: res.filePath };
});

function buildMenu() {
  const isMac = process.platform === "darwin";
  const send = (channel: string) => () => mainWindow?.webContents.send(channel);
  const template: Electron.MenuItemConstructorOptions[] = [
    ...(isMac
      ? [{ role: "appMenu" as const }]
      : []),
    {
      label: "File",
      submenu: [
        { label: "New", accelerator: "CmdOrCtrl+N", click: send("menu:new") },
        { label: "Open…", accelerator: "CmdOrCtrl+O", click: send("menu:open") },
        { type: "separator" as const },
        { label: "Save", accelerator: "CmdOrCtrl+S", click: send("menu:save") },
        { label: "Save As…", accelerator: "CmdOrCtrl+Shift+S", click: send("menu:saveAs") },
        { type: "separator" as const },
        { label: "Import CSV…", click: send("menu:importCsv") },
        { label: "Export CSV…", click: send("menu:exportCsv") },
        { label: "Export PDF…", accelerator: "CmdOrCtrl+P", click: send("menu:exportPdf") },
        { type: "separator" as const },
        isMac ? { role: "close" as const } : { role: "quit" as const },
      ],
    },
    {
      label: "Edit",
      submenu: [
        { label: "Undo", accelerator: "CmdOrCtrl+Z", click: send("menu:undo") },
        { label: "Redo", accelerator: "CmdOrCtrl+Shift+Z", click: send("menu:redo") },
        { type: "separator" as const },
        { role: "cut" as const },
        { role: "copy" as const },
        { role: "paste" as const },
      ],
    },
    {
      label: "View",
      submenu: [
        { role: "reload" as const },
        { role: "forceReload" as const },
        { role: "toggleDevTools" as const },
        { type: "separator" as const },
        // Zoom the worksheet only (not the whole UI), like Excel.
        { label: "Zoom In", accelerator: "CmdOrCtrl+Plus", click: send("menu:zoomIn") },
        { label: "Zoom In", accelerator: "CmdOrCtrl+=", visible: false, click: send("menu:zoomIn") },
        { label: "Zoom Out", accelerator: "CmdOrCtrl+-", click: send("menu:zoomOut") },
        { label: "Actual Size", accelerator: "CmdOrCtrl+0", click: send("menu:zoomReset") },
        { type: "separator" as const },
        { role: "togglefullscreen" as const },
      ],
    },
    { role: "windowMenu" as const },
    {
      role: "help" as const,
      submenu: [
        {
          label: "Learn More",
          click: () => shell.openExternal("https://github.com"),
        },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

app.whenReady().then(() => {
  createWindow();
  buildMenu();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
