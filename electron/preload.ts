import { contextBridge, ipcRenderer } from "electron";

// The only bridge between the sandboxed renderer and Node. Everything the app
// needs from the OS (file dialogs, reading/writing files, menu events) goes here.
export type OpenResult = { path: string; data: ArrayBuffer } | null;
export type SaveResult = { path: string } | null;

const api = {
  openXlsx: (): Promise<OpenResult> => ipcRenderer.invoke("dialog:openXlsx"),
  saveXlsxAs: (data: ArrayBuffer, defaultPath?: string): Promise<SaveResult> =>
    ipcRenderer.invoke("dialog:saveXlsx", { data, defaultPath }),
  saveFile: (path: string, data: ArrayBuffer): Promise<SaveResult> =>
    ipcRenderer.invoke("file:save", { path, data }),
  openCsv: (): Promise<{ path: string; text: string } | null> => ipcRenderer.invoke("dialog:openCsv"),
  saveCsv: (text: string, defaultPath?: string): Promise<SaveResult> =>
    ipcRenderer.invoke("dialog:saveCsv", { text, defaultPath }),
  printPdf: (defaultPath?: string): Promise<SaveResult> =>
    ipcRenderer.invoke("print:pdf", { defaultPath }),
  onMenu: (channel: string, cb: () => void) => {
    const valid = [
      "menu:new",
      "menu:open",
      "menu:save",
      "menu:saveAs",
      "menu:undo",
      "menu:redo",
      "menu:importCsv",
      "menu:exportCsv",
      "menu:exportPdf",
      "menu:zoomIn",
      "menu:zoomOut",
      "menu:zoomReset",
    ];
    if (!valid.includes(channel)) return;
    ipcRenderer.on(channel, () => cb());
  },
};

contextBridge.exposeInMainWorld("native", api);

export type NativeApi = typeof api;
