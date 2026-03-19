"use strict";
const electron = require("electron");
const api = {
  ipc: (channel) => electron.ipcRenderer.invoke(channel),
  logResults: (lines) => electron.ipcRenderer.invoke("log-results", lines)
};
electron.contextBridge.exposeInMainWorld("api", api);
