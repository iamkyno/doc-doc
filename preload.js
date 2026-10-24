const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  openFile: () => ipcRenderer.invoke('open-file'),
  saveFile: (payload) => ipcRenderer.invoke('save-file', payload),
  pickImage: () => ipcRenderer.invoke('pick-image'),
  exportPdf: (name) => ipcRenderer.invoke('export-pdf', name),
  print: () => ipcRenderer.invoke('print'),
  setTitle: (t) => ipcRenderer.send('set-title', t),
  setDirty: (d) => ipcRenderer.send('set-dirty', d),
  onMenu: (cb) => ipcRenderer.on('menu', (_e, cmd) => cb(cmd)),
  edit: (cmd) => ipcRenderer.send('edit', cmd),
  docxBasic: (bytes) => ipcRenderer.invoke('docx-basic', bytes),
  listFonts: () => ipcRenderer.invoke('list-fonts'),
  closeNow: () => ipcRenderer.send('close-now'),
  onOpenPath: (cb) => ipcRenderer.on('open-path-result', (_e, r) => cb(r)),
});
