const { app, BrowserWindow, Menu, dialog, ipcMain, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

// DocDoc stores an exact copy of the editor content inside each .docx it saves,
// so formatting (colours, sizes, shapes) survives reopening in DocDoc.
// If the file was later edited in Word, the copy is ignored.
const EMBED_PART = 'docdoc/content.json';
const sha1 = (b) => crypto.createHash('sha1').update(b).digest('hex');

let win;
let dirty = false;
let forceClose = false;

const FILTERS_OPEN = [
  { name: 'Documents', extensions: ['docx', 'html', 'htm', 'txt'] },
  { name: 'Word Document', extensions: ['docx'] },
  { name: 'Web Page', extensions: ['html', 'htm'] },
  { name: 'Text', extensions: ['txt'] },
];
const FILTERS_SAVE = [
  { name: 'Word Document (*.docx)', extensions: ['docx'] },
  { name: 'Web Page (*.html)', extensions: ['html'] },
  { name: 'Plain Text (*.txt)', extensions: ['txt'] },
];

function escapeHtml(s) {
  return s.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
}

async function readDocument(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.docx') {
    try {
      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
      const part = zip.file(EMBED_PART);
      const docXml = zip.file('word/document.xml');
      if (part && docXml) {
        const data = JSON.parse(await part.async('string'));
        if (data.check === sha1(await docXml.async('nodebuffer'))) {
          return { filePath, html: data.html, page: data.page };
        }
      }
    } catch (_) { /* fall back to plain conversion */ }
    const mammoth = require('mammoth');
    const result = await mammoth.convertToHtml({ path: filePath });
    return { filePath, html: result.value };
  }
  const text = fs.readFileSync(filePath, 'utf8');
  if (ext === '.html' || ext === '.htm') {
    const m = text.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    return { filePath, html: m ? m[1] : text };
  }
  const html = text.split(/\r?\n/).map((l) => `<p>${escapeHtml(l) || '<br>'}</p>`).join('');
  return { filePath, html };
}

function send(cmd) {
  if (win) win.webContents.send('menu', cmd);
}

function buildMenu() {
  const template = [
    {
      label: 'File',
      submenu: [
        { label: 'New', accelerator: 'CmdOrCtrl+N', click: () => send('new') },
        { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: () => send('open') },
        { type: 'separator' },
        { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => send('save') },
        { label: 'Save As…', accelerator: 'F12', click: () => send('saveAs') },
        { label: 'Save As… ', accelerator: 'CmdOrCtrl+Shift+S', visible: false, click: () => send('saveAs') },
        { label: 'Export as PDF…', click: () => send('pdf') },
        { type: 'separator' },
        { label: 'Print…', accelerator: 'CmdOrCtrl+P', click: () => send('print') },
        { type: 'separator' },
        { label: 'Exit', accelerator: 'Alt+F4', click: () => win && win.close() },
      ],
    },
    {
      label: 'Edit',
      submenu: [
        { role: 'undo', accelerator: 'CmdOrCtrl+Z' },
        { role: 'redo', accelerator: 'CmdOrCtrl+Y' },
        { type: 'separator' },
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        { label: 'Paste as Plain Text', accelerator: 'CmdOrCtrl+Shift+V', role: 'pasteAndMatchStyle' },
        { role: 'selectAll' },
        { type: 'separator' },
        { label: 'Find…', accelerator: 'CmdOrCtrl+F', click: () => send('find') },
        { label: 'Replace…', accelerator: 'CmdOrCtrl+H', click: () => send('replace') },
      ],
    },
    {
      label: 'View',
      submenu: [
        { label: 'Zoom In  (Ctrl + Mouse Wheel)', click: () => send('zoomIn') },
        { label: 'Zoom Out', click: () => send('zoomOut') },
        { label: 'Reset Zoom (100%)', click: () => send('zoomReset') },
        { type: 'separator' },
        { role: 'togglefullscreen' },
      ],
    },
    {
      label: 'Help',
      submenu: [{ label: 'Keyboard Shortcuts', accelerator: 'F1', click: () => send('shortcuts') }],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

function createWindow() {
  win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 800,
    minHeight: 500,
    title: 'Document1 - DocDoc',
    icon: path.join(__dirname, 'build', 'icon.png'),
    backgroundColor: '#e9ebee',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      spellcheck: true,
    },
  });
  win.loadFile(path.join(__dirname, 'src', 'index.html'));

  // Open links in the system browser, never inside the editor.
  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e) => e.preventDefault());

  // Spellcheck suggestions on right click.
  win.webContents.on('context-menu', (_e, params) => {
    const items = params.dictionarySuggestions.map((s) => ({
      label: s,
      click: () => win.webContents.replaceMisspelling(s),
    }));
    if (items.length) items.push({ type: 'separator' });
    items.push({ role: 'cut' }, { role: 'copy' }, { role: 'paste' }, { type: 'separator' }, { role: 'selectAll' });
    Menu.buildFromTemplate(items).popup();
  });

  win.on('close', (e) => {
    if (!dirty || forceClose) return;
    const choice = dialog.showMessageBoxSync(win, {
      type: 'warning',
      buttons: ['Save', "Don't Save", 'Cancel'],
      defaultId: 0,
      cancelId: 2,
      message: 'Do you want to save changes to your document?',
    });
    if (choice === 2) e.preventDefault();
    else if (choice === 0) {
      e.preventDefault();
      send('saveThenClose');
    }
  });

  // Open a file passed on the command line (double-clicking a .docx).
  win.webContents.once('did-finish-load', async () => {
    const arg = process.argv.slice(1).find((a) => /\.(docx|html?|txt)$/i.test(a) && fs.existsSync(a));
    if (arg) {
      try {
        win.webContents.send('open-path-result', await readDocument(arg));
      } catch (err) {
        dialog.showErrorBox('Could not open file', String(err.message || err));
      }
    }
  });
}

ipcMain.on('set-title', (_e, t) => win && win.setTitle(t));
ipcMain.on('set-dirty', (_e, d) => { dirty = !!d; });
ipcMain.on('edit', (_e, cmd) => {
  if (win && ['cut', 'copy', 'paste'].includes(cmd)) win.webContents[cmd]();
});
ipcMain.on('close-now', () => { forceClose = true; win && win.close(); });

ipcMain.handle('open-file', async () => {
  const r = await dialog.showOpenDialog(win, { properties: ['openFile'], filters: FILTERS_OPEN });
  if (r.canceled || !r.filePaths.length) return null;
  try {
    return await readDocument(r.filePaths[0]);
  } catch (err) {
    dialog.showErrorBox('Could not open file', String(err.message || err));
    return null;
  }
});

ipcMain.handle('pick-image', async () => {
  const r = await dialog.showOpenDialog(win, {
    properties: ['openFile'],
    filters: [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg'] }],
  });
  if (r.canceled || !r.filePaths.length) return null;
  const p = r.filePaths[0];
  const ext = path.extname(p).slice(1).toLowerCase();
  const mime = ext === 'svg' ? 'image/svg+xml' : ext === 'jpg' ? 'image/jpeg' : `image/${ext}`;
  return `data:${mime};base64,${fs.readFileSync(p).toString('base64')}`;
});

ipcMain.handle('save-file', async (_e, { filePath, html, editorHtml, text, saveAs, page }) => {
  let target = filePath;
  if (!target || saveAs) {
    const r = await dialog.showSaveDialog(win, {
      defaultPath: target || 'Document1.docx',
      filters: FILTERS_SAVE,
    });
    if (r.canceled || !r.filePath) return null;
    target = r.filePath;
  }
  const ext = path.extname(target).toLowerCase();
  try {
    if (ext === '.html' || ext === '.htm') {
      fs.writeFileSync(target, `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(path.basename(target))}</title></head><body>${editorHtml || html}</body></html>`);
    } else if (ext === '.txt') {
      fs.writeFileSync(target, text);
    } else {
      if (ext !== '.docx') target += '.docx';
      const HTMLtoDOCX = require('@turbodocx/html-to-docx');
      const twip = (mm) => Math.round(mm * 56.6929);
      const buf = await HTMLtoDOCX(`<!doctype html><html><body>${html}</body></html>`, null, {
        orientation: page.landscape ? 'landscape' : 'portrait',
        // portrait dimensions; the library swaps them for landscape
        pageSize: { width: twip(page.widthMm), height: twip(page.heightMm) },
        margins: {
          top: twip(page.marginMm), bottom: twip(page.marginMm),
          left: twip(page.marginMm), right: twip(page.marginMm),
        },
        font: 'Calibri',
        fontSize: 22,
        table: { row: { cantSplit: true } },
      });
      const JSZip = require('jszip');
      const zip = await JSZip.loadAsync(Buffer.from(buf));
      const docXml = await zip.file('word/document.xml').async('nodebuffer');
      zip.file(EMBED_PART, JSON.stringify({ check: sha1(docXml), html: editorHtml || html, page }), { createFolders: false });
      const ct = await zip.file('[Content_Types].xml').async('string');
      if (!/Extension="json"/i.test(ct)) {
        zip.file('[Content_Types].xml', ct.replace('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="json" ContentType="application/json"/>'));
      }
      fs.writeFileSync(target, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
    }
    return target;
  } catch (err) {
    dialog.showErrorBox('Could not save file', String(err.message || err));
    return null;
  }
});

ipcMain.handle('export-pdf', async (_e, name) => {
  const r = await dialog.showSaveDialog(win, {
    defaultPath: (name || 'Document1').replace(/\.\w+$/, '') + '.pdf',
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  });
  if (r.canceled || !r.filePath) return null;
  const data = await win.webContents.printToPDF({ printBackground: true, preferCSSPageSize: true });
  fs.writeFileSync(r.filePath, data);
  shell.openPath(r.filePath);
  return r.filePath;
});

ipcMain.handle('print', () => new Promise((resolve) => {
  win.webContents.print({ printBackground: true }, (ok) => resolve(ok));
}));

app.whenReady().then(() => {
  buildMenu();
  createWindow();
});
app.on('window-all-closed', () => app.quit());
