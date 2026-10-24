const { app, BrowserWindow, Menu, dialog, ipcMain, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFile } = require('child_process');
const crypto = require('crypto');

// DocDoc stores an exact copy of the editor content inside each .docx it saves,
// so formatting (colours, sizes, shapes) survives reopening in DocDoc.
// If the file was later edited in Word, the copy is ignored.
const EMBED_PART = 'docdoc/content.json';
const sha1 = (b) => crypto.createHash('sha1').update(b).digest('hex');

let win;
let dirty = false;
let forceClose = false;

const WORD_EXT = ['docx', 'docm', 'dotx', 'dotm'];
// Older or other formats, converted to .docx with Word or LibreOffice if installed.
const CONVERT_EXT = ['doc', 'dot', 'rtf', 'odt', 'ott', 'wpd', 'wps', 'xml'];
const FILTERS_OPEN = [
  { name: 'All Documents', extensions: [...WORD_EXT, ...CONVERT_EXT, 'html', 'htm', 'md', 'markdown', 'txt'] },
  { name: 'Word Documents', extensions: [...WORD_EXT, 'doc', 'dot'] },
  { name: 'Rich Text / OpenDocument', extensions: ['rtf', 'odt', 'ott'] },
  { name: 'Web Page', extensions: ['html', 'htm'] },
  { name: 'Markdown', extensions: ['md', 'markdown'] },
  { name: 'Text', extensions: ['txt'] },
  { name: 'All Files', extensions: ['*'] },
];
const FILTERS_SAVE = [
  { name: 'Word Document (*.docx)', extensions: ['docx'] },
  { name: 'Web Page (*.html)', extensions: ['html'] },
  { name: 'Plain Text (*.txt)', extensions: ['txt'] },
];

function escapeHtml(s) {
  return s.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
}

function run(cmd, args, opts = {}) {
  return new Promise((resolve, reject) => {
    execFile(cmd, args, { timeout: 180000, windowsHide: true, maxBuffer: 64 * 1024 * 1024, ...opts }, (err, stdout, stderr) => {
      if (err) reject(new Error((stderr || err.message || '').toString().trim()));
      else resolve(stdout.toString());
    });
  });
}

function findLibreOffice() {
  const candidates = [];
  if (process.platform === 'win32') {
    for (const base of [process.env.ProgramFiles, process.env['ProgramFiles(x86)'], 'C:\\Program Files', 'C:\\Program Files (x86)']) {
      if (base) candidates.push(path.join(base, 'LibreOffice', 'program', 'soffice.exe'));
    }
  } else {
    candidates.push('/usr/bin/soffice', '/usr/local/bin/soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice');
  }
  return candidates.find((c) => fs.existsSync(c));
}

// Convert any document Word or LibreOffice can read into .docx bytes.
async function convertToDocx(filePath) {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'docdoc-'));
  const out = path.join(tmp, path.parse(filePath).name + '.docx');
  const errors = [];
  try {
    if (process.platform === 'win32') {
      // Microsoft Word, if installed (best results for .doc files)
      const ps = [
        "$ErrorActionPreference = 'Stop'",
        '$w = New-Object -ComObject Word.Application',
        '$w.Visible = $false',
        '$w.DisplayAlerts = 0',
        'try { $d = $w.Documents.Open($env:DD_IN, $false, $true, $false); $d.SaveAs2($env:DD_OUT, 16); $d.Close($false) } finally { $w.Quit() }',
      ].join('; ');
      try {
        await run('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', ps], {
          env: { ...process.env, DD_IN: filePath, DD_OUT: out },
        });
        if (fs.existsSync(out)) return fs.readFileSync(out);
      } catch (e) {
        errors.push('Word: ' + e.message.split('\n')[0]);
      }
    }
    const soffice = findLibreOffice();
    if (soffice) {
      try {
        await run(soffice, ['--headless', '--norestore', `-env:UserInstallation=file:///${tmp.replace(/\\/g, '/')}/profile`,
          '--convert-to', 'docx:MS Word 2007 XML', '--outdir', tmp, filePath]);
        if (fs.existsSync(out)) return fs.readFileSync(out);
      } catch (e) {
        errors.push('LibreOffice: ' + e.message.split('\n')[0]);
      }
    }
    const err = new Error(
      `DocDoc needs Microsoft Word or the free LibreOffice (libreoffice.org) installed to open ${path.extname(filePath)} files. ` +
      'Install either one, then try again. Or open the file in another program and save it as .docx.' +
      (errors.length ? '\n\nDetails: ' + errors.join('; ') : '')
    );
    err.friendly = true;
    throw err;
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
}

async function readDocument(filePath) {
  const ext = path.extname(filePath).slice(1).toLowerCase();
  const base = path.parse(filePath).name;
  // Saving must never overwrite templates, macro files or old formats with .docx content.
  const asDocx = (bytes, keepPath) => ({
    filePath: keepPath ? filePath : null,
    suggestedName: keepPath ? null : path.join(path.dirname(filePath), base + '.docx'),
    docx: bytes,
  });

  if (WORD_EXT.includes(ext)) {
    const buf = fs.readFileSync(filePath);
    if (ext === 'docx') {
      try {
        const JSZip = require('jszip');
        const zip = await JSZip.loadAsync(buf);
        const part = zip.file(EMBED_PART);
        const docXml = zip.file('word/document.xml');
        if (part && docXml) {
          const data = JSON.parse(await part.async('string'));
          if (data.check === sha1(await docXml.async('nodebuffer'))) {
            return { filePath, html: data.html, page: data.page };
          }
        }
      } catch (_) { /* not a DocDoc file; render it normally */ }
    }
    return asDocx(buf, ext === 'docx');
  }
  if (CONVERT_EXT.includes(ext)) {
    return asDocx(await convertToDocx(filePath), false);
  }
  const text = fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, '');
  if (ext === 'html' || ext === 'htm') {
    const m = text.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    return { filePath, html: m ? m[1] : text };
  }
  if (ext === 'md' || ext === 'markdown') {
    const { marked } = require('marked');
    return { filePath: null, suggestedName: path.join(path.dirname(filePath), base + '.docx'), html: marked.parse(text) };
  }
  const html = text.split(/\r?\n/).map((l) => `<p>${escapeHtml(l) || '<br>'}</p>`).join('');
  return { filePath, html };
}

// Tidy tables in the generated Word XML:
//  - alignment per table (the converter centres them all), in document order
//  - no extra space after paragraphs inside cells, so mark boxes and grids stay compact
function fixTables(xml, aligns) {
  let n = 0;
  xml = xml.replace(/<w:tbl>(\s*)<w:tblPr>([\s\S]*?)<\/w:tblPr>/g, (m, ws, inner) => {
    const jc = aligns[n++] || 'left';
    const body = /<w:jc [^>]*\/>/.test(inner) ? inner.replace(/<w:jc [^>]*\/>/, `<w:jc w:val="${jc}"/>`) : inner + `<w:jc w:val="${jc}"/>`;
    return `<w:tbl>${ws}<w:tblPr>${body}</w:tblPr>`;
  });
  // Column grid: use the real cell widths (the converter writes equal, fractional columns).
  xml = xml.replace(/<w:tblGrid>([\s\S]*?)<\/w:tblGrid>([\s\S]*?)<\/w:tbl>/g, (m, grid, rest) => {
    if (rest.includes('<w:tbl>')) return m; // nested table: leave as is
    const n = (grid.match(/<w:gridCol /g) || []).length;
    const rows = rest.split('<w:tr>').slice(1);
    for (const row of rows) {
      const cells = row.split('<w:tc>').slice(1);
      if (cells.length !== n || cells.some((c) => /<w:gridSpan /.test(c))) continue;
      const widths = cells.map((c) => (c.match(/<w:tcW w:w="([\d.]+)"/) || [])[1]);
      if (widths.some((w) => !w)) continue;
      const cols = widths.map((w) => `<w:gridCol w:w="${Math.round(+w)}"/>`).join('');
      return `<w:tblGrid>${cols}</w:tblGrid>${rest}</w:tbl>`;
    }
    return m;
  });
  // Word expects whole numbers for widths.
  xml = xml.replace(/w:w="(\d+)\.\d+"/g, (m, i) => `w:w="${i}"`);
  let depth = 0;
  return xml.replace(/<w:tc>|<\/w:tc>|<w:spacing ([^>]*?)\/>/g, (m, attrs) => {
    if (m === '<w:tc>') { depth++; return m; }
    if (m === '</w:tc>') { depth--; return m; }
    if (depth > 0 && !/w:after=/.test(attrs)) return `<w:spacing w:after="0" ${attrs}/>`;
    return m;
  });
}

// Fallback reader if the high-fidelity renderer fails on a file.
ipcMain.handle('docx-basic', async (_e, bytes) => {
  const mammoth = require('mammoth');
  const r = await mammoth.convertToHtml({ buffer: Buffer.from(bytes) });
  return r.value;
});

// Every font family installed on this computer.
let fontCache = null;
ipcMain.handle('list-fonts', async () => {
  if (fontCache) return fontCache;
  let names = [];
  try {
    if (process.platform === 'win32') {
      const out = await run('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command',
        '[Console]::OutputEncoding = [Text.Encoding]::UTF8; Add-Type -AssemblyName System.Drawing; (New-Object System.Drawing.Text.InstalledFontCollection).Families | ForEach-Object { $_.Name }']);
      names = out.split(/\r?\n/);
    } else {
      const out = await run('fc-list', [':', 'family']);
      names = out.split('\n').map((l) => l.split(',')[0]);
    }
  } catch (_) { /* no list available */ }
  fontCache = [...new Set(names.map((n) => n.trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  return fontCache;
});

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
    const arg = process.argv.slice(1).find((a) => /\.(docx|docm|dotx|dotm|doc|dot|rtf|odt|ott|wpd|wps|html?|md|markdown|txt)$/i.test(a) && fs.existsSync(a));
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
    dialog.showErrorBox('Could not open file', err.friendly ? err.message : `${path.basename(r.filePaths[0])} could not be opened.\n\n${err.message || err}`);
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

ipcMain.handle('save-file', async (_e, { filePath, suggestedName, html, editorHtml, text, saveAs, page, tableAlign }) => {
  let target = filePath;
  if (!target || saveAs) {
    const r = await dialog.showSaveDialog(win, {
      defaultPath: target || suggestedName || 'Document1.docx',
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
      zip.file('word/document.xml', fixTables(await zip.file('word/document.xml').async('string'), tableAlign || []));
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
  const { session } = require('electron');
  const allowed = new Set(['local-fonts', 'clipboard-read', 'clipboard-sanitized-write']);
  session.defaultSession.setPermissionRequestHandler((_wc, perm, cb) => cb(allowed.has(perm)));
  session.defaultSession.setPermissionCheckHandler((_wc, perm) => allowed.has(perm));
  buildMenu();
  createWindow();
});
app.on('window-all-closed', () => app.quit());
