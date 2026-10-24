const { app, BrowserWindow, Menu, dialog, ipcMain, shell, net } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFile } = require('child_process');
const crypto = require('crypto');

// DocDoc stores an exact copy of the editor content inside each .docx it saves,
// so formatting (colours, sizes, shapes) survives reopening in DocDoc.
// If the file was later edited in Word, the copy is ignored.
const EMBED_PART = 'docdoc/content.json';
const { readHeaderFooter, writeHeaderFooter } = require('./lib/headerfooter');
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
  const asDocx = async (bytes, keepPath) => {
    let headerFooter = null;
    try {
      headerFooter = await readHeaderFooter(await require('jszip').loadAsync(bytes));
    } catch (_) { /* no header/footer */ }
    return {
      filePath: keepPath ? filePath : null,
      suggestedName: keepPath ? null : path.join(path.dirname(filePath), base + '.docx'),
      docx: bytes,
      headerFooter,
    };
  };

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
        {
          label: 'Open Recent',
          submenu: loadRecent().length
            ? [
              ...loadRecent().map((f) => ({ label: path.basename(f), sublabel: path.dirname(f), click: () => send('openRecent:' + f) })),
              { type: 'separator' },
              { label: 'Clear Recent Files', click: () => { saveRecent([]); app.clearRecentDocuments(); buildMenu(); send('recentChanged'); } },
            ]
            : [{ label: 'No recent files', enabled: false }],
        },
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
        // the page handles Ctrl+Z / Ctrl+Y itself (its own history covers every change)
        { label: 'Undo', accelerator: 'CmdOrCtrl+Z', registerAccelerator: false, click: () => send('undo') },
        { label: 'Redo', accelerator: 'CmdOrCtrl+Y', registerAccelerator: false, click: () => send('redo') },
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
      submenu: [
        { label: 'Keyboard Shortcuts', accelerator: 'F1', click: () => send('shortcuts') },
        { type: 'separator' },
        { label: 'Check for Updates…', click: () => checkForUpdates(true) },
        { label: `About DocDoc (version ${app.getVersion()})`, click: () => send('about') },
      ],
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

  // If the page crashes, reload it; the autosaved copy is offered straight away.
  win.webContents.on('render-process-gone', (_e, details) => {
    if (details.reason === 'clean-exit') return;
    rendererCrashed = true;
    setTimeout(() => { if (win && !win.isDestroyed()) win.reload(); }, 300);
  });

  // Spellcheck suggestions on right click.
  win.webContents.on('context-menu', (_e, params) => {
    const ctx = lastContext;
    lastContext = {};
    const items = params.dictionarySuggestions.map((s) => ({
      label: s,
      click: () => win.webContents.replaceMisspelling(s),
    }));
    if (items.length) items.push({ type: 'separator' });
    if (ctx.table) {
      const t = (label, cmd) => ({ label, click: () => send('table:' + cmd) });
      items.push(
        { label: 'Insert', submenu: [t('Row Above', 'rowAbove'), t('Row Below', 'rowBelow'), t('Column Left', 'colLeft'), t('Column Right', 'colRight')] },
        { label: 'Delete', submenu: [t('Row', 'delRow'), t('Column', 'delCol'), t('Whole Table', 'delTable')] },
        t('Merge Cells', 'merge'),
        t('Split Cell', 'split'),
        { type: 'separator' },
      );
    }
    if (ctx.object) {
      const o = (label, cmd) => ({ label, click: () => send('obj:' + cmd) });
      items.push(o('Bring to Front', 'front'), o('Send Behind Text', 'back'), o('Delete', 'delete'), { type: 'separator' });
    }
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
    else if (choice === 1) clearRecovery(); // Don't Save: the user chose to discard
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
        addRecent(arg);
      } catch (err) {
        dialog.showErrorBox('Could not open file', String(err.message || err));
      }
    }
  });
}

// What was right-clicked (sent by the page just before the menu opens).
let lastContext = {};
ipcMain.on('context-info', (e, info) => {
  lastContext = info || {};
  e.returnValue = true;
});

// Open a file dropped onto the window.
ipcMain.handle('open-path', async (_e, filePath) => {
  if (!filePath || !fs.existsSync(filePath)) {
    dialog.showErrorBox('File not found', `${filePath || 'The file'} could not be found. It may have been moved or deleted.`);
    removeRecent(filePath);
    return null;
  }
  try {
    const doc = await readDocument(filePath);
    addRecent(filePath);
    return doc;
  } catch (err) {
    dialog.showErrorBox('Could not open file', err.friendly ? err.message : `${path.basename(filePath)} could not be opened.\n\n${err.message || err}`);
    return null;
  }
});

// Download a picture pasted from a web page and return it embedded (data URL).
ipcMain.handle('fetch-image', async (_e, url) => {
  if (!/^https?:\/\//i.test(url || '')) return null;
  try {
    const res = await net.fetch(url);
    if (!res.ok) return null;
    const type = (res.headers.get('content-type') || '').split(';')[0].trim();
    if (!type.startsWith('image/')) return null;
    const buf = Buffer.from(await res.arrayBuffer());
    if (buf.length > 15 * 1024 * 1024) return null;
    return `data:${type};base64,${buf.toString('base64')}`;
  } catch (_) {
    return null;
  }
});

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
    const doc = await readDocument(r.filePaths[0]);
    addRecent(r.filePaths[0]);
    return doc;
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
      await writeHeaderFooter(zip, page.hf, page, path.parse(target).name);
      const docXml = await zip.file('word/document.xml').async('nodebuffer');
      zip.file(EMBED_PART, JSON.stringify({ check: sha1(docXml), html: editorHtml || html, page }), { createFolders: false });
      const ct = await zip.file('[Content_Types].xml').async('string');
      if (!/Extension="json"/i.test(ct)) {
        zip.file('[Content_Types].xml', ct.replace('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="json" ContentType="application/json"/>'));
      }
      fs.writeFileSync(target, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
    }
    addRecent(target);
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

/* ================= Autosave & recovery =================
 * While a document has unsaved changes, the page sends a copy every few seconds. It is
 * written to <userData>/recovery/<session>.json and deleted once the work is saved or
 * discarded. If DocDoc or Windows crashes, the copy is offered the next time DocDoc starts.
 */
const RECOVERY_DIR = path.join(app.getPath('userData'), 'recovery');
const SESSION = `${Date.now()}-${process.pid}`;
const sessionFile = () => path.join(RECOVERY_DIR, `${SESSION}.json`);
let rendererCrashed = false;
function clearRecovery() {
  try { fs.rmSync(sessionFile(), { force: true }); } catch (_) { /* ignore */ }
}
function processAlive(pid) {
  try { process.kill(pid, 0); return true; } catch (e) { return e.code === 'EPERM'; }
}
ipcMain.handle('recovery-save', (_e, data) => {
  try {
    fs.mkdirSync(RECOVERY_DIR, { recursive: true });
    const tmp = sessionFile() + '.tmp';
    fs.writeFileSync(tmp, JSON.stringify({ ...data, pid: process.pid }));
    fs.renameSync(tmp, sessionFile()); // never leave a half-written backup
    return true;
  } catch (_) {
    return false;
  }
});
ipcMain.handle('recovery-clear', () => clearRecovery());
// Backups left behind by DocDoc sessions that did not close properly.
ipcMain.handle('recovery-list', () => {
  let names = [];
  try { names = fs.readdirSync(RECOVERY_DIR).filter((n) => n.endsWith('.json')); } catch (_) { return []; }
  const out = [];
  for (const n of names) {
    const id = n.slice(0, -5);
    if (id === SESSION && !rendererCrashed) continue;
    try {
      const data = JSON.parse(fs.readFileSync(path.join(RECOVERY_DIR, n), 'utf8'));
      if (id !== SESSION && data.pid && data.pid !== process.pid && processAlive(data.pid)) continue; // another open DocDoc window
      out.push({ id, name: data.name || 'Document', savedAt: data.savedAt || 0, filePath: data.filePath || null });
    } catch (_) { /* unreadable backup */ }
  }
  rendererCrashed = false;
  return out.sort((a, b) => b.savedAt - a.savedAt);
});
ipcMain.handle('recovery-load', (_e, id) => {
  if (!/^[\w.-]+$/.test(id || '')) return null;
  try { return JSON.parse(fs.readFileSync(path.join(RECOVERY_DIR, id + '.json'), 'utf8')); } catch (_) { return null; }
});
ipcMain.handle('recovery-delete', (_e, id) => {
  if (!/^[\w.-]+$/.test(id || '') || id === SESSION) return;
  try { fs.rmSync(path.join(RECOVERY_DIR, id + '.json'), { force: true }); } catch (_) { /* ignore */ }
});

/* ================= Recent files ================= */
const RECENT_FILE = path.join(app.getPath('userData'), 'recent.json');
function loadRecent() {
  try { return JSON.parse(fs.readFileSync(RECENT_FILE, 'utf8')).filter((f) => typeof f === 'string'); } catch (_) { return []; }
}
function saveRecent(list) {
  try { fs.writeFileSync(RECENT_FILE, JSON.stringify(list.slice(0, 10))); } catch (_) { /* ignore */ }
}
function addRecent(file) {
  if (!file) return;
  const full = path.resolve(file);
  saveRecent([full, ...loadRecent().filter((f) => f.toLowerCase() !== full.toLowerCase())]);
  try { app.addRecentDocument(full); } catch (_) { /* not supported */ }
  buildMenu();
  send('recentChanged');
}
function removeRecent(file) {
  if (!file) return;
  saveRecent(loadRecent().filter((f) => f !== file));
  buildMenu();
  send('recentChanged');
}
ipcMain.handle('recent-list', () => loadRecent().filter((f) => fs.existsSync(f)).map((f) => ({ path: f, name: path.basename(f), dir: path.dirname(f) })));

/* ================= Automatic updates =================
 * New versions are published as GitHub Releases (see .github/workflows/build-windows.yml).
 * The installed app checks for a newer release on start-up and every few hours, then the
 * page shows "Update available" -> download progress -> "Restart to update".
 */
let updater = null;
let manualCheck = false;
function sendUpdate(info) {
  if (win && !win.isDestroyed()) win.webContents.send('update-status', info);
}
// GitHub gives release notes as HTML; the page shows plain text.
function plainNotes(notes) {
  const text = Array.isArray(notes) ? notes.map((n) => n.note || '').join('\n') : String(notes || '');
  return text
    .replace(/<\/(p|li|h\d|div)>|<br\s*\/?>/gi, '\n')
    .replace(/<li[^>]*>/gi, '• ')
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, '\n\n')
    .trim()
    .slice(0, 4000);
}
function setupUpdates() {
  if (!app.isPackaged) return; // only the installed app updates itself
  const { autoUpdater } = require('electron-updater');
  updater = autoUpdater;
  autoUpdater.autoDownload = false; // ask first
  autoUpdater.autoInstallOnAppQuit = true; // a downloaded update installs when DocDoc closes
  autoUpdater.on('update-available', (info) => {
    sendUpdate({ state: 'available', version: info.version, current: app.getVersion(), notes: plainNotes(info.releaseNotes) });
  });
  autoUpdater.on('update-not-available', () => {
    if (manualCheck) sendUpdate({ state: 'none', current: app.getVersion() });
    manualCheck = false;
  });
  autoUpdater.on('download-progress', (p) => sendUpdate({ state: 'progress', percent: Math.round(p.percent || 0) }));
  autoUpdater.on('update-downloaded', (info) => sendUpdate({ state: 'downloaded', version: info.version }));
  autoUpdater.on('error', (err) => {
    // stay quiet on automatic checks (e.g. offline); explain on a manual check or a failed download
    sendUpdate({ state: 'error', quiet: !manualCheck, message: String((err && err.message) || err).split('\n')[0] });
    manualCheck = false;
  });
  setTimeout(() => checkForUpdates(false), 8000);
  setInterval(() => checkForUpdates(false), 4 * 60 * 60 * 1000);
}
function checkForUpdates(manual) {
  if (!updater) {
    if (manual) sendUpdate({ state: 'dev', current: app.getVersion() });
    return;
  }
  manualCheck = manual;
  if (manual) sendUpdate({ state: 'checking' });
  updater.checkForUpdates().catch(() => { /* reported through the error event */ });
}
ipcMain.handle('app-version', () => app.getVersion());
ipcMain.on('update-check', () => checkForUpdates(true));
ipcMain.on('update-download', () => {
  if (updater) updater.downloadUpdate().catch(() => { /* reported through the error event */ });
});
ipcMain.on('update-install', () => {
  // the page has already asked about unsaved work; install silently and reopen DocDoc
  if (!updater) return;
  forceClose = true;
  clearRecovery();
  updater.quitAndInstall(true, true);
});

app.whenReady().then(() => {
  const { session } = require('electron');
  const allowed = new Set(['local-fonts', 'clipboard-read', 'clipboard-sanitized-write']);
  session.defaultSession.setPermissionRequestHandler((_wc, perm, cb) => cb(allowed.has(perm)));
  session.defaultSession.setPermissionCheckHandler((_wc, perm) => allowed.has(perm));
  buildMenu();
  createWindow();
  setupUpdates();
});
app.on('window-all-closed', () => app.quit());
