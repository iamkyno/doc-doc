/* DocDoc editor logic */
(() => {
  'use strict';

  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
  const editor = $('#editor');
  const page = $('#page');
  const workspace = $('#workspace');
  const api = window.api || {};

  const PAGE_SIZES = { A4: [210, 297], Letter: [215.9, 279.4], Legal: [215.9, 355.6] };
  const FONT_STEPS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
  const BLOCK_SEL = 'p,h1,h2,h3,h4,h5,h6,li,blockquote,div:not(.shape):not(.page-break),pre';

  const state = {
    filePath: null,
    dirty: false,
    zoom: 1,
    page: { size: 'A4', landscape: false, marginMm: 25.4 },
    savedRange: null,
    selectedObj: null,
    painter: null,
  };

  document.execCommand('styleWithCSS', false, true);
  document.execCommand('defaultParagraphSeparator', false, 'p');

  /* ================= Selection helpers ================= */
  function inEditor(node) {
    return node && (node === editor || editor.contains(node));
  }
  document.addEventListener('selectionchange', () => {
    const sel = window.getSelection();
    if (sel.rangeCount && inEditor(sel.anchorNode)) {
      state.savedRange = sel.getRangeAt(0).cloneRange();
      updateToolbarState();
    }
  });
  function restoreSelection() {
    const sel = window.getSelection();
    // Selection is still live in the editor: leave it alone.
    if (document.activeElement && inEditor(document.activeElement) && sel.rangeCount && inEditor(sel.anchorNode)) return;
    editor.focus({ preventScroll: true });
    if (state.savedRange) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(state.savedRange);
    }
  }
  function exec(cmd, val = null) {
    restoreSelection();
    document.execCommand(cmd, false, val);
    markDirty();
    updateToolbarState();
  }
  function currentRange() {
    const sel = window.getSelection();
    if (sel.rangeCount && inEditor(sel.anchorNode)) return sel.getRangeAt(0);
    return state.savedRange;
  }
  function anchorElement() {
    const r = currentRange();
    if (!r) return editor;
    let n = r.startContainer;
    if (n.nodeType === 3) n = n.parentNode;
    return inEditor(n) ? n : editor;
  }
  function selectedBlocks() {
    const r = currentRange();
    if (!r) return [];
    // innermost blocks touched by the selection
    const blocks = $$(BLOCK_SEL, editor).filter((b) => r.intersectsNode(b) && !b.querySelector('p,h1,h2,h3,h4,h5,h6,li,blockquote,pre'));
    if (blocks.length) return blocks;
    const b = anchorElement().closest(BLOCK_SEL);
    return b && inEditor(b) && b !== editor ? [b] : [];
  }
  function ensureParagraph() {
    // Wrap loose text in the editor root in a <p>, so block formatting always has a target.
    if (!editor.firstChild) editor.innerHTML = '<p><br></p>';
  }

  /* ================= Dirty / title ================= */
  function docName() {
    const p = state.filePath || state.suggestedName;
    return p ? p.split(/[\\/]/).pop() : 'Document1';
  }
  function updateTitle() {
    const name = docName();
    $('#docname').textContent = name + (state.dirty ? ' •' : '');
    $('#saveState').textContent = state.dirty ? 'Unsaved changes' : (state.filePath ? 'Saved' : '');
    api.setTitle && api.setTitle(`${name}${state.dirty ? '*' : ''} - DocDoc`);
    api.setDirty && api.setDirty(state.dirty);
  }
  function markDirty() {
    relayoutFloats();
    if (!state.dirty) {
      state.dirty = true;
      updateTitle();
    }
    scheduleStats();
  }
  editor.addEventListener('input', () => {
    ensureParagraph();
    markDirty();
    positionHandles();
  });

  /* ================= Ribbon tabs ================= */
  $$('.tab').forEach((t) =>
    t.addEventListener('click', () => {
      $$('.tab').forEach((x) => x.classList.toggle('active', x === t));
      $$('.panel').forEach((p) => p.classList.toggle('active', p.dataset.panel === t.dataset.tab));
    })
  );

  // Keep editor selection when clicking ribbon buttons.
  $('#ribbon').addEventListener('mousedown', (e) => {
    if (e.target.closest('button') && !e.target.closest('.tab')) e.preventDefault();
  });

  /* ================= Commands ================= */
  const commands = {
    new: newDoc,
    open: openDoc,
    save: () => saveDoc(false),
    saveAs: () => saveDoc(true),
    pdf: exportPdf,
    print: printDoc,
    paste: () => clip('paste'),
    cut: () => clip('cut'),
    copy: () => clip('copy'),
    selectAll: () => { editor.focus(); document.execCommand('selectAll'); },
    bold: () => exec('bold'),
    italic: () => exec('italic'),
    underline: () => exec('underline'),
    strikeThrough: () => exec('strikeThrough'),
    subscript: () => exec('subscript'),
    superscript: () => exec('superscript'),
    insertUnorderedList: () => { exec('insertUnorderedList'); fixLists(); },
    insertOrderedList: () => { exec('insertOrderedList'); fixLists(); },
    justifyLeft: () => align('left'),
    justifyCenter: () => align('center'),
    justifyRight: () => align('right'),
    justifyFull: () => align('justify'),
    indent: () => indent(1),
    outdent: () => indent(-1),
    growFont: () => stepFont(1),
    shrinkFont: () => stepFont(-1),
    clearFormat: () => { exec('removeFormat'); exec('unlink'); },
    changeCase,
    foreColor: () => exec('foreColor', $('#fcColor').value),
    highlight: () => exec('hiliteColor', $('#hlColor').value),
    shading: () => {
      restoreSelection();
      selectedBlocks().forEach((b) => (b.style.backgroundColor = $('#shColor').value));
      markDirty();
    },
    formatPainter,
    find: () => openFind(false),
    replace: () => openFind(true),
    pageBreak: insertPageBreak,
    image: insertImage,
    textBox: insertTextBox,
    link: insertLink,
    unlink: () => exec('unlink'),
    hr: () => insertBlock('<hr><p><br></p>'),
    date: () => exec('insertText', new Date().toLocaleDateString(undefined, { day: 'numeric', month: 'long', year: 'numeric' })),
    zoomIn: () => setZoom(state.zoom + 0.1),
    zoomOut: () => setZoom(state.zoom - 0.1),
    zoomReset: () => setZoom(1),
    shortcuts: showShortcuts,
    saveThenClose: async () => {
      const ok = await saveDoc(false);
      if (ok) api.closeNow && api.closeNow();
    },
  };
  function clip(cmd) {
    restoreSelection();
    if (api.edit) api.edit(cmd);
    else document.execCommand(cmd);
    if (cmd !== 'copy') setTimeout(markDirty, 50);
  }
  const lastRun = {};
  function run(cmd) {
    // File shortcuts can arrive from both the menu and the page; run once.
    const now = Date.now();
    if (['new', 'open', 'save', 'saveAs', 'print', 'pdf'].includes(cmd)) {
      if (now - (lastRun[cmd] || 0) < 400) return;
      lastRun[cmd] = now;
    }
    if (commands[cmd]) commands[cmd]();
  }
  $$('[data-cmd]').forEach((b) => b.addEventListener('click', () => run(b.dataset.cmd)));
  api.onMenu && api.onMenu(run);

  // Colour bars reflect chosen colour; picking a colour applies it straight away.
  function syncBars() {
    $('#fcBar').style.background = $('#fcColor').value;
    $('#hlBar').style.background = $('#hlColor').value;
    $('#shBar').style.background = $('#shColor').value;
  }
  syncBars();
  $('#fcColor').addEventListener('change', () => { syncBars(); run('foreColor'); });
  $('#hlColor').addEventListener('change', () => { syncBars(); run('highlight'); });
  $('#shColor').addEventListener('change', () => { syncBars(); run('shading'); });

  /* ---------- Paragraph ---------- */
  function align(dir) {
    if (state.selectedObj) {
      // Align a selected picture/shape by its wrap mode.
      arrange(dir === 'left' ? 'left' : dir === 'right' ? 'right' : 'inline');
      if (dir === 'center') {
        const p = state.selectedObj.closest(BLOCK_SEL);
        if (p) p.style.textAlign = 'center';
      }
      return;
    }
    restoreSelection();
    const blocks = selectedBlocks();
    if (!blocks.length) {
      exec({ left: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', justify: 'justifyFull' }[dir]);
      return;
    }
    blocks.forEach((b) => {
      b.style.textAlign = dir === 'left' ? '' : dir;
      if (!b.getAttribute('style')) b.removeAttribute('style');
    });
    markDirty();
    updateToolbarState();
  }
  // Chrome sometimes nests a list inside a <p>; lift it out.
  function fixLists() {
    $$('p > ul, p > ol', editor).forEach((l) => {
      const p = l.parentElement;
      if (p && p.parentNode) p.replaceWith(...p.childNodes);
    });
  }
  function indent(dir) {
    restoreSelection();
    if (anchorElement().closest('li')) {
      document.execCommand(dir > 0 ? 'indent' : 'outdent');
    } else {
      selectedBlocks().forEach((b) => {
        const cur = parseFloat(b.style.marginLeft) || 0;
        const next = Math.max(0, cur + dir * 36);
        b.style.marginLeft = next ? next + 'px' : '';
      });
    }
    markDirty();
  }
  function setLineSpacing(v) {
    restoreSelection();
    selectedBlocks().forEach((b) => (b.style.lineHeight = v));
    $('#lineSpacing').value = Number(v).toFixed(v == 1.15 ? 2 : 1);
    markDirty();
  }
  $('#lineSpacing').addEventListener('change', (e) => setLineSpacing(e.target.value));

  /* ---------- Styles ---------- */
  function setBlock(tag) {
    exec('formatBlock', tag);
  }
  $$('#styles [data-block]').forEach((b) => b.addEventListener('click', () => setBlock(b.dataset.block)));

  /* ---------- Fonts ---------- */
  function applyFontSize(pt) {
    restoreSelection();
    const r = currentRange();
    if (!r) return;
    document.execCommand('styleWithCSS', false, false);
    document.execCommand('fontSize', false, '7');
    document.execCommand('styleWithCSS', false, true);
    const made = [];
    $$('font[size="7"]', editor).forEach((f) => {
      const span = document.createElement('span');
      span.style.fontSize = pt + 'pt';
      while (f.firstChild) span.appendChild(f.firstChild);
      f.replaceWith(span);
      made.push(span);
      // drop nested sizes so the new size wins
      $$('span', span).forEach((s) => {
        s.style.fontSize = '';
        if (!s.getAttribute('style')) s.replaceWith(...s.childNodes);
      });
    });
    if (made.length) {
      // keep the text selected so the size can be changed again
      const nr = document.createRange();
      nr.setStartBefore(made[0].firstChild || made[0]);
      nr.setEndAfter(made[made.length - 1].lastChild || made[made.length - 1]);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(nr);
      state.savedRange = nr.cloneRange();
    }
    // Chrome may emit spans with font-size: xxx-large when styleWithCSS sticks
    $$('span[style*="xxx-large"]', editor).forEach((s) => (s.style.fontSize = pt + 'pt'));
    markDirty();
    updateToolbarState();
  }
  function currentFontPt() {
    const px = parseFloat(getComputedStyle(anchorElement()).fontSize) || 14.67;
    return Math.round((px * 72) / 96 * 2) / 2;
  }
  function stepFont(dir, exact) {
    const cur = currentFontPt();
    let next;
    if (exact) next = Math.max(1, cur + dir);
    else if (dir > 0) next = FONT_STEPS.find((s) => s > cur) || cur + 10;
    else next = [...FONT_STEPS].reverse().find((s) => s < cur) || Math.max(1, cur - 1);
    applyFontSize(next);
  }
  /* ---------- Font picker ---------- */
  const fontInput = $('#fontName');
  const fontList = $('#fontList');
  const fonts = {
    bundled: (window.BUNDLED_FONTS || []).map((f) => f.family).sort((a, b) => a.localeCompare(b)),
    installed: [],
    recent: [],
    loaded: false,
  };
  try { fonts.recent = JSON.parse(localStorage.getItem('docdoc-recent-fonts') || '[]'); } catch (_) { /* no storage */ }
  let fontItems = [];
  let fontHot = -1;
  let fontBefore = '';

  async function loadInstalledFonts() {
    const set = new Set();
    try {
      if (window.queryLocalFonts) (await window.queryLocalFonts()).forEach((f) => set.add(f.family));
    } catch (_) { /* not allowed here; use the system list */ }
    try {
      if (api.listFonts) (await api.listFonts()).forEach((f) => set.add(f));
    } catch (_) { /* ignore */ }
    fonts.installed = [...set].filter((f) => f && !f.startsWith('.')).sort((a, b) => a.localeCompare(b));
    fonts.loaded = true;
    if (!fontList.hidden) renderFontList(fontInput.value === fontBefore ? '' : fontInput.value);
  }
  setTimeout(loadInstalledFonts, 300);

  function fontRow(name) {
    const d = document.createElement('div');
    d.className = 'fi';
    d.textContent = name;
    d.title = name;
    d.style.fontFamily = `${JSON.stringify(name)}, "Segoe UI", sans-serif`;
    if (name.toLowerCase() === fontBefore.toLowerCase()) d.classList.add('current');
    fontItems.push(d);
    return d;
  }
  function renderFontList(filter = '') {
    const q = filter.trim().toLowerCase();
    const pick = (arr) => {
      const m = arr.filter((f) => !q || f.toLowerCase().includes(q));
      return q ? [...m.filter((f) => f.toLowerCase().startsWith(q)), ...m.filter((f) => !f.toLowerCase().startsWith(q))] : m;
    };
    const bundledSet = new Set(fonts.bundled.map((f) => f.toLowerCase()));
    const sections = [
      ['Recently used', fonts.recent],
      ['DocDoc fonts (built in)', fonts.bundled],
      ['Installed on this PC', fonts.installed.filter((f) => !bundledSet.has(f.toLowerCase()))],
    ];
    fontList.innerHTML = '';
    fontItems = [];
    for (const [title, arr] of sections) {
      const list = pick(arr);
      if (!list.length) continue;
      const h = document.createElement('div');
      h.className = 'fl-head';
      h.textContent = title;
      fontList.appendChild(h);
      list.forEach((f) => fontList.appendChild(fontRow(f)));
    }
    if (!fonts.loaded) {
      const d = document.createElement('div');
      d.className = 'fl-empty';
      d.textContent = 'Finding fonts on this PC…';
      fontList.appendChild(d);
    }
    if (!fontItems.length && q) {
      const d = document.createElement('div');
      d.className = 'fl-empty';
      d.textContent = `No match. Press Enter to use "${filter.trim()}".`;
      fontList.appendChild(d);
    }
    fontHot = q ? (fontItems.length ? 0 : -1) : fontItems.findIndex((d) => d.classList.contains('current'));
    markFontHot();
  }
  function markFontHot() {
    fontItems.forEach((d, i) => d.classList.toggle('hot', i === fontHot));
    if (fontItems[fontHot]) fontItems[fontHot].scrollIntoView({ block: 'nearest' });
  }
  function openFontList() {
    fontList.hidden = false;
    renderFontList();
  }
  function closeFontList() {
    fontList.hidden = true;
  }
  function applyFont(name) {
    name = (name || '').trim();
    closeFontList();
    if (!name) { fontInput.value = fontBefore; return; }
    fontInput.value = name;
    fontBefore = name;
    exec('fontName', name);
    fonts.recent = [name, ...fonts.recent.filter((f) => f.toLowerCase() !== name.toLowerCase())].slice(0, 6);
    try { localStorage.setItem('docdoc-recent-fonts', JSON.stringify(fonts.recent)); } catch (_) { /* ignore */ }
    editor.focus({ preventScroll: true });
  }
  fontInput.addEventListener('focus', () => {
    fontBefore = fontInput.value;
    fontInput.select();
    openFontList();
  });
  fontInput.addEventListener('input', () => {
    fontList.hidden = false;
    renderFontList(fontInput.value);
  });
  fontInput.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      if (fontList.hidden) openFontList();
      fontHot = Math.max(0, Math.min(fontItems.length - 1, fontHot + (e.key === 'ArrowDown' ? 1 : -1)));
      markFontHot();
      if (fontItems[fontHot]) fontInput.value = fontItems[fontHot].textContent;
    } else if (e.key === 'Enter') {
      e.preventDefault();
      const typed = fontInput.value.trim();
      const hotName = fontItems[fontHot] && fontItems[fontHot].textContent;
      // exact typed name wins; otherwise take the highlighted suggestion
      const exact = fontItems.find((d) => d.textContent.toLowerCase() === typed.toLowerCase());
      applyFont(exact ? exact.textContent : hotName || typed);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      fontInput.value = fontBefore;
      closeFontList();
      restoreSelection();
    } else if (e.key === 'Tab') {
      closeFontList();
    }
  });
  fontInput.addEventListener('blur', () => {
    setTimeout(() => {
      if (document.activeElement === fontInput) return;
      closeFontList();
      fontInput.value = fontBefore;
    }, 150);
  });
  fontList.addEventListener('mousedown', (e) => e.preventDefault());
  fontList.addEventListener('click', (e) => {
    const it = e.target.closest('.fi');
    if (it) applyFont(it.textContent);
  });
  $('#fontArrow').addEventListener('click', () => {
    if (!fontList.hidden && document.activeElement === fontInput) { closeFontList(); return; }
    fontInput.focus();
  });
  $('#fontSize').addEventListener('change', (e) => applyFontSize(e.target.value));

  function changeCase() {
    restoreSelection();
    const sel = window.getSelection();
    const t = sel.toString();
    if (!t) return;
    let out;
    if (t === t.toLowerCase()) out = t.toUpperCase();
    else if (t === t.toUpperCase()) out = t.toLowerCase().replace(/\b\p{L}/gu, (c) => c.toUpperCase());
    else out = t.toLowerCase();
    document.execCommand('insertText', false, out);
    // reselect so it can be cycled again
    const r = sel.getRangeAt(0);
    const node = r.startContainer;
    if (node.nodeType === 3 && r.startOffset >= out.length) {
      const nr = document.createRange();
      nr.setStart(node, r.startOffset - out.length);
      nr.setEnd(node, r.startOffset);
      sel.removeAllRanges();
      sel.addRange(nr);
    }
    markDirty();
  }

  /* ---------- Format painter ---------- */
  function formatPainter() {
    if (state.painter) {
      state.painter = null;
      $('#painterBtn').classList.remove('on');
      return;
    }
    const cs = getComputedStyle(anchorElement());
    state.painter = {
      bold: parseInt(cs.fontWeight, 10) >= 600,
      italic: cs.fontStyle === 'italic',
      underline: cs.textDecorationLine.includes('underline'),
      strike: cs.textDecorationLine.includes('line-through'),
      font: cs.fontFamily.split(',')[0].replace(/["']/g, ''),
      size: Math.round((parseFloat(cs.fontSize) * 72) / 96),
      color: cs.color,
    };
    $('#painterBtn').classList.add('on');
  }
  editor.addEventListener('mouseup', () => {
    if (!state.painter) return;
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return;
    const p = state.painter;
    document.execCommand('removeFormat');
    if (p.bold) document.execCommand('bold');
    if (p.italic) document.execCommand('italic');
    if (p.underline) document.execCommand('underline');
    if (p.strike) document.execCommand('strikeThrough');
    document.execCommand('fontName', false, p.font);
    document.execCommand('foreColor', false, p.color);
    state.savedRange = sel.getRangeAt(0).cloneRange();
    applyFontSize(p.size);
    state.painter = null;
    $('#painterBtn').classList.remove('on');
  });

  /* ---------- Toolbar state ---------- */
  function updateToolbarState() {
    $$('[data-state]').forEach((b) => {
      let on = false;
      try { on = document.queryCommandState(b.dataset.state); } catch (_) { /* ignore */ }
      b.classList.toggle('on', on);
    });
    const el = anchorElement();
    const cs = getComputedStyle(el);
    const block = el.closest(BLOCK_SEL);
    const ta = block ? getComputedStyle(block).textAlign : 'left';
    const map = { left: 'justifyLeft', start: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', end: 'justifyRight', justify: 'justifyFull' };
    $$('[data-state^="justify"]').forEach((b) => b.classList.toggle('on', map[ta] === b.dataset.state));
    const fam = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    if (document.activeElement !== fontInput) fontInput.value = fam;
    const pt = currentFontPt();
    const fs = $('#fontSize');
    const ptStr = String(pt);
    if (![...fs.options].some((o) => o.value === ptStr)) {
      $$('option.custom', fs).forEach((o) => o.remove());
      const o = new Option(ptStr);
      o.className = 'custom';
      fs.add(o);
    }
    fs.value = ptStr;
  }

  /* ================= Insert ================= */
  // Insert block-level HTML (tables, breaks) after the current paragraph.
  function insertBlock(html) {
    restoreSelection();
    const tpl = document.createElement('template');
    tpl.innerHTML = html;
    const nodes = Array.from(tpl.content.childNodes);
    let block = anchorElement().closest('p,h1,h2,h3,h4,h5,h6,blockquote,ul,ol,table,pre,.textbox > *');
    while (block && block.parentElement && block.parentElement !== editor && !block.parentElement.classList.contains('textbox')) {
      block = block.parentElement.closest('p,h1,h2,h3,h4,h5,h6,blockquote,ul,ol,table,pre') || block.parentElement;
      if (block === editor) break;
    }
    if (!block || block === editor) editor.append(...nodes);
    else if (block.tagName === 'P' && !block.textContent.trim() && !block.querySelector('img,.shape')) block.replaceWith(...nodes);
    else block.after(...nodes);
    const last = nodes[nodes.length - 1];
    const target = last && last.nodeType === 1 && last.tagName === 'P' ? last : (nodes[0].querySelector && nodes[0].querySelector('td')) || last;
    const r = document.createRange();
    r.selectNodeContents(target);
    r.collapse(true);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
    state.savedRange = r.cloneRange();
    markDirty();
  }
  function insertPageBreak() {
    insertBlock('<div class="page-break" contenteditable="false"></div><p><br></p>');
  }

  async function insertImage() {
    if (!api.pickImage) return;
    const src = await api.pickImage();
    if (!src) return;
    const img = new Image();
    img.onload = () => {
      const maxW = editor.clientWidth;
      const w = Math.min(img.naturalWidth, maxW, 480);
      exec('insertHTML', `<img src="${src}" style="width:${w}px">`);
    };
    img.src = src;
  }

  function insertLink() {
    const r = currentRange();
    const text = r ? r.toString() : '';
    const existing = anchorElement().closest('a');
    prompt2('Insert Link', [
      { id: 'text', label: 'Text to display', value: existing ? existing.textContent : text },
      { id: 'url', label: 'Address', value: existing ? existing.getAttribute('href') : 'https://' },
    ]).then((v) => {
      if (!v || !v.url || v.url === 'https://') return;
      let url = v.url.trim();
      if (!/^(https?:|mailto:|file:)/i.test(url)) url = 'https://' + url;
      if (existing) {
        existing.href = url;
        existing.textContent = v.text || url;
        markDirty();
        return;
      }
      const a = `<a href="${escapeAttr(url)}">${escapeHtml(v.text || url)}</a>`;
      exec('insertHTML', a);
    });
  }

  // Ctrl+click opens links like Word.
  editor.addEventListener('click', (e) => {
    const a = e.target.closest('a');
    if (a && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      window.open(a.href);
    }
  });

  /* ---------- Table picker ---------- */
  const cells = $('#tableCells');
  for (let r = 0; r < 8; r++) {
    for (let c = 0; c < 10; c++) {
      const d = document.createElement('div');
      d.dataset.r = r + 1;
      d.dataset.c = c + 1;
      cells.appendChild(d);
    }
  }
  cells.addEventListener('mouseover', (e) => {
    const t = e.target.dataset;
    if (!t.r) return;
    $$('div', cells).forEach((d) => d.classList.toggle('hot', +d.dataset.r <= +t.r && +d.dataset.c <= +t.c));
    $('#tableLabel').textContent = `${t.c} × ${t.r} Table`;
  });
  cells.addEventListener('mousedown', (e) => e.preventDefault());
  cells.addEventListener('click', (e) => {
    const t = e.target.dataset;
    if (!t.r) return;
    closeDropdowns();
    let html = '<table><tbody>';
    for (let r = 0; r < +t.r; r++) {
      html += '<tr>' + '<td><br></td>'.repeat(+t.c) + '</tr>';
    }
    html += '</tbody></table><p><br></p>';
    insertBlock(html);
    // put the caret in the first cell
    const sel = window.getSelection();
    const p = sel.anchorNode && (sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentNode);
    const table = p && p.previousElementSibling;
    if (table && table.tagName === 'TABLE') {
      const r = document.createRange();
      r.selectNodeContents(table.querySelector('td'));
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
    }
  });

  /* ---------- Dropdowns ---------- */
  function closeDropdowns() {
    $$('.dropdown.open').forEach((d) => d.classList.remove('open'));
  }
  $$('[data-dd]').forEach((b) =>
    b.addEventListener('click', (e) => {
      e.stopPropagation();
      const dd = b.closest('.dropdown');
      const was = dd.classList.contains('open');
      closeDropdowns();
      if (!was) dd.classList.add('open');
    })
  );
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.dropdown')) closeDropdowns();
  });

  /* ---------- Symbols ---------- */
  const SYMBOLS = '©®™§¶•–—…°±×÷≠≤≥≈∞√∑πΩµ€£¥¢←→↑↓↔✓✗★☆♥☺½¼¾¹²³‰«»“”‘’'.split('');
  const symMenu = $('#symbolMenu');
  SYMBOLS.forEach((s) => {
    const b = document.createElement('button');
    b.textContent = s;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => { closeDropdowns(); exec('insertText', s); });
    symMenu.appendChild(b);
  });

  /* ---------- Shapes ---------- */
  const SHAPES = {
    rect: '<rect x="1" y="1" width="98" height="98"/>',
    rounded: '<rect x="1" y="1" width="98" height="98" rx="14" ry="14"/>',
    ellipse: '<ellipse cx="50" cy="50" rx="49" ry="49"/>',
    triangle: '<polygon points="50,1 99,99 1,99"/>',
    rtriangle: '<polygon points="1,1 99,99 1,99"/>',
    diamond: '<polygon points="50,1 99,50 50,99 1,50"/>',
    pentagon: '<polygon points="50,1 99,38 80,99 20,99 1,38"/>',
    hexagon: '<polygon points="25,1 75,1 99,50 75,99 25,99 1,50"/>',
    star: '<polygon points="50,1 62,37 99,37 69,60 80,97 50,75 20,97 31,60 1,37 38,37"/>',
    arrowR: '<polygon points="1,30 60,30 60,5 99,50 60,95 60,70 1,70"/>',
    arrowL: '<polygon points="99,30 40,30 40,5 1,50 40,95 40,70 99,70"/>',
    arrowU: '<polygon points="30,99 30,40 5,40 50,1 95,40 70,40 70,99"/>',
    arrowD: '<polygon points="30,1 30,60 5,60 50,99 95,60 70,60 70,1"/>',
    plus: '<polygon points="35,1 65,1 65,35 99,35 99,65 65,65 65,99 35,99 35,65 1,65 1,35 35,35"/>',
    heart: '<path d="M50,95 C10,65 1,45 1,28 C1,12 13,1 27,1 C38,1 46,8 50,18 C54,8 62,1 73,1 C87,1 99,12 99,28 C99,45 90,65 50,95 Z"/>',
    callout: '<path d="M1,1 H99 V70 H40 L20,99 L25,70 H1 Z"/>',
    line: '<line x1="1" y1="99" x2="99" y2="1" fill="none"/>',
    arrowLine: '<line x1="1" y1="50" x2="88" y2="50" fill="none"/><polygon points="99,50 85,42 85,58" data-solid="1"/>',
  };
  const shapeMenu = $('#shapeMenu');
  function menuSection(title, cls) {
    const h = document.createElement('div');
    h.className = 'sec';
    h.textContent = title;
    const g = document.createElement('div');
    g.className = cls;
    shapeMenu.append(h, g);
    return g;
  }
  const shapeGrid = menuSection('Shapes', 'shape-grid');
  Object.keys(SHAPES).forEach((k) => {
    const b = document.createElement('button');
    b.title = k;
    b.innerHTML = `<svg viewBox="0 0 100 100" preserveAspectRatio="none"><g fill="#cfd9ec" stroke="#2f528f" stroke-width="6">${SHAPES[k]}</g></svg>`;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => { closeDropdowns(); insertShape(k); });
    shapeGrid.appendChild(b);
  });

  /* ---------- Exam paper boxes ---------- */
  const EXAM = [
    { id: 'markBox', name: 'Mark box', hint: 'Empty box at the right margin',
      icon: '<rect x="12" y="4" width="26" height="26" fill="#fff" stroke="#000" stroke-width="2"/>' },
    { id: 'splitBox', name: 'Mark out of', hint: 'Box with space and "/ 5"',
      icon: '<rect x="3" y="7" width="40" height="20" fill="#fff" stroke="#000" stroke-width="2"/><line x1="23" y1="7" x2="23" y2="27" stroke="#000" stroke-width="2"/><text x="26" y="22" font-size="11" font-family="Arial">5</text>' },
    { id: 'allocation', name: 'Mark allocation', hint: '"(5)" at the right margin',
      icon: '<text x="9" y="23" font-size="15" font-family="Arial">(5)</text>' },
    { id: 'markCircle', name: 'Mark circle', hint: 'Round box for a score',
      icon: '<circle cx="23" cy="17" r="13" fill="#fff" stroke="#000" stroke-width="2"/>' },
    { id: 'totalBox', name: 'Total box', hint: 'TOTAL: ____ / 100',
      icon: '<rect x="2" y="8" width="42" height="18" fill="#fff" stroke="#000" stroke-width="2.5"/><text x="5" y="21" font-size="8" font-weight="bold" font-family="Arial">TOTAL</text>' },
    { id: 'tickRow', name: 'Tick boxes', hint: 'A  B  C  D, click to tick',
      icon: '<rect x="2" y="11" width="10" height="10" fill="#fff" stroke="#000" stroke-width="1.5"/><rect x="18" y="11" width="10" height="10" fill="#fff" stroke="#000" stroke-width="1.5"/><path d="M19 16 l3 3 l5 -7" stroke="#000" stroke-width="1.6" fill="none"/><rect x="34" y="11" width="10" height="10" fill="#fff" stroke="#000" stroke-width="1.5"/>' },
    { id: 'answerLines', name: 'Answer lines', hint: 'Lined space for written answers',
      icon: '<g stroke="#000" stroke-width="1.5"><line x1="3" y1="9" x2="43" y2="9"/><line x1="3" y1="17" x2="43" y2="17"/><line x1="3" y1="25" x2="43" y2="25"/></g>' },
    { id: 'dottedLines', name: 'Dotted lines', hint: 'Dotted answer lines',
      icon: '<g stroke="#000" stroke-width="1.6" stroke-dasharray="1.5 2.5"><line x1="3" y1="9" x2="43" y2="9"/><line x1="3" y1="17" x2="43" y2="17"/><line x1="3" y1="25" x2="43" y2="25"/></g>' },
    { id: 'answerBox', name: 'Answer box', hint: 'Large bordered space',
      icon: '<rect x="3" y="4" width="40" height="26" fill="#fff" stroke="#000" stroke-width="2"/>' },
    { id: 'examinerTable', name: "Examiner's use table", hint: 'Question / Max / Mark + Total',
      icon: '<g fill="#fff" stroke="#000" stroke-width="1.3"><rect x="4" y="3" width="38" height="28"/><line x1="4" y1="9" x2="42" y2="9"/><line x1="4" y1="15" x2="42" y2="15"/><line x1="4" y1="21" x2="42" y2="21"/><line x1="4" y1="26" x2="42" y2="26"/><line x1="17" y1="9" x2="17" y2="31"/><line x1="29" y1="9" x2="29" y2="31"/></g>' },
    { id: 'candidateBox', name: 'Candidate details', hint: 'Name, number boxes, class, date',
      icon: '<g fill="#fff" stroke="#000" stroke-width="1.3"><rect x="3" y="4" width="40" height="26"/><line x1="3" y1="13" x2="43" y2="13"/><line x1="3" y1="22" x2="43" y2="22"/><line x1="17" y1="4" x2="17" y2="30"/></g><g stroke="#000" stroke-width="1"><line x1="23" y1="13" x2="23" y2="22"/><line x1="29" y1="13" x2="29" y2="22"/><line x1="35" y1="13" x2="35" y2="22"/></g>' },
    { id: 'instructionBox', name: 'Instructions box', hint: 'Boxed "Instructions" panel',
      icon: '<rect x="3" y="4" width="40" height="26" fill="#f2f2f2" stroke="#000" stroke-width="1.8"/><text x="5" y="12" font-size="5" font-weight="bold" font-family="Arial">INSTRUCTIONS</text><g stroke="#666"><line x1="6" y1="19" x2="38" y2="19"/><line x1="6" y1="24" x2="32" y2="24"/></g>' },
  ];
  const examGrid = menuSection('Exam paper', 'exam-grid');
  EXAM.forEach((x) => {
    const b = document.createElement('button');
    b.title = x.hint;
    b.innerHTML = `<span class="prev"><svg viewBox="0 0 46 34">${x.icon}</svg></span><span>${x.name}<small>${x.hint}</small></span>`;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => { closeDropdowns(); insertExam(x.id); });
    examGrid.appendChild(b);
  });

  // Vertical position (editor coordinates) of the caret line, for boxes that sit at the right margin.
  function caretTop() {
    const r = state.savedRange;
    if (r && inEditor(r.startContainer)) {
      let rect = r.getBoundingClientRect();
      if (!rect.height) {
        const el = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
        rect = el.getBoundingClientRect();
      }
      if (rect.height) return Math.max(0, (rect.top - editor.getBoundingClientRect().top) / state.zoom);
    }
    return visibleTop();
  }
  function floatBox(html, { w, h, left, top, style = '' }) {
    const el = document.createElement('div');
    el.className = 'textbox floating exam';
    el.contentEditable = 'false';
    el.style.cssText = `left:${left}px;top:${top}px;width:${w}px;${h ? `min-height:${h}px;` : ''}${style}`;
    el.innerHTML = textBoxHtml(html);
    placeFloating(el);
    reanchor(el);
    selectObj(el, true);
    markDirty();
    return el;
  }
  const B = '1.5px solid #000000';
  const cell = (content, extra = '') => `<td style="border:${B};padding:3px 6px;${extra}">${content || '<br>'}</td>`;
  function linesTable(n, dotted) {
    const line = dotted ? '1.5px dotted #000000' : '1px solid #000000';
    let rows = '';
    for (let i = 0; i < n; i++) rows += `<tr style="height:30px"><td style="border:none;border-bottom:${line};height:30px;padding:0 2px;vertical-align:bottom"><br></td></tr>`;
    return `<table class="exam" style="width:100%;border-collapse:collapse;margin:0 0 8pt"><tbody>${rows}</tbody></table><p><br></p>`;
  }
  function insertExam(id) {
    const right = (w) => Math.max(0, editor.clientWidth - w);
    const centered = 'display:flex;flex-direction:column;justify-content:center;';
    switch (id) {
      case 'markBox':
        floatBox('<p style="text-align:center;margin:0"><br></p>', { w: 44, h: 44, left: right(44), top: caretTop(), style: `padding:0;border:${B};${centered}` });
        break;
      case 'splitBox':
        floatBox(`<table style="border-collapse:collapse;width:100%;height:100%;margin:0"><tbody><tr>${cell('', 'width:50%;height:34px;text-align:center;vertical-align:middle')}${cell('5', 'width:50%;text-align:center;vertical-align:middle;font-weight:bold')}</tr></tbody></table>`,
          { w: 84, h: 36, left: right(84), top: caretTop(), style: 'padding:0;border:none;background:transparent' });
        break;
      case 'allocation':
        floatBox('<p style="text-align:right;margin:0">(5)</p>', { w: 50, left: right(50), top: caretTop(), style: 'padding:0;border:none;background:transparent' });
        break;
      case 'markCircle':
        floatBox('<p style="text-align:center;margin:0"><br></p>', { w: 52, h: 52, left: right(52), top: caretTop(), style: `padding:0;border:${B};border-radius:999px;${centered}` });
        break;
      case 'totalBox':
        floatBox('<p style="margin:0"><b>TOTAL:</b>&nbsp;&nbsp;_______ / 100</p>', { w: 200, h: 34, left: right(200), top: caretTop(), style: `padding:4px 10px;border:2.5px solid #000000;${centered}` });
        break;
      case 'tickRow':
        exec('insertHTML', ['A', 'B', 'C', 'D'].map((l) => `<span class="tickbox" contenteditable="false">☐</span>&nbsp;${l}`).join('&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;') + '&nbsp;');
        break;
      case 'answerLines':
        insertBlock(linesTable(5, false));
        break;
      case 'dottedLines':
        insertBlock(linesTable(5, true));
        break;
      case 'answerBox':
        insertBlock(`<table class="exam" style="width:100%;border-collapse:collapse;margin:0 0 8pt"><tbody><tr><td style="border:${B};height:160px;padding:6px;vertical-align:top"><br></td></tr></tbody></table><p><br></p>`);
        break;
      case 'examinerTable': {
        let rows = '';
        const cw = Math.round(editor.clientWidth * 0.55 / 3);
        const c3 = (t, extra = '') => cell(t, `width:${cw}px;text-align:center;${extra}`);
        for (let i = 1; i <= 5; i++) rows += `<tr>${c3(String(i))}${c3('')}${c3('')}</tr>`;
        insertBlock(`<table class="exam" style="width:${cw * 3}px;margin:0 0 8pt auto;border-collapse:collapse"><tbody>
<tr><td colspan="3" style="border:${B};padding:4px 6px;text-align:center;font-weight:bold;background:#d9d9d9">FOR EXAMINER'S USE ONLY</td></tr>
<tr>${c3('<b>Question</b>', 'background:#f2f2f2')}${c3('<b>Max</b>', 'background:#f2f2f2')}${c3('<b>Mark</b>', 'background:#f2f2f2')}</tr>
${rows}
<tr>${c3('<b>TOTAL</b>', 'background:#f2f2f2')}${c3('')}${c3('')}</tr>
</tbody></table><p><br></p>`);
        break;
      }
      case 'candidateBox': {
        // one table, no nesting: the number row has 10 small boxes, other rows span them
        // fixed pixel widths: Word ignores percentage widths on cells
        const W = editor.clientWidth;
        const lw = Math.round(W * 0.3);
        const bw = Math.floor((W - lw) / 10);
        const label = (t) => cell(`<b>${t}</b>`, `width:${lw}px;background:#f2f2f2`);
        const wide = `<td colspan="10" style="border:${B};padding:3px 6px;width:${bw * 10}px"><br></td>`;
        const boxes = Array.from({ length: 10 }, () => `<td style="border:${B};width:${bw}px;height:26px;padding:0;text-align:center"><br></td>`).join('');
        insertBlock(`<table class="exam" style="width:100%;border-collapse:collapse;margin:0 0 8pt"><tbody>
<tr>${label('NAME')}${wide}</tr>
<tr>${label('SURNAME')}${wide}</tr>
<tr style="height:28px">${label('CANDIDATE NUMBER')}${boxes}</tr>
<tr>${label('CLASS / GRADE')}${wide}</tr>
<tr>${label('DATE')}${wide}</tr>
</tbody></table><p><br></p>`);
        break;
      }
      case 'instructionBox':
        insertBlock(`<div class="para-box" style="border:${B};padding:6px 10px;background:#f2f2f2"><p style="margin:0 0 4pt"><b>INSTRUCTIONS AND INFORMATION</b></p><ol style="margin:0"><li>Answer ALL the questions.</li><li>Write neatly and legibly.</li><li>Show ALL calculations.</li></ol></div><p><br></p>`);
        break;
      default:
    }
  }

  // Click a tick box to cycle: empty, ticked, crossed.
  editor.addEventListener('click', (e) => {
    const t = e.target.closest && e.target.closest('.tickbox');
    if (!t) return;
    t.textContent = { '☐': '☑', '☑': '☒', '☒': '☐' }[t.textContent] || '☐';
    markDirty();
  });

  // Put a floating object in the document right after the paragraph holding the caret,
  // so it stays with that paragraph (and exports next to it) instead of piling up at the end.
  function placeFloating(el) {
    const r = state.savedRange;
    let n = r && inEditor(r.startContainer) ? r.startContainer : null;
    while (n && n.parentNode !== editor) n = n.parentNode;
    if (n && n !== editor && !(n.classList && n.classList.contains('floating'))) n.after(el);
    else editor.appendChild(el);
  }
  // Floating objects follow their paragraph: we store their offset from it (data-dy)
  // and re-position them whenever the text above changes, like Word's anchors.
  // Each floating object remembers its paragraph by id (data-anchor -> data-aid).
  let aidSeq = 0;
  function newAid() {
    return 'a' + Date.now().toString(36) + (aidSeq++).toString(36);
  }
  function prevBlock(el) {
    let a = el.previousElementSibling;
    while (a && a.classList.contains('floating')) a = a.previousElementSibling;
    return a;
  }
  function anchorOf(el) {
    const id = el.dataset.anchor;
    if (id) {
      for (const b of editor.children) if (b.dataset.aid === id && !b.classList.contains('floating')) return b;
    }
    return prevBlock(el);
  }
  function reanchor(el) {
    if (!el || !el.classList.contains('floating') || el.parentNode !== editor) return;
    const top = parseFloat(el.style.top) || 0;
    let anchor = null;
    for (const b of editor.children) {
      if (b.classList.contains('floating')) continue;
      if (b.offsetTop <= top + 4) anchor = b;
      else break;
    }
    if (anchor) {
      if (!anchor.dataset.aid) anchor.dataset.aid = newAid();
      el.dataset.anchor = anchor.dataset.aid;
      if (prevBlock(el) !== anchor && !el.contains(document.activeElement)) anchor.after(el);
      el.dataset.dy = String(Math.round(top - anchor.offsetTop));
    } else {
      delete el.dataset.anchor;
      el.dataset.dy = String(Math.round(top));
    }
  }
  function relayoutFloats() {
    // Splitting a paragraph copies its id to the new half. The first half with text keeps it
    // (Enter at the start of a question leaves the box with the question, not the blank line).
    const groups = new Map();
    for (const b of editor.children) {
      const id = b.dataset.aid;
      if (!id || b.classList.contains('floating')) continue;
      if (!groups.has(id)) groups.set(id, []);
      groups.get(id).push(b);
    }
    for (const list of groups.values()) {
      if (list.length < 2) continue;
      const keep = list.find((b) => b.textContent.trim()) || list[0];
      list.forEach((b) => { if (b !== keep) delete b.dataset.aid; });
    }
    for (const el of editor.children) {
      if (!el.classList.contains('floating') || el.dataset.dy == null) continue;
      const a = anchorOf(el);
      if (a && !a.dataset.aid) a.dataset.aid = newAid();
      if (a) el.dataset.anchor = a.dataset.aid;
      const top = Math.max(0, (a ? a.offsetTop : 0) + (+el.dataset.dy));
      if (Math.round(parseFloat(el.style.top) || 0) !== Math.round(top)) el.style.top = Math.round(top) + 'px';
    }
  }
  // Text boxes: a non-editable frame with its own editable body, so the caret can't wander in.
  function textBoxHtml(inner) {
    return `<div class="tb-body" contenteditable="true">${inner}</div>`;
  }

  let insertCount = 0;
  function visibleTop() {
    // y (in editor coordinates) near the top of the visible part of the page
    const er = editor.getBoundingClientRect();
    const wr = workspace.getBoundingClientRect();
    const step = (insertCount++ % 6) * 24;
    return Math.max(10, (wr.top - er.top) / state.zoom + 60) + step;
  }
  function shapeSvg(kind, fill, stroke) {
    const isLine = kind === 'line' || kind === 'arrowLine';
    return `<svg viewBox="0 0 100 100" preserveAspectRatio="none"><g fill="${isLine ? 'none' : fill}" stroke="${stroke}" stroke-width="2" vector-effect="non-scaling-stroke">${SHAPES[kind].replace(/<(\w+)/g, '<$1 vector-effect="non-scaling-stroke"')}</g></svg>`;
  }
  function insertShape(kind) {
    const fill = '#4472c4';
    const stroke = '#2f528f';
    const el = document.createElement('span');
    el.className = 'shape floating';
    el.contentEditable = 'false';
    el.dataset.shape = kind;
    const isLine = kind === 'line' || kind === 'arrowLine';
    const top = visibleTop();
    el.style.cssText = `left:${60 + (insertCount % 6) * 24}px;top:${top}px;width:${isLine ? 160 : 140}px;height:${kind === 'arrowLine' ? 24 : isLine ? 100 : 110}px`;
    el.innerHTML = shapeSvg(kind, fill, stroke);
    $$('[data-solid]', el).forEach((p) => p.setAttribute('fill', stroke));
    placeFloating(el);
    reanchor(el);
    selectObj(el, true);
    markDirty();
  }
  function insertTextBox() {
    const el = document.createElement('div');
    el.className = 'textbox floating';
    el.contentEditable = 'false';
    el.style.cssText = `left:60px;top:${visibleTop()}px;width:220px`;
    el.innerHTML = textBoxHtml('<p>Type your text here</p>');
    placeFloating(el);
    reanchor(el);
    selectObj(el, true);
    const r = document.createRange();
    r.selectNodeContents(el.querySelector('p'));
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
    markDirty();
  }

  /* ================= Object selection / move / resize ================= */
  const handle = document.createElement('div');
  handle.id = 'handle';
  const mover = document.createElement('div');
  mover.id = 'mover';
  mover.textContent = '✥ Move';
  page.append(handle, mover);

  function objFrom(target) {
    if (!target || !inEditor(target)) return null;
    if (target.tagName === 'IMG') return target;
    return target.closest('.shape, .textbox');
  }
  function selectObj(el, justInserted) {
    if (state.selectedObj === el) {
      if (justInserted) showTab('format');
      return positionHandles();
    }
    if (state.selectedObj) state.selectedObj.classList.remove('selected-obj');
    state.selectedObj = el;
    if (el) el.classList.add('selected-obj');
    // Word-style contextual tab
    const tab = $('#formatTab');
    tab.hidden = !el;
    if (el) {
      syncFormat(el);
      if (justInserted) showTab('format');
    } else if (tab.classList.contains('active')) showTab('home');
    positionHandles();
  }
  function showTab(name) {
    $$('.tab').forEach((x) => x.classList.toggle('active', x.dataset.tab === name));
    $$('.panel').forEach((p) => p.classList.toggle('active', p.dataset.panel === name));
  }
  function positionHandles() {
    const el = state.selectedObj;
    if (!el || !editor.contains(el)) {
      handle.style.display = mover.style.display = 'none';
      if (el) state.selectedObj = null;
      return;
    }
    const pr = page.getBoundingClientRect();
    const r = el.getBoundingClientRect();
    const z = state.zoom;
    const x = (r.left - pr.left) / z;
    const y = (r.top - pr.top) / z;
    const w = r.width / z;
    const h = r.height / z;
    handle.style.display = 'block';
    handle.style.left = x + w - 6 + 'px';
    handle.style.top = y + h - 6 + 'px';
    const floating = el.classList.contains('floating');
    mover.style.display = floating ? 'block' : 'none';
    mover.style.left = x + 'px';
    mover.style.top = y - 22 + 'px';
  }
  editor.addEventListener('mousedown', (e) => {
    const o = objFrom(e.target);
    selectObj(o);
    if (o && o.tagName === 'IMG') {
      // select the image as a node so copy/cut/delete work naturally
      const r = document.createRange();
      r.selectNode(o);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
      e.preventDefault();
    }
    if (o && o.classList.contains('shape') && o.classList.contains('floating')) {
      startDrag(e, o, 'move');
    }
  });
  document.addEventListener('mousedown', (e) => {
    if (!e.target.closest('#editor, #handle, #mover, #ribbon, #modal')) selectObj(null);
  });
  handle.addEventListener('mousedown', (e) => startDrag(e, state.selectedObj, 'resize'));
  mover.addEventListener('mousedown', (e) => startDrag(e, state.selectedObj, 'move'));

  function startDrag(e, el, mode) {
    if (!el) return;
    e.preventDefault();
    e.stopPropagation();
    const z = state.zoom;
    const sx = e.clientX;
    const sy = e.clientY;
    const rect = el.getBoundingClientRect();
    const sw = rect.width / z;
    const sh = rect.height / z;
    const sl = parseFloat(el.style.left) || 0;
    const st = parseFloat(el.style.top) || 0;
    const ratio = sw / sh;
    let moved = false;
    const onMove = (ev) => {
      const dx = (ev.clientX - sx) / z;
      const dy = (ev.clientY - sy) / z;
      if (Math.abs(dx) + Math.abs(dy) > 1) moved = true;
      if (mode === 'move') {
        el.style.left = Math.round(sl + dx) + 'px';
        el.style.top = Math.max(0, Math.round(st + dy)) + 'px';
      } else {
        let w = Math.max(16, sw + dx);
        let h = Math.max(8, sh + dy);
        const keepRatio = el.tagName === 'IMG' ? !ev.shiftKey : ev.shiftKey;
        if (keepRatio) h = w / ratio;
        el.style.width = Math.round(w) + 'px';
        if (el.tagName === 'IMG') el.style.height = keepRatio ? '' : Math.round(h) + 'px';
        else if (el.classList.contains('textbox')) el.style.minHeight = Math.round(h) + 'px';
        else el.style.height = Math.round(h) + 'px';
      }
      positionHandles();
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      if (moved && mode === 'move') reanchor(el);
      if (moved) markDirty();
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }
  workspace.addEventListener('scroll', positionHandles);
  window.addEventListener('resize', positionHandles);

  function arrange(mode) {
    const el = state.selectedObj;
    if (!el) return;
    if (mode === 'delete') {
      el.remove();
      selectObj(null);
      markDirty();
      return;
    }
    if (mode === 'front' || mode === 'back') {
      el.style.zIndex = mode === 'front' ? 10 : 0;
      if (mode === 'back' && el.classList.contains('floating')) el.style.zIndex = 0;
    } else {
      el.classList.remove('wrap-left', 'wrap-right');
      if (el.classList.contains('floating') && el.tagName !== 'IMG') {
        // move a floating shape/text box back into the text flow at the caret
        el.classList.remove('floating');
        el.style.left = el.style.top = '';
        delete el.dataset.dy;
        delete el.dataset.anchor;
        const r = state.savedRange;
        if (r && !el.contains(r.startContainer)) {
          r.collapse(true);
          r.insertNode(el);
        }
      }
      if (mode === 'left') el.classList.add('wrap-left');
      if (mode === 'right') el.classList.add('wrap-right');
    }
    markDirty();
    positionHandles();
  }
  $$('[data-obj]').forEach((b) => b.addEventListener('click', () => arrange(b.dataset.obj)));
  /* ---------- Shape Format tab ---------- */
  const PX_PER_CM = 96 / 2.54;
  const DASH_CSS = { solid: 'solid', dashed: 'dashed', dotted: 'dotted', long: 'dashed', double: 'double' };
  function objKind(el) {
    if (!el) return null;
    if (el.tagName === 'IMG') return 'img';
    if (el.classList.contains('shape')) return 'shape';
    return 'textbox';
  }
  function nearestOption(sel, v) {
    let best = sel.options[0].value;
    for (const o of sel.options) if (Math.abs(+o.value - v) < Math.abs(+best - v)) best = o.value;
    sel.value = best;
  }
  function hexToRgba(hex, alpha) {
    const n = parseInt(hex.slice(1), 16);
    return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${alpha})`;
  }
  // Read the selected object's look into the Shape Format controls.
  function syncFormat(el) {
    const kind = objKind(el);
    const cs = getComputedStyle(el);
    const d = el.dataset;
    $('#fmtFillGroup').classList.toggle('off', kind === 'img' || (kind === 'shape' && /line/i.test(d.shape)));
    $('#fmtTextGroup').classList.toggle('off', kind !== 'textbox');
    $('#fmtRadius').closest('label').style.opacity = kind === 'shape' ? '.4' : '';
    $('#fmtRadius').disabled = kind === 'shape';
    if (kind === 'shape') {
      const g = el.querySelector('g');
      if (g.getAttribute('fill') !== 'none') $('#fmtFill').value = toHex(g.getAttribute('fill'));
      else if (d.fill) $('#fmtFill').value = d.fill;
      $('#fmtLine').value = toHex(g.getAttribute('stroke') === 'none' ? d.line || '#000000' : g.getAttribute('stroke'));
      $('#fmtWeight').value = d.weight || (g.getAttribute('stroke') === 'none' ? '0' : '1.5');
      if (!d.weight) nearestOption($('#fmtWeight'), g.getAttribute('stroke') === 'none' ? 0 : (parseFloat(g.getAttribute('stroke-width')) || 2) * 0.75);
    } else {
      $('#fmtFill').value = d.fill || (TRANSPARENT_BG(cs.backgroundColor) ? '#ffffff' : toHex(cs.backgroundColor));
      $('#fmtLine').value = toHex(cs.borderTopColor);
      nearestOption($('#fmtWeight'), cs.borderTopStyle === 'none' ? 0 : parseFloat(cs.borderTopWidth) * 0.75);
      nearestOption($('#fmtRadius'), parseFloat(cs.borderTopLeftRadius) || 0);
    }
    $('#fmtAlpha').value = d.alpha || 0;
    $('#fmtDash').value = d.dash || (cs.borderTopStyle === 'dotted' ? 'dotted' : cs.borderTopStyle === 'dashed' ? 'dashed' : cs.borderTopStyle === 'double' ? 'double' : 'solid');
    $('#fmtShadow').checked = d.shadow === '1';
    $('#fmtRotate').value = d.rot || '0';
    if (kind === 'textbox') {
      nearestOption($('#fmtPadding'), parseFloat(cs.paddingLeft) || 0);
      $('#fmtVAlign').value = cs.display === 'flex' ? cs.justifyContent.replace('normal', 'flex-start') : 'flex-start';
      if (!$('#fmtVAlign').value) $('#fmtVAlign').value = 'flex-start';
    }
    const r = el.getBoundingClientRect();
    $('#fmtW').value = (r.width / state.zoom / PX_PER_CM).toFixed(2);
    $('#fmtH').value = (r.height / state.zoom / PX_PER_CM).toFixed(2);
  }
  function TRANSPARENT_BG(v) {
    return !v || v === 'transparent' || v === 'rgba(0, 0, 0, 0)';
  }
  // Write the controls' values onto the selected object.
  function applyFormat(changed) {
    const el = state.selectedObj;
    if (!el) return;
    const kind = objKind(el);
    const d = el.dataset;
    if (changed === 'fill') delete d.nofill;
    if (changed === 'nofill') d.nofill = '1';
    if (changed === 'line' && $('#fmtWeight').value === '0') $('#fmtWeight').value = '1';
    const fill = $('#fmtFill').value;
    const alpha = +$('#fmtAlpha').value;
    const line = $('#fmtLine').value;
    const weight = +$('#fmtWeight').value;
    const dash = $('#fmtDash').value;
    d.fill = fill;
    d.line = line;
    d.weight = String(weight);
    d.dash = dash;
    if (alpha) d.alpha = String(alpha); else delete d.alpha;
    d.shadow = $('#fmtShadow').checked ? '1' : '0';
    const rot = +$('#fmtRotate').value;
    d.rot = String(rot);
    el.style.transform = rot ? `rotate(${rot}deg)` : '';

    if (kind === 'shape') {
      const g = el.querySelector('g');
      const isLine = /line/i.test(d.shape);
      if (!isLine) {
        g.setAttribute('fill', d.nofill ? 'none' : fill);
        if (alpha) g.setAttribute('fill-opacity', String(1 - alpha / 100)); else g.removeAttribute('fill-opacity');
      }
      const sw = weight * 4 / 3;
      g.setAttribute('stroke', weight ? line : 'none');
      g.setAttribute('stroke-width', String(sw || 1));
      $$('[data-solid]', el).forEach((p) => p.setAttribute('fill', line));
      const arr = { dashed: `${3 * sw} ${2 * sw}`, dotted: `${sw} ${1.6 * sw}`, long: `${8 * sw} ${3 * sw}` }[dash];
      if (arr) g.setAttribute('stroke-dasharray', arr); else g.removeAttribute('stroke-dasharray');
      g.setAttribute('stroke-linecap', dash === 'dotted' ? 'round' : 'butt');
      el.style.filter = d.shadow === '1' ? 'drop-shadow(2px 3px 3px rgba(0,0,0,.35))' : '';
    } else {
      if (kind === 'textbox') {
        el.style.background = d.nofill ? 'transparent' : (alpha ? hexToRgba(fill, 1 - alpha / 100) : fill);
        el.style.padding = $('#fmtPadding').value + 'px';
        const va = $('#fmtVAlign').value;
        if (va === 'flex-start') {
          el.style.display = el.style.flexDirection = el.style.justifyContent = '';
        } else {
          el.style.display = 'flex';
          el.style.flexDirection = 'column';
          el.style.justifyContent = va;
        }
      }
      el.style.border = weight ? `${weight}pt ${DASH_CSS[dash]} ${line}` : 'none';
      el.style.borderRadius = +$('#fmtRadius').value ? $('#fmtRadius').value + 'px' : '';
      el.style.boxShadow = d.shadow === '1' ? '2px 3px 6px rgba(0,0,0,.35)' : '';
    }
    markDirty();
    positionHandles();
  }
  function applySize() {
    const el = state.selectedObj;
    if (!el) return;
    const w = Math.max(0.2, parseFloat($('#fmtW').value) || 0) * PX_PER_CM;
    const h = Math.max(0.2, parseFloat($('#fmtH').value) || 0) * PX_PER_CM;
    el.style.width = Math.round(w) + 'px';
    if (objKind(el) === 'textbox') el.style.minHeight = Math.round(h) + 'px';
    else el.style.height = Math.round(h) + 'px';
    markDirty();
    positionHandles();
  }
  $('#fmtFill').addEventListener('input', () => applyFormat('fill'));
  $('[data-fmt="nofill"]').addEventListener('click', () => applyFormat('nofill'));
  $('#fmtAlpha').addEventListener('input', () => applyFormat('alpha'));
  $('#fmtLine').addEventListener('input', () => applyFormat('line'));
  ['#fmtWeight', '#fmtDash', '#fmtRadius', '#fmtShadow', '#fmtRotate', '#fmtPadding', '#fmtVAlign'].forEach((id) =>
    $(id).addEventListener('change', () => applyFormat(id)));
  ['#fmtW', '#fmtH'].forEach((id) => $(id).addEventListener('change', applySize));

  /* ---------- Paragraph borders ---------- */
  function topLevelBlocks(blocks) {
    const out = [];
    blocks.forEach((b) => {
      let n = b;
      while (n.parentElement && n.parentElement !== editor && !n.parentElement.classList.contains('textbox') && !n.parentElement.classList.contains('para-box') && !/^(TD|TH)$/.test(n.parentElement.tagName)) n = n.parentElement;
      if (!out.includes(n)) out.push(n);
    });
    return out;
  }
  function paraBorder(kind) {
    restoreSelection();
    const blocks = selectedBlocks();
    if (!blocks.length) return;
    const line = '1px solid #000000';
    if (kind === 'bottom' || kind === 'top') {
      blocks.forEach((b) => {
        const side = kind === 'bottom' ? 'Bottom' : 'Top';
        b.style['border' + side] = line;
        b.style['padding' + side] = '2px';
      });
    } else if (kind === 'box') {
      const tops = topLevelBlocks(blocks);
      if (tops[0].parentElement.classList.contains('para-box')) return;
      const box = document.createElement('div');
      box.className = 'para-box';
      box.style.border = '1.5px solid #000000';
      box.style.padding = '6px 8px';
      tops[0].before(box);
      tops.forEach((t) => box.appendChild(t));
    } else {
      blocks.forEach((b) => {
        ['borderTop', 'borderBottom', 'borderLeft', 'borderRight', 'border', 'paddingTop', 'paddingBottom'].forEach((p) => (b.style[p] = ''));
        if (!b.getAttribute('style')) b.removeAttribute('style');
        const box = b.closest('.para-box');
        if (box) box.replaceWith(...box.childNodes);
      });
    }
    markDirty();
  }
  $$('[data-border]').forEach((b) => b.addEventListener('click', () => { closeDropdowns(); paraBorder(b.dataset.border); }));

  /* ================= Layout ================= */
  function applyPage() {
    const [w, h] = PAGE_SIZES[state.page.size];
    const pw = state.page.landscape ? h : w;
    const ph = state.page.landscape ? w : h;
    page.style.setProperty('--pw', pw + 'mm');
    page.style.setProperty('--ph', ph + 'mm');
    page.style.setProperty('--margin', state.page.marginMm + 'mm');
    let ps = $('#printStyle');
    if (!ps) {
      ps = document.createElement('style');
      ps.id = 'printStyle';
      document.head.appendChild(ps);
    }
    ps.textContent = `@media print { @page { size: ${pw}mm ${ph}mm; margin: ${state.page.marginMm}mm; } .page { width: auto !important; padding: 0 !important; } }`;
    positionHandles();
    scheduleStats();
  }
  $('#margins').addEventListener('change', (e) => { state.page.marginMm = +e.target.value; applyPage(); markDirty(); });
  $('#orientation').addEventListener('change', (e) => { state.page.landscape = e.target.value === 'landscape'; applyPage(); markDirty(); });
  $('#pageSize').addEventListener('change', (e) => { state.page.size = e.target.value; applyPage(); markDirty(); });
  $('#columns').addEventListener('change', (e) => {
    editor.style.columnCount = e.target.value === '1' ? '' : e.target.value;
    editor.style.columnGap = '12mm';
    markDirty();
  });
  $('#pageColor').addEventListener('input', (e) => {
    page.style.backgroundColor = e.target.value;
    markDirty();
  });

  /* ================= View ================= */
  function setZoom(z) {
    state.zoom = Math.min(3, Math.max(0.3, Math.round(z * 10) / 10));
    document.documentElement.style.setProperty('--zoom', state.zoom);
    $('#zoomRange').value = Math.round(state.zoom * 100);
    $('#zoomLabel').textContent = Math.round(state.zoom * 100) + '%';
    positionHandles();
  }
  $('#zoomRange').addEventListener('input', (e) => setZoom(e.target.value / 100));
  workspace.addEventListener('wheel', (e) => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    setZoom(state.zoom + (e.deltaY < 0 ? 0.1 : -0.1));
  }, { passive: false });
  $('#spellToggle').addEventListener('change', (e) => (editor.spellcheck = e.target.checked));

  /* ================= Stats ================= */
  let statsTimer;
  function scheduleStats() {
    clearTimeout(statsTimer);
    statsTimer = setTimeout(updateStats, 200);
  }
  function updateStats() {
    const text = editor.innerText || '';
    const words = (text.match(/[\p{L}\p{N}'’-]+/gu) || []).length;
    $('#wordCount').textContent = `${words.toLocaleString()} word${words === 1 ? '' : 's'}`;
    $('#charCount').textContent = `${text.replace(/\n/g, '').length.toLocaleString()} characters`;
    const mmPx = 96 / 25.4;
    const [w, h] = PAGE_SIZES[state.page.size];
    const ph = (state.page.landscape ? w : h) * mmPx;
    const breaks = $$('.page-break', editor).length;
    const pages = Math.max(1, Math.ceil((page.scrollHeight - 4) / ph), breaks + 1);
    const wr = workspace.getBoundingClientRect();
    const pr = page.getBoundingClientRect();
    const cur = Math.min(pages, Math.max(1, Math.floor((wr.top + wr.height / 3 - pr.top) / state.zoom / ph) + 1));
    $('#pageInfo').textContent = `Page ${cur} of ${pages}`;
  }
  workspace.addEventListener('scroll', scheduleStats);

  /* ================= Find & replace ================= */
  const findbar = $('#findbar');
  function openFind(withReplace) {
    findbar.hidden = false;
    findbar.classList.toggle('find-only', !withReplace);
    const t = window.getSelection().toString();
    if (t && t.length < 100) $('#findInput').value = t;
    $('#findInput').focus();
    $('#findInput').select();
    countMatches();
  }
  function closeFind() {
    findbar.hidden = true;
    restoreSelection();
  }
  function textNodes() {
    const out = [];
    const w = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
    while (w.nextNode()) out.push(w.currentNode);
    return out;
  }
  function countMatches() {
    const q = $('#findInput').value;
    if (!q) { $('#findCount').textContent = ''; return 0; }
    const text = editor.textContent.toLowerCase();
    let n = 0;
    let i = text.indexOf(q.toLowerCase());
    while (i !== -1) { n++; i = text.indexOf(q.toLowerCase(), i + q.length); }
    $('#findCount').textContent = n ? `${n} result${n === 1 ? '' : 's'}` : 'No results';
    return n;
  }
  function findNext(backwards = false) {
    const q = $('#findInput').value;
    if (!q) return false;
    editor.focus();
    let found = window.find(q, false, backwards, true, false, false, false);
    if (found && !inEditor(window.getSelection().anchorNode)) found = window.find(q, false, backwards, true, false, false, false);
    if (found) {
      const sel = window.getSelection();
      state.savedRange = sel.getRangeAt(0).cloneRange();
      const rect = sel.getRangeAt(0).getBoundingClientRect();
      const wr = workspace.getBoundingClientRect();
      if (rect.top < wr.top || rect.bottom > wr.bottom) workspace.scrollTop += rect.top - wr.top - wr.height / 3;
    }
    return found;
  }
  function replaceOne() {
    const q = $('#findInput').value;
    const sel = window.getSelection();
    if (sel.toString().toLowerCase() === q.toLowerCase() && inEditor(sel.anchorNode)) {
      document.execCommand('insertText', false, $('#replaceInput').value);
      markDirty();
    }
    findNext();
    countMatches();
  }
  function replaceAll() {
    const q = $('#findInput').value;
    if (!q) return;
    const rep = $('#replaceInput').value;
    const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    let n = 0;
    textNodes().forEach((t) => {
      const v = t.nodeValue;
      const nv = v.replace(re, () => { n++; return rep; });
      if (nv !== v) t.nodeValue = nv;
    });
    if (n) markDirty();
    $('#findCount').textContent = `Replaced ${n}`;
  }
  $('#findInput').addEventListener('input', countMatches);
  $('#findInput').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); findNext(e.shiftKey); $('#findInput').focus(); }
  });
  $('#replaceInput').addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); replaceOne(); }
  });
  $('#findNext').addEventListener('click', () => findNext());
  $('#replaceOne').addEventListener('click', replaceOne);
  $('#replaceAll').addEventListener('click', replaceAll);
  $('#findClose').addEventListener('click', closeFind);

  /* ================= Modal helpers ================= */
  function modal(html, { ok = 'OK', cancel = 'Cancel', extra } = {}) {
    return new Promise((resolve) => {
      const m = $('#modal');
      $('#modalBody').innerHTML = html;
      const okB = $('#modalOk');
      const cB = $('#modalCancel');
      okB.textContent = ok;
      cB.textContent = cancel;
      cB.style.display = cancel ? '' : 'none';
      let exB = null;
      if (extra) {
        exB = document.createElement('button');
        exB.textContent = extra;
        okB.after(exB);
      }
      m.hidden = false;
      const first = $('#modalBody input');
      (first || okB).focus();
      const done = (v) => {
        m.hidden = true;
        okB.onclick = cB.onclick = null;
        m.onkeydown = null;
        if (exB) exB.remove();
        resolve(v);
      };
      okB.onclick = () => done('ok');
      cB.onclick = () => done(null);
      if (exB) exB.onclick = () => done('extra');
      m.onkeydown = (e) => {
        if (e.key === 'Escape') done(null);
        if (e.key === 'Enter') { e.preventDefault(); done('ok'); }
      };
    });
  }
  async function prompt2(title, fields) {
    const html = `<h3 style="margin-top:0">${title}</h3>` + fields
      .map((f) => `<label>${f.label}<input id="pf_${f.id}" value="${escapeAttr(f.value || '')}"></label>`)
      .join('');
    const r = await modal(html);
    const vals = {};
    fields.forEach((f) => (vals[f.id] = $(`#pf_${f.id}`).value));
    restoreSelection();
    return r === 'ok' ? vals : null;
  }

  function showShortcuts() {
    const rows = [
      ['Ctrl+N / Ctrl+O', 'New / Open'], ['Ctrl+S / F12', 'Save / Save As'], ['Ctrl+P', 'Print'],
      ['Ctrl+Z / Ctrl+Y', 'Undo / Redo'], ['Ctrl+X / C / V', 'Cut / Copy / Paste'], ['Ctrl+Shift+V', 'Paste as plain text'],
      ['Ctrl+A', 'Select all'], ['Ctrl+F / Ctrl+H', 'Find / Replace'],
      ['Ctrl+B / I / U', 'Bold / Italic / Underline'], ['Ctrl+= / Ctrl+Shift+=', 'Subscript / Superscript'],
      ['Ctrl+Shift+> / <', 'Grow / Shrink font'], ['Ctrl+] / Ctrl+[', 'Font size +1 / −1 pt'],
      ['Ctrl+Space', 'Clear character formatting'], ['Shift+F3', 'Change case'],
      ['Ctrl+L / E / R / J', 'Align left / center / right / justify'],
      ['Ctrl+1 / 2 / 5', 'Line spacing 1.0 / 2.0 / 1.5'], ['Ctrl+M / Ctrl+Shift+M', 'Indent / Outdent'],
      ['Ctrl+Shift+L', 'Bullet list'], ['Ctrl+Alt+1 / 2 / 3', 'Heading 1 / 2 / 3'], ['Ctrl+Shift+N', 'Normal style'],
      ['Ctrl+K', 'Insert link (Ctrl+Click to open)'], ['Ctrl+Enter', 'Page break'], ['Alt+Shift+D', 'Insert date'],
      ['Tab / Shift+Tab', 'Tab, list level, next table cell'], ['Delete', 'Delete selected picture or shape'],
      ['Ctrl+Shift+F / P', 'Jump to font / size box'], ['Ctrl+Mouse wheel', 'Zoom'],
    ];
    modal('<h3 style="margin-top:0">Keyboard Shortcuts</h3><table>' + rows.map(([k, v]) => `<tr><td>${k}</td><td>${v}</td></tr>`).join('') + '</table>', { cancel: null });
  }

  /* ================= Keyboard shortcuts ================= */
  document.addEventListener('keydown', (e) => {
    if (!$('#modal').hidden) return;
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key.toLowerCase();
    const code = e.code;
    let handled = true;

    if (e.key === 'Escape') {
      if (!findbar.hidden) closeFind();
      else if (state.painter) formatPainter();
      else selectObj(null);
      closeDropdowns();
      return;
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && state.selectedObj && !e.target.closest('input,select,textarea')) {
      const el = state.selectedObj;
      const sel = window.getSelection();
      const typingInBox = el.classList.contains('textbox') && sel.rangeCount && el.contains(sel.anchorNode);
      if (!typingInBox) {
        e.preventDefault();
        el.remove();
        selectObj(null);
        markDirty();
        return;
      }
    }
    if (e.key === 'F3' && e.shiftKey) { e.preventDefault(); changeCase(); return; }
    if (e.key === 'F1') { e.preventDefault(); showShortcuts(); return; }
    if (e.key === 'F12') { e.preventDefault(); run('saveAs'); return; }
    if (e.altKey && e.shiftKey && !ctrl && k === 'd') { e.preventDefault(); run('date'); return; }

    if (e.key === 'Tab' && inEditor(document.activeElement)) {
      e.preventDefault();
      const el = anchorElement();
      const cell = el.closest('td,th');
      if (cell) {
        const all = $$('td,th', cell.closest('table'));
        let i = all.indexOf(cell) + (e.shiftKey ? -1 : 1);
        if (i >= all.length) {
          const row = cell.closest('tr');
          const nr = row.cloneNode(true);
          $$('td,th', nr).forEach((c) => (c.innerHTML = '<br>'));
          row.after(nr);
          markDirty();
          i = all.length;
        }
        const target = $$('td,th', cell.closest('table'))[Math.max(0, i)];
        const r = document.createRange();
        r.selectNodeContents(target);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(r);
      } else if (el.closest('li')) {
        document.execCommand(e.shiftKey ? 'outdent' : 'indent');
        markDirty();
      } else if (!e.shiftKey) {
        document.execCommand('insertHTML', false, '<span style="white-space:pre">\t</span>');
        markDirty();
      }
      return;
    }

    if (!ctrl) return;

    if (e.altKey && ['1', '2', '3'].includes(e.key)) setBlock('h' + e.key);
    else if (e.altKey) handled = false;
    else if (e.shiftKey) {
      if (k === 'l') run('insertUnorderedList');
      else if (k === 'n') setBlock('p');
      else if (k === 'm') indent(-1);
      else if (e.key === '>' || code === 'Period') stepFont(1);
      else if (e.key === '<' || code === 'Comma') stepFont(-1);
      else if (e.key === '+' || code === 'Equal') run('superscript');
      else if (k === 'f') { $('.tab[data-tab="home"]').click(); $('#fontName').focus(); }
      else if (k === 'p') { $('.tab[data-tab="home"]').click(); $('#fontSize').focus(); }
      else if (k === 's') run('saveAs');
      else if (k === 'c') formatPainter();
      else handled = false;
    } else {
      switch (k) {
        case 'b': run('bold'); break;
        case 'i': run('italic'); break;
        case 'u': run('underline'); break;
        case 'l': run('justifyLeft'); break;
        case 'e': run('justifyCenter'); break;
        case 'r': run('justifyRight'); break;
        case 'j': run('justifyFull'); break;
        case 'm': indent(1); break;
        case 'k': run('link'); break;
        case 'n': run('new'); break;
        case 'o': run('open'); break;
        case 's': run('save'); break;
        case 'p': run('print'); break;
        case 'f': run('find'); break;
        case 'h': run('replace'); break;
        case '=': run('subscript'); break;
        case ']': stepFont(1, true); break;
        case '[': stepFont(-1, true); break;
        case ' ': exec('removeFormat'); break;
        case '1': setLineSpacing(1.0); break;
        case '2': setLineSpacing(2.0); break;
        case '5': setLineSpacing(1.5); break;
        case 'enter': insertPageBreak(); break;
        default: handled = false;
      }
    }
    if (handled) e.preventDefault();
  });

  /* ================= File operations ================= */
  async function confirmDiscard() {
    if (!state.dirty) return true;
    const r = await modal('<h3 style="margin-top:0">Save changes?</h3><p>Do you want to save the changes to ' + escapeHtml(docName()) + '?</p>', {
      ok: 'Save', extra: "Don't Save", cancel: 'Cancel',
    });
    if (r === 'ok') return saveDoc(false);
    return r === 'extra';
  }
  // Work out page size/orientation from saved or imported settings.
  function resolvePage(pg) {
    if (PAGE_SIZES[pg.size]) return { size: pg.size, landscape: !!pg.landscape };
    if (pg.widthMm && pg.heightMm) {
      const landscape = pg.widthMm > pg.heightMm + 1;
      const short = Math.min(pg.widthMm, pg.heightMm);
      const long = Math.max(pg.widthMm, pg.heightMm);
      const known = Object.keys(PAGE_SIZES).find((k) => k !== 'Custom' && Math.abs(PAGE_SIZES[k][0] - short) < 2.5 && Math.abs(PAGE_SIZES[k][1] - long) < 2.5);
      if (known) return { size: known, landscape };
      PAGE_SIZES.Custom = [Math.round(short * 10) / 10, Math.round(long * 10) / 10];
      const sel = $('#pageSize');
      let opt = sel.querySelector('option[value="Custom"]');
      if (!opt) { opt = new Option('', 'Custom'); sel.add(opt); }
      opt.textContent = `Custom (${(short / 10).toFixed(1)} × ${(long / 10).toFixed(1)} cm)`;
      return { size: 'Custom', landscape };
    }
    return { size: 'A4', landscape: false };
  }
  function setMarginOption(mm) {
    const sel = $('#margins');
    const v = String(mm);
    if (![...sel.options].some((o) => o.value === v)) {
      $$('option.custom', sel).forEach((o) => o.remove());
      const o = new Option(`Custom (${(mm / 10).toFixed(2)} cm)`, v);
      o.className = 'custom';
      sel.add(o);
    }
    sel.value = v;
  }
  function loadHtml(html, filePath, pg, suggestedName) {
    editor.innerHTML = html && html.trim() ? html : '<p><br></p>';
    // restore page layout saved with the document (or defaults)
    pg = pg || {};
    const ps = resolvePage(pg);
    state.page.size = ps.size;
    state.page.landscape = ps.landscape;
    state.page.marginMm = pg.marginMm || 25.4;
    state.suggestedName = suggestedName || null;
    $('#pageSize').value = state.page.size;
    $('#orientation').value = state.page.landscape ? 'landscape' : 'portrait';
    setMarginOption(state.page.marginMm);
    $('#columns').value = pg.columns || '1';
    editor.style.columnCount = pg.columns && pg.columns !== '1' ? pg.columns : '';
    $('#pageColor').value = pg.color || '#ffffff';
    page.style.backgroundColor = pg.color && pg.color !== '#ffffff' ? pg.color : '';
    applyPage();
    // older files: make shapes non-editable again
    $$('.shape', editor).forEach((s) => (s.contentEditable = 'false'));
    $$('.page-break', editor).forEach((s) => (s.contentEditable = 'false'));
    state.filePath = filePath || null;
    state.dirty = false;
    selectObj(null);
    updateTitle();
    updateStats();
    workspace.scrollTop = 0;
    editor.focus();
    const r = document.createRange();
    r.setStart(editor, 0);
    r.collapse(true);
    window.getSelection().removeAllRanges();
    window.getSelection().addRange(r);
  }
  async function newDoc() {
    if (!(await confirmDiscard())) return;
    loadHtml('');
  }
  async function openDoc() {
    if (!(await confirmDiscard())) return;
    const r = api.openFile && (await api.openFile());
    if (r) await loadResult(r);
  }
  api.onOpenPath && api.onOpenPath((r) => r && loadResult(r));

  // Show a document returned by the main process (HTML, or Word bytes to render).
  async function loadResult(r) {
    if (!r.docx) {
      loadHtml(r.html, r.filePath, r.page, r.suggestedName);
      return;
    }
    $('#saveState').textContent = 'Opening…';
    document.body.style.cursor = 'progress';
    try {
      const res = await window.DocImport.importDocx(r.docx);
      loadHtml(res.html, r.filePath, res.page, r.suggestedName);
    } catch (err) {
      console.error(err);
      const html = api.docxBasic ? await api.docxBasic(r.docx) : '';
      loadHtml(html, r.filePath, null, r.suggestedName);
      $('#saveState').textContent = 'Opened in simple mode';
    } finally {
      document.body.style.cursor = '';
    }
  }

  // Turn shapes into PNG pictures so they survive in Word.
  function shapeToPng(el) {
    return new Promise((resolve) => {
      const w = el.offsetWidth;
      const h = el.offsetHeight;
      const svg = el.querySelector('svg').cloneNode(true);
      svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      svg.setAttribute('width', w);
      svg.setAttribute('height', h);
      const url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(new XMLSerializer().serializeToString(svg));
      const img = new Image();
      img.onload = () => {
        const c = document.createElement('canvas');
        c.width = w * 2;
        c.height = h * 2;
        const ctx = c.getContext('2d');
        ctx.scale(2, 2);
        ctx.drawImage(img, 0, 0, w, h);
        resolve(`<img src="${c.toDataURL('image/png')}" width="${w}" height="${h}" style="width:${w}px;height:${h}px">`);
      };
      img.onerror = () => resolve('');
      img.src = url;
    });
  }
  async function htmlForDocx() {
    // tag objects so each copy can be matched to its on-screen original
    const liveObjs = $$('.shape, .textbox, .para-box', editor);
    liveObjs.forEach((e, i) => (e.dataset.xid = i));
    const clone = editor.cloneNode(true);
    liveObjs.forEach((e) => delete e.dataset.xid);
    const liveOf = (c) => liveObjs[+c.dataset.xid];
    // Word has no free-floating boxes here: move each one next to the paragraph it sits beside.
    const liveKids = Array.from(editor.children);
    const cloneKids = Array.from(clone.children);
    liveKids.forEach((el, i) => {
      if (!el.classList.contains('floating')) return;
      const top = el.offsetTop;
      let anchor = -1;
      liveKids.forEach((b, j) => {
        if (!b.classList.contains('floating') && b.offsetTop <= top + 4) anchor = j;
      });
      if (anchor >= 0) cloneKids[anchor].after(cloneKids[i]);
      else clone.prepend(cloneKids[i]);
    });
    const cl = $$('.shape', clone);
    for (let i = 0; i < cl.length; i++) {
      const png = await shapeToPng(liveOf(cl[i]));
      const p = document.createElement('p');
      p.innerHTML = png;
      if (cl[i].classList.contains('floating')) cl[i].replaceWith(p);
      else cl[i].outerHTML = png;
    }
    // Word export cannot read a <div> inside a table cell, so unwrap text box bodies.
    $$('.tb-body', clone).forEach((b) => b.replaceWith(...b.childNodes));
    // A box that only holds a table (e.g. "mark out of") exports as that table.
    $$('.textbox', clone).forEach((tb) => {
      const kids = Array.from(tb.children);
      if (kids.length === 1 && kids[0].tagName === 'TABLE') {
        const t = kids[0];
        t.style.width = tb.style.width;
        t.style.height = '';
        if (tb.classList.contains('floating') && parseFloat(tb.style.left) > editor.clientWidth / 2) t.style.marginLeft = 'auto';
        tb.replaceWith(t);
      }
    });
    // Text boxes and boxed paragraphs become one-cell tables in Word, keeping border and fill.
    $$('.textbox, .para-box', clone).reverse().forEach((tb) => {
      const cs = getComputedStyle(liveOf(tb));
      const border = cs.borderTopStyle === 'none' ? 'none' : `${cs.borderTopWidth} ${cs.borderTopStyle} ${cs.borderTopColor}`;
      const bg = TRANSPARENT_BG(cs.backgroundColor) ? '' : `background-color:${cs.backgroundColor};`;
      const t = document.createElement('table');
      t.style.borderCollapse = 'collapse';
      t.style.width = tb.classList.contains('para-box') ? '100%' : (tb.style.width || '220px');
      if (tb.classList.contains('floating') && parseFloat(tb.style.left) > editor.clientWidth / 2) t.style.marginLeft = 'auto';
      t.innerHTML = `<tbody><tr><td style="border:${border};padding:${cs.paddingTop} ${cs.paddingRight};${bg}vertical-align:top">${tb.innerHTML}</td></tr></tbody>`;
      tb.replaceWith(t);
    });
    $$('.page-break', clone).forEach((pb) => {
      pb.outerHTML = '<div class="page-break" style="page-break-after: always;"></div>';
    });
    $$('img', clone).forEach((img) => {
      const live = img.style.width;
      if (live && !img.getAttribute('width')) {
        const w = parseFloat(live);
        img.setAttribute('width', w);
        if (img.style.height) img.setAttribute('height', parseFloat(img.style.height));
      }
    });
    $$('.selected-obj', clone).forEach((x) => x.classList.remove('selected-obj'));
    $$('td', clone).forEach((td) => { if (!/border/.test(td.getAttribute('style') || '')) td.style.border = '1px solid #999999'; });
    $$('[contenteditable]', clone).forEach((x) => x.removeAttribute('contenteditable'));
    $$('[data-xid]', clone).forEach((x) => x.removeAttribute('data-xid'));
    // A lone <br> becomes an extra line in Word; empty blocks/cells are already one line tall.
    $$('p,li,td,th,h1,h2,h3,h4,h5,h6', clone).forEach((el) => {
      if (el.childNodes.length === 1 && el.firstChild.nodeName === 'BR') el.firstChild.remove();
    });
    // Table alignment for Word (the converter centres every table otherwise).
    const tableAlign = $$('table', clone).map((t) => {
      const l = t.style.marginLeft;
      const r = t.style.marginRight;
      if (l === 'auto' && r === 'auto') return 'center';
      return l === 'auto' ? 'right' : 'left';
    });
    return { html: clone.innerHTML, tableAlign };
  }
  async function saveDoc(saveAs) {
    if (!api.saveFile) return false;
    selectObj(null);
    const target = state.filePath || '';
    const isDocx = saveAs || !target || /\.docx$/i.test(target);
    const clean = editor.cloneNode(true);
    $$('.selected-obj', clean).forEach((x) => x.classList.remove('selected-obj'));
    const [w, h] = PAGE_SIZES[state.page.size];
    $('#saveState').textContent = 'Saving…';
    const result = await api.saveFile({
      filePath: state.filePath,
      suggestedName: state.suggestedName,
      saveAs,
      ...(isDocx ? await htmlForDocx() : { html: clean.innerHTML }),
      editorHtml: clean.innerHTML,
      text: editor.innerText,
      page: {
        widthMm: w, heightMm: h, marginMm: state.page.marginMm, landscape: state.page.landscape,
        size: state.page.size, columns: $('#columns').value, color: $('#pageColor').value,
      },
    });
    if (!result) { updateTitle(); return false; }
    state.suggestedName = null;
    state.filePath = result;
    state.dirty = false;
    updateTitle();
    return true;
  }
  async function exportPdf() {
    selectObj(null);
    applyPage();
    api.exportPdf && (await api.exportPdf(docName()));
  }
  function printDoc() {
    selectObj(null);
    applyPage();
    api.print ? api.print() : window.print();
  }

  /* ================= Utils ================= */
  function escapeHtml(s) {
    return String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }
  function escapeAttr(s) { return escapeHtml(s).replace(/'/g, '&#39;'); }
  function toHex(c) {
    if (!c) return '#000000';
    if (c[0] === '#') return c.length === 4 ? '#' + [...c.slice(1)].map((x) => x + x).join('') : c;
    const m = c.match(/\d+/g);
    return m ? '#' + m.slice(0, 3).map((n) => (+n).toString(16).padStart(2, '0')).join('') : '#000000';
  }

  // Paste images from clipboard as embedded pictures.
  editor.addEventListener('paste', (e) => {
    const items = e.clipboardData && e.clipboardData.items;
    if (!items) return;
    for (const it of items) {
      if (it.type.startsWith('image/')) {
        e.preventDefault();
        const f = it.getAsFile();
        const fr = new FileReader();
        fr.onload = () => exec('insertHTML', `<img src="${fr.result}" style="width:${Math.min(480, editor.clientWidth)}px">`);
        fr.readAsDataURL(f);
        return;
      }
    }
  });
  // Drop image files onto the page.
  editor.addEventListener('drop', (e) => {
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (!f || !f.type.startsWith('image/')) return;
    e.preventDefault();
    const pos = document.caretRangeFromPoint(e.clientX, e.clientY);
    if (pos) state.savedRange = pos;
    const fr = new FileReader();
    fr.onload = () => exec('insertHTML', `<img src="${fr.result}" style="width:${Math.min(480, editor.clientWidth)}px">`);
    fr.readAsDataURL(f);
  });

  /* ================= Start ================= */
  applyPage();
  setZoom(1);
  updateTitle();
  updateStats();
  editor.focus();
})();
