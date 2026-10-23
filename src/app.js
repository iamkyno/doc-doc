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
    return state.filePath ? state.filePath.split(/[\\/]/).pop() : 'Document1';
  }
  function updateTitle() {
    const name = docName();
    $('#docname').textContent = name + (state.dirty ? ' •' : '');
    $('#saveState').textContent = state.dirty ? 'Unsaved changes' : (state.filePath ? 'Saved' : '');
    api.setTitle && api.setTitle(`${name}${state.dirty ? '*' : ''} - DocDoc`);
    api.setDirty && api.setDirty(state.dirty);
  }
  function markDirty() {
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
  $('#fontName').addEventListener('change', (e) => exec('fontName', e.target.value));
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
    const fn = $('#fontName');
    if (![...fn.options].some((o) => o.value === fam)) fn.add(new Option(fam));
    fn.value = fam;
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
  Object.keys(SHAPES).forEach((k) => {
    const b = document.createElement('button');
    b.title = k;
    b.innerHTML = `<svg viewBox="0 0 100 100" preserveAspectRatio="none"><g fill="#cfd9ec" stroke="#2f528f" stroke-width="6">${SHAPES[k]}</g></svg>`;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => { closeDropdowns(); insertShape(k); });
    shapeMenu.appendChild(b);
  });

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
    const fill = $('#shapeFill').value;
    const stroke = $('#shapeStroke').value;
    const el = document.createElement('span');
    el.className = 'shape floating';
    el.contentEditable = 'false';
    el.dataset.shape = kind;
    const isLine = kind === 'line' || kind === 'arrowLine';
    const top = visibleTop();
    el.style.cssText = `left:${60 + (insertCount % 6) * 24}px;top:${top}px;width:${isLine ? 160 : 140}px;height:${kind === 'arrowLine' ? 24 : isLine ? 100 : 110}px`;
    el.innerHTML = shapeSvg(kind, fill, stroke);
    $$('[data-solid]', el).forEach((p) => p.setAttribute('fill', stroke));
    editor.appendChild(el);
    selectObj(el);
    markDirty();
  }
  function insertTextBox() {
    const el = document.createElement('div');
    el.className = 'textbox floating';
    el.style.cssText = `left:60px;top:${visibleTop()}px;width:220px`;
    el.innerHTML = '<p>Type your text here</p>';
    editor.appendChild(el);
    selectObj(el);
    const r = document.createRange();
    r.selectNodeContents(el.firstChild);
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
  function selectObj(el) {
    if (state.selectedObj === el) return positionHandles();
    if (state.selectedObj) state.selectedObj.classList.remove('selected-obj');
    state.selectedObj = el;
    if (el) el.classList.add('selected-obj');
    $('#objGroup').classList.toggle('disabled', !el);
    if (el && el.classList.contains('shape')) {
      const g = el.querySelector('g');
      if (g.getAttribute('fill') !== 'none') $('#shapeFill').value = toHex(g.getAttribute('fill'));
      $('#shapeStroke').value = toHex(g.getAttribute('stroke'));
    }
    positionHandles();
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
    if (mode === 'front' || mode === 'back') {
      el.style.zIndex = mode === 'front' ? 10 : 0;
      if (mode === 'back' && el.classList.contains('floating')) el.style.zIndex = 0;
    } else {
      el.classList.remove('wrap-left', 'wrap-right');
      if (el.classList.contains('floating') && el.tagName !== 'IMG') {
        // move a floating shape/text box back into the text flow at the caret
        el.classList.remove('floating');
        el.style.left = el.style.top = '';
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
  $('#shapeFill').addEventListener('input', (e) => {
    const el = state.selectedObj;
    if (!el) return;
    if (el.classList.contains('shape')) {
      const g = el.querySelector('g');
      if (g.getAttribute('fill') !== 'none') g.setAttribute('fill', e.target.value);
    } else if (el.classList.contains('textbox')) el.style.background = e.target.value;
    markDirty();
  });
  $('#shapeStroke').addEventListener('input', (e) => {
    const el = state.selectedObj;
    if (!el) return;
    if (el.classList.contains('shape')) {
      el.querySelector('g').setAttribute('stroke', e.target.value);
      $$('[data-solid]', el).forEach((p) => p.setAttribute('fill', e.target.value));
    } else if (el.classList.contains('textbox')) el.style.borderColor = e.target.value;
    markDirty();
  });
  $('#objGroup').classList.add('disabled');

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

    if (e.key === 'Tab' && document.activeElement === editor) {
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
  function loadHtml(html, filePath, pg) {
    editor.innerHTML = html && html.trim() ? html : '<p><br></p>';
    // restore page layout saved with the document (or defaults)
    pg = pg || {};
    state.page.size = PAGE_SIZES[pg.size] ? pg.size : 'A4';
    state.page.landscape = !!pg.landscape;
    state.page.marginMm = pg.marginMm || 25.4;
    $('#pageSize').value = state.page.size;
    $('#orientation').value = state.page.landscape ? 'landscape' : 'portrait';
    $('#margins').value = String(state.page.marginMm);
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
    if (r) loadHtml(r.html, r.filePath, r.page);
  }
  api.onOpenPath && api.onOpenPath((r) => r && loadHtml(r.html, r.filePath, r.page));

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
    const clone = editor.cloneNode(true);
    const live = $$('.shape', editor);
    const cl = $$('.shape', clone);
    for (let i = 0; i < cl.length; i++) {
      const png = await shapeToPng(live[i]);
      const p = document.createElement('p');
      p.innerHTML = png;
      if (cl[i].classList.contains('floating')) cl[i].replaceWith(p);
      else cl[i].outerHTML = png;
    }
    $$('.textbox', clone).forEach((tb) => {
      const t = document.createElement('table');
      t.setAttribute('border', '1');
      t.style.width = tb.style.width || '220px';
      t.innerHTML = `<tbody><tr><td style="border:1px solid #000;padding:6px;${tb.style.background ? 'background:' + tb.style.background : ''}">${tb.innerHTML}</td></tr></tbody>`;
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
    $$('td', clone).forEach((td) => { if (!td.style.border) td.style.border = '1px solid #999999'; });
    $$('[contenteditable]', clone).forEach((x) => x.removeAttribute('contenteditable'));
    return clone.innerHTML;
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
      saveAs,
      html: isDocx ? await htmlForDocx() : clean.innerHTML,
      editorHtml: clean.innerHTML,
      text: editor.innerText,
      page: {
        widthMm: w, heightMm: h, marginMm: state.page.marginMm, landscape: state.page.landscape,
        size: state.page.size, columns: $('#columns').value, color: $('#pageColor').value,
      },
    });
    if (!result) { updateTitle(); return false; }
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
