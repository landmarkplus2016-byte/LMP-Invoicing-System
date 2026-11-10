// =============================================================================
// LMP Invoicing System — file-sources.js  ("My Files")
// Remembers where the user's files and folders are (File System Access API,
// Chrome / Edge) and loads them into every tab with one click.
//
//  • Each source is a file or folder handle saved in IndexedDB (this browser,
//    this computer). Nothing is uploaded anywhere.
//  • "Load My Files" re-reads the current version from disk and puts it into
//    each tab's own upload input — the same way drag & drop does (DataTransfer
//    + change event, or the `folderdrop` event for folder inputs) — so every
//    app loads it with its existing code.
//  • The tracking workbook is read and parsed once for all tabs: buffers of
//    source files are cached per file, and XLSX.read is cached per buffer.
// =============================================================================
(function () {
'use strict';

const SUPPORTED = typeof window.showOpenFilePicker === 'function' && typeof window.showDirectoryPicker === 'function';
const $ = id => document.getElementById(id);

// Sources and the inputs each one fills
const SOURCES = [
  { id: 'tracking', kind: 'file', label: 'Task Tracking File', hint: 'Total Task Tracking New 2026.xlsm', sheets: ['Invoicing Track'],
    usedBy: 'TSR Sub Prep · Validation · Sub Order · Planner · Contractor · Finance · Acceptance Check',
    inputs: ['tracking-input', 'tsrval-input-trk', 'tso-input-track', 'tsp-input-track',
             'con-track-input', 'con2-track-input', 'contxrx-track-input', 'fin-track-input', 'acc-input-track'] },
  { id: 'mails', kind: 'dir', label: 'Acceptance Mails Folder', hint: 'Signed PDF › 2026 (Upper-Cairo / Alex / Delta)',
    usedBy: 'Sub Order · Planner', folderInputs: ['tso-input-folder', 'tsp-input-folder'] },
  { id: 'accsheet', kind: 'file', label: 'Acceptance Sheet', hint: 'Total Acceptance Tracking.xlsx', sheets: ['Total'],
    usedBy: 'Acceptance Check', inputs: ['acc-input-acc'] },
  { id: 'poc', kind: 'file', label: 'POC3 Tracking File', hint: 'POC3 Tracking', sheets: ['POC3 Tracking'],
    usedBy: 'POC Invoice Prep · Contractor (POC) · Finance (POC)', inputs: ['fileInput', 'con-poc-track-input', 'poc-track-input'] },
  { id: 'tsrroot', kind: 'dir', label: 'TSR Root Folder', hint: 'Alya Sheets › TSR — then pick the TSR file below',
    usedBy: 'TSR file: Sub Prep · Validation · Sub Order · Planner — its Sub folder: Validation mails' }
];
const TSR_INPUTS    = ['tsr-input', 'tsrval-input-tsr', 'tso-input-tsr', 'tsp-input-tsr'];
const TSR_FOLDER_IN = ['tsrval-input-folder'];
const TSR_SHEETS    = ['Request Form - VF', 'PO Break Down- Contractor & Acc'];

// A wrong pick (e.g. another workbook from the same folder) must never reach the
// tabs: check the sheet names first (bookSheets = names only, fast even for 5 MB).
async function missingSheets(file, need, buf) {
  if (!need || !need.length || !window.XLSX) return { missing: [], names: [] };
  const data = buf || await _origArrayBuffer.call(file);
  let names = [];
  try { names = XLSX.read(data, { type: 'array', bookSheets: true }).SheetNames || []; } catch { names = []; }
  const have = new Set(names.map(n => n.trim().toLowerCase()));
  return { missing: need.filter(n => !have.has(n.toLowerCase())), names };
}
function wrongFileText(file, need, check) {
  return 'Wrong file: "' + file.name + '" has no ' + check.missing.map(m => '"' + m + '"').join(' / ') +
    ' sheet' + (check.names.length ? ' (it has: ' + check.names.slice(0, 6).join(', ') + (check.names.length > 6 ? '…' : '') + ')' : '') +
    '. Click Change and pick the right file.';
}

// ---------------------------------------------------------------------------
// IndexedDB — handles are structured-cloneable
// ---------------------------------------------------------------------------
const DB_NAME = 'lmp-invoicing', STORE = 'sources';
function db() {
  return new Promise((res, rej) => {
    const r = indexedDB.open(DB_NAME, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(STORE);
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function idb(mode, fn) {
  const d = await db();
  return new Promise((res, rej) => {
    const tx = d.transaction(STORE, mode);
    const req = fn(tx.objectStore(STORE));
    tx.oncomplete = () => res(req && req.result);
    tx.onerror = () => rej(tx.error);
  });
}
const idbGet = k => idb('readonly', s => s.get(k));
const idbSet = (k, v) => idb('readwrite', s => s.put(v, k));
const idbDel = k => idb('readwrite', s => s.delete(k));

// ---------------------------------------------------------------------------
// Shared read: one buffer per source file, one parse per buffer + options
// ---------------------------------------------------------------------------
const _buffers = new Map();          // "name|size|lastModified" → Promise<ArrayBuffer>
const _shared  = new WeakSet();      // buffers that belong to source files
const _parsed  = new WeakMap();      // buffer → Map(optionsKey → workbook)
const fileKey  = f => f.name + '|' + f.size + '|' + f.lastModified;

const _origArrayBuffer = Blob.prototype.arrayBuffer;
File.prototype.arrayBuffer = function () {
  const p = _buffers.get(fileKey(this));
  return p || _origArrayBuffer.call(this);
};
function registerFile(file) {
  if (_buffers.has(fileKey(file))) return;
  const p = _origArrayBuffer.call(file).then(buf => { _shared.add(buf); return buf; });
  _buffers.set(fileKey(file), p);
}
function wrapXlsxRead() {
  if (!window.XLSX || XLSX.__lmpShared) return;
  const orig = XLSX.read;
  XLSX.read = function (data, opts) {
    if (!(data instanceof ArrayBuffer) || !_shared.has(data)) return orig.call(this, data, opts);
    const key = JSON.stringify(opts || {});
    let m = _parsed.get(data);
    if (!m) { m = new Map(); _parsed.set(data, m); }
    if (!m.has(key)) m.set(key, orig.call(this, data, opts));
    return m.get(key);
  };
  XLSX.__lmpShared = true;
}

// ---------------------------------------------------------------------------
// State + permissions
// ---------------------------------------------------------------------------
const state = {};                    // id → { handle, status, note }
let _tsrList = [];                   // [{ path, handle, dirHandle, dirPath, mtime }]
let _tsrSelected = '';

async function ensurePermission(handle, ask) {
  const opts = { mode: 'read' };
  try {
    if ((await handle.queryPermission(opts)) === 'granted') return 'granted';
    if (!ask) return 'prompt';
    return (await handle.requestPermission(opts)) === 'granted' ? 'granted' : 'denied';
  } catch (e) {
    return e && e.name === 'SecurityError' ? 'needs-click' : 'error';
  }
}

async function loadSaved() {
  for (const s of SOURCES) {
    const handle = await idbGet('h:' + s.id).catch(() => null);
    state[s.id] = { handle: handle || null, status: handle ? 'saved' : 'unset', note: '' };
  }
  _tsrSelected = (await idbGet('tsr:selected').catch(() => '')) || '';
}

// ---------------------------------------------------------------------------
// Folder reading
// ---------------------------------------------------------------------------
async function walk(dir, prefix, depth, out, filter) {
  for await (const entry of dir.values()) {
    if (entry.name.startsWith('~$') || entry.name.startsWith('.')) continue;
    if (entry.kind === 'file') {
      if (!filter || filter(entry.name)) out.push({ entry, path: prefix + '/' + entry.name });
    } else if (depth > 0) {
      await walk(entry, prefix + '/' + entry.name, depth - 1, out, filter);
    }
  }
  return out;
}

// File objects with webkitRelativePath = "Root/sub/…/name" (what folder inputs give)
async function folderFiles(dirHandle, rootName, depth, onlySub) {
  const found = [];
  for await (const entry of dirHandle.values()) {
    if (entry.name.startsWith('~$') || entry.name.startsWith('.')) continue;
    if (entry.kind === 'directory') {
      if (onlySub && !onlySub(entry.name)) continue;
      await walk(entry, rootName + '/' + entry.name, depth - 1, found);
    } else if (!onlySub) {
      found.push({ entry, path: rootName + '/' + entry.name });
    }
  }
  const files = [];
  for (const f of found) {
    const file = await f.entry.getFile();
    Object.defineProperty(file, 'webkitRelativePath', { value: f.path, configurable: true });
    files.push(file);
  }
  return files;
}

// TSR files under the TSR root: Root / <TSR#> / <Sub-n> / *.xlsb (up to 3 levels)
async function scanTsr(root) {
  const out = [];
  const found = await walk(root, '', 2, [], n => /\.xlsb$/i.test(n));
  for (const f of found) {
    const parts = f.path.split('/').filter(Boolean);
    let dirHandle = root;
    for (const p of parts.slice(0, -1)) dirHandle = await dirHandle.getDirectoryHandle(p);
    const file = await f.entry.getFile();
    out.push({ path: parts.join('/'), handle: f.entry, dirHandle, dirPath: parts.slice(0, -1).join('/'), mtime: file.lastModified });
  }
  out.sort((a, b) => b.mtime - a.mtime);
  return out;
}

// ---------------------------------------------------------------------------
// Filling the tabs' inputs
// ---------------------------------------------------------------------------
function setInputFile(id, file) {
  const input = $(id);
  if (!input) return false;
  const dt = new DataTransfer();
  dt.items.add(file);
  input.files = dt.files;
  input.dispatchEvent(new Event('change', { bubbles: true }));
  return true;
}
function setInputFolder(id, files) {
  const input = $(id);
  if (!input) return false;
  input.dispatchEvent(new CustomEvent('folderdrop', { detail: files }));
  return true;
}
const pause = ms => new Promise(r => setTimeout(r, ms));

// ---------------------------------------------------------------------------
// Load everything
// ---------------------------------------------------------------------------
let _busy = false;
async function loadAll() {
  if (!SUPPORTED) { showTab(); return; }
  if (_busy) return;
  const set = SOURCES.filter(s => state[s.id] && state[s.id].handle);
  if (!set.length) { showTab(); setStatus('Set your files first — pick each one once below.'); return; }
  _busy = true;
  wrapXlsxRead();
  setBusy(true);
  try {
    // 1. permissions (one browser prompt per source the first time in a session)
    const blocked = [];
    for (const s of set) {
      const p = await ensurePermission(state[s.id].handle, true);
      state[s.id].status = p === 'granted' ? 'saved' : 'blocked';
      state[s.id].note = p === 'granted' ? '' : (p === 'needs-click' ? 'click Load again to allow' : 'access not allowed');
      if (p !== 'granted') blocked.push(s.label);
    }
    render();
    if (blocked.length && blocked.length === set.length) {
      setStatus('Chrome needs your OK to read: ' + blocked.join(', ') + '. Click Load My Files again and choose Allow.');
      return;
    }

    // 2. read files from disk (once) and hand them to every tab
    let done = 0;
    const log = [];
    for (const s of set) {
      const st = state[s.id];
      if (st.status !== 'saved') continue;
      try {
        if (s.kind === 'file') {
          setStatus('Reading ' + st.handle.name + '…');
          const file = await st.handle.getFile();
          registerFile(file);
          const buf = await file.arrayBuffer();           // read once, now
          const chk = await missingSheets(file, s.sheets, buf);
          if (chk.missing.length) { st.status = 'error'; st.note = wrongFileText(file, s.sheets, chk); render(); continue; }
          for (const id of s.inputs) { setInputFile(id, file); await pause(30); }
          st.note = 'loaded ' + fmtTime(file.lastModified);
          log.push(file.name);
        } else if (s.id === 'mails') {
          setStatus('Reading mails folder ' + st.handle.name + '…');
          const files = await folderFiles(st.handle, st.handle.name, 3);
          for (const id of s.folderInputs) { setInputFolder(id, files); await pause(30); }
          st.note = files.length + ' files';
          log.push(st.handle.name + ' (' + files.length + ' files)');
        } else if (s.id === 'tsrroot') {
          setStatus('Looking for TSR files in ' + st.handle.name + '…');
          _tsrList = await scanTsr(st.handle);
          renderTsrSelect();
          const pick = _tsrList.find(t => t.path === _tsrSelected) || _tsrList[0];
          if (!pick) { st.note = 'no .xlsb file found'; continue; }
          const file = await pick.handle.getFile();
          registerFile(file);
          const buf = await file.arrayBuffer();
          const chk = await missingSheets(file, TSR_SHEETS, buf);
          if (chk.missing.length) { st.status = 'error'; st.note = wrongFileText(file, TSR_SHEETS, chk) + ' (or choose another TSR file below)'; render(); continue; }
          for (const id of TSR_INPUTS) { setInputFile(id, file); await pause(30); }
          // its Sub folder's numbered mail folders → TSR Sub Validation
          const subName = pick.dirPath.split('/').pop() || st.handle.name;
          const subFiles = await folderFiles(pick.dirHandle, subName, 2, n => /^\d+$/.test(n.trim()));
          if (subFiles.length) for (const id of TSR_FOLDER_IN) setInputFolder(id, subFiles);
          st.note = pick.path + (subFiles.length ? ' + ' + new Set(subFiles.map(f => f.webkitRelativePath.split('/')[1])).size + ' mail folders' : '');
          log.push(pick.path);
        }
        st.status = 'loaded';
        done++;
      } catch (e) {
        st.status = 'error';
        st.note = e && e.name === 'NotFoundError' ? 'not found — moved or renamed? Set it again' : (e.message || String(e));
      }
      render();
    }
    const when = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    setStatus('Loaded ' + done + ' of ' + set.length + ' at ' + when + (blocked.length ? ' — still needs your OK: ' + blocked.join(', ') : '') + '.');
    $('btn-load-sources-note').textContent = 'Loaded ' + when;
  } finally {
    _busy = false;
    setBusy(false);
    render();
  }
}

// ---------------------------------------------------------------------------
// Picking / clearing sources
// ---------------------------------------------------------------------------
async function pick(id) {
  const s = SOURCES.find(x => x.id === id);
  try {
    let handle;
    if (s.kind === 'file') {
      [handle] = await window.showOpenFilePicker({ id: 'lmp-' + id, multiple: false,
        types: [{ description: 'Excel', accept: { 'application/octet-stream': ['.xlsm', '.xlsx', '.xls', '.xlsb'] } }] });
    } else {
      handle = await window.showDirectoryPicker({ id: 'lmp-' + id, mode: 'read' });
    }
    if (s.kind === 'file' && s.sheets) {
      setStatus('Checking ' + handle.name + '…');
      const file = await handle.getFile();
      const chk = await missingSheets(file, s.sheets);
      if (chk.missing.length) {
        const msg = wrongFileText(file, s.sheets, chk).replace('Click Change and pick the right file.', 'Not saved — pick ' + s.hint + '.');
        if (state[id]) state[id].note = msg;
        setStatus('✖ ' + s.label + ': ' + msg);
        render();
        return;
      }
    }
    await idbSet('h:' + id, handle);
    state[id] = { handle, status: 'saved', note: 'set — load to read it' };
    setStatus('✔ ' + s.label + ' set to ' + handle.name + '.');
    if (id === 'tsrroot') { _tsrList = await scanTsr(handle); renderTsrSelect(); }
    render();
  } catch (e) {
    if (e && e.name !== 'AbortError') setStatus('Could not set ' + s.label + ': ' + e.message);
  }
}
async function clear(id) {
  await idbDel('h:' + id);
  state[id] = { handle: null, status: 'unset', note: '' };
  if (id === 'tsrroot') { _tsrList = []; renderTsrSelect(); }
  render();
}

// ---------------------------------------------------------------------------
// UI
// ---------------------------------------------------------------------------
function fmtTime(ms) {
  const d = new Date(ms);
  return d.toLocaleDateString([], { day: '2-digit', month: 'short' }) + ' ' + d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
}
function setStatus(t) { const e = $('src-status'); if (e) e.textContent = t; }
function setBusy(b) {
  [$('btn-load-sources'), $('src-btn-load')].forEach(x => { if (x) { x.disabled = b; x.classList.toggle('is-busy', b); } });
}
function showTab() {
  const btn = document.querySelector('.nav-item[data-tab="sources"]');
  if (btn) btn.click();
}
function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

const BADGE = { unset: ['Not set', 'src-badge-off'], saved: ['Saved', 'src-badge-saved'], loaded: ['Loaded', 'src-badge-ok'],
                blocked: ['Needs OK', 'src-badge-warn'], error: ['Problem', 'src-badge-bad'] };
function render() {
  const tb = $('src-rows');
  if (!tb) return;
  tb.innerHTML = SOURCES.map(s => {
    const st = state[s.id] || { status: 'unset' };
    const [txt, cls] = BADGE[st.status] || BADGE.unset;
    const where = st.handle ? (s.kind === 'dir' ? '📁 ' : '📄 ') + esc(st.handle.name) : '<span class="acc-sub">' + esc(s.hint) + '</span>';
    return '<tr><td><strong>' + esc(s.label) + '</strong><div class="acc-sub">' + esc(s.usedBy) + '</div></td>' +
      '<td>' + where + (st.note ? '<div class="acc-note">' + esc(st.note) + '</div>' : '') + '</td>' +
      '<td><span class="src-badge ' + cls + '">' + txt + '</span></td>' +
      '<td class="src-actions"><button class="btn-secondary src-pick" data-id="' + s.id + '">' + (st.handle ? 'Change' : 'Set') + '</button>' +
      (st.handle ? ' <button class="btn-secondary src-clear" data-id="' + s.id + '">Clear</button>' : '') + '</td></tr>';
  }).join('');
  tb.querySelectorAll('.src-pick').forEach(b => b.addEventListener('click', () => pick(b.dataset.id)));
  tb.querySelectorAll('.src-clear').forEach(b => b.addEventListener('click', () => clear(b.dataset.id)));
}
function renderTsrSelect() {
  const sel = $('src-tsr-select');
  if (!sel) return;
  const wrap = $('src-tsr-wrap');
  wrap.style.display = state.tsrroot && state.tsrroot.handle ? '' : 'none';
  if (!_tsrList.length) {
    sel.innerHTML = '<option value="">' + (state.tsrroot && state.tsrroot.handle ? 'Click Load My Files to list the TSR files' : '—') + '</option>';
    return;
  }
  const cur = _tsrList.find(t => t.path === _tsrSelected) ? _tsrSelected : _tsrList[0].path;
  sel.innerHTML = _tsrList.map(t => '<option value="' + esc(t.path) + '"' + (t.path === cur ? ' selected' : '') + '>' +
    esc(t.path.split('/').join('  ›  ')) + '   (' + fmtTime(t.mtime) + ')</option>').join('');
}

function init() {
  const panel = $('panel-sources');
  if (!panel) return;
  if (!SUPPORTED) {
    $('src-unsupported').style.display = 'block';
    $('src-main').style.display = 'none';
    const sb = $('btn-load-sources');
    if (sb) sb.title = 'Remembering files needs Chrome or Edge';
    return;
  }
  $('btn-load-sources').addEventListener('click', loadAll);
  $('src-btn-load').addEventListener('click', loadAll);
  $('src-tsr-select').addEventListener('change', async e => {
    _tsrSelected = e.target.value;
    await idbSet('tsr:selected', _tsrSelected);
    setStatus('TSR file set to ' + _tsrSelected.split('/').join(' › ') + ' — click Load My Files to load it.');
  });
  loadSaved().then(() => { render(); renderTsrSelect(); }).catch(e => setStatus('Saved files could not be read: ' + e.message));
}

// test hook (headless browser tests set handles without the picker)
window.LMPSources = {
  async setHandle(id, handle) { await idbSet('h:' + id, handle); state[id] = { handle, status: 'saved', note: '' }; render(); },
  async selectTsr(path) { _tsrSelected = path; await idbSet('tsr:selected', path); },
  loadAll, get state() { return state; }, get tsrList() { return _tsrList; }
};

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
