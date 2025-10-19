import { FORM_SCHEMA } from './formschema.js';
import { loadPrefill } from '../Imput/storage.js';
import { renderAccordion, $$, download } from './uiHelpers.js';

const DEFAULT_EVIDENCE_LAYOUT = {
  maxFilesPerSection: 10,
  defaults: {
    sheet: 'Evidence',
    startCell: 'B2',
    rowStride: 18,
    imageSize: { width: 320, height: 180 },
    linkCellOffset: { columns: 2, rows: 0 },
    header: { enabled: true, textPrefix: 'Section: ', sheet: null },
  },
  sections: {},
};

const readEvidenceLayout = () => {
  try {
    const cfg = window.ISA315_EVIDENCE_LAYOUT;
    return (cfg && typeof cfg === 'object') ? cfg : null;
  } catch (_) {
    return null;
  }
};

const ensurePositiveInteger = (value, fallback) => {
  const num = Number(value);
  if (!Number.isFinite(num)) return fallback;
  const floored = Math.floor(num);
  return floored > 0 ? floored : fallback;
};

const mergeObjects = (base, override) => {
  if (!override || typeof override !== 'object') return { ...base };
  return { ...base, ...override };
};

const sanitizeSegment = (input='') => {
  return String(input)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '') || 'section';
};

const buildSlotId = (sectionId, index) => `${sanitizeSegment(sectionId)}_${index + 1}`;

const sanitizeFileName = (name='', fallbackExt='') => {
  const cleaned = String(name).trim().replace(/[^a-z0-9_.-]+/gi, '_') || 'file';
  const hasExt = /\.[a-z0-9]+$/i.test(cleaned);
  if (hasExt) return cleaned;
  return fallbackExt ? `${cleaned}.${fallbackExt}` : cleaned;
};

const numberToColumn = (num=1) => {
  let col = '';
  let n = Math.max(1, Math.floor(num));
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col || 'A';
};

const columnToNumber = (col='') => {
  return col.toUpperCase().split('').reduce((sum, ch) => sum * 26 + (ch.charCodeAt(0) - 64), 0);
};

const parseCellAddress = (addr='') => {
  const match = /^([A-Z]+)(\d+)$/i.exec((addr || '').toString().trim());
  if (!match) return { col: 1, row: 1 };
  return { col: columnToNumber(match[1]), row: parseInt(match[2], 10) || 1 };
};

const arrayBufferToBase64 = (buffer) => {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
};

const getColumnOffset = (offset) => {
  if (!offset || typeof offset !== 'object') return 0;
  for (const key of ['columns', 'column', 'col', 'columnOffset']) {
    if (Number.isFinite(offset[key])) return offset[key];
  }
  return 0;
};

const getRowOffset = (offset) => {
  if (!offset || typeof offset !== 'object') return 0;
  for (const key of ['rows', 'row', 'rowOffset']) {
    if (Number.isFinite(offset[key])) return offset[key];
  }
  return 0;
};

const computeDefaultSlot = (sectionId, index, base) => {
  const anchor = parseCellAddress(base.startCell || 'B2');
  const stride = Number.isFinite(base.rowStride) ? base.rowStride : 18;
  const row = anchor.row + 1 + stride * index;
  const col = anchor.col;
  const colOffset = getColumnOffset(base.linkOffset);
  const rowOffset = getRowOffset(base.linkOffset);
  const linkCol = Math.max(1, col + colOffset);
  const linkRow = Math.max(1, row + rowOffset);
  return {
    id: buildSlotId(sectionId, index),
    sheet: base.sheet || 'Evidence',
    imageCell: `${numberToColumn(col)}${row}`,
    linkCell: `${numberToColumn(linkCol)}${linkRow}`,
    size: base.imageSize,
  };
};

const buildEvidenceLayout = (rawLayout) => {
  const layout = {
    maxFilesPerSection: ensurePositiveInteger(rawLayout?.maxFilesPerSection, DEFAULT_EVIDENCE_LAYOUT.maxFilesPerSection),
    defaults: { ...DEFAULT_EVIDENCE_LAYOUT.defaults },
    sections: {},
  };

  if (rawLayout?.defaults) {
    const overrides = rawLayout.defaults;
    if (overrides.sheet) layout.defaults.sheet = overrides.sheet;
    if (overrides.startCell) layout.defaults.startCell = overrides.startCell;
    if (Number.isFinite(overrides.rowStride)) layout.defaults.rowStride = overrides.rowStride;
    if (overrides.imageSize) layout.defaults.imageSize = mergeObjects(DEFAULT_EVIDENCE_LAYOUT.defaults.imageSize, overrides.imageSize);
    if (overrides.linkCellOffset) layout.defaults.linkCellOffset = mergeObjects(DEFAULT_EVIDENCE_LAYOUT.defaults.linkCellOffset, overrides.linkCellOffset);
    if (overrides.header) layout.defaults.header = mergeObjects(DEFAULT_EVIDENCE_LAYOUT.defaults.header, overrides.header);
  }

  if (!layout.defaults.imageSize) layout.defaults.imageSize = { width: 320, height: 180 };
  if (!layout.defaults.linkCellOffset) layout.defaults.linkCellOffset = { columns: 2, rows: 0 };
  if (!layout.defaults.header) layout.defaults.header = { enabled: true, textPrefix: 'Section: ', sheet: null };

  if (rawLayout?.sections && typeof rawLayout.sections === 'object') {
    Object.keys(rawLayout.sections).forEach(sectionId => {
      const section = rawLayout.sections[sectionId];
      if (!section || typeof section !== 'object') return;
      const copy = { ...section };
      if (Array.isArray(section.slots)) {
        copy.slots = section.slots.map(slot => (slot && typeof slot === 'object') ? { ...slot } : slot);
      }
      if (section.imageSize) copy.imageSize = mergeObjects(layout.defaults.imageSize, section.imageSize);
      if (section.linkCellOffset) copy.linkCellOffset = mergeObjects(layout.defaults.linkCellOffset, section.linkCellOffset);
      if (section.header) copy.header = mergeObjects(layout.defaults.header, section.header);
      layout.sections[sectionId] = copy;
    });
  }

  return layout;
};

const computeLayoutSignature = (rawLayout) => {
  try {
    return JSON.stringify(rawLayout || {});
  } catch (_) {
    return null;
  }
};

let CACHED_LAYOUT_SIGNATURE = null;
let CACHED_EVIDENCE_LAYOUT = buildEvidenceLayout(readEvidenceLayout() || {});
const SECTION_LAYOUT_CACHE = new Map();

const resolveEvidenceLayout = () => {
  const rawLayout = readEvidenceLayout() || {};
  const signature = computeLayoutSignature(rawLayout);
  if (!CACHED_EVIDENCE_LAYOUT || signature !== CACHED_LAYOUT_SIGNATURE) {
    CACHED_EVIDENCE_LAYOUT = buildEvidenceLayout(rawLayout);
    CACHED_EVIDENCE_LAYOUT.__signature = signature;
    CACHED_LAYOUT_SIGNATURE = signature;
    SECTION_LAYOUT_CACHE.clear();
  }
  return CACHED_EVIDENCE_LAYOUT;
};

const getMaxEvidenceFiles = (layout = resolveEvidenceLayout()) => {
  return ensurePositiveInteger(layout?.maxFilesPerSection, DEFAULT_EVIDENCE_LAYOUT.maxFilesPerSection);
};

try {
  if (typeof window !== 'undefined') {
    window.__ISA315_GET_EFFECTIVE_LAYOUT = resolveEvidenceLayout;
  }
} catch (_) { /* noop */ }

const getSectionLayout = (sectionId, layoutArg) => {
  const effectiveLayout = layoutArg || resolveEvidenceLayout();
  const signature = effectiveLayout.__signature || CACHED_LAYOUT_SIGNATURE || 'default';
  const cacheKey = `${signature}::${sectionId}`;
  if (SECTION_LAYOUT_CACHE.has(cacheKey)) return SECTION_LAYOUT_CACHE.get(cacheKey);
  const defaults = effectiveLayout.defaults || {};
  const override = (effectiveLayout.sections && effectiveLayout.sections[sectionId]) || {};
  const sheet = override.sheet || defaults.sheet || 'Evidence';
  const startCell = override.startCell || defaults.startCell || 'B2';
  const rowStride = Number.isFinite(override.rowStride) ? override.rowStride : (Number.isFinite(defaults.rowStride) ? defaults.rowStride : 18);
  const imageSize = override.imageSize || defaults.imageSize || { width: 320, height: 180 };
  const linkOffset = override.linkCellOffset || defaults.linkCellOffset || { columns: 2, rows: 0 };
  const headerDefaults = defaults.header || {};
  const headerOverride = override.header || {};
  const header = {
    enabled: headerOverride.enabled != null ? !!headerOverride.enabled : (headerDefaults.enabled != null ? !!headerDefaults.enabled : true),
    textPrefix: headerOverride.textPrefix || headerDefaults.textPrefix || 'Section: ',
    cell: headerOverride.cell || startCell,
    sheet: headerOverride.sheet || override.sheet || headerDefaults.sheet || sheet,
  };

  const slots = [];
  const overrideSlots = Array.isArray(override.slots) ? override.slots : [];

  const maxFiles = getMaxEvidenceFiles(effectiveLayout);

  for (let idx = 0; idx < maxFiles; idx++) {
    const slotOverride = overrideSlots[idx];
    if (slotOverride && typeof slotOverride === 'object') {
      const slotSize = slotOverride.size || slotOverride.imageSize || imageSize;
      const slotSheet = slotOverride.sheet || slotOverride.worksheet || sheet;
      const slotImageCell = slotOverride.imageCell || slotOverride.cell || startCell;
      let slotLinkCell = slotOverride.linkCell || slotOverride.link;
      if (!slotLinkCell) {
        const anchor = parseCellAddress(slotImageCell);
        const colOffset = getColumnOffset(linkOffset);
        const rowOffset = getRowOffset(linkOffset);
        slotLinkCell = `${numberToColumn(Math.max(1, anchor.col + colOffset))}${Math.max(1, anchor.row + rowOffset)}`;
      }
      slots.push({
        id: slotOverride.id || buildSlotId(sectionId, idx),
        sheet: slotSheet,
        imageCell: slotImageCell,
        linkCell: slotLinkCell,
        size: slotSize,
      });
    } else {
      slots.push(computeDefaultSlot(sectionId, idx, {
        sheet,
        startCell,
        rowStride,
        imageSize,
        linkOffset,
      }));
    }
  }

  const layout = { sheet, startCell, rowStride, imageSize, linkOffset, header, slots };
  SECTION_LAYOUT_CACHE.set(cacheKey, layout);
  return layout;
};

const splitFileName = (name='') => {
  const idx = name.lastIndexOf('.');
  if (idx <= 0) return { base: name, ext: '' };
  return { base: name.slice(0, idx), ext: name.slice(idx) };
};

const ensureUniqueFileName = (registry, sectionId, desiredName) => {
  const key = sectionId || '__default__';
  let bucket = registry.get(key);
  if (!bucket) {
    bucket = new Set();
    registry.set(key, bucket);
  }
  const lower = (candidate) => candidate.toLowerCase();
  let candidate = desiredName;
  const parts = splitFileName(desiredName);
  let counter = 2;
  while (bucket.has(lower(candidate))) {
    const suffix = ` (${counter++})`;
    candidate = `${parts.base || 'file'}${suffix}${parts.ext || ''}`;
  }
  bucket.add(lower(candidate));
  return candidate;
};

const collectEvidenceSections = (layoutArg) => {
  const layout = layoutArg || resolveEvidenceLayout();
  const maxFiles = getMaxEvidenceFiles(layout);
  const sections = [];
  $$('.accordion .acc-item').forEach(item => {
    const secId = item.getAttribute('data-sec') || '';
    if (!secId || secId === 'project_basics') return;
    const attach = item.querySelector('.q-attach');
    const files = attach?.__evidenceFiles ? Array.from(attach.__evidenceFiles).slice(0, maxFiles) : [];
    if (!files.length) return;
    const title = item.querySelector('.acc-title')?.textContent || secId;
    sections.push({ id: secId, title, files });
  });
  return sections;
};

const prepareEvidencePayload = async (sections, layoutArg) => {
  const layout = layoutArg || resolveEvidenceLayout();
  const maxFiles = getMaxEvidenceFiles(layout);
  const prepared = [];
  const usedNames = new Map();
  for (const section of sections) {
    const items = [];
    const sectionLayout = getSectionLayout(section.id, layout);
    const limit = Math.min(section.files.length, maxFiles);
    for (let idx = 0; idx < limit; idx++) {
      const file = section.files[idx];
      if (!file) continue;
      const slot = sectionLayout.slots[idx] || computeDefaultSlot(section.id, idx, {
        sheet: sectionLayout.sheet,
        startCell: sectionLayout.startCell,
        rowStride: sectionLayout.rowStride,
        imageSize: sectionLayout.imageSize,
        linkOffset: sectionLayout.linkOffset,
      });
      const buffer = await file.arrayBuffer();
      const type = file.type || '';
      const extGuess = type.startsWith('image/') ? (type.split('/')[1] || '').toLowerCase() : (file.name.split('.').pop() || '').toLowerCase();
      const normalizedExt = extGuess === 'jpg' ? 'jpeg' : extGuess;
      const fallbackExt = normalizedExt || (type === 'application/pdf' ? 'pdf' : '');
      const desiredName = sanitizeFileName(file.name, fallbackExt);
      const sectionFolder = sanitizeSegment(section.id || 'section');
      const finalName = ensureUniqueFileName(usedNames, section.id, desiredName);
      const relPath = `evidence/${sectionFolder}/${finalName}`;
      const entry = {
        file,
        buffer,
        type,
        extension: normalizedExt,
        finalName,
        relativePath: relPath,
        sectionId: section.id,
        sectionTitle: section.title,
        isImage: type.startsWith('image/'),
        name: file.name || finalName,
        sectionFolder,
        slotIndex: idx,
        slotId: slot.id || buildSlotId(section.id, idx),
      };
      if (entry.isImage) {
        entry.base64 = arrayBufferToBase64(buffer);
      }
      items.push(entry);
    }
    if (items.length) {
      prepared.push({ sectionId: section.id, sectionTitle: section.title, items });
    }
  }
  return prepared;
};

const applyEvidenceToWorkbook = (workbook, evidence, layoutArg) => {
  if (!Array.isArray(evidence) || !evidence.length) return;
  const layout = layoutArg || resolveEvidenceLayout();
  const sheetCache = new Map();

  const getSheet = (name) => {
    const key = name || 'Sheet1';
    if (sheetCache.has(key)) return sheetCache.get(key);
    let ws = workbook.getWorksheet(key);
    if (!ws) ws = workbook.addWorksheet(key);
    sheetCache.set(key, ws);
    return ws;
  };

  evidence.forEach(section => {
    const sectionLayout = getSectionLayout(section.sectionId, layout);
    const header = sectionLayout.header || {};
    if (header.enabled) {
      const headerSheet = getSheet(header.sheet || sectionLayout.sheet);
      const headerAddr = header.cell || sectionLayout.startCell;
      const headerCell = headerSheet.getCell(headerAddr);
      headerCell.value = `${header.textPrefix || 'Section: '}${section.sectionTitle}`;
      headerCell.font = { bold: true };
    }

    section.items.forEach((item, idx) => {
      const slot = sectionLayout.slots[idx] || computeDefaultSlot(section.sectionId, idx, {
        sheet: sectionLayout.sheet,
        startCell: sectionLayout.startCell,
        rowStride: sectionLayout.rowStride,
        imageSize: sectionLayout.imageSize,
        linkOffset: sectionLayout.linkOffset,
      });
      const ws = getSheet(slot.sheet || sectionLayout.sheet);
      const anchor = parseCellAddress(slot.imageCell || sectionLayout.startCell);
      const zeroAnchor = { col: Math.max(anchor.col - 1, 0), row: Math.max(anchor.row - 1, 0) };
      const size = slot.size || sectionLayout.imageSize || layout.defaults?.imageSize || { width: 320, height: 180 };
      const linkAddress = slot.linkCell || `${numberToColumn(Math.max(1, anchor.col + getColumnOffset(sectionLayout.linkOffset)))}${Math.max(1, anchor.row + getRowOffset(sectionLayout.linkOffset))}`;
      const linkCell = ws.getCell(linkAddress);
      linkCell.value = { text: item.name, hyperlink: item.relativePath };
      linkCell.font = { color: { argb: 'FF1F4E79' }, underline: true };
      linkCell.note = item.type || '';
      if (slot.id) item.slotId = slot.id;

      if (item.isImage && item.base64) {
        const imageId = workbook.addImage({ base64: item.base64, extension: item.extension || 'png' });
        ws.addImage(imageId, {
          tl: { col: zeroAnchor.col, row: zeroAnchor.row },
          ext: { width: size?.width || 320, height: size?.height || 180 },
        });
        const rowsCovered = Math.max(1, Math.ceil((size?.height || 180) / 20));
        for (let r = 0; r < rowsCovered; r++) {
          const excelRow = ws.getRow(anchor.row + r);
          if (!excelRow.height || excelRow.height < 60) excelRow.height = 60;
        }
      } else {
        const fallbackCell = ws.getCell(slot.imageCell || sectionLayout.startCell);
        fallbackCell.value = 'Preview not available';
        fallbackCell.font = { italic: true, color: { argb: 'FF6E6E6E' } };
      }
    });
  });
};

export function onManuel(){
  const root = document.querySelector('#manualRoot');
  if(!root) return;
  root.innerHTML = '';
  const PREFILL = loadPrefill();

  // Small debug line to help verify transport
  try{
    const info = document.createElement('div');
    info.className = 'small';
    const src = window.__PREFILL_SOURCE || 'none';
    const cnt = Object.keys(PREFILL||{}).length;
    const ts = localStorage.getItem('isa315_prefill_last_saved_at') || '';
    info.id = 'prefillDebug';
    info.textContent = `Prefill: ${cnt} items from ${src}${ts?` (last saved ${ts})`:''}`;
    root.appendChild(info);
  }catch(_){ }

  const acc = renderAccordion(FORM_SCHEMA, PREFILL);
  root.appendChild(acc);

  // Build section index with progress badges
  const secIndex = document.getElementById('sectionIndex');
  const secLinkMap = new Map();
  if (secIndex) {
    secIndex.innerHTML = '';
    $$('.accordion .acc-item').forEach(it => {
      const id = it.getAttribute('data-sec') || '';
      const title = it.querySelector('.acc-title')?.textContent || id;
      const a = document.createElement('a');
      a.href = `#${id}`; a.dataset.sec = id; a.textContent = title + ' ';
      const badge = document.createElement('span'); badge.className = 'badge'; badge.textContent = '';
      a.appendChild(badge);
      a.addEventListener('click', (e)=>{
        e.preventDefault();
        it.classList.add('open');
        it.scrollIntoView({ behavior:'smooth', block:'start' });
      });
      secIndex.appendChild(a);
      secLinkMap.set(id, { link:a, badge });
    });
    // update handler from renderer events
    document.addEventListener('sec-progress', (e)=>{
      const { id, filled, total, complete } = e.detail || {};
      const rec = secLinkMap.get(id);
      if (rec) {
        rec.badge.textContent = `${filled||0}/${total||0}`;
        rec.link.classList.toggle('complete', !!complete);
      }
    });
  }

  // Wire controls
  const expandAll = document.getElementById('btnExpandAll');
  const collapseAll = document.getElementById('btnCollapseAll');
  const chkCompact = document.getElementById('chkCompact');
  const saveJson = document.getElementById('btnSaveJson');
  const exportCsv = document.getElementById('btnExportCsv');
  const exportXlsx = document.getElementById('btnExportXlsx');

  expandAll?.addEventListener('click', ()=>{
    $$('.accordion .acc-item').forEach(it=> it.classList.add('open'));
  });
  collapseAll?.addEventListener('click', ()=>{
    $$('.accordion .acc-item').forEach(it=> it.classList.remove('open'));
  });
  chkCompact?.addEventListener('change', (e)=>{
    if(e.target.checked){
      acc.classList.add('compact');
      document.body.classList.add('compact','one-open');
    } else {
      acc.classList.remove('compact');
      document.body.classList.remove('compact','one-open');
    }
  });
  // ensure initial state reflects checkbox (in case of persisted state later)
  if (chkCompact && chkCompact.checked) {
    acc.classList.add('compact');
    document.body.classList.add('compact','one-open');
  } else {
    document.body.classList.remove('one-open');
  }

  const collectAnswers = () => {
    const out = {};
    $$('.accordion .acc-panel input, .accordion .acc-panel textarea, .accordion .acc-panel select').forEach(el=>{
      if(el.id) out[el.id] = el.value;
    });
    return out;
  };

  saveJson?.addEventListener('click', ()=>{
    const data = collectAnswers();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type:'application/json' });
    const url = URL.createObjectURL(blob);
    const name = `isa315_answers_${new Date().toISOString().slice(0,10)}.json`;
    download(url, name);
    setTimeout(()=> URL.revokeObjectURL(url), 3000);
  });

  // Optional stubs for future enhancements
  exportCsv?.addEventListener('click', ()=>{
    alert('CSV export coming soon. Use "Save answers (JSON)" for now.');
  });
  exportXlsx?.addEventListener('click', async ()=>{
    try{
      if (typeof ExcelJS === 'undefined'){
        alert('ExcelJS library not loaded. Please check your internet connection.');
        return;
      }
      if (typeof JSZip === 'undefined'){
        alert('JSZip library not loaded. Please check your internet connection.');
        return;
      }
      // Load default template from assets (always use repository template)
      const answers = collectAnswers();
      const layout = resolveEvidenceLayout();
      const evidenceSections = collectEvidenceSections(layout);
      const preparedEvidence = await prepareEvidencePayload(evidenceSections, layout);

      // Build sequential Question-N map based on schema
      const idsInOrder = FORM_SCHEMA
        .flatMap(s => (s.questions||[]).map(q => q.id));
      const numToId = new Map(); // 1-based
      idsInOrder.forEach((id, idx) => numToId.set(idx+1, id));

      // Helper: normalize Yes/No similar to import
      const normalizeYesNo = (v)=>{
        const s = String(v==null? '': v).trim().toLowerCase();
        if(['yes','y','true','ja','1'].includes(s)) return 'Yes';
        if(['no','n','false','nein','0'].includes(s)) return 'No';
        return String(v==null? '': v);
      };

      // 2) Load workbook
      const wb = new ExcelJS.Workbook();
      const loadTemplateBuffer = async ()=>{
        const buildCandidates = ()=>{
          const rel = ['assets/Output/Template.xlsx','./../assets/Output/Template.xlsx'];
          const abs = ['assets/Output/Template.xlsx'];
          try{
            abs.push(new URL('assets/Output/Template.xlsx', location.href).href);
            abs.push(new URL('./assets/Output/Template.xlsx', location.href).href);
          }catch(_){ }
          const all = [...rel, ...abs];
          return Array.from(new Set(all));
        };
        const candidates = buildCandidates();
        // Try fetch first
        for (const url of candidates){
          try{
            const resp = await fetch(url, { cache: 'no-cache' });
            if (resp && resp.ok){
              return await resp.arrayBuffer();
            }
          }catch(e){ /* try next */ }
        }
        // Fallback: try XMLHttpRequest (helps in some file:// contexts)
        const tryXhr = (url) => new Promise((resolve, reject) => {
          try{
            const xhr = new XMLHttpRequest();
            xhr.open('GET', url, true);
            xhr.responseType = 'arraybuffer';
            xhr.onload = function(){
              if (xhr.status === 200 || (xhr.status === 0 && xhr.response)){
                resolve(xhr.response);
              } else {
                reject(new Error('XHR status '+xhr.status));
              }
            };
            xhr.onerror = function(){ reject(new Error('XHR error')); };
            xhr.send();
          }catch(err){ reject(err); }
        });
        for (const url of candidates){
          try{
            const buf = await tryXhr(url);
            if (buf) return buf;
          }catch(e){ /* continue */ }
        }
        // Final fallback for file:// contexts: ask user to pick the template locally
        if (location.protocol === 'file:'){
          try{
            const buf = await new Promise((resolve, reject)=>{
              const inp = document.createElement('input');
              inp.type = 'file'; inp.accept = '.xlsx'; inp.style.display = 'none';
              inp.addEventListener('change', async ()=>{
                try{
                  if (!inp.files || !inp.files.length) return reject(new Error('no file'));
                  const file = inp.files[0];
                  const arr = await file.arrayBuffer();
                  resolve(arr);
                }catch(err){ reject(err); }
                finally { inp.remove(); }
              });
              document.body.appendChild(inp);
              inp.click();
              setTimeout(()=>{ /* user canceled */ try{ inp.remove(); }catch(_){} reject(new Error('picker canceled')); }, 30000);
            });
            if (buf) return buf;
          }catch(_){ /* ignore and fall through */ }
        }
        console.error('[ISA315] Default template not found via any path or method. Tried:', candidates);
        if (window.showError) { window.showError({ title:'Default template not found', message:'We could not locate assets/Output/Template.xlsx using any known path.', hint:'Ensure the template file exists in the assets folder. If you run this via file://, try a local webserver or pick the file manually when prompted.', code:'EXP-TPL-404' }); } else { alert('Default template not found: assets/Output/Template.xlsx'); }
        throw new Error('Default template missing');
      };
      const buf = await loadTemplateBuffer();
      await wb.xlsx.load(buf);

      // 3) Replace placeholders across all worksheets
      const reQ = /Question\s*(\d{1,3})/g; // e.g., "Question 1", "Question 01", "Question 123"

      wb.worksheets.forEach(ws => {
        ws.eachRow({ includeEmpty: true }, (row) => {
          row.eachCell({ includeEmpty: true }, (cell) => {
            if (cell == null) return;
            // Simple string values
            if (typeof cell.value === 'string'){
              let text = cell.value;
              const replaced = text.replace(reQ, (match, g1)=>{
                const n = parseInt(g1, 10);
                const id = numToId.get(n);
                if(!id) return match; // leave untouched if unknown
                const val = normalizeYesNo(answers[id] ?? '');
                return val || '';
              });
              if (replaced !== text) cell.value = replaced;
            } else if (cell.value && typeof cell.value === 'object' && cell.value.richText){
              // If richText, flatten to plain text, do replacement, then set as plain string (styling lost)
              const text = (cell.value.richText||[]).map(rt=> rt.text || '').join('');
              const replaced = text.replace(reQ, (match, g1)=>{
                const n = parseInt(g1, 10);
                const id = numToId.get(n);
                if(!id) return match;
                const val = normalizeYesNo(answers[id] ?? '');
                return val || '';
              });
              if (replaced !== text) cell.value = replaced;
            }
          });
        });
      });

      // 4) Apply evidence (images + hyperlinks)
      applyEvidenceToWorkbook(wb, preparedEvidence, layout);

      // 5) Prepare download bundle (Excel + evidence files)
      const outBuf = await wb.xlsx.writeBuffer();
      const zip = new JSZip();
      const dateTag = new Date().toISOString().slice(0,10);
      const excelName = `isa315_filled_${dateTag}.xlsx`;
      zip.file(excelName, outBuf);
      if (preparedEvidence.length){
        const evidenceRoot = zip.folder('evidence');
        preparedEvidence.forEach(section => {
          const secFolder = evidenceRoot.folder(section.items[0]?.sectionFolder || sanitizeSegment(section.sectionId));
          section.items.forEach(item => {
            if (!item || !item.buffer) return;
            secFolder.file(item.finalName, item.buffer);
          });
        });
      }

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const zipUrl = URL.createObjectURL(zipBlob);
      const zipName = `isa315_export_${dateTag}.zip`;
      download(zipUrl, zipName);
      setTimeout(()=> URL.revokeObjectURL(zipUrl), 3000);

    }catch(err){
      console.error('[ISA315][ExportXlsx] error:', err);
      if (window.showError) { window.showError({ title:'Excel export failed', message:'We could not generate the Excel file.', hint:'Verify Template.xlsx exists and is accessible. If running locally, try via a local webserver. Then try exporting again.', code:'EXP-WRITE-001', cause: (typeof err!=='undefined'?err:undefined) }); } else { alert('Excel export failed. Please check the template and try again.'); }
    }
  });
}

// Expose for non-module fallback
if (typeof window !== 'undefined') {
  window.onManuel = onManuel;
}
