import { FORM_SCHEMA } from './formschema.js';
import { loadPrefill } from '../Imput/storage.js';
import { renderAccordion, $$, download } from './uiHelpers.js';

// Evidence export configuration – adjust to control placement and sizing in the template
const EVIDENCE_EXPORT_CONFIG = {
  maxFilesPerSection: 10,
  defaultSheet: 'Evidence',
  defaultStartCell: 'B2',
  defaultImageSize: { width: 320, height: 180 },
  defaultRowStride: 18,
  linkColumnOffset: 2,
  sections: {
    // Example:
    // 'IT Environment Overview': { sheet: 'Evidence', startCell: 'B2', imageSize: { width: 300, height: 170 }, rowStride: 18, linkColumnOffset: 2 },
  },
};

const sanitizeSegment = (input='') => {
  return String(input)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '') || 'section';
};

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
  const match = /^([A-Z]+)(\d+)$/i.exec(addr.trim());
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

const collectEvidenceSections = () => {
  const sections = [];
  $$('.accordion .acc-item').forEach(item => {
    const secId = item.getAttribute('data-sec') || '';
    if (!secId || secId === 'project_basics') return;
    const attach = item.querySelector('.q-attach');
    const files = attach?.__evidenceFiles ? Array.from(attach.__evidenceFiles).slice(0, EVIDENCE_EXPORT_CONFIG.maxFilesPerSection) : [];
    if (!files.length) return;
    const title = item.querySelector('.acc-title')?.textContent || secId;
    sections.push({ id: secId, title, files });
  });
  return sections;
};

const prepareEvidencePayload = async (sections) => {
  const prepared = [];
  for (const section of sections) {
    const items = [];
    for (let idx = 0; idx < section.files.length && idx < EVIDENCE_EXPORT_CONFIG.maxFilesPerSection; idx++) {
      const file = section.files[idx];
      if (!file) continue;
      const buffer = await file.arrayBuffer();
      const type = file.type || '';
      const extGuess = type.startsWith('image/') ? (type.split('/')[1] || '').toLowerCase() : (file.name.split('.').pop() || '').toLowerCase();
      const normalizedExt = extGuess === 'jpg' ? 'jpeg' : extGuess;
      const finalName = sanitizeFileName(file.name, normalizedExt || (type === 'application/pdf' ? 'pdf' : ''));
      const sectionFolder = sanitizeSegment(section.id || 'section');
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

const applyEvidenceToWorkbook = (workbook, evidence, config = EVIDENCE_EXPORT_CONFIG) => {
  if (!Array.isArray(evidence) || !evidence.length) return;
  const sheetCache = new Map();
  const defaults = {
    sheet: config.defaultSheet,
    startCell: config.defaultStartCell,
    imageSize: config.defaultImageSize,
    rowStride: config.defaultRowStride,
    linkColumnOffset: config.linkColumnOffset,
  };

  const defaultAnchor = parseCellAddress(config.defaultStartCell || 'B2');
  let nextDefaultRow = defaultAnchor.row;

  const getSheet = (name) => {
    if (sheetCache.has(name)) return sheetCache.get(name);
    let ws = workbook.getWorksheet(name);
    if (!ws) ws = workbook.addWorksheet(name);
    sheetCache.set(name, ws);
    return ws;
  };

  const getSettings = (sectionId) => {
    const override = config.sections?.[sectionId] || {};
    const settings = {
      sheet: override.sheet || defaults.sheet,
      startCell: override.startCell || defaults.startCell,
      imageSize: override.imageSize || defaults.imageSize,
      rowStride: override.rowStride || defaults.rowStride,
      linkColumnOffset: typeof override.linkColumnOffset === 'number' ? override.linkColumnOffset : defaults.linkColumnOffset,
    };
    if (!override.startCell) {
      settings.startCell = `${numberToColumn(defaultAnchor.col)}${nextDefaultRow}`;
    }
    return settings;
  };

  const toZeroBasedAnchor = (addr) => {
    const { col, row } = parseCellAddress(addr || 'A1');
    return { col: Math.max(col - 1, 0), row: Math.max(row - 1, 0) };
  };

  evidence.forEach(section => {
    const settings = getSettings(section.sectionId);
    const ws = getSheet(settings.sheet);
    const start = parseCellAddress(settings.startCell);
    const headerCell = ws.getCell(start.row, start.col);
    headerCell.value = `Section: ${section.sectionTitle}`;
    headerCell.font = { bold: true };
    const stride = settings.rowStride || defaults.rowStride || 18;
    const hasCustomStart = Boolean(config.sections?.[section.sectionId]?.startCell);

    section.items.forEach((item, idx) => {
      const rowOffset = stride * idx;
      const baseRow = start.row + 1 + rowOffset;
      const baseCol = start.col;
      const previewAnchor = toZeroBasedAnchor(`${numberToColumn(baseCol)}${baseRow}`);
      const size = settings.imageSize || defaults.imageSize;
      const linkOffset = Number.isFinite(settings.linkColumnOffset) ? Math.max(0, settings.linkColumnOffset) : 0;
      const linkCol = Math.max(1, baseCol + linkOffset);
      const linkCell = ws.getCell(baseRow, linkCol);
      linkCell.value = { text: item.name, hyperlink: item.relativePath };
      linkCell.font = { color: { argb: 'FF1F4E79' }, underline: true };
      linkCell.note = item.type || '';

      if (item.isImage && item.base64) {
        const imageId = workbook.addImage({ base64: item.base64, extension: item.extension || 'png' });
        ws.addImage(imageId, {
          tl: { col: previewAnchor.col, row: previewAnchor.row },
          ext: { width: size?.width || 320, height: size?.height || 180 },
        });
        const rowsCovered = Math.ceil((size?.height || 180) / 20);
        for (let r = 0; r < rowsCovered; r++) {
          const excelRow = ws.getRow(baseRow + r);
          if (!excelRow.height || excelRow.height < 60) excelRow.height = 60;
        }
      } else {
        const noteCell = ws.getCell(baseRow, baseCol);
        noteCell.value = 'Preview not available';
        noteCell.font = { italic: true, color: { argb: 'FF6E6E6E' } };
      }
    });

    if (!hasCustomStart) {
      nextDefaultRow = start.row + 1 + stride * Math.max(section.items.length, 1) + 4;
    }
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
      const evidenceSections = collectEvidenceSections();
      const preparedEvidence = await prepareEvidencePayload(evidenceSections);

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
      applyEvidenceToWorkbook(wb, preparedEvidence, EVIDENCE_EXPORT_CONFIG);

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
