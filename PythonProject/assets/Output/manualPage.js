import { FORM_SCHEMA } from './formschema.js';
import { loadPrefill } from '../Imput/storage.js';
import { renderAccordion, $$, download } from './uiHelpers.js';

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
      // Load default template from assets (always use repository template)
      const answers = collectAnswers();

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
          const rel = ['assets/Output/Template.xlsx','../Output/Template.xlsx'];
          const abs = ['assets/Output/Template.xlsx'];
          try{
            abs.push(new URL('assets/Output/Template.xlsx', location.href).href);
            abs.push(new URL('./assets/Template.xlsx', location.href).href);
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
        if (window.showError) { window.showError({ title:'Default template not found', message:'We could not locate assets/Template.xlsx using any known path.', hint:'Ensure the template file exists in the assets folder. If you run this via file://, try a local webserver or pick the file manually when prompted.', code:'EXP-TPL-404' }); } else { alert('Default template not found: assets/Template.xlsx'); }
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

      // 4) Download filled workbook
      const outBuf = await wb.xlsx.writeBuffer();
      const blob = new Blob([outBuf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const name = `isa315_filled_${new Date().toISOString().slice(0,10)}.xlsx`;
      download(url, name);
      setTimeout(()=> URL.revokeObjectURL(url), 3000);

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
