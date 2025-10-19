import { FORM_SCHEMA, FORM_QUESTION_INDEX } from '../formschema.js';
import { savePrefill } from './storage.js';
import { $, $$ } from '../Output/uiHelpers.js';

export function onOverview() {
  const dz      = $('#dropZone');
  const input   = $('#excelInput');
  const fileLbl = $('#dzFileName');
  const useBtn  = $('#btnExcelUse');
  const hint    = $('#excelHint');
  const jsonInput  = $('#manualJsonInput');
  const jsonHint   = $('#manualJsonHint');
  const jsonStatus = $('#manualJsonStatus');

  if(!dz || !input || !useBtn) return;

  let selectedFile = null;

  const setJsonStatus = (message='', type='info') => {
    if (!jsonStatus) return;
    jsonStatus.textContent = message;
    if (jsonStatus.style){
      if (type === 'error') jsonStatus.style.color = '#d9534f';
      else if (type === 'success') jsonStatus.style.color = '#3fa34d';
      else jsonStatus.style.color = '';
    }
  };

  const setFile = (file) => {
    selectedFile = file || null;
    if(selectedFile){
      fileLbl.textContent = `Selected: ${selectedFile.name}`;
      useBtn.disabled = false;
      hint.textContent = 'Ready to parse.';
    }else{
      fileLbl.textContent = '';
      useBtn.disabled = true;
      hint.textContent = 'Accepted: .xlsx, .xlsm, .csv';
    }
  };

  input.addEventListener('change', (e)=> setFile(e.target.files?.[0] || null));

  ;['dragenter','dragover'].forEach(evt=>{
    dz.addEventListener(evt, (e)=>{ e.preventDefault(); e.stopPropagation(); dz.classList.add('is-drag'); });
  });
  ;['dragleave','drop'].forEach(evt=>{
    dz.addEventListener(evt, (e)=>{ e.preventDefault(); e.stopPropagation(); dz.classList.remove('is-drag'); });
  });
  dz.addEventListener('drop', (e)=>{
    const f = e.dataTransfer?.files?.[0];
    if(!f) return;
    const ok = /\.(xlsx|xlsm|csv)$/i.test(f.name);
    if(!ok){ hint.textContent = 'Use .xlsx, .xlsm or .csv'; return; }
    setFile(f);
  });
  dz.addEventListener('keydown', (e)=>{ if(e.key==='Enter' || e.key===' '){ e.preventDefault(); input.click(); } });

  jsonInput?.addEventListener('change', async (e)=>{
    const file = e.target.files?.[0];
    if (!file) return;
    try{
      setJsonStatus('', 'info');
      const text = await file.text();
      const data = JSON.parse(text);
      if (!data || typeof data !== 'object' || Array.isArray(data)){
        throw new Error('Invalid payload');
      }
      savePrefill(data);
      const count = Object.keys(data).length;
      if (jsonHint){ jsonHint.textContent = 'JSON loaded. Opening manual entry…'; }
      setJsonStatus(`Loaded ${count} answers. Redirecting to manual entry…`, 'success');
      setTimeout(()=>{ location.href = 'manuel.html' + location.hash; }, 400);
    }catch(err){
      console.error('[ISA315][JSONImport] Failed to load JSON', err);
      setJsonStatus('Could not load the JSON file. Ensure it was exported from manual entry.', 'error');
    }finally{
      e.target.value = '';
    }
  });

  useBtn.addEventListener('click', async ()=>{
    if(!selectedFile) return alert('Please select a file.');

    try{
      if (typeof XLSX === 'undefined') {
        alert('Excel parser (XLSX) not available. Please ensure internet connection or open via a local webserver.');
        return;
      }
      const data = await selectedFile.arrayBuffer();
      const wb = XLSX.read(data, { type:'array' });
      const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes('questionare')) || wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'' });

      const result = parseISA315(rows);
      // Optional preview
      showPreview(result);
      // Inline status banner
      try{
        let status = document.getElementById('excelStatus');
        if(!status){
          status = document.createElement('section');
          status.className = 'card';
          status.id = 'excelStatus';
          const main = document.querySelector('main.container');
          if(main) main.insertBefore(status, document.getElementById('excelPreview')?.nextSibling || null);
        }
        status.innerHTML = `<h3 style="margin-top:0">Import summary</h3><p class="small">Imported <strong>${Object.keys(result).length}</strong> items. You can proceed to review.</p>`;
      }catch(_){ }
      // Save and redirect to review page
      savePrefill(result);
      try{ console.log('[ISA315] Prefill saved (module path):', { count:Object.keys(result||{}).length, meta: JSON.parse(localStorage.getItem('isa315_prefill_meta')||'{}') }); }catch(_){ }
      alert(`Imported ${Object.keys(result).length} items. Redirecting to review…`);
      // Preserve the base64 prefill hash set by savePrefill so results page can decode it reliably
      location.href = 'results.html' + location.hash;
    }catch(err){
      console.error(err);
      alert('Error reading Excel file.');
    }
  });
}

function parseISA315(rows){
  // Fixed layout per provided sheet: Questions in column B (index 1), Answers in column C (index 2), starting at row 6
  let startRow = 5, colQ = 1, colA = 2;
  // Try to detect header and column positions dynamically; only override if a clear header row is found
  outer: for(let r=0; r<Math.min(25, rows.length); r++){
    const row = rows[r]||[];
    for(let c=0; c<row.length; c++){
      const v = String(row[c]||'').toLowerCase();
      if(v.includes('question')){ colQ = c; }
      if(v.includes('answer') || v.includes('response')){ colA = c; }
    }
    if(String((rows[r]||[])[colQ]||'').toLowerCase().includes('question')){ startRow = r+1; break outer; }
  }

  const skipPhrases = [
    'IT Environment Overview',
    'Application Profile',
    'Managing Vendor Supplied Change',
    'Managing Entity-Programmed Change',
    'Managing Security Settings',
    'Managing User Access',
    'Job Scheduling and Monitoring'
  ].map(s=>s.toLowerCase());

  // Sequential mapping through schema (excluding project_basics)
  const list = FORM_SCHEMA.filter(s=>s.id!=='project_basics').flatMap(s=>s.questions.map(q=>q.id));
  const out = {}; let idx = 0;

  const isLabelToken = (s)=>{
    const t = String(s||'').trim();
    if(!t) return false;
    if(/^[A-Za-z ]{1,20}:$/.test(t)) return true; // e.g., "Date:", "Answer:"
    return false;
  };
  const pickAnswer = (row)=>{
    const candidates = [colA, colA-1, colA+1, 4,5,6].filter(i=> i>=0 && i<row.length);
    for(const c of candidates){
      const val = String(row[c]||'').trim();
      if(!val) continue;
      if(isLabelToken(val)) continue;
      return val;
    }
    return '';
  };

  for(let r=startRow; r<rows.length && idx<list.length; r++){
    const row = rows[r]||[];
    const q = String(row[colQ]||'').trim();
    const a = pickAnswer(row);
    if(!q) continue;
    if(skipPhrases.some(s=>q.toLowerCase().includes(s))) continue;
    out[list[idx++]] = normalizeYesNo(a);
  }
  return out;
}

function normalizeYesNo(v){
  const s = String(v).trim().toLowerCase();
  if(['yes','y','true','ja','1'].includes(s)) return 'Yes';
  if(['no','n','false','nein','0'].includes(s)) return 'No';
  return v;
}

function showPreview(map){
  let box = document.getElementById('excelPreview');
  if(!box){
    box = document.createElement('section');
    box.className = 'card';
    box.id = 'excelPreview';
    const main = document.querySelector('main.container');
    if(!main) return;
    main.appendChild(box);
  }
  const rows = Object.entries(map);
  const table = document.createElement('table');
  table.className = 'table';
  const thead = document.createElement('thead');
  thead.innerHTML = '<tr><th>Question ID</th><th>Answer</th></tr>';
  const tbody = document.createElement('tbody');
  rows.slice(0, 50).forEach(([k,v])=>{
    const tr = document.createElement('tr');
    const td1 = document.createElement('td'); td1.textContent = k;
    const td2 = document.createElement('td'); td2.textContent = String(v);
    tr.appendChild(td1); tr.appendChild(td2); tbody.appendChild(tr);
  });
  table.appendChild(thead); table.appendChild(tbody);
  box.innerHTML = '<h3 style="margin-top:0">Preview (first 50)</h3>';
  box.appendChild(table);
}
