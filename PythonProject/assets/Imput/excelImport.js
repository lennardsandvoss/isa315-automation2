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
  const jsonHintDefault = jsonHint?.textContent || '';

  const sanitizeKey = (value, fallback = '') => {
    const base = (value == null ? '' : String(value)).trim();
    const safe = base
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '_')
      .replace(/^_+|_+$/g, '');
    return safe || (fallback ? String(fallback) : 'item');
  };

  const questionIndex = (() => {
    if (Array.isArray(FORM_QUESTION_INDEX) && FORM_QUESTION_INDEX.length) {
      return FORM_QUESTION_INDEX.map((entry, idx) => ({
        id: entry.id,
        legacyId: entry.legacyId || entry.id,
        slug: entry.slug || sanitizeKey(entry.id, `q_${idx + 1}`),
        label: entry.label || entry.id,
        index: idx,
      }));
    }
    const flat = [];
    FORM_SCHEMA.forEach((section) => {
      (section.questions || []).forEach((question) => {
        flat.push({
          id: question.id,
          legacyId: question.legacyId || question.id,
          slug: question.slug || sanitizeKey(question.id, `q_${flat.length + 1}`),
          label: question.label || question.id,
          index: flat.length,
        });
      });
    });
    return flat;
  })();

  const simplifyString = (value) => {
    if (value == null) return '';
    return String(value)
      .replace(/[“”]/g, '"')
      .replace(/[’]/g, "'")
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  };

  const questionMatchers = questionIndex.map((entry) => ({
    id: simplifyString(entry.id),
    legacyId: simplifyString(entry.legacyId),
    slug: simplifyString(entry.slug),
    label: simplifyString(entry.label),
  }));

  const matchesQuestionLabel = (value) => {
    const simplified = simplifyString(value);
    if (!simplified) return false;
    return questionMatchers.some((entry) => (
      entry.id === simplified ||
      entry.legacyId === simplified ||
      entry.slug === simplified ||
      entry.label === simplified
    ));
  };

  const unwrapValue = (value) => {
    if (value == null) return value;
    if (typeof value === 'object') {
      if (Object.prototype.hasOwnProperty.call(value, 'answer')) return value.answer;
      if (Object.prototype.hasOwnProperty.call(value, 'value')) return value.value;
    }
    return value;
  };

  const collectSequentialAnswers = (values) => {
    if (!Array.isArray(values) || !values.length) return {};
    const out = {};
    let cursor = 0;
    values.forEach((value) => {
      if (cursor >= questionIndex.length) return;
      if (matchesQuestionLabel(value)) return;
      const answer = unwrapValue(value);
      out[questionIndex[cursor].id] = answer;
      cursor += 1;
    });
    return out;
  };

  const normaliseJsonAnswers = (raw) => {
    if (!raw) return {};

    if (Array.isArray(raw)) {
      const candidates = [];
      raw.forEach((entry) => {
        if (entry == null) return;
        if (Array.isArray(entry)) {
          if (!entry.length) return;
          if (entry.length >= 2) {
            const answer = unwrapValue(entry[1]);
            if (answer !== undefined) {
              candidates.push(answer);
              return;
            }
          }
          const last = entry[entry.length - 1];
          if (!matchesQuestionLabel(last)) candidates.push(unwrapValue(last));
          return;
        }
        if (typeof entry === 'object') {
          const answerCandidate = entry.answer ?? entry.Answer ?? entry.value ?? entry.Value ?? entry.response ?? entry.Response;
          if (answerCandidate !== undefined) {
            candidates.push(unwrapValue(answerCandidate));
            return;
          }
          const keys = Object.keys(entry);
          for (let i = 0; i < keys.length; i += 1) {
            const key = keys[i];
            if (matchesQuestionLabel(key)) {
              const value = entry[key];
              if (Array.isArray(value) && value.length) {
                const lastValue = value[value.length - 1];
                candidates.push(unwrapValue(lastValue));
                return;
              }
              if (!matchesQuestionLabel(value)) {
                candidates.push(unwrapValue(value));
                return;
              }
            }
          }
          return;
        }
        if (!matchesQuestionLabel(entry)) {
          candidates.push(unwrapValue(entry));
        }
      });
      const seqOut = collectSequentialAnswers(candidates);
      if (Object.keys(seqOut).length) return seqOut;
    }

    const isPlainObject = Object.prototype.toString.call(raw) === '[object Object]';
    if (!isPlainObject) return {};

    const directMatches = {};
    let matched = 0;
    const lowerMap = new Map();
    Object.keys(raw).forEach((key) => {
      lowerMap.set(simplifyString(key), raw[key]);
    });

    questionIndex.forEach((entry) => {
      const candidates = [
        entry.id,
        entry.legacyId,
        entry.slug,
        entry.label,
      ];
      for (const key of candidates) {
        if (!key) continue;
        if (Object.prototype.hasOwnProperty.call(raw, key)) {
          directMatches[entry.id] = unwrapValue(raw[key]);
          matched += 1;
          return;
        }
        const lowered = simplifyString(key);
        if (lowerMap.has(lowered)) {
          directMatches[entry.id] = unwrapValue(lowerMap.get(lowered));
          matched += 1;
          return;
        }
      }
    });

    if (matched > 0) {
      return directMatches;
    }

    const entries = Object.entries(raw);
    const sequentialCandidates = [];
    entries.forEach(([key, value]) => {
      if (Array.isArray(value) && value.length >= 2 && matchesQuestionLabel(value[0])) {
        sequentialCandidates.push(unwrapValue(value[1]));
        return;
      }
      if (typeof value === 'object' && value !== null) {
        if (Object.prototype.hasOwnProperty.call(value, 'answer') || Object.prototype.hasOwnProperty.call(value, 'value')) {
          sequentialCandidates.push(unwrapValue(value));
          return;
        }
      }
      if (matchesQuestionLabel(key)) {
        if (!matchesQuestionLabel(value)) {
          sequentialCandidates.push(unwrapValue(value));
        }
        return;
      }
      if (!matchesQuestionLabel(value)) {
        sequentialCandidates.push(unwrapValue(value));
      }
    });
    const seqOut = collectSequentialAnswers(sequentialCandidates);
    if (Object.keys(seqOut).length) {
      return seqOut;
    }

    const answersSource = Array.isArray(raw.answers)
      ? Array.from(raw.answers)
      : (Object.prototype.toString.call(raw.answers) === '[object Object]' ? Object.values(raw.answers) : Object.values(raw));

    return collectSequentialAnswers(answersSource);
  };

  if(!dz || !input || !useBtn) return;

  let selectedFile = null;

  const isJsonFile = (file) => {
    if (!file) return false;
    const name = (file.name || '').toLowerCase();
    if (name.endsWith('.json')) return true;
    const type = (file.type || '').toLowerCase();
    return type === 'application/json' || type === 'text/json';
  };

  const isSpreadsheetFile = (file) => {
    if (!file) return false;
    const name = (file.name || '').toLowerCase();
    if (/\.(xlsx|xlsm|csv)$/i.test(name)) return true;
    const type = (file.type || '').toLowerCase();
    return type.includes('spreadsheet') || type === 'text/csv';
  };

  const setJsonStatus = (message='', type='info') => {
    if (!jsonStatus) return;
    jsonStatus.textContent = message;
    if (jsonStatus.style){
      if (type === 'error') jsonStatus.style.color = '#d9534f';
      else if (type === 'success') jsonStatus.style.color = '#3fa34d';
      else jsonStatus.style.color = '';
    }
  };

  const processJsonPrefill = async (file, origin='dropzone') => {
    try{
      if (!file) return;
      setJsonStatus('Loading JSON answers…', 'info');
      if (jsonHint){ jsonHint.textContent = 'Importing answers into manual entry…'; }
      const text = await file.text();
      const raw = JSON.parse(text);
      const answers = normaliseJsonAnswers(raw);
      if (!answers || typeof answers !== 'object' || !Object.keys(answers).length) {
        throw new Error('No answers found in JSON file');
      }
      savePrefill(answers);
      const count = Object.keys(answers).length;
      if (jsonHint){ jsonHint.textContent = 'JSON loaded. Opening manual entry…'; }
      setJsonStatus(`Loaded ${count} answers. Redirecting to manual entry…`, 'success');
      setTimeout(()=>{ location.href = 'manuel.html' + location.hash; }, origin === 'button' ? 250 : 400);
    }catch(err){
      console.error('[ISA315][JSONImport] Failed to load JSON', err);
      setJsonStatus('Could not load the JSON file. Ensure it was exported from manual entry.', 'error');
      if (jsonHint){ jsonHint.textContent = jsonHintDefault || 'Restore a session by selecting the JSON file exported from manual entry.'; }
    }finally{
      if (jsonInput) jsonInput.value = '';
      setFile(null);
    }
  };

  const setFile = (file) => {
    selectedFile = null;
    if (file && isJsonFile(file)) {
      fileLbl.textContent = `Selected: ${file.name}`;
      useBtn.disabled = true;
      hint.textContent = 'Processing JSON answers…';
      processJsonPrefill(file, 'dropzone');
      return;
    }
    if (file && !isSpreadsheetFile(file)) {
      hint.textContent = 'Use .xlsx, .xlsm, .csv, or .json';
      fileLbl.textContent = '';
      useBtn.disabled = true;
      return;
    }
    selectedFile = file || null;
    if(selectedFile){
      fileLbl.textContent = `Selected: ${selectedFile.name}`;
      useBtn.disabled = false;
      hint.textContent = 'Ready to parse.';
    }else{
      fileLbl.textContent = '';
      useBtn.disabled = true;
      hint.textContent = 'Accepted: .xlsx, .xlsm, .csv, .json';
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
    if (!isSpreadsheetFile(f) && !isJsonFile(f)) {
      hint.textContent = 'Use .xlsx, .xlsm, .csv, or .json';
      return;
    }
    setFile(f);
  });
  dz.addEventListener('keydown', (e)=>{ if(e.key==='Enter' || e.key===' '){ e.preventDefault(); input.click(); } });

  jsonInput?.addEventListener('change', async (e)=>{
    const file = e.target.files?.[0];
    if (!file) return;
    await processJsonPrefill(file, 'button');
  });

  useBtn.addEventListener('click', async ()=>{
    if(!selectedFile){
      alert('Please select a file.');
      return;
    }

    if (isJsonFile(selectedFile)) {
      await processJsonPrefill(selectedFile, 'button');
      return;
    }

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
