// Simple DOM helpers
export const $  = (s, root=document) => root.querySelector(s);
export const $$ = (s, root=document) => Array.from(root.querySelectorAll(s));

// Header: back button + dropdown nav
export function wireHeader(){
  const back = $('#btnBack');
  if(back) back.addEventListener('click', ()=> history.back());

  const btn   = $('#navToggle');
  const panel = $('#navPanel');
  if(!btn || !panel) return;

  const here = (location.pathname.split('/').pop() || 'index.html').toLowerCase();
  $$('#navPanel .navitem').forEach(a => { if((a.getAttribute('href')||'').toLowerCase()===here) a.classList.add('active'); });

  const onDocClick = (e) => { if(!panel.contains(e.target) && !btn.contains(e.target)) close(); };
  const onKey = (e) => {
    const items = $$('#navPanel .navitem');
    const idx = items.indexOf(document.activeElement);
    switch(e.key){
      case 'Escape': e.preventDefault(); close(); break;
      case 'ArrowDown': e.preventDefault(); (items[idx+1] || items[0])?.focus(); break;
      case 'ArrowUp': e.preventDefault(); (items[idx-1] || items[items.length-1])?.focus(); break;
      case 'Home': e.preventDefault(); items[0]?.focus(); break;
      case 'End': e.preventDefault(); items[items.length-1]?.focus(); break;
    }
  };
  const open = () => {
    btn.setAttribute('aria-expanded','true');
    panel.hidden = false;
    requestAnimationFrame(()=> panel.classList.add('open'));
    (panel.querySelector('.navitem'))?.focus({ preventScroll:true });
    document.addEventListener('click', onDocClick);
    document.addEventListener('keydown', onKey);
  };
  const close = () => {
    btn.setAttribute('aria-expanded','false');
    panel.classList.remove('open');
    setTimeout(()=> panel.hidden = true, 150);
    document.removeEventListener('click', onDocClick);
    document.removeEventListener('keydown', onKey);
  };
  const toggle = () => (btn.getAttribute('aria-expanded')==='true' ? close() : open());
  btn.addEventListener('click', toggle);
  btn.addEventListener('keydown', (e)=>{ if(e.key==='ArrowDown'){ e.preventDefault(); if(btn.getAttribute('aria-expanded')!=='true') open(); }});
}

// Minimal accordion renderer for Manual page
export function renderAccordion(schema, prefill={}){
  const wrap = document.createElement('div');
  wrap.className = 'accordion';

  // dispatch progress event helper
  const dispatchProgress = (detail) => {
    try{ document.dispatchEvent(new CustomEvent('sec-progress', { detail })); }catch(_){ /* ignore */ }
  };

  // helper to wire a lightweight evidence uploader (drag&drop + preview list)
  const wireEvidence = (rootEl) => {
    const drop = rootEl.querySelector('.q-drop');
    const input = rootEl.querySelector('input[type="file"]');
    const list  = rootEl.querySelector('.q-files');
    if(!drop || !input || !list) return;

    const maxFiles = (()=>{
      try{
        const cfg = window.ISA315_EVIDENCE_LAYOUT;
        if (cfg && typeof cfg.maxFilesPerSection === 'number' && cfg.maxFilesPerSection > 0){
          return Math.min(50, Math.max(1, cfg.maxFilesPerSection));
        }
      }catch(_){ }
      return 10;
    })();

    let currentFiles = [];

    const renderFiles = (files) => {
      list.innerHTML = '';
      Array.from(files).forEach(f => {
        const row = document.createElement('div'); row.className = 'q-file';
        const thumb = document.createElement('div'); thumb.className = 'q-thumb';
        const meta = document.createElement('div'); meta.className = 'q-fmeta';
        const name = document.createElement('div'); name.className = 'q-fname'; name.textContent = f.name;
        const note = document.createElement('div'); note.className = 'q-fnote'; note.textContent = `${Math.round(f.size/1024)} KB`;
        meta.appendChild(name); meta.appendChild(note);
        // Preview for images
        if (f.type.startsWith('image/')){
          const img = document.createElement('img');
          img.className = 'q-thumb';
          img.alt = '';
          img.src = URL.createObjectURL(f);
          thumb.innerHTML = '';
          thumb.appendChild(img);
        } else {
          thumb.textContent = '📎';
        }
        row.appendChild(thumb); row.appendChild(meta);
        list.appendChild(row);
      });
      if (files.length >= maxFiles){
        const info = document.createElement('div');
        info.className = 'q-fnote limit';
        info.textContent = `Showing first ${maxFiles} files. Remove one to add more.`;
        list.appendChild(info);
      }
    };

    const setFiles = (files) => {
      currentFiles = Array.from(files).slice(0, maxFiles);
      rootEl.__evidenceFiles = currentFiles;
      renderFiles(currentFiles);
    };

    const appendFiles = (files) => {
      if(!files || !files.length) return;
      const merged = currentFiles.slice();
      Array.from(files).forEach(file => {
        const duplicate = merged.some(existing => (
          existing.name === file.name &&
          existing.size === file.size &&
          existing.lastModified === file.lastModified
        ));
        if (!duplicate) merged.push(file);
      });
      setFiles(merged);
    };

    ;['dragenter','dragover'].forEach(evt=>{
      drop.addEventListener(evt, e=>{ e.preventDefault(); e.stopPropagation(); drop.classList.add('is-drag'); });
    });
    ;['dragleave','drop'].forEach(evt=>{
      drop.addEventListener(evt, e=>{ e.preventDefault(); e.stopPropagation(); drop.classList.remove('is-drag'); });
    });
    drop.addEventListener('drop', e=>{
      const f = e.dataTransfer?.files; if(!f || !f.length) return;
      appendFiles(f);
    });
    input.addEventListener('change', e=>{
      if(e.target.files && e.target.files.length){ appendFiles(e.target.files); }
      e.target.value = '';
    });
  };

  schema.forEach(section => {
    const item = document.createElement('div');
    item.className = 'acc-item';
    item.setAttribute('data-sec', section.id || '');

    const hdr = document.createElement('button');
    hdr.className = 'acc-header';
    hdr.type = 'button';
    const title = document.createElement('span');
    title.className = 'acc-title';
    title.textContent = section.title || section.id;
    const chev = document.createElement('span');
    chev.className = 'acc-chev';
    hdr.appendChild(title);
    hdr.appendChild(chev);

    const body = document.createElement('div');
    body.className = 'acc-panel';

    // progress badge in header
    const badge = document.createElement('span');
    badge.className = 'badge';
    badge.textContent = '';
    hdr.insertBefore(badge, chev);

    const updateProgress = () => {
      const inputs = body.querySelectorAll('input, textarea, select');
      const total = inputs.length;
      let filled = 0;
      inputs.forEach(el=>{ const v = (el.value||'').toString().trim(); if(v) filled++; });
      badge.textContent = `${filled}/${total}`;
      const complete = total>0 && filled===total;
      item.classList.toggle('complete', complete);
      dispatchProgress({ id: section.id, filled, total, complete });
    };

    (section.questions||[]).forEach(q => {
      const row = document.createElement('div');
      row.className = 'q-item' + (section.id==='project_basics' ? ' no-attach' : '');

      const colLabel = document.createElement('div');
      colLabel.className = 'q-label';
      const label = document.createElement('label');
      const LMAP = (typeof window !== 'undefined' && window.FORM_LABEL_OVERRIDES) ? window.FORM_LABEL_OVERRIDES : null;
      const labTxt = (LMAP && Object.prototype.hasOwnProperty.call(LMAP, q.id)) ? LMAP[q.id] : (q.label || q.id);
      label.textContent = labTxt;
      label.htmlFor = q.id;
      colLabel.appendChild(label);

      const colInput = document.createElement('div');
      colInput.className = 'q-input';
      let input;
      if(q.type==='textarea'){
        input = document.createElement('textarea');
        input.rows = 10; // larger default typing area
        input.className = 'input';
      }else if(q.type==='yesno'){
        // Minimal yes/no as select; styled via .input
        input = document.createElement('select');
        input.className = 'input';
        ['','Yes','No'].forEach(v=>{ const o=document.createElement('option'); o.value=v; o.textContent=v||'—'; input.appendChild(o); });
      }else{
        input = document.createElement('input');
        input.type = 'text';
        input.className = 'input';
      }
      input.id = q.id;
      input.value = prefill[q.id] ?? '';
      colInput.appendChild(input);
      // update on change
      ['change','input'].forEach(evt=> input.addEventListener(evt, updateProgress));

      // Place label on the left (narrow), input on the right (wide)
      row.appendChild(colLabel);
      row.appendChild(colInput);
      body.appendChild(row);
    });

    // Add Evidence row at end of each section except the first (project_basics)
    if ((section.id||'') !== 'project_basics'){
      const erow = document.createElement('div');
      erow.className = 'q-item';
      const lcol = document.createElement('div'); lcol.className = 'q-label';
      const l = document.createElement('label'); l.textContent = 'Evidence'; lcol.appendChild(l);
      const icol = document.createElement('div'); icol.className = 'q-input';
      const attach = document.createElement('div'); attach.className = 'q-attach';
      const drop = document.createElement('div'); drop.className = 'q-drop'; drop.textContent = 'Drop files here or browse';
      const browse = document.createElement('label'); browse.className = 'btn q-browse'; browse.textContent = 'Browse…';
      const inp = document.createElement('input'); inp.type = 'file'; inp.accept = 'image/*,.png,.jpg,.jpeg,.gif,.webp,.bmp,.tiff,.svg,.pdf'; inp.multiple = true; inp.hidden = true; inp.id = `${section.id}_evidence_input`;
      browse.setAttribute('for', inp.id);
      const files = document.createElement('div'); files.className = 'q-files';
      attach.appendChild(drop); attach.appendChild(browse); attach.appendChild(inp); attach.appendChild(files);
      icol.appendChild(attach);
      erow.appendChild(lcol); erow.appendChild(icol);
      body.appendChild(erow);
      // wire evidence behaviors
      wireEvidence(attach);
    }

    hdr.addEventListener('click', ()=>{
      const oneOpen = document.body.classList.contains('one-open');
      if (oneOpen) {
        Array.from(wrap.querySelectorAll('.acc-item.open')).forEach(it=>{ if(it!==item) it.classList.remove('open'); });
      }
      item.classList.toggle('open');
    });
    if(section.defaultOpen) item.classList.add('open');

    // initial progress calc
    updateProgress();

    item.appendChild(hdr);
    item.appendChild(body);
    wrap.appendChild(item);
  });

  return wrap;
}

// Small utility to download a blob/url
export function download(url, filename){
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.rel='noopener';
  document.body.appendChild(a); a.click(); a.remove();
}

// Lightweight inline banner area
function ensureBannerRoot(){
  try{
    let root = document.getElementById('appMessages');
    if (root) return root;
    const main = document.querySelector('main.container') || document.body;
    root = document.createElement('section');
    root.id = 'appMessages';
    root.setAttribute('aria-live','polite');
    root.style.marginBottom = '12px';
    root.className = 'card';
    // Insert at top of main
    if (main.firstChild) main.insertBefore(root, main.firstChild); else main.appendChild(root);
    return root;
  }catch(_){ return null; }
}

function renderMessage({ type='info', title='', message='', hint='', code='', details='' }){
  const root = ensureBannerRoot();
  if (!root){
    // Fallback to alert if no DOM available
    const txt = [title, message, hint, code && `(code: ${code})`].filter(Boolean).join('\n');
    try{ alert(txt); }catch(_){}
    return;
  }
  const box = document.createElement('div');
  box.className = 'msg ' + type;
  const h = document.createElement('h3'); h.style.marginTop='0'; h.textContent = title || (type==='error' ? 'Something went wrong' : 'Notice');
  const p = document.createElement('p'); p.className = 'small'; p.textContent = message || '';
  box.appendChild(h); box.appendChild(p);
  if (hint){ const ph = document.createElement('p'); ph.className = 'small'; ph.textContent = `Hint: ${hint}`; box.appendChild(ph); }
  if (code){ const pc = document.createElement('p'); pc.className = 'small'; pc.textContent = `Error code: ${code}`; box.appendChild(pc); }
  if (details){ const pre = document.createElement('pre'); pre.className = 'small'; pre.textContent = details; pre.style.whiteSpace='pre-wrap'; pre.style.background='#f7f7f7'; pre.style.padding='6px'; pre.style.borderRadius='6px'; box.appendChild(pre); }
  // Close button
  const close = document.createElement('button'); close.className = 'btn'; close.textContent = 'Dismiss'; close.style.float='right'; close.addEventListener('click', ()=> box.remove());
  box.appendChild(close);
  root.appendChild(box);
}

export function showError(opts={}){
  const { title, message, hint, code, cause } = opts || {};
  try{
    console.error('[ISA315][UIError]', { title, message, hint, code, cause });
  }catch(_){ }
  const details = (function(){
    try{
      if (!cause) return '';
      if (typeof cause === 'string') return cause;
      const parts = [];
      if (cause.message) parts.push(cause.message);
      if (cause.name) parts.push(`(${cause.name})`);
      if (cause.stack) parts.push('\n'+cause.stack);
      return parts.join(' ');
    }catch(_){ return ''; }
  })();
  renderMessage({ type:'error', title: title||'Error', message: message||'', hint: hint||'', code: code||'', details });
}

export function showInfo(opts={}){
  const { title, message, hint } = opts || {};
  try{ console.log('[ISA315][Info]', { title, message, hint }); }catch(_){ }
  renderMessage({ type:'info', title: title||'Info', message: message||'', hint: hint||'' });
}

// Expose for non-module fallback
if (typeof window !== 'undefined') {
  window.$ = $;
  window.$$ = $$;
  window.wireHeader = wireHeader;
  window.renderAccordion = renderAccordion;
  window.download = download;
  window.showError = showError;
  window.showInfo = showInfo;
}