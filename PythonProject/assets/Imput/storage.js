export function savePrefill(obj){
  const json = JSON.stringify(obj);
  // Persist in multiple places
  window.name = '__ISA315__'+json;
  localStorage.setItem('isa315_prefill_v1', json);
  // Store meta for debugging/verification
  try{
    const meta = { count: Object.keys(obj||{}).length, savedAt: new Date().toISOString() };
    localStorage.setItem('isa315_prefill_count', String(meta.count));
    localStorage.setItem('isa315_prefill_last_saved_at', meta.savedAt);
    localStorage.setItem('isa315_prefill_meta', JSON.stringify(meta));
  }catch(_){ }
  // Reflect in URL (optional base64 in hash)
  try{
    const b64 = btoa(unescape(encodeURIComponent(json)));
    location.hash = `prefill=${b64}`;
  }catch(_){ /* ignore */ }
}

export function loadPrefill(){
  let source = 'none';
  let data = {};
  // Try hash first, but fall back gracefully if it is not a valid base64 payload
  const m = location.hash.match(/prefill=([^&]+)/);
  if (m) {
    try {
      data = JSON.parse(decodeURIComponent(escape(atob(m[1]))));
      source = 'hash';
    } catch (e) {
      // Ignore invalid hash and continue to other sources
    }
  }
  if (source==='none'){
    try{
      if (typeof window.name === 'string' && window.name.startsWith('__ISA315__')){
        data = JSON.parse(window.name.slice('__ISA315__'.length));
        source = 'window.name';
      }
    }catch(e){ /* ignore window.name parse errors */ }
  }
  if (source==='none'){
    try{
      const ls = localStorage.getItem('isa315_prefill_v1');
      if (ls) { data = JSON.parse(ls); source = 'localStorage'; }
    }catch(e){ /* ignore localStorage parse errors */ }
  }
  try{ window.__PREFILL_SOURCE = source; window.__PREFILL = data; }catch(_){ }
  return data || {};
}

// Expose for non-module fallback
if (typeof window !== 'undefined') {
  window.savePrefill = savePrefill;
  window.loadPrefill = loadPrefill;
}
