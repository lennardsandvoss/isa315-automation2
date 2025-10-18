// Fallback runtime for environments where ES modules are blocked (e.g., file://)
(function(){
  if (window.__ISA315_MODULE_ACTIVE === true) return; // Module app is running
  if (window.__ISA315_LEGACY_ACTIVE === true) return; // Legacy listener is handling boot

  document.addEventListener('DOMContentLoaded', function(){
    try{
      // Wire navbar
      if (typeof window.wireHeader === 'function') window.wireHeader();

      // Page-specific boot
      var page = document.body && document.body.getAttribute('data-page');
      if (page === 'manuel' && typeof window.onManuel === 'function') {
        window.onManuel();
      }
      // We can later add: if (page==='overview') window.onOverview();
    }catch(e){
      console.error('Fallback boot error:', e);
    }
  });
})();
