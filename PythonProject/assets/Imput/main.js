import { onOverview } from './excelImport.js';
import { onManuel } from '../Output/manualPage.js';
import { wireHeader } from '../Output/uiHelpers.js';

document.addEventListener('DOMContentLoaded', () => {
  // Mark module-based app as active so legacy listener.js can skip boot
  try{ window.__ISA315_MODULE_ACTIVE = true; }catch(_){}
  const page = document.body.dataset.page;

  // On manual page, let legacy listener render and wire header to avoid duplicates
  if (page !== 'manuel') {
    wireHeader();
  }

  switch(page){
    case 'overview': onOverview(); break;
    case 'results':  if (typeof window.onResults === 'function') window.onResults(); break;
    case 'manuel':   /* legacy listener handles rendering */ break;
    case 'index':    document.querySelector('#btnGetStarted')
                      ?.addEventListener('click', ()=>location.href='overview.html');
  }
});
