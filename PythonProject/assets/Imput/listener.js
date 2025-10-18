// Legacy listener shim: integrates navbar and manual categories rendering
// Works as a non-module script and cooperates with the ES-module app.
(function(){
  // If module app is active, skip on non-manual pages to avoid duplicate handlers
  var __page = (typeof document !== 'undefined' && document.body) ? (document.body.getAttribute('data-page')||'') : '';
  if (typeof window !== 'undefined' && window.__ISA315_MODULE_ACTIVE === true && String(__page).toLowerCase() !== 'manuel') return;

  // Mark legacy active so other fallbacks can skip
  try{ window.__ISA315_LEGACY_ACTIVE = true; }catch(_){}

  // Inline fallback of FORM_SCHEMA for environments where fetch/module fails (e.g., file://)
  var __INLINE_FORM_SCHEMA = [
    {
      id: 'project_basics',
    title: 'Project basics',
    defaultOpen: true,
    questions: [
      { id:'Client Name',  label:'Client Name', type:'text' },
      { id:'Questionnaire Person', label:'Questionnaire Person', type:'text' },
      { id:'Walkthrough Date', label:'Walkthrough Date', type:'text' },
      { id:'Vendor Supported', label:'Vendor Supported', type:'yesno' },
      { id:'Job Monitoring/Scheduling Present', label:'Job Monitoring/Scheduling Present', type:'yesno' },
      { id:'Application Version', label:'Application Version', type:'text' },
      { id:'Database Management System', label:'Database Management System', type:'text' },
      { id:'Operating System', label:'Operating System', type:'text' },
      { id:'Network Software', label:'Network Software', type:'text' },
      { id:'Client POC', label:'Client POC', type:'text' },
      { id:'Application Name', label:'Application Name', type:'text' },
      { id:'IT Management Team', label:'IT Management Team', type:'text' },
      { id:'Managed Internally', label:'Managed Internally', type:'yesno' },
      { id:'Applicable Forms (x, y, z….)', label:'Applicable Forms (x, y, z….)', type:'text' }
    ]
  },

  /* ======================= IT Environment Overview ======================= */
  {
    id: 'IT Environment Overview',
    title: 'IT Environment Overview',
    questions: [
      { id: 'Identify the key individuals responsible for the entity’s IT organization. Indicate the name and role in the organization.',
        label: 'Identify the key individuals responsible for the entity’s IT organization. Indicate the name and role in the organization.',
        type: 'textarea' },

      { id: 'Describe the extent of centralization of the IT organization, including the use of common policies, personnel, technology and level of oversight. - a central IT function for the entity - distributed IT functions in major business components or locations - hybrid of both, etc.',
        label: 'Describe the extent of centralization of the IT organization, including the use of common policies, personnel, technology and level of oversight. (central IT / distributed / hybrid)',
        type: 'textarea' },

      { id: 'Describe the entity’s strategic IT plans and priorities (e.g., entity’s consideration for system upgrades and new system implementations, moving on-prem solutions to cloud-based environments, and other major IT transformation).',
        label: 'Describe the entity’s strategic IT plans and priorities (upgrades, new implementations, cloud migration, transformations).',
        type: 'textarea' },

      { id: 'Describe the extent to which the entity is planning to use new technologies or technical practices (e.g., Robotic Process Automation (RPA), blockchain, including smart contracts, artificial intelligence, establishing zero trust, DevOps).',
        label: 'Describe planned use of new technologies/practices (RPA, blockchain/smart contracts, AI, zero trust, DevOps, etc.).',
        type: 'textarea' },

      { id: 'Have there been changes to management of the IT organization in the current audit period? Are such changes planned for the near future?',
        label: 'Have there been changes to IT management this period (or planned soon)?',
        type: 'yesno' },

        { id: 'If yes, describe the changes and timeline (if available).',
        label: 'If yes, describe the changes and timeline (if available).',
        type: 'textarea' },

      { id: 'Have there been significant modifications (e.g., software, people, data or processes) to the IT applications relevant to the audit in the current audit period? Are such changes planned for the near future?',
        label: 'Significant modifications to relevant IT applications this period (or planned)?',
        type: 'yesno' },

        { id: 'If yes, describe the significant modifications made and timeline (if available)',
        label: 'If yes, describe the significant modifications made and timeline (if available)',
        type: 'textarea' },

      { id: 'Have there been significant issues with the relevant IT applications during the audit period?',
        label: 'Significant issues with relevant IT applications during the period?',
        type: 'yesno' },

        { id: 'If yes, describe the significant issues and dates the issues occurred.',
        label: 'If yes, describe the significant issues and dates the issues occurred.',
        type: 'textarea' },

      { id: 'Have there been any extended outages of the relevant IT applications for which backup systems or data were required in order to restore operations?',
        label: 'Any extended outages requiring backups to restore operations?',
        type: 'yesno' },

        { id: 'If yes, describe the cause of the outage and the period the outage occurred',
        label: 'If yes, describe the cause of the outage and the period the outage occurred',
        type: 'textarea' }
    ]
  },

  /* ======================= Application Profile ======================= */
  {
    id: 'Application Profile',
    title: 'Application Profile',
    questions: [
      { id: 'Application Type', label: 'Application Type', type: 'textarea' },
      { id: 'Application Version Release (indicate the full name of the IT application, including the version/release)',
        label: 'Application Version Release (indicate the full name of the IT application, including the version/release)',
        type: 'textarea' },
      { id: 'Database Management System2', label: 'Database Management System', type: 'textarea' },
      { id: 'Operating System2', label: 'Operating System', type: 'textarea' },
      { id: 'Network Software2', label: 'Network Software', type: 'textarea' },

      { id: 'Describe the business processes supported by this IT application (e.g., GL system, accepts orders, creates invoices, tracks inventory, etc.)',
        label: 'Describe the business processes supported by this IT application (e.g., GL, order entry, invoicing, inventory, etc.)',
        type: 'textarea' },

      { id: 'Is a service organisation (SO)/third-party/vendor involved in operating or maintaining the IT environment?',
        label: 'Is a service organisation (SO)/third-party/vendor involved in operating or maintaining the IT environment?',
        type: 'yesno' },

      { id: 'If yes, indicate the name of the SO/Third-party/Vendor and the extent of their involvement in the IT processes (e.g., application development and maintenance, implementation of changes, troubleshooting, user access maintenance (creation, modification, deletion, review of access rights), problem and incident management, etc.)',
        label: 'If yes, indicate the name of the SO/Third-party/Vendor and the extent of their involvement in the IT processes (dev, maintenance, changes, troubleshooting, access maintenance, incident mgmt., etc.)',
        type: 'textarea' },

      { id: 'If yes, how are changes to the IT application (e.g., modifications, updates, customisations, etc.) performed and implemented? Vendor initiates changes, supplies and/or maintains the IT application through issuance of update application package.',
        label: 'If yes, how are changes performed and implemented? — Vendor initiates changes via update packages.',
        type: 'yesno' },

      { id: 'The entity (local IT) performs some programming (in addition to the vendor-supplied programs), and/or to create or maintain custom reports or interfaces.',
        label: 'The entity (local IT) performs some programming and/or creates/maintains custom reports or interfaces.',
        type: 'yesno' },

      { id: 'The entity (local IT) creates and maintains the IT application (e.g., additional functionality, reports, interfaces with other IT applications).',
        label: 'The entity (local IT) creates and maintains the IT application (functionality, reports, interfaces).',
        type: 'yesno' },

      { id: 'Does the service organisation/third-party/vendor have a Service Organisation Controls (SOC) Report to support the effective operation of its IT general controls?',
        label: 'Does the service organisation/third-party/vendor have a SOC report to support effective operation of its ITGCs?',
        type: 'yesno' },

      { id: 'Is one or more of the IT environment components secured by settings (including passwords)? [Guidance: IT environment components refer to the IT application and its supporting database, operating system, network components.]',
        label: 'Is one or more of the IT environment components secured by settings (including passwords)?',
        type: 'yesno' },

      { id: 'Do users (including privileged users) from the entity have access to one or more of the IT environment components?',
        label: 'Do users (including privileged users) from the entity have access to one or more IT environment components?',
        type: 'yesno' },

      { id: 'Is the operation of programs in this IT application scheduled and monitored by the IT department? [Programs refer to tasks or jobs that are scheduled to automatically run at defined triggers or timeline].',
        label: 'Is the operation of programs in this IT application scheduled and monitored by IT?',
        type: 'yesno' }
    ]
  },

  /* =================== Managing Vendor Supplied Change =================== */
  {
    id: 'Managing Vendor Supplied Change',
    title: 'Managing Vendor Supplied Change',
    questions: [
      { id: 'Describe the process on how the entity learns about the change and how to determine if the change is necessary or appropriate for the entity’s environment',
        label: 'Describe how the entity learns about vendor changes and decides if they are necessary/appropriate.',
        type: 'textarea' },

      { id: 'Does an appropriate business or IT personnel from the entity perform testing of the functional change before it is implemented to the production environment? (Note: A functional change is a change which affects what the IT application does or how it operates or functions.  Not all changes are functional updates, for example some changes may be only security patches.)',
        label: 'Does appropriate business/IT personnel test functional changes before production?',
        type: 'yesno' },

      { id: 'If No, do users formally review each functional change within a few days of the change being implemented (i.e., report whether any issues are identified)?',
        label: 'If No, do users formally review each functional change within a few days of implementation?',
        type: 'yesno' },

      { id: 'If Yes, describe how the non-production environment is maintained and how management makes sure it represents the production environment at the time of testing.',
        label: 'If Yes, describe how non-production is maintained to mirror production at testing time.',
        type: 'textarea' },

      { id: 'Does the testing occur in a non-production environment that mirrors the production environment?',
        label: 'Does testing occur in a non-production environment that mirrors production?',
        type: 'yesno' },

      { id: 'Are changes to the IT applications approved by appropriate management (of the business or IT areas) prior to implementation?',
        label: 'Are changes approved by appropriate management prior to implementation?',
        type: 'yesno' }
    ]
  },

  /* ================= Managing Entity-Programmed Change ================== */
  {
    id: 'Managing Entity-Programmed Change',
    title: 'Managing Entity-Programmed Change',
    questions: [
      { id: 'Describe what action initiates the development process (e.g., a form completed in a help desk system, a free-form email to the IT applications director, a telephone call to the IT applications director)',
        label: 'Describe what action initiates the development process.',
        type: 'textarea' },

      { id: 'Change Ticket created in a helpdesk or ticketing system (e.g., ServiceNow, JIRA, etc.)',
        label: 'Change Ticket created in a helpdesk/ticketing system (e.g., ServiceNow, JIRA).',
        type: 'yesno' },

      { id: 'A manually-prepared Change Request Form',
        label: 'A manually-prepared Change Request Form',
        type: 'yesno' },

      { id: 'A Free Form email sent to IT Department',
        label: 'A Free Form email sent to IT Department',
        type: 'yesno' },

      { id: 'Via telephone call',
        label: 'Via telephone call',
        type: 'yesno' },

      { id: 'Others (please describe in the Notes section)',
        label: 'Others (please describe in the Notes section)',
        type: 'yesno' },

      { id: 'Are new IT applications or changes to existing IT applications tested by appropriate business or IT personnel prior to implementation?',
        label: 'Are new IT applications or changes to existing ones tested by appropriate personnel before implementation?',
        type: 'yesno' },

      { id: 'If No, do users formally review the expected effects of a change within a few days of the change being implemented (i.e., report whether any issues are identified)?',
        label: 'If No, do users formally review expected effects within a few days of implementation?',
        type: 'textarea' },

      { id: 'Does the testing occur in a non-production environment that mirrors the production environment?',
        label: 'Does the testing occur in a non-production environment that mirrors production?',
        type: 'yesno' },

      { id: 'If Yes, describe how the non-production environment is maintained and how management makes sure it represents the production environment at the time of testing.',
        label: 'If Yes, describe how non-production is maintained to mirror production at testing time.',
        type: 'textarea' },

      { id: 'Describe the types of changes that are not tested (if any)',
        label: 'Describe the types of changes that are not tested (if any).',
        type: 'textarea' },

      { id: 'Are changes to the IT applications approved by appropriate management (of the business or IT areas) prior to implementation?',
        label: 'Are changes to the IT applications approved by appropriate management prior to implementation?',
        type: 'yesno' },

      { id: 'If Yes, describe the basis for the approval decision',
        label: 'If Yes, describe the basis for the approval decision.',
        type: 'textarea' },

      { id: 'Do developers have more than read-only access to the production environment (either permanently or as needed)?',
        label: 'Do developers have more than read-only access to production (permanently or as needed)?',
        type: 'yesno' },

      { id: 'If yes, are the changes implemented to the production environment by these developers reviewed on a periodic basis by an appropriate individual (e.g., changes are logged, reviewed and traced back to approved change requests)?',
        label: 'If yes, are changes implemented by these developers periodically reviewed and traced back to approved change requests?',
        type: 'yesno' },

      { id: 'Does the entity operate multiple instances (i.e., separate copies) of this IT application? (e.g., an entity may operate multiple instances of an IT application to serve different geographical markets)',
        label: 'Does the entity operate multiple instances (separate copies) of this IT application?',
        type: 'yesno' },

      { id: 'If yes, does the entity intend that these instances operate identically?',
        label: 'If yes, does the entity intend that these instances operate identically?',
        type: 'yesno' },

      { id: 'If yes, how does the entity make sure all the instances remain identical?',
        label: 'If yes, how does the entity make sure all the instances remain identical?',
        type: 'textarea' }
    ]
  },

  /* ======================= Managing Security Settings ======================= */
  {
    id: 'Managing Security Settings',
    title: 'Managing Security Settings',
    questions: [
      { id: 'How does the entity decide on the appropriate application and network security settings (e.g., use the default settings, rely on information in installation manuals, involve individuals experienced with the software)?',
        label: 'How does the entity decide on appropriate application/network security settings?',
        type: 'textarea' },

      { id: 'What password requirements, or other authentication mechanisms (e.g., single sign-on, etc.), are implemented for this IT application and its supporting database, and for IT administrators under the entity’s control? Password complexity includes password length, composition (letters, numbers, symbols), maximum password age, account lockout process, etc.)',
        label: 'What password requirements or other authentication mechanisms are implemented (incl. admins)?',
        type: 'textarea' },

      { id: 'Has the entity disabled or changed the passwords for the default accounts that come with IT application, database, operating system and network software?',
        label: 'Has the entity disabled or changed the passwords for default accounts (application, DB, OS, network)?',
        type: 'yesno' }
    ]
  },

  /* ========================= Managing User Access ========================= */
  {
    id: 'Managing User Access',
    title: 'Managing User Access',
    questions: [
      { id: 'How is user access (including privileged access) to the IT application and its supporting IT environment components requested? A Free Form email sent to IT Department',
        label: 'How is user access requested? — A Free Form email sent to IT Department',
        type: 'yesno' },

      { id: 'How is user access (including privileged access) to the IT application and its supporting IT environment components requested? Via telephone call',
        label: 'How is user access requested? — Via telephone call',
        type: 'yesno' },

      { id: 'How is user access (including privileged access) to the IT application and its supporting IT environment components requested? Access request ticket (e.g., Service Requests)',
        label: 'How is user access requested? — Access request ticket (e.g., Service Requests)',
        type: 'yesno' },

      { id: 'How is user access (including privileged access) to the IT application and its supporting IT environment components requested? A manually-prepared access request form',
        label: 'How is user access requested? — A manually-prepared access request form',
        type: 'yesno' },

      { id: 'How is user access (including privileged access) to the IT application and its supporting IT environment components requested? Others (please describe in the Notes section)',
        label: 'How is user access requested? — Others (please describe in the Notes section)',
        type: 'yesno' },

      { id: 'Who approves requested new or additional access and where is that approval documented? [e.g., approval is provided by line manager via email, request is automatically routed to line manager via the ticketing tool, etc.]',
        label: 'Who approves requested new/additional access and where is that approval documented?',
        type: 'textarea' },

      { id: 'What factors are considered in providing that approval? [e.g., access is appropriate based on company User Access Matrix, Access request complies with segregation of duties policies, knowledge of the approver, etc]',
        label: 'What factors are considered in providing that approval (UAM matrix, SoD policies, knowledge of approver, etc.)?',
        type: 'textarea' },

      { id: 'How are access rights managed? (e.g., access rights are assigned individually to each user or access rights are setup in a role/profile that is assigned to a user, etc.)',
        label: 'How are access rights managed (individual vs. role/profile, etc.)?',
        type: 'textarea' },

      { id: 'Who is responsible to create, modify and terminate users in the IT application and its supporting IT environment components?',
        label: 'Who is responsible to create, modify and terminate users (application and supporting components)?',
        type: 'textarea' },

      { id: 'Are access rights managed (create, modify and terminate) by individuals who are independent of the business user functions? (e.g., Finance, HR, etc.)',
        label: 'Are access rights managed by individuals independent of the business user functions?',
        type: 'yesno' },

      { id: 'If no, is the access right created/modified subjected to verification or review for validity and accuracy by another individual who does not have access management responsibility?',
        label: 'If no, is the access right creation/modification verified/reviewed by someone without access management responsibility?',
        type: 'yesno' },

      { id: 'How is the request to terminate access rights (business and privileged users) in IT application and its supporting IT environment component initiated? [For example, HR notifies IT department of resigning employees via email upon receipt of resignation letter]',
        label: 'How is the request to terminate access rights initiated (business and privileged users)?',
        type: 'textarea' },

      { id: 'Does the entity perform a periodic (at least annually) review of users with privileged access?',
        label: 'Does the entity perform a periodic (at least annually) review of users with privileged access?',
        type: 'yesno' },

      { id: 'If yes, describe the scope of the review and criteria used in identifying privileged access. Describe also the basis for evaluating the appropriateness of users with privileged access.',
        label: 'If yes, describe the scope/criteria and basis for evaluating appropriateness of privileged access.',
        type: 'textarea' }
    ]
  },

  /* ==================== Job Scheduling and Monitoring ==================== */
  {
    id: 'Job Scheduling and Monitoring',
    title: 'Job Scheduling and Monitoring',
    questions: [
      { id: 'Who is responsible to configure and/or maintain scheduled programs in the IT application? [Guidance: Programs refer to tasks or jobs that are scheduled to automatically run at defined triggers or timeline].',
        label: 'Who is responsible to configure and/or maintain scheduled programs in the IT application?',
        type: 'textarea' },

      { id: 'Do individuals outside the IT department, or inappropriate individuals inside the IT department, have access to the change the scheduled programs or jobs?',
        label: 'Do individuals outside IT, or inappropriate individuals inside IT, have access to change scheduled programs/jobs?',
        type: 'yesno' },

      { id: 'How does the entity make sure scheduled jobs are configured accurately? (i.e., aligned with business requirements and/or policies)',
        label: 'How does the entity ensure scheduled jobs are configured accurately (aligned with requirements/policies)?',
        type: 'textarea' },

      { id: 'How is the successful and timely operation of the job schedule monitored?  [e.g., a software or tool that is configured to send notifications to IT operators of any failed job run and these job logs are reviewed for issues and when issues are identified they are addressed.]',
        label: 'How is successful and timely operation of the job schedule monitored?',
        type: 'textarea' },

      { id: 'What do the IT operations people personnel do when jobs stop or complete with errors? What documentation exists to assist IT operations personnel in resolving issues with jobs?',
        label: 'What do IT operations personnel do when jobs stop or error, and what documentation assists them?',
        type: 'textarea' }
    ]
  }
];

  // Ensure FORM_SCHEMA is available even when ES modules are blocked (file://)
  function ensureFormSchema(done){
    try{
      if (typeof window.FORM_SCHEMA !== 'undefined' && Array.isArray(window.FORM_SCHEMA) && window.FORM_SCHEMA.length){
        return done();
      }
      // Try to dynamically load the module file and extract the array literal
      fetch('./assets/formschema.js')
        .then(function(r){ return r.ok ? r.text() : Promise.reject(new Error('HTTP '+r.status)); })
        .then(function(txt){
          var m = txt.match(/export\s+const\s+FORM_SCHEMA\s*=\s*(\[[\s\S]*?\]);/);
          if(!m){ throw new Error('Schema array not found'); }
          var arr = (new Function('return ' + m[1]))();
          window.FORM_SCHEMA = arr;
        })
        .catch(function(err){
          console.error('Could not load FORM_SCHEMA:', err);
          // Fallback to inline schema if fetch/parse failed
          if (typeof window.FORM_SCHEMA === 'undefined' || !Array.isArray(window.FORM_SCHEMA) || !window.FORM_SCHEMA.length){
            window.FORM_SCHEMA = __INLINE_FORM_SCHEMA;
          }
        })
        .finally(function(){
          // Ensure we have at least the inline fallback
          if (typeof window.FORM_SCHEMA === 'undefined' || !Array.isArray(window.FORM_SCHEMA) || !window.FORM_SCHEMA.length){
            window.FORM_SCHEMA = __INLINE_FORM_SCHEMA;
          }
          done();
        });
    }catch(e){
      console.error('ensureFormSchema error:', e);
      if (typeof window.FORM_SCHEMA === 'undefined' || !Array.isArray(window.FORM_SCHEMA) || !window.FORM_SCHEMA.length){
        window.FORM_SCHEMA = __INLINE_FORM_SCHEMA;
      }
      done();
    }
  }

  function wireHeaderFallback(){
    if (typeof window.wireHeader === 'function') return window.wireHeader();
    // Minimal inline header wiring if helper is unavailable
    var back = document.getElementById('btnBack');
    if (back) back.addEventListener('click', function(){ history.back(); });
    var btn = document.getElementById('navToggle');
    var panel = document.getElementById('navPanel');
    if (!btn || !panel) return;
    function here(){
      var p = (location.pathname.split('/').pop() || 'index.html').toLowerCase();
      return p;
    }
    Array.prototype.forEach.call(panel.querySelectorAll('.navitem'), function(a){
      if ((a.getAttribute('href')||'').toLowerCase() === here()) a.classList.add('active');
    });
    function close(){ btn.setAttribute('aria-expanded','false'); panel.classList.remove('open'); setTimeout(function(){ panel.hidden = true; }, 150); document.removeEventListener('click', onDocClick); }
    function onDocClick(e){ if(!panel.contains(e.target) && !btn.contains(e.target)) close(); }
    function open(){ btn.setAttribute('aria-expanded','true'); panel.hidden = false; requestAnimationFrame(function(){ panel.classList.add('open'); }); document.addEventListener('click', onDocClick); }
    btn.addEventListener('click', function(){ (btn.getAttribute('aria-expanded')==='true') ? close() : open(); });
  }

  function renderManual(){
    // Prefer existing manual renderer if available
    if (typeof window.onManuel === 'function') return window.onManuel();

    // Fallback: render a very simple accordion from FORM_SCHEMA if present
    var root = document.getElementById('manualRoot');
    var SCHEMA = (typeof window.FORM_SCHEMA !== 'undefined') ? window.FORM_SCHEMA : [];
    if (!root) return;
    // Load prefill values even if modules are not active
    function __tryLoadPrefill(){
      // Prefer shared loader if available
      if (typeof window.loadPrefill === 'function') {
        try { return window.loadPrefill(); } catch(_) {}
      }
      // Inline minimal loader compatible with storage.js format
      var m = location.hash.match(/prefill=([^&]+)/);
      if (m) {
        try { return JSON.parse(decodeURIComponent(escape(atob(m[1])))); } catch(_) { /* ignore invalid hash */ }
      }
      try{
        if (typeof window.name === 'string' && window.name.indexOf('__ISA315__')===0){
          return JSON.parse(window.name.slice('__ISA315__'.length));
        }
      }catch(_){ }
      try{
        var ls = localStorage.getItem('isa315_prefill_v1');
        if (ls) return JSON.parse(ls);
      }catch(_){ }
      return {};
    }
    var PREFILL = __tryLoadPrefill();
    if (!SCHEMA.length){
      // If schema is still missing, try inline fallback installed by ensureFormSchema
      if (typeof window.FORM_SCHEMA !== 'undefined' && Array.isArray(window.FORM_SCHEMA)) SCHEMA = window.FORM_SCHEMA;
    }
    if (!SCHEMA.length) return;

    function el(tag, cls, text){ var n=document.createElement(tag); if(cls) n.className=cls; if(text) n.textContent=text; return n; }
    var wrap = el('div','accordion');
    SCHEMA.forEach(function(sec){
      var item = el('div','acc-item');
      item.setAttribute('data-sec', sec.id || '');

      var hdr = el('button','acc-header');
      hdr.type = 'button';
      var title = el('span','acc-title', sec.title || sec.id);
      var chev = el('span','acc-chev');
      hdr.appendChild(title);
      hdr.appendChild(chev);

      var body = el('div','acc-panel');

      // progress badge in header
      var badge = el('span','badge','');
      hdr.insertBefore(badge, chev);
      function updateProgress(){
        var inputs = body.querySelectorAll('input, textarea, select');
        var total = inputs.length; var filled = 0;
        Array.prototype.forEach.call(inputs, function(el){ var v=(el.value||'').toString().trim(); if(v) filled++; });
        badge.textContent = filled + '/' + total;
        var complete = (total>0 && filled===total);
        item.classList.toggle('complete', complete);
        try{ document.dispatchEvent(new CustomEvent('sec-progress', { detail:{ id: sec.id, filled:filled, total:total, complete:complete } })); }catch(_){ }
      }

      // helper to wire evidence widget
      function wireEvidence(rootEl){
        var drop = rootEl.querySelector('.q-drop');
        var input = rootEl.querySelector('input[type="file"]');
        var list  = rootEl.querySelector('.q-files');
        if(!drop || !input || !list) return;
        function renderFiles(files){
          list.innerHTML = '';
          Array.prototype.forEach.call(files, function(f){
            var row = el('div','q-file');
            var thumb = el('div','q-thumb');
            var meta = el('div','q-fmeta');
            var name = el('div','q-fname', f.name);
            var note = el('div','q-fnote', Math.round(f.size/1024)+' KB');
            meta.appendChild(name); meta.appendChild(note);
            if ((f.type||'').indexOf('image/')===0){
              var img = document.createElement('img'); img.className='q-thumb'; img.alt=''; img.src = URL.createObjectURL(f); thumb.innerHTML=''; thumb.appendChild(img);
            } else { thumb.textContent = '📎'; }
            row.appendChild(thumb); row.appendChild(meta);
            list.appendChild(row);
          });
        }
        function setFiles(files){ rootEl.__evidenceFiles = files; renderFiles(files); }
        ;['dragenter','dragover'].forEach(function(evt){ drop.addEventListener(evt, function(e){ e.preventDefault(); e.stopPropagation(); drop.classList.add('is-drag'); }); });
        ;['dragleave','drop'].forEach(function(evt){ drop.addEventListener(evt, function(e){ e.preventDefault(); e.stopPropagation(); drop.classList.remove('is-drag'); }); });
        drop.addEventListener('drop', function(e){ var f=e.dataTransfer&&e.dataTransfer.files; if(f&&f.length) setFiles(f); });
        input.addEventListener('change', function(e){ if(e.target.files && e.target.files.length) setFiles(e.target.files); });
      }

      (sec.questions||[]).forEach(function(q){
        var row = el('div','q-item' + (sec.id==='project_basics' ? ' no-attach' : ''));

        var colLabel = el('div','q-label');
        var label = document.createElement('label');
        try{
          var LMAP = (typeof window !== 'undefined' && window.FORM_LABEL_OVERRIDES) ? window.FORM_LABEL_OVERRIDES : null;
          var labTxt = (LMAP && Object.prototype.hasOwnProperty.call(LMAP, q.id)) ? LMAP[q.id] : (q.label || q.id);
          label.textContent = labTxt;
        }catch(_){
          label.textContent = q.label || q.id;
        }
        label.htmlFor = q.id; colLabel.appendChild(label);

        var colInput = el('div','q-input');
        var input;
        if(q.type==='textarea'){ input=document.createElement('textarea'); input.rows=10; input.className='input'; }
        else if(q.type==='yesno'){ input=document.createElement('select'); input.className='input'; ['', 'Yes','No'].forEach(function(v){ var o=document.createElement('option'); o.value=v; o.textContent=v||'—'; input.appendChild(o); }); }
        else { input=document.createElement('input'); input.type='text'; input.className='input'; }
        input.id = q.id;
        if (PREFILL && Object.prototype.hasOwnProperty.call(PREFILL, q.id)) input.value = PREFILL[q.id] || '';
        colInput.appendChild(input);
        ;['change','input'].forEach(function(evt){ input.addEventListener(evt, updateProgress); });

        // Place label on the left (narrow), input on the right (wide)
        row.appendChild(colLabel);
        row.appendChild(colInput);
        body.appendChild(row);
      });

      // Append Evidence row to all sections except the first (project_basics)
      if ((sec.id||'') !== 'project_basics'){
        var erow = el('div','q-item');
        var lcol = el('div','q-label'); lcol.appendChild(el('label','', 'Evidence'));
        var icol = el('div','q-input');
        var attach = el('div','q-attach');
        var drop = el('div','q-drop', 'Drop files here or browse');
        var browse = el('label','btn q-browse','Browse…');
        var inp = document.createElement('input'); inp.type='file'; inp.accept='image/*,.png,.jpg,.jpeg,.gif,.webp,.bmp,.tiff,.svg,.pdf'; inp.multiple=true; inp.hidden=true; inp.id = (sec.id||'sec')+'_evidence_input';
        browse.setAttribute('for', inp.id);
        var files = el('div','q-files');
        attach.appendChild(drop); attach.appendChild(browse); attach.appendChild(inp); attach.appendChild(files);
        icol.appendChild(attach);
        erow.appendChild(lcol); erow.appendChild(icol);
        body.appendChild(erow);
        wireEvidence(attach);
      }

      hdr.addEventListener('click', function(){
        var oneOpen = document.body && document.body.classList.contains('one-open');
        if (oneOpen) {
          Array.prototype.forEach.call(wrap.querySelectorAll('.acc-item.open'), function(it){ if(it!==item) it.classList.remove('open'); });
        }
        item.classList.toggle('open');
      });
      if (sec.defaultOpen) item.classList.add('open');
      // initial progress calc
      updateProgress();
      item.appendChild(hdr); item.appendChild(body); wrap.appendChild(item);
    });
    root.innerHTML = '';
    root.appendChild(wrap);
  }

  document.addEventListener('DOMContentLoaded', function(){
    try{
      wireHeaderFallback();
      var page = document.body && document.body.getAttribute('data-page');
      if (page === 'manuel') {
        ensureFormSchema(function(){
          renderManual();
          // build section index with progress badges
          try{
            var secIndex = document.getElementById('sectionIndex');
            if (secIndex){
              secIndex.innerHTML='';
              var map = {};
              Array.prototype.forEach.call(document.querySelectorAll('.accordion .acc-item'), function(it){
                var id = it.getAttribute('data-sec')||'';
                var title = (it.querySelector('.acc-title')||{}).textContent || id;
                var a = document.createElement('a'); a.href='#'+id; a.dataset.sec=id; a.textContent = title + ' ';
                var b = document.createElement('span'); b.className='badge'; b.textContent=''; a.appendChild(b);
                a.addEventListener('click', function(e){ e.preventDefault(); it.classList.add('open'); it.scrollIntoView({behavior:'smooth', block:'start'}); });
                secIndex.appendChild(a);
                map[id] = { link:a, badge:b };
              });
              document.addEventListener('sec-progress', function(e){
                var d = e.detail||{}; var rec = map[d.id]; if(rec){ rec.badge.textContent = (d.filled||0)+'/'+(d.total||0); rec.link.classList.toggle('complete', !!d.complete); }
              });
            }
          }catch(_){ }

          // simple expand/collapse controls
          var ex = document.getElementById('btnExpandAll');
          var cl = document.getElementById('btnCollapseAll');
          if (ex) ex.addEventListener('click', function(){ Array.prototype.forEach.call(document.querySelectorAll('.acc-item'), function(i){ i.classList.add('open'); }); });
          if (cl) cl.addEventListener('click', function(){ Array.prototype.forEach.call(document.querySelectorAll('.acc-item'), function(i){ i.classList.remove('open'); }); });

          // compact / one-open toggle
          var chk = document.getElementById('chkCompact');
          if (chk){
            chk.addEventListener('change', function(e){
              var acc = document.querySelector('.accordion');
              if(e.target.checked){ if(acc) acc.classList.add('compact'); document.body.classList.add('compact','one-open'); }
              else { if(acc) acc.classList.remove('compact'); document.body.classList.remove('compact','one-open'); }
            });
            if (chk.checked){ var acc0=document.querySelector('.accordion'); if(acc0) acc0.classList.add('compact'); document.body.classList.add('compact','one-open'); }
          }

          // Save answers (JSON)
          var btnSave = document.getElementById('btnSaveJson');
          if (btnSave){
            btnSave.addEventListener('click', function(){
              try{
                var out = {};
                Array.prototype.forEach.call(document.querySelectorAll('.acc-panel input, .acc-panel textarea, .acc-panel select'), function(el){ if(el.id) out[el.id] = el.value; });
                var blob = new Blob([JSON.stringify(out, null, 2)], { type:'application/json' });
                var url = URL.createObjectURL(blob);
                var a = document.createElement('a'); a.href=url; a.download='isa315_answers_'+(new Date().toISOString().slice(0,10))+'.json'; document.body.appendChild(a); a.click(); a.remove();
                setTimeout(function(){ URL.revokeObjectURL(url); }, 3000);
              }catch(e){ alert('Could not create JSON: '+e.message); }
            });
          }
          // CSV remains a stub; XLSX now implemented using default template in assets
          var btnCsv = document.getElementById('btnExportCsv');
          var btnXlsx = document.getElementById('btnExportXlsx');
          if (btnCsv) btnCsv.addEventListener('click', function(){ alert('CSV export coming soon. Use "Save answers (JSON)" for now.'); });
          if (btnXlsx) btnXlsx.addEventListener('click', async function(){
            try{
              if (typeof ExcelJS === 'undefined'){
                alert('ExcelJS library not loaded. Please check your internet connection.');
                return;
              }
              // Collect current answers from the manual form
              var answers = {};
              Array.prototype.forEach.call(document.querySelectorAll('.acc-panel input, .acc-panel textarea, .acc-panel select'), function(el){ if(el.id) answers[el.id] = el.value; });

              // Build sequential Question-N map (including first section)
              var idsInOrder = [];
              try{
                (window.FORM_SCHEMA||[]).forEach(function(s){ (s.questions||[]).forEach(function(q){ idsInOrder.push(q.id); }); });
              }catch(_){ }
              var numToId = new Map(); // 1-based
              idsInOrder.forEach(function(id, idx){ numToId.set(idx+1, id); });

              function normalizeYesNo(v){
                var s = String(v==null? '': v).trim().toLowerCase();
                if(['yes','y','true','ja','1'].indexOf(s)>=0) return 'Yes';
                if(['no','n','false','nein','0'].indexOf(s)>=0) return 'No';
                return String(v==null? '': v);
              }

              // Load workbook from default template in assets
              var wb = new ExcelJS.Workbook();
              async function loadTemplateBuffer(){
                function buildCandidates(){
                  var rel = ['assets/Output/Template.xlsx','./../assets/Output/Template.xlsx'];
                  var abs = ['assets/Output/Template.xlsx'];
                  try{
                    abs.push(new URL('assets/Output/Template.xlsx', location.href).href);
                    abs.push(new URL('./assets/Output/Template.xlsx', location.href).href);
                  }catch(_){ }
                  var all = rel.concat(abs);
                  // de-dup
                  return Array.from(new Set(all));
                }
                var candidates = buildCandidates();
                // Try fetch first
                for (var i=0;i<candidates.length;i++){
                  var url = candidates[i];
                  try{
                    var resp = await fetch(url, { cache: 'no-cache' });
                    if(resp && resp.ok){ return await resp.arrayBuffer(); }
                  }catch(e){ /* try next */ }
                }
                // Fallback: try XMLHttpRequest (helps in some file:// contexts)
                function tryXhr(url){
                  return new Promise(function(resolve, reject){
                    try{
                      var xhr = new XMLHttpRequest();
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
                }
                for (var j=0;j<candidates.length;j++){
                  var url2 = candidates[j];
                  try{
                    var buf = await tryXhr(url2);
                    if (buf) return buf;
                  }catch(e){ /* continue */ }
                }
                // Final fallback for file:// contexts: prompt user to pick the template file locally
                if (location.protocol === 'file:'){
                  try{
                    var picked = await new Promise(function(resolve, reject){
                      var inp = document.createElement('input');
                      inp.type = 'file'; inp.accept = '.xlsx'; inp.style.display = 'none';
                      inp.addEventListener('change', async function(){
                        try{
                          if (!inp.files || !inp.files.length) return reject(new Error('no file'));
                          var file = inp.files[0];
                          var arr = await file.arrayBuffer();
                          resolve(arr);
                        }catch(err){ reject(err); }
                        finally { try{ inp.remove(); }catch(_){ } }
                      });
                      document.body.appendChild(inp);
                      inp.click();
                      setTimeout(function(){ try{ inp.remove(); }catch(_){ } reject(new Error('picker canceled')); }, 30000);
                    });
                    if (picked) return picked;
                  }catch(_){ /* ignore and fall through */ }
                }
                console.error('[ISA315][Legacy] Default template not found via any path or method. Tried:', candidates);
                if (window.showError) { window.showError({ title:'Default template not found', message:'We could not locate assets/Template.xlsx using any known path.', hint:'Make sure the Template.xlsx file exists in the assets folder or pick it manually when prompted. If running via file://, try a local webserver.', code:'EXP-TPL-404' }); } else { alert('Default template not found: assets/Template.xlsx'); }
                throw new Error('Default template missing');
              }
              var buf = await loadTemplateBuffer();
              await wb.xlsx.load(buf);

              // Replace placeholders across all worksheets
              var reQ = /Question\s*(\d{1,3})/g;
              wb.worksheets.forEach(function(ws){
                ws.eachRow({ includeEmpty: true }, function(row){
                  row.eachCell({ includeEmpty: true }, function(cell){
                    if (cell == null) return;
                    if (typeof cell.value === 'string'){
                      var text = cell.value;
                      var replaced = text.replace(reQ, function(match, g1){
                        var n = parseInt(g1, 10);
                        var id = numToId.get(n);
                        if(!id) return match;
                        var val = normalizeYesNo(answers[id] || '');
                        return val || '';
                      });
                      if (replaced !== text) cell.value = replaced;
                    } else if (cell.value && typeof cell.value === 'object' && cell.value.richText){
                      var text2 = (cell.value.richText||[]).map(function(rt){ return rt.text || ''; }).join('');
                      var replaced2 = text2.replace(reQ, function(match, g1){
                        var n2 = parseInt(g1, 10);
                        var id2 = numToId.get(n2);
                        if(!id2) return match;
                        var val2 = normalizeYesNo(answers[id2] || '');
                        return val2 || '';
                      });
                      if (replaced2 !== text2) cell.value = replaced2;
                    }
                  });
                });
              });

              // Download filled workbook
              var outBuf = await wb.xlsx.writeBuffer();
              var blob = new Blob([outBuf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
              var url = URL.createObjectURL(blob);
              var a = document.createElement('a');
              a.href = url; a.download = 'isa315_filled_'+(new Date().toISOString().slice(0,10))+'.xlsx'; a.rel = 'noopener';
              document.body.appendChild(a); a.click(); a.remove();
              setTimeout(function(){ URL.revokeObjectURL(url); }, 3000);
            }catch(err){
              console.error('[ISA315][Legacy ExportXlsx] error:', err);
              if (window.showError) { window.showError({ title:'Excel export failed', message:'We could not generate the Excel file.', hint:'Verify Template.xlsx exists and is accessible. If running locally, try via a local webserver. Then try exporting again.', code:'EXP-WRITE-001', cause: (typeof err!=='undefined'?err:undefined) }); } else { alert('Excel export failed. Please check the template and try again.'); }
            }
          });
        });
      } else if (page === 'overview') {
        // Legacy overview: wire Excel upload and parse → savePrefill → redirect
        ensureFormSchema(function(){
          try{ wireOverview(); }catch(e){ console.error('Overview wiring failed:', e); }
        });
      }
    }catch(e){ console.error('Legacy listener boot error:', e); }
  });
})();


// ===== Legacy Overview wiring (Excel upload/parse) =====
function wireOverview(){
  try{
    var dz      = document.getElementById('dropZone');
    var input   = document.getElementById('excelInput');
    var fileLbl = document.getElementById('dzFileName');
    var useBtn  = document.getElementById('btnExcelUse');
    var hint    = document.getElementById('excelHint');
    if(!dz || !input || !useBtn) return;

    var selectedFile = null;
    function setFile(file){
      selectedFile = file || null;
      if(selectedFile){
        if(fileLbl) fileLbl.textContent = 'Selected: ' + selectedFile.name;
        useBtn.disabled = false;
        if(hint) hint.textContent = 'Ready to parse.';
      } else {
        if(fileLbl) fileLbl.textContent = '';
        useBtn.disabled = true;
        if(hint) hint.textContent = 'Accepted: .xlsx, .xlsm, .csv';
      }
    }

    input.addEventListener('change', function(e){ setFile((e.target.files && e.target.files[0]) || null); });

    ['dragenter','dragover'].forEach(function(evt){ dz.addEventListener(evt, function(e){ e.preventDefault(); e.stopPropagation(); dz.classList.add('is-drag'); }); });
    ['dragleave','drop'].forEach(function(evt){ dz.addEventListener(evt, function(e){ e.preventDefault(); e.stopPropagation(); dz.classList.remove('is-drag'); }); });
    dz.addEventListener('drop', function(e){ var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]; if(!f) return; if(!/\.(xlsx|xlsm|csv)$/i.test(f.name)){ if(hint) hint.textContent='Use .xlsx, .xlsm or .csv'; return; } setFile(f); });
    dz.addEventListener('keydown', function(e){ if(e.key==='Enter' || e.key===' '){ e.preventDefault(); input.click(); } });

    useBtn.addEventListener('click', function(){
      if(!selectedFile){ if (window.showError) { window.showError({ title:'No file selected', message:'Please select a file to continue.', hint:'Click “Browse file” and choose a .xlsx, .xlsm or .csv file.', code:'OVW-SEL-001' }); } else { alert('Please select a file.'); } return; }
      try{
        if (typeof XLSX === 'undefined') { if (window.showError) { window.showError({ title:'Excel parser not available', message:'The XLSX library could not be loaded.', hint:'Ensure you have internet connectivity or open this page via a local webserver so CDN scripts can load.', code:'OVW-XLSX-NA' }); } else { alert('Excel parser (XLSX) not available. Please ensure internet connection or open via a local webserver.'); } return; }
        var fr = new FileReader();
        fr.onload = function(){
          try{
            var data = fr.result;
            var wb = XLSX.read(data, { type:'array' });
            var sheetName = (wb.SheetNames || []).find(function(n){ return String(n||'').toLowerCase().indexOf('question')>=0; }) || wb.SheetNames[0];
            var sheet = wb.Sheets[sheetName];
            var rows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'' });
            var result = parseISA315Legacy(rows);
            // Remove project basics (section 1) so prefill starts from section 2 (from question 15)
            try{
              var schema = window.FORM_SCHEMA || [];
              var pb = (schema || []).find(function(s){ return s && s.id === 'project_basics'; });
              if (pb && Array.isArray(pb.questions)){
                var toDrop = new Set(pb.questions.map(function(q){ return q && q.id; }).filter(Boolean));
                var filtered = {};
                Object.keys(result||{}).forEach(function(k){ if(!toDrop.has(k)) filtered[k] = result[k]; });
                result = filtered;
              }
            }catch(_){ /* keep as-is if schema unavailable */ }
            showPreviewLegacy(result);
            // Inline status banner (legacy)
            try{
              var status = document.getElementById('excelStatus');
              if(!status){
                status = document.createElement('section');
                status.className = 'card';
                status.id = 'excelStatus';
                var main = document.querySelector('main.container');
                if(main){ main.insertBefore(status, document.getElementById('excelPreview') && document.getElementById('excelPreview').nextSibling || null); }
              }
              status.innerHTML = '<h3 style="margin-top:0">Import summary</h3><p class="small">Imported <strong>'+Object.keys(result).length+'</strong> items. You can proceed to review.</p>';
            }catch(_){ }
            if (typeof window.savePrefill === 'function') {
              window.savePrefill(result);
            } else {
              // minimal inline save (same as storage.js)
              try{
                var json = JSON.stringify(result);
                window.name = '__ISA315__'+json;
                localStorage.setItem('isa315_prefill_v1', json);
                // store meta for debug
                try{
                  var meta = { count:Object.keys(result||{}).length, savedAt: new Date().toISOString() };
                  localStorage.setItem('isa315_prefill_count', String(meta.count));
                  localStorage.setItem('isa315_prefill_last_saved_at', meta.savedAt);
                  localStorage.setItem('isa315_prefill_meta', JSON.stringify(meta));
                }catch(_){ }
                var b64 = btoa(unescape(encodeURIComponent(json)));
                // Redirect directly with the base64 payload to the Manual page (skip results)
                location.href = 'manuel.html#prefill='+b64;
                return;
              }catch(_){ }
            }
            try{ console.log('[ISA315] Prefill saved:', { count:Object.keys(result||{}).length }); }catch(_){ }
            alert('Imported ' + Object.keys(result).length + ' items. Redirecting to review…');
            // If savePrefill() was used, it already set location.hash to base64 on this page → carry it over to Manual
            location.href = 'manuel.html' + location.hash;
          }catch(err){ console.error(err); if (window.showError) { window.showError({ title:'Could not parse Excel', message:'The file was loaded but could not be parsed as a valid Excel workbook.', hint:'Make sure you selected a .xlsx/.xlsm file with a visible worksheet and supported content. CSV files should contain a header row.', code:'OVW-PARSE-001', cause: err }); } else { alert('Error reading Excel file.'); } }
        };
        fr.onerror = function(){ if (window.showError) { window.showError({ title:'Could not read file', message:'The selected file could not be read by the browser.', hint:'Please try selecting the file again. If the issue persists, try a different browser or save the Excel under a new name.', code:'OVW-READ-001' }); } else { alert('Could not read file.'); } };
        fr.readAsArrayBuffer(selectedFile);
      }catch(e){ console.error(e); if (window.showError) { window.showError({ title:'Unexpected error', message:(e && e.message) ? e.message : 'An unexpected error occurred while processing your file.', hint:'Try again. If it keeps happening, refresh the page and retry. If you are offline, reconnect first.', code:'OVW-UNEXP-001', cause:e }); } else { alert('Unexpected error: '+e.message); } }
    });
  }catch(e){ console.error('wireOverview error:', e); }
}

function parseISA315Legacy(rows){
  try{
    // actual layout: Questions in column B (index 1), Answers in column C (index 2), starting at Excel row 7 (index 6)
    var startRow = 6, colQ=2, colA=2;
    try{ console.log('[ISA315][Import] Using fixed start (row=6, colQ=B, colA=C)'); }catch(_){ }

    // Do NOT skip by hard-coded row numbers anymore. Instead, detect section headers by content.
    var skipPhrases = [
      'IT Environment Overview',
      'Application Profile',
      'Managing Vendor Supplied Change',
      'Managing Entity-Programmed Change',
      'Managing Security Settings',
      'Managing User Access',
      'Job Scheduling and Monitoring'
    ];
    var list = [];
    var numToId = new Map(); // 1-based numbering from the full schema (includes section 1)
    try{
      var SCHEMA = (typeof window.FORM_SCHEMA !== 'undefined') ? window.FORM_SCHEMA : [];
      // Build sequential list excluding section 1 (for fallback), but also build a number→id map for all sections
      var idsInOrderAll = [];
      (SCHEMA||[]).forEach(function(s){ (s.questions||[]).forEach(function(q){ idsInOrderAll.push(q.id); }); });
      idsInOrderAll.forEach(function(id, idx){ numToId.set(idx+1, id); });

      if (SCHEMA && SCHEMA.filter && Array.prototype.flatMap){
        list = SCHEMA.filter(function(s){ return s.id!=='project_basics'; }).flatMap(function(s){ return (s.questions||[]).map(function(q){ return q.id; }); });
      } else {
        // IE11-safe fallback
        list = []; (SCHEMA||[]).forEach(function(s){ if(s.id!=='project_basics'){ (s.questions||[]).forEach(function(q){ list.push(q.id); }); }});
      }
    }catch(_){ list=[]; numToId = new Map(); }

    var out = {}; var idx=0;
    function isLabelToken(s){
      var t = String(s||'').trim();
      if(!t) return false;
      if(/^[A-Za-z ]{1,20}:$/.test(t)) return true; // e.g., "Date:", "Answer:"
      return false;
    }
    function pickAnswer(row){
      var candidates = [colA, colA-1, colA+1, 4,5,6].filter(function(i){ return i>=0 && i<row.length; });
      for(var k=0;k<candidates.length;k++){
        var i = candidates[k];
        var vv = String((row && row[i])||'').trim();
        if(!vv) continue;
        if(isLabelToken(vv)) continue;
        return vv;
      }
      return '';
    }
    var anchorMatched = false;
    var anchorText = 'Identify the key individuals responsible for the entity’s IT organization. Indicate the name and role in the organization.'; // first question of Section 2
    for(var r=startRow; r<rows.length; r++){
      var row = rows[r] || [];
      var qRaw = String(row[colQ] || '').trim();
      var aRaw = pickAnswer(row);
      // If entire row is empty, skip
      var rowHasAny = row.some ? row.some(function(c){ return String(c||'').trim().length>0; }) : (qRaw || aRaw);
      if(!rowHasAny) continue;

      // Detect section headers anywhere in the row (merged header cells may not be in column B)
      var rowLower = (row && row.length ? row.map(function(c){ return String(c||'').toLowerCase(); }).join(' | ') : '').toLowerCase();
      var isHeader = false; for(var i2=0;i2<skipPhrases.length;i2++){ if(rowLower.indexOf(skipPhrases[i2].toLowerCase())>=0){ isHeader=true; break; } }
      if(isHeader){ try{ console.log('[ISA315][Import] Skipping header row', { r: r+1 }); }catch(_){ } continue; }

      // Require some question text in the expected column to proceed
      if(!qRaw){ continue; }
      var lower = qRaw.toLowerCase();

      // Hard anchor: If this looks like the first question of Section 2, force-map to the first id (key_individuals)
      if (!anchorMatched && idx===0 && lower.indexOf(anchorText) >= 0){
        if (list.length > 0){ out[list[0]] = normalizeYesNoLegacy(aRaw); idx = 1; anchorMatched = true; try{ console.log('[ISA315][Import] Anchor matched at row', r+1); }catch(_){ } }
        continue;
      }

      // Prefer explicit numbering like "Question 15" to anchor mapping, but only if it maps to Section 2+
      var m = qRaw.match(/question\s*(\d{1,3})/i);
      if (m){
        var n = parseInt(m[1], 10);
        var idByNum = numToId.get(n);
        // Guard: ensure we don't accidentally fill Section 1 via numbering; only accept ids present in our Section2+ list
        if (idByNum && list.indexOf(idByNum) >= 0){ out[idByNum] = normalizeYesNoLegacy(aRaw); if(!anchorMatched && idByNum===list[0]){ anchorMatched=true; idx = Math.max(idx, 1); } try{ if(idx<5) console.log('[ISA315][Import] Map by number', { n:n, id:idByNum, row:r+1 }); }catch(_){ } continue; }
      }
      // Fallback: sequential mapping within Section 2+
      if (idx < list.length){ var targetId=list[idx++]; out[targetId] = normalizeYesNoLegacy(aRaw); try{ if(idx<=5) console.log('[ISA315][Import] Map sequential', { id:targetId, row:r+1 }); }catch(_){ } }
    }
    return out;
  }catch(e){ console.error('parseISA315Legacy error:', e); return {}; }
}

function normalizeYesNoLegacy(v){
  var s = String(v||'').trim().toLowerCase();
  if(['yes','y','true','ja','1'].indexOf(s)>=0) return 'Yes';
  if(['no','n','false','nein','0'].indexOf(s)>=0) return 'No';
  return v;
}

function showPreviewLegacy(map){
  try{
    var box = document.getElementById('excelPreview');
    if(!box){
      box = document.createElement('section'); box.className='card'; box.id='excelPreview';
      var main = document.querySelector('main.container'); if(!main) return; main.appendChild(box);
    }
    var rows = Object.keys(map).map(function(k){ return [k, map[k]]; });
    var table = document.createElement('table'); table.className='table';
    var thead = document.createElement('thead'); thead.innerHTML = '<tr><th>Question ID</th><th>Answer</th></tr>';
    var tbody = document.createElement('tbody');
    rows.slice(0,50).forEach(function(kv){
      var tr = document.createElement('tr');
      var td1 = document.createElement('td'); td1.textContent = kv[0];
      var td2 = document.createElement('td'); td2.textContent = String(kv[1]);
      tr.appendChild(td1); tr.appendChild(td2); tbody.appendChild(tr);
    });
    table.appendChild(thead); table.appendChild(tbody);
    box.innerHTML = '<h3 style="margin-top:0">Preview (first 50)</h3>';
    box.appendChild(table);
  }catch(e){ console.error('showPreviewLegacy error:', e); }
}
