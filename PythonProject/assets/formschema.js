export const FORM_SCHEMA = [
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

const slugify = (value, fallback = '') => {
  const base = (value == null ? '' : String(value)).trim();
  const safe = base
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
  return safe || (fallback ? String(fallback) : 'item');
};

FORM_SCHEMA.forEach((section, sectionIndex) => {
  const sectionSlug = slugify(section.id || section.title || `section_${sectionIndex + 1}`, `section_${sectionIndex + 1}`);
  section.slug = sectionSlug;
  if (!section.title) section.title = section.id || `Section ${sectionIndex + 1}`;
  (section.questions || []).forEach((question, questionIndex) => {
    const legacyId = question.id || `${section.id || section.title || 'question'} ${questionIndex + 1}`;
    question.legacyId = legacyId;
    question.slug = slugify(legacyId, `${sectionSlug}_${questionIndex + 1}`);
    if (!question.label) question.label = legacyId;
  });
});

export const FORM_QUESTION_INDEX = [];
FORM_SCHEMA.forEach((section) => {
  (section.questions || []).forEach((question, questionIndex) => {
    FORM_QUESTION_INDEX.push({
      sectionId: section.id,
      sectionTitle: section.title,
      sectionSlug: section.slug,
      id: question.id,
      legacyId: question.legacyId,
      slug: question.slug,
      label: question.label,
      index: questionIndex,
    });
  });
});

export const FORM_QUESTION_LOOKUP = new Map(FORM_QUESTION_INDEX.map((entry) => [entry.id, entry]));
export const FORM_QUESTION_SLUG_LOOKUP = new Map(FORM_QUESTION_INDEX.map((entry) => [entry.slug, entry.id]));
export const FORM_QUESTION_ORDER = FORM_QUESTION_INDEX.map((entry) => entry.id);
