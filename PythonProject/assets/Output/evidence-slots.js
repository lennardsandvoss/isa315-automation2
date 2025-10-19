(function(global){
  // Update this file to control where each section's evidence previews and links land in the Excel export.
  // Adjust the sheet names, cells, or image sizes below to match your customised Template.xlsx layout.
  const MAX_FILES = 10;
  const DEFAULT_IMAGE_SIZE = { width: 320, height: 180 };
  const DEFAULT_ROW_STRIDE = 18;
  const DEFAULT_LINK_SHIFT = { columns: 4, rows: 0 };

  const SECTION_BLUEPRINTS = {
    'IT Environment Overview': {
      sheet: 'IT Environment Overview',
      rowStride: DEFAULT_ROW_STRIDE,
      header: { cell: 'B3', textPrefix: 'IT Environment Overview – ' },
      slots: [
        { id: 'it_environment_overview_1', sheet: 'IT Environment Overview', imageCell: 'B5', linkCell: 'F5', size: DEFAULT_IMAGE_SIZE },
        { id: 'it_environment_overview_2', sheet: 'IT Environment Overview', imageCell: 'B23', linkCell: 'F23', size: DEFAULT_IMAGE_SIZE },
        { id: 'it_environment_overview_3', sheet: 'IT Environment Overview', imageCell: 'B41', linkCell: 'F41', size: DEFAULT_IMAGE_SIZE },
        { id: 'it_environment_overview_4', sheet: 'IT Environment Overview', imageCell: 'B59', linkCell: 'F59', size: DEFAULT_IMAGE_SIZE },
        { id: 'it_environment_overview_5', sheet: 'IT Environment Overview', imageCell: 'B77', linkCell: 'F77', size: DEFAULT_IMAGE_SIZE },
        { id: 'it_environment_overview_6', sheet: 'IT Environment Overview', imageCell: 'B95', linkCell: 'F95', size: DEFAULT_IMAGE_SIZE },
        { id: 'it_environment_overview_7', sheet: 'IT Environment Overview', imageCell: 'B113', linkCell: 'F113', size: DEFAULT_IMAGE_SIZE },
        { id: 'it_environment_overview_8', sheet: 'IT Environment Overview', imageCell: 'B131', linkCell: 'F131', size: DEFAULT_IMAGE_SIZE },
        { id: 'it_environment_overview_9', sheet: 'IT Environment Overview', imageCell: 'B149', linkCell: 'F149', size: DEFAULT_IMAGE_SIZE },
        { id: 'it_environment_overview_10', sheet: 'IT Environment Overview', imageCell: 'B167', linkCell: 'F167', size: DEFAULT_IMAGE_SIZE },
      ],
    },
    'Application Profile': {
      sheet: 'Application Profile',
      rowStride: DEFAULT_ROW_STRIDE,
      header: { cell: 'B3', textPrefix: 'Application Profile – ' },
      slots: [
        { id: 'application_profile_1', sheet: 'Application Profile', imageCell: 'B5', linkCell: 'F5', size: DEFAULT_IMAGE_SIZE },
        { id: 'application_profile_2', sheet: 'Application Profile', imageCell: 'B23', linkCell: 'F23', size: DEFAULT_IMAGE_SIZE },
        { id: 'application_profile_3', sheet: 'Application Profile', imageCell: 'B41', linkCell: 'F41', size: DEFAULT_IMAGE_SIZE },
        { id: 'application_profile_4', sheet: 'Application Profile', imageCell: 'B59', linkCell: 'F59', size: DEFAULT_IMAGE_SIZE },
        { id: 'application_profile_5', sheet: 'Application Profile', imageCell: 'B77', linkCell: 'F77', size: DEFAULT_IMAGE_SIZE },
        { id: 'application_profile_6', sheet: 'Application Profile', imageCell: 'B95', linkCell: 'F95', size: DEFAULT_IMAGE_SIZE },
        { id: 'application_profile_7', sheet: 'Application Profile', imageCell: 'B113', linkCell: 'F113', size: DEFAULT_IMAGE_SIZE },
        { id: 'application_profile_8', sheet: 'Application Profile', imageCell: 'B131', linkCell: 'F131', size: DEFAULT_IMAGE_SIZE },
        { id: 'application_profile_9', sheet: 'Application Profile', imageCell: 'B149', linkCell: 'F149', size: DEFAULT_IMAGE_SIZE },
        { id: 'application_profile_10', sheet: 'Application Profile', imageCell: 'B167', linkCell: 'F167', size: DEFAULT_IMAGE_SIZE },
      ],
    },
    'Managing Vendor Supplied Change': {
      sheet: 'Managing Vendor Supplied Change',
      rowStride: DEFAULT_ROW_STRIDE,
      header: { cell: 'B3', textPrefix: 'Managing Vendor Supplied Change – ' },
      slots: [
        { id: 'managing_vendor_supplied_change_1', sheet: 'Managing Vendor Supplied Change', imageCell: 'B5', linkCell: 'F5', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_vendor_supplied_change_2', sheet: 'Managing Vendor Supplied Change', imageCell: 'B23', linkCell: 'F23', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_vendor_supplied_change_3', sheet: 'Managing Vendor Supplied Change', imageCell: 'B41', linkCell: 'F41', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_vendor_supplied_change_4', sheet: 'Managing Vendor Supplied Change', imageCell: 'B59', linkCell: 'F59', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_vendor_supplied_change_5', sheet: 'Managing Vendor Supplied Change', imageCell: 'B77', linkCell: 'F77', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_vendor_supplied_change_6', sheet: 'Managing Vendor Supplied Change', imageCell: 'B95', linkCell: 'F95', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_vendor_supplied_change_7', sheet: 'Managing Vendor Supplied Change', imageCell: 'B113', linkCell: 'F113', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_vendor_supplied_change_8', sheet: 'Managing Vendor Supplied Change', imageCell: 'B131', linkCell: 'F131', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_vendor_supplied_change_9', sheet: 'Managing Vendor Supplied Change', imageCell: 'B149', linkCell: 'F149', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_vendor_supplied_change_10', sheet: 'Managing Vendor Supplied Change', imageCell: 'B167', linkCell: 'F167', size: DEFAULT_IMAGE_SIZE },
      ],
    },
    'Managing Entity-Programmed Change': {
      sheet: 'Managing Entity-Programmed Change',
      rowStride: DEFAULT_ROW_STRIDE,
      header: { cell: 'B3', textPrefix: 'Managing Entity-Programmed Change – ' },
      slots: [
        { id: 'managing_entity_programmed_change_1', sheet: 'Managing Entity-Programmed Change', imageCell: 'B5', linkCell: 'F5', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_entity_programmed_change_2', sheet: 'Managing Entity-Programmed Change', imageCell: 'B23', linkCell: 'F23', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_entity_programmed_change_3', sheet: 'Managing Entity-Programmed Change', imageCell: 'B41', linkCell: 'F41', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_entity_programmed_change_4', sheet: 'Managing Entity-Programmed Change', imageCell: 'B59', linkCell: 'F59', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_entity_programmed_change_5', sheet: 'Managing Entity-Programmed Change', imageCell: 'B77', linkCell: 'F77', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_entity_programmed_change_6', sheet: 'Managing Entity-Programmed Change', imageCell: 'B95', linkCell: 'F95', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_entity_programmed_change_7', sheet: 'Managing Entity-Programmed Change', imageCell: 'B113', linkCell: 'F113', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_entity_programmed_change_8', sheet: 'Managing Entity-Programmed Change', imageCell: 'B131', linkCell: 'F131', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_entity_programmed_change_9', sheet: 'Managing Entity-Programmed Change', imageCell: 'B149', linkCell: 'F149', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_entity_programmed_change_10', sheet: 'Managing Entity-Programmed Change', imageCell: 'B167', linkCell: 'F167', size: DEFAULT_IMAGE_SIZE },
      ],
    },
    'Managing Security Settings': {
      sheet: 'Managing Security Settings',
      rowStride: DEFAULT_ROW_STRIDE,
      header: { cell: 'B3', textPrefix: 'Managing Security Settings – ' },
      slots: [
        { id: 'managing_security_settings_1', sheet: 'Managing Security Settings', imageCell: 'B5', linkCell: 'F5', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_security_settings_2', sheet: 'Managing Security Settings', imageCell: 'B23', linkCell: 'F23', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_security_settings_3', sheet: 'Managing Security Settings', imageCell: 'B41', linkCell: 'F41', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_security_settings_4', sheet: 'Managing Security Settings', imageCell: 'B59', linkCell: 'F59', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_security_settings_5', sheet: 'Managing Security Settings', imageCell: 'B77', linkCell: 'F77', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_security_settings_6', sheet: 'Managing Security Settings', imageCell: 'B95', linkCell: 'F95', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_security_settings_7', sheet: 'Managing Security Settings', imageCell: 'B113', linkCell: 'F113', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_security_settings_8', sheet: 'Managing Security Settings', imageCell: 'B131', linkCell: 'F131', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_security_settings_9', sheet: 'Managing Security Settings', imageCell: 'B149', linkCell: 'F149', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_security_settings_10', sheet: 'Managing Security Settings', imageCell: 'B167', linkCell: 'F167', size: DEFAULT_IMAGE_SIZE },
      ],
    },
    'Managing User Access': {
      sheet: 'Managing User Access',
      rowStride: DEFAULT_ROW_STRIDE,
      header: { cell: 'B3', textPrefix: 'Managing User Access – ' },
      slots: [
        { id: 'managing_user_access_1', sheet: 'Managing User Access', imageCell: 'B5', linkCell: 'F5', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_user_access_2', sheet: 'Managing User Access', imageCell: 'B23', linkCell: 'F23', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_user_access_3', sheet: 'Managing User Access', imageCell: 'B41', linkCell: 'F41', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_user_access_4', sheet: 'Managing User Access', imageCell: 'B59', linkCell: 'F59', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_user_access_5', sheet: 'Managing User Access', imageCell: 'B77', linkCell: 'F77', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_user_access_6', sheet: 'Managing User Access', imageCell: 'B95', linkCell: 'F95', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_user_access_7', sheet: 'Managing User Access', imageCell: 'B113', linkCell: 'F113', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_user_access_8', sheet: 'Managing User Access', imageCell: 'B131', linkCell: 'F131', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_user_access_9', sheet: 'Managing User Access', imageCell: 'B149', linkCell: 'F149', size: DEFAULT_IMAGE_SIZE },
        { id: 'managing_user_access_10', sheet: 'Managing User Access', imageCell: 'B167', linkCell: 'F167', size: DEFAULT_IMAGE_SIZE },
      ],
    },
    'Job Scheduling and Monitoring': {
      sheet: 'Job Scheduling and Monitoring',
      rowStride: DEFAULT_ROW_STRIDE,
      header: { cell: 'B3', textPrefix: 'Job Scheduling and Monitoring – ' },
      slots: [
        { id: 'job_scheduling_and_monitoring_1', sheet: 'Job Scheduling and Monitoring', imageCell: 'B5', linkCell: 'F5', size: DEFAULT_IMAGE_SIZE },
        { id: 'job_scheduling_and_monitoring_2', sheet: 'Job Scheduling and Monitoring', imageCell: 'B23', linkCell: 'F23', size: DEFAULT_IMAGE_SIZE },
        { id: 'job_scheduling_and_monitoring_3', sheet: 'Job Scheduling and Monitoring', imageCell: 'B41', linkCell: 'F41', size: DEFAULT_IMAGE_SIZE },
        { id: 'job_scheduling_and_monitoring_4', sheet: 'Job Scheduling and Monitoring', imageCell: 'B59', linkCell: 'F59', size: DEFAULT_IMAGE_SIZE },
        { id: 'job_scheduling_and_monitoring_5', sheet: 'Job Scheduling and Monitoring', imageCell: 'B77', linkCell: 'F77', size: DEFAULT_IMAGE_SIZE },
        { id: 'job_scheduling_and_monitoring_6', sheet: 'Job Scheduling and Monitoring', imageCell: 'B95', linkCell: 'F95', size: DEFAULT_IMAGE_SIZE },
        { id: 'job_scheduling_and_monitoring_7', sheet: 'Job Scheduling and Monitoring', imageCell: 'B113', linkCell: 'F113', size: DEFAULT_IMAGE_SIZE },
        { id: 'job_scheduling_and_monitoring_8', sheet: 'Job Scheduling and Monitoring', imageCell: 'B131', linkCell: 'F131', size: DEFAULT_IMAGE_SIZE },
        { id: 'job_scheduling_and_monitoring_9', sheet: 'Job Scheduling and Monitoring', imageCell: 'B149', linkCell: 'F149', size: DEFAULT_IMAGE_SIZE },
        { id: 'job_scheduling_and_monitoring_10', sheet: 'Job Scheduling and Monitoring', imageCell: 'B167', linkCell: 'F167', size: DEFAULT_IMAGE_SIZE },
      ],
    },
  };

  global.ISA315_EVIDENCE_SLOTS = {
    maxFiles: MAX_FILES,
    defaults: {
      imageSize: DEFAULT_IMAGE_SIZE,
      linkShift: DEFAULT_LINK_SHIFT,
      rowStride: DEFAULT_ROW_STRIDE,
    },
    sections: SECTION_BLUEPRINTS,
  };
})(typeof window !== 'undefined' ? window : globalThis);
