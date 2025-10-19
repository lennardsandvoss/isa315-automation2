(function(global){
  const config = {
    // Maximum number of evidence files stored per section. Update if you need more slots.
    maxFilesPerSection: 10,
    defaults: {
      // Sheet used when a section does not override it.
      sheet: 'Evidence',
      // Cell used as the anchor for the automatically generated slots.
      startCell: 'B2',
      // Distance in rows between automatically generated slots.
      rowStride: 18,
      // Default preview size for images (in pixels). Adjust to match your template layout.
      imageSize: { width: 320, height: 180 },
      // Hyperlink placement relative to the preview cell. Positive column values shift to the right.
      linkCellOffset: { columns: 2, rows: 0 },
      // Header configuration written when no custom slots are supplied.
      header: { enabled: true, textPrefix: 'Section: ' }
    },
    sections: {
      // Example override:
      // 'IT Environment Overview': {
      //   sheet: 'IT Overview',
      //   startCell: 'C5',
      //   rowStride: 16,
      //   imageSize: { width: 280, height: 160 },
      //   header: { enabled: false },
      //   slots: [
      //     { id: 'it_environment_overview_1', imageCell: 'C5', linkCell: 'F5', size: { width: 280, height: 160 } },
      //     { id: 'it_environment_overview_2', imageCell: 'C21', linkCell: 'F21' },
      //     // ...add up to 10 slots per section
      //   ]
      // }
    }
  };

  global.ISA315_EVIDENCE_LAYOUT = config;
})(typeof window !== 'undefined' ? window : globalThis);
