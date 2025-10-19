(function(global){
  const slotsConfig = global.ISA315_EVIDENCE_SLOTS || {};
  const sectionBlueprints = slotsConfig.sections || {};
  const maxFiles = Number.isFinite(slotsConfig.maxFiles) && slotsConfig.maxFiles > 0 ? slotsConfig.maxFiles : 10;

  const defaults = {
    sheet: 'Evidence',
    startCell: 'B2',
    rowStride: slotsConfig.defaults?.rowStride || 18,
    imageSize: slotsConfig.defaults?.imageSize || { width: 320, height: 180 },
    linkCellOffset: slotsConfig.defaults?.linkShift || { columns: 2, rows: 0 },
    header: { enabled: true, textPrefix: 'Section: ' },
  };

  const sections = {};

  const sanitize = (value='') => String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '') || 'section';

  Object.keys(sectionBlueprints).forEach((sectionId) => {
    const blueprint = sectionBlueprints[sectionId] || {};
    const entry = {
      sheet: blueprint.sheet || defaults.sheet,
      startCell: blueprint.startCell || defaults.startCell,
      rowStride: Number.isFinite(blueprint.rowStride) ? blueprint.rowStride : defaults.rowStride,
      imageSize: blueprint.imageSize || defaults.imageSize,
      linkCellOffset: blueprint.linkShift || defaults.linkCellOffset,
      header: {
        enabled: blueprint.header?.enabled !== undefined ? !!blueprint.header.enabled : true,
        textPrefix: blueprint.header?.textPrefix || `${sectionId} – `,
        cell: blueprint.header?.cell || blueprint.startCell || defaults.startCell,
        sheet: blueprint.header?.sheet || blueprint.sheet || defaults.sheet,
      },
      slots: [],
    };

    const slots = Array.isArray(blueprint.slots) ? blueprint.slots : [];
    for (let index = 0; index < maxFiles; index++) {
      const slot = slots[index] || {};
      entry.slots.push({
        id: slot.id || `${sanitize(sectionId)}_${index + 1}`,
        sheet: slot.sheet || blueprint.sheet || defaults.sheet,
        imageCell: slot.imageCell || blueprint.startCell || defaults.startCell,
        linkCell: slot.linkCell || null,
        size: slot.size || blueprint.imageSize || defaults.imageSize,
      });
    }

    sections[sectionId] = entry;
  });

  global.ISA315_EVIDENCE_LAYOUT = {
    maxFilesPerSection: maxFiles,
    defaults,
    sections,
  };
})(typeof window !== 'undefined' ? window : globalThis);
