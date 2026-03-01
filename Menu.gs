function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Narrative')
    .addItem('Open Narrative Builder', 'openNarrativeBuilder')
    .addSeparator()
    .addItem('Quick build: Generic JSON  (all quests)', 'lounchGenericLayerExport')
    .addItem('Quick build: Twine / Twee  (all quests)', 'lounchTweeExport')
    .addItem('Quick build: Mermaid graph (all quests)', 'lounchMermaidExport')
    .addToUi();
}

function showBuilderSidebar_(defaultTarget, initialLink) {
  const html = HtmlService.createTemplateFromFile('LinkPanel');
  html.defaultTarget = defaultTarget || 'twee';
  html.initialLink = initialLink || '';
  const output = html.evaluate().setTitle('Narrative Builder').setWidth(360);
  SpreadsheetApp.getUi().showSidebar(output);
}

function openNarrativeBuilder() {
  showBuilderSidebar_('twee', '');
}


function getAvailableQuests() {
  const values = getSheetValues();
  validateData(values);

  const headers = values[0].map(h => String(h).trim());
  const questColIndex = headers.indexOf(String(COL_QUEST).trim());
  if (questColIndex < 0) {
    throw new Error('Could not find a "' + COL_QUEST + '" column in the header row.');
  }

  const set = {};
  for (let r = 1; r < values.length; r++) {
    const q = String(values[r][questColIndex] ?? '').trim();
    if (!isBlank_(q)) set[q] = true;
  }
  return Object.keys(set);
}

function buildNarrative(options) {
  options = options || {};
  const target = String(options.target || 'twee');
  const buildAll = Boolean(options.buildAll);
  const quests = Array.isArray(options.quests) ? options.quests.map(q => String(q).trim()).filter(q => q) : [];

  const ss = SpreadsheetApp.getActive();
  const originalSheet = ss.getActiveSheet();

  if (buildAll || quests.length === 0) {
    return runExportForTarget_(target);
  }

  const values = getSheetValues();
  validateData(values);
  const filtered = filterValuesByQuests_(values, quests);

  if (filtered.length < 2) {
    throw new Error('No rows found for the selected quests.');
  }

  const tmpName = 'TMP_Narrative_Builder_' + new Date().getTime();
  const tmpSheet = ss.insertSheet(tmpName);

  try {
    tmpSheet.getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);
    ss.setActiveSheet(tmpSheet);
    return runExportForTarget_(target);
  } finally {
    ss.setActiveSheet(originalSheet);
    ss.deleteSheet(tmpSheet);
  }
}

function runExportForTarget_(target) {
  switch (target) {
    case 'twee':
      return exportDialoguesToTwee();

    case 'generic':
      return exportDialoguesToGenericNarrativeModel();

    case 'mermaid':
      return exportDialoguesToMermaid();

    default:
      throw new Error('Unknown target: ' + target);
  }
}

function filterValuesByQuests_(values, quests) {
  const headers = values[0].map(h => String(h).trim());
  const questColIndex = headers.indexOf(String(COL_QUEST).trim());
  if (questColIndex < 0) {
    throw new Error('Could not find a "' + COL_QUEST + '" column in the header row.');
  }

  const questSet = {};
  quests.forEach(q => questSet[q] = true);

  const out = [values[0]];
  for (let r = 1; r < values.length; r++) {
    const q = String(values[r][questColIndex] ?? '').trim();
    if (questSet[q]) out.push(values[r]);
  }
  return out;
}

function showLinkPanel(link) {
  showBuilderSidebar_('twee', link);
}

function lounchTweeExport() {
  const link = runExportForTarget_('twee');
  showBuilderSidebar_('twee', link);
}

function lounchGenericLayerExport() {
  const link = runExportForTarget_('generic');
  showBuilderSidebar_('generic', link);
}

function lounchTweeExport() {
  const link = runExportForTarget_('mermaid');
  showBuilderSidebar_('mermaid', link);
}

