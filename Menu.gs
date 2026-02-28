function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const narrativeMenu = ui.createMenu("Narrative");
  narrativeMenu.addItem("Build Generic Narrative Layer", "lounchGenericLayerExport").addToUi();
  narrativeMenu.addItem("Build Twine", "lounchTweeExport").addToUi();
}

function showLinkPanel(link) {  
  const html = HtmlService.createTemplateFromFile("LinkPanel");
  html.linkGenerated = link;

  const output = html.evaluate().setTitle("Link Generated").setWidth(300);
  SpreadsheetApp.getUi().showSidebar(output);
}

function lounchTweeExport(){
  const link = exportDialoguesToTwee();
  showLinkPanel(link);
}

function lounchGenericLayerExport(){
  const link = exportDialoguesToGenericNarrativeModel();
  showLinkPanel(link);
}