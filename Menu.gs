function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Narrative").addItem("Build Twine", "showLinkPanel").addToUi();
}

function showLinkPanel() {
  const link = exportDialoguesToTwee();
  
  const html = HtmlService.createTemplateFromFile("LinkPanel");
  html.linkGenerated = link;

  const output = html.evaluate().setTitle("Link Generated").setWidth(300);
  SpreadsheetApp.getUi().showSidebar(output);
}