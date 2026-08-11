function onOpen(e) {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem('Open Docs Assistant', 'showSidebar')
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Docs Assistant');
  DocumentApp.getUi().showSidebar(html);
}

function showSettings() {
  const html = HtmlService.createTemplateFromFile('Settings')
    .evaluate()
    .setWidth(460)
    .setHeight(330);
  DocumentApp.getUi().showModalDialog(html, 'Docs Assistant Settings');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDocumentContext() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  const cursor = doc.getCursor();
  let selectedText = '';

  if (selection) {
    selection.getRangeElements().forEach(re => {
      const el = re.getElement();
      if (el.editAsText) {
        const t = el.asText();
        if (re.isPartial()) {
          selectedText += t.getText().substring(
            re.getStartOffset(),
            re.getEndOffsetInclusive() + 1
          ) + '\n';
        } else {
          selectedText += t.getText() + '\n';
        }
      }
    });
  }

  return {
    hasSelection: !!selection,
    hasCursor: !!cursor,
    selectedText: selectedText.trim(),
    wordCount: selectedText.trim() ? selectedText.trim().split(/\s+/).length : 0
  };
}
