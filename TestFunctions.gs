/**
 * TEST LAB
 *
 * TEST-021
 * Crea un inciso nativo nuevo a) en la pestaña activa.
 *
 * Seleccionar únicamente la palabra "Octubre".
 */

function runCurrentTest() {
  return testCreateIndependentNativeIncisoWithApi_();
}


function testCreateIndependentNativeIncisoWithApi_() {
  const started = Date.now();
  const document = DocumentApp.getActiveDocument();
  const documentId = document.getId();

  const activeTab = document.getActiveTab();
  const activeTabId = activeTab.getId();
  const documentTab = activeTab.asDocumentTab();

  let selection = document.getSelection();

  if (!selection) {
    throw new Error(
      'Select only the word Octubre.'
    );
  }

  const rangeElements = selection.getRangeElements();

  if (rangeElements.length === 0) {
    throw new Error('The selection is empty.');
  }

  const targetParagraph =
    testFindParagraphAncestor_(
      rangeElements[0].getElement()
    );

  if (!targetParagraph) {
    throw new Error(
      'No paragraph was found in the selection.'
    );
  }

  if (
    targetParagraph.getType() ===
    DocumentApp.ElementType.LIST_ITEM
  ) {
    throw new Error(
      'Octubre must be a normal paragraph for this test.'
    );
  }

  /*
   * Limpia residuos de una ejecución anterior fallida.
   */
  const targetText = targetParagraph.editAsText();

  if (
    targetText.getText().charAt(0) === '\t'
  ) {
    targetText.deleteText(0, 0);
  }

  const staleRanges =
    documentTab.getNamedRanges();

  for (let i = 0; i < staleRanges.length; i++) {
    if (
      staleRanges[i]
        .getName()
        .indexOf('DOCSASSISTANT_TEST_') === 0
    ) {
      staleRanges[i].remove();
    }
  }

  /*
   * Inserta un párrafo invisible antes de Octubre.
   * Así la API no conecta la lista nueva con la anterior.
   */
  const parent = targetParagraph.getParent();
  const targetIndex =
    parent.getChildIndex(targetParagraph);

  parent.insertParagraph(
    targetIndex,
    '\u200B'
  );

  /*
   * Recupera la selección después de modificar la estructura.
   */
  selection =
    document.getSelection() || selection;

  const markerName =
    'DOCSASSISTANT_TEST_' +
    Utilities.getUuid();

  const namedRange =
    documentTab.addNamedRange(
      markerName,
      selection
    );

  const namedRangeId =
    namedRange.getId();

  /*
   * Un tabulador inicial genera el nivel 1:
   * a), b), c)...
   */
  targetParagraph
    .editAsText()
    .insertText(0, '\t');

  document.saveAndClose();

  const readStarted = Date.now();

  /*
   * Recupera únicamente:
   * - ID de las pestañas.
   * - Rangos nombrados.
   *
   * No descarga las 295 páginas.
   */
  const apiDocument =
    Docs.Documents.get(
      documentId,
      {
        includeTabsContent: true,
        fields:
          'tabs(' +
            'tabProperties(tabId),' +
            'documentTab(namedRanges),' +
            'childTabs(' +
              'tabProperties(tabId),' +
              'documentTab(namedRanges)' +
            ')' +
          ')'
      }
    );

  const apiReadMs =
    Date.now() - readStarted;

  const apiNamedRange =
    testFindNamedRangeInTabs_(
      apiDocument.tabs || [],
      activeTabId,
      markerName,
      namedRangeId
    );

  if (
    !apiNamedRange ||
    !apiNamedRange.ranges ||
    apiNamedRange.ranges.length === 0
  ) {
    throw new Error(
      'The named range was not found in the active Docs tab.'
    );
  }

  const sourceRange =
    apiNamedRange.ranges[0];

  const requestRange = {
    startIndex: sourceRange.startIndex,
    endIndex: sourceRange.endIndex,
    tabId: activeTabId
  };

  const writeStarted = Date.now();

  Docs.Documents.batchUpdate(
    {
      requests: [
        {
          createParagraphBullets: {
            range: requestRange,
            bulletPreset:
              'NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS'
          }
        }
      ]
    },
    documentId
  );

  const apiWriteMs =
    Date.now() - writeStarted;

  /*
   * Reabre exactamente la misma pestaña.
   */
  const reopenedDocument =
    DocumentApp.openById(documentId);

  const reopenedTab =
    reopenedDocument
      .getTab(activeTabId)
      .asDocumentTab();

  const savedNamedRange =
    reopenedTab.getNamedRangeById(
      namedRangeId
    );

  if (!savedNamedRange) {
    throw new Error(
      'The named range could not be reopened.'
    );
  }

  const updatedElements =
    savedNamedRange
      .getRange()
      .getRangeElements();

  let listItem = null;

  for (
    let i = 0;
    i < updatedElements.length;
    i++
  ) {
    listItem =
      testFindListItemAncestor_(
        updatedElements[i].getElement()
      );

    if (listItem) {
      break;
    }
  }

  if (!listItem) {
    throw new Error(
      'The API did not create a native list item.'
    );
  }

  listItem
    .setIndentFirstLine(18)
    .setIndentStart(36)
    .setIndentEnd(0);

  /*
   * Retira el separador invisible.
   * Los List ID ya fueron creados y permanecen separados.
   */
  const previousSibling =
    listItem.getPreviousSibling();

  let separatorRemoved = false;

  if (
    previousSibling &&
    previousSibling.getType() ===
      DocumentApp.ElementType.PARAGRAPH &&
    previousSibling.getText() === '\u200B'
  ) {
    previousSibling.removeFromParent();
    separatorRemoved = true;
  }

  const result = {
    testId:
      'TEST-021-INDEPENDENT-NATIVE-INCISO',

    ok: true,
    text: listItem.getText(),
    listId: listItem.getListId(),
    glyphType:
      String(listItem.getGlyphType()),
    nestingLevel:
      listItem.getNestingLevel(),
    isNative:
      listItem.getType() ===
      DocumentApp.ElementType.LIST_ITEM,
    separatorRemoved: separatorRemoved,
    indentFirstLine:
      listItem.getIndentFirstLine(),
    indentStart:
      listItem.getIndentStart(),
    indentEnd:
      listItem.getIndentEnd(),
    apiReadMs: apiReadMs,
    apiWriteMs: apiWriteMs,
    elapsedMs: Date.now() - started
  };

  savedNamedRange.remove();
  reopenedDocument.saveAndClose();

  return result;
}


function testFindParagraphAncestor_(element) {
  let current = element;

  while (current) {
    const type = current.getType();

    if (
      type ===
        DocumentApp.ElementType.PARAGRAPH ||
      type ===
        DocumentApp.ElementType.LIST_ITEM
    ) {
      return current;
    }

    current = current.getParent();
  }

  return null;
}


function testFindListItemAncestor_(element) {
  let current = element;

  while (current) {
    if (
      current.getType() ===
      DocumentApp.ElementType.LIST_ITEM
    ) {
      return current.asListItem();
    }

    current = current.getParent();
  }

  return null;
}


function testFindNamedRangeInTabs_(
  tabs,
  activeTabId,
  markerName,
  namedRangeId
) {
  for (let i = 0; i < tabs.length; i++) {
    const tab = tabs[i];
    const tabId =
      tab.tabProperties &&
      tab.tabProperties.tabId;

    if (
      tabId === activeTabId &&
      tab.documentTab &&
      tab.documentTab.namedRanges
    ) {
      const group =
        tab.documentTab
          .namedRanges[markerName];

      const namedRanges =
        group &&
        group.namedRanges
          ? group.namedRanges
          : [];

      for (
        let j = 0;
        j < namedRanges.length;
        j++
      ) {
        if (
          namedRanges[j].namedRangeId ===
          namedRangeId
        ) {
          return namedRanges[j];
        }
      }
    }

    const childResult =
      testFindNamedRangeInTabs_(
        tab.childTabs || [],
        activeTabId,
        markerName,
        namedRangeId
      );

    if (childResult) {
      return childResult;
    }
  }

  return null;
}
