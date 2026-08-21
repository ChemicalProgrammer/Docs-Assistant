/**
 * TEST LAB
 *
 * TEST-020:
 * Crear una lista nativa nueva a) usando:
 * - NamedRange local.
 * - Lectura parcial de la API.
 * - NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS.
 *
 * Seleccionar únicamente una línea normal: Octubre.
 */

function runCurrentTest() {
  return testCreateNewNativeInciso_();
}


function testCreateNewNativeInciso_() {
  const started = Date.now();
  const document = DocumentApp.getActiveDocument();
  const documentId = document.getId();
  const selection = document.getSelection();

  if (!selection) {
    throw new Error(
      'Select the normal paragraph Octubre.'
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
      'For this test, select a normal paragraph without numbering.'
    );
  }

  const markerName =
    'DOCSASSISTANT_TEST_' +
    Utilities.getUuid();

  const namedRange = document.addNamedRange(
    markerName,
    selection
  );

  const namedRangeId = namedRange.getId();

  /*
   * Un tabulador inicial indica nivel 1.
   * En el preset seleccionado, el nivel 1 es a).
   * La API elimina automáticamente el tabulador.
   */
  targetParagraph
    .editAsText()
    .insertText(0, '\t');

  document.saveAndClose();

  const readStarted = Date.now();

  /*
   * Solo recupera los rangos nombrados.
   * No descarga el contenido completo del documento.
   */
  const apiDocument = Docs.Documents.get(
    documentId,
    {
      fields: 'namedRanges'
    }
  );

  const apiReadMs =
    Date.now() - readStarted;

  const apiNamedRange =
    testFindApiNamedRange_(
      apiDocument,
      markerName,
      namedRangeId
    );

  if (
    !apiNamedRange ||
    !apiNamedRange.ranges ||
    apiNamedRange.ranges.length === 0
  ) {
    throw new Error(
      'The temporary named range was not returned by the Docs API.'
    );
  }

  const sourceRange =
    apiNamedRange.ranges[0];

  const requestRange = {
    startIndex: sourceRange.startIndex,
    endIndex: sourceRange.endIndex
  };

  if (sourceRange.segmentId) {
    requestRange.segmentId =
      sourceRange.segmentId;
  }

  if (sourceRange.tabId) {
    requestRange.tabId =
      sourceRange.tabId;
  }

  const updateStarted = Date.now();

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
    Date.now() - updateStarted;

  /*
   * Reabre el documento para aplicar las sangrías
   * mediante DocumentApp y retirar el rango auxiliar.
   */
  const reopened =
    DocumentApp.openById(documentId);

  const savedNamedRange =
    reopened.getNamedRangeById(namedRangeId);

  if (!savedNamedRange) {
    throw new Error(
      'The named range could not be reopened after the API update.'
    );
  }

  const updatedRangeElements =
    savedNamedRange
      .getRange()
      .getRangeElements();

  let insertedListItem = null;

  for (
    let i = 0;
    i < updatedRangeElements.length;
    i++
  ) {
    insertedListItem =
      testFindListItemAncestor_(
        updatedRangeElements[i].getElement()
      );

    if (insertedListItem) {
      break;
    }
  }

  if (!insertedListItem) {
    throw new Error(
      'The API request did not create a native list item.'
    );
  }

  insertedListItem
    .setIndentFirstLine(18)
    .setIndentStart(36)
    .setIndentEnd(0);

  const result = {
    testId:
      'TEST-020-NEW-NATIVE-INCISO',

    ok: true,

    text:
      insertedListItem.getText(),

    listId:
      insertedListItem.getListId(),

    glyphType:
      String(insertedListItem.getGlyphType()),

    nestingLevel:
      insertedListItem.getNestingLevel(),

    isNative:
      insertedListItem.getType() ===
      DocumentApp.ElementType.LIST_ITEM,

    indentFirstLine:
      insertedListItem.getIndentFirstLine(),

    indentStart:
      insertedListItem.getIndentStart(),

    indentEnd:
      insertedListItem.getIndentEnd(),

    apiReadMs: apiReadMs,
    apiWriteMs: apiWriteMs,
    elapsedMs: Date.now() - started
  };

  savedNamedRange.remove();
  reopened.saveAndClose();

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


function testFindApiNamedRange_(
  apiDocument,
  markerName,
  namedRangeId
) {
  if (
    !apiDocument.namedRanges ||
    !apiDocument.namedRanges[markerName]
  ) {
    return null;
  }

  const group =
    apiDocument.namedRanges[markerName];

  const ranges = group.namedRanges || [];

  for (let i = 0; i < ranges.length; i++) {
    if (
      ranges[i].namedRangeId ===
      namedRangeId
    ) {
      return ranges[i];
    }
  }

  return null;
}
