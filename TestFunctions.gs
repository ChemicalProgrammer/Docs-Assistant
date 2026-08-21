/**
 * TEST LAB
 *
 * TEST-018:
 * Crear una lista nativa a) independiente utilizando
 * un documento temporal como puente.
 *
 * Seleccionar únicamente la palabra "Julio".
 */

function runCurrentTest() {
  return testCreateIndependentNativeInciso_();
}


function testCreateIndependentNativeInciso_() {
  const started = Date.now();
  const document = DocumentApp.getActiveDocument();
  const selection = document.getSelection();

  if (!selection) {
    throw new Error(
      'Select the word Julio inside the native a) item.'
    );
  }

  const rangeElements = selection.getRangeElements();

  if (rangeElements.length === 0) {
    throw new Error('The selection is empty.');
  }

  let source = rangeElements[0].getElement();

  while (
    source &&
    source.getType() !==
      DocumentApp.ElementType.LIST_ITEM
  ) {
    source = source.getParent();
  }

  if (
    !source ||
    source.getType() !==
      DocumentApp.ElementType.LIST_ITEM
  ) {
    throw new Error(
      'The selection is not inside a native list item.'
    );
  }

  source = source.asListItem();

  const activeParent = source.getParent();
  const sourceIndex = activeParent.getChildIndex(source);

  const sourceListId = source.getListId();
  const sourceGlyphType = String(source.getGlyphType());

  let temporaryDocumentId = null;
  let temporaryFileTrashed = false;
  let cleanupError = null;
  let result = null;

  try {
    /*
     * Copia el inciso a otro documento.
     * Google debe crear un List ID válido para ese documento.
     */
    const temporaryDocument = DocumentApp.create(
      'DocsAssistant Native List Bridge ' + Date.now()
    );

    temporaryDocumentId = temporaryDocument.getId();

    const temporaryBody = temporaryDocument.getBody();

    const temporaryItem = temporaryBody.insertListItem(
      0,
      source.copy()
    );

    temporaryItem.setText('NATIVE LIST TEMPLATE');

    const temporaryListId = temporaryItem.getListId();
    const bridgeCopy = temporaryItem.copy();

    temporaryDocument.saveAndClose();

    /*
     * Regresa la copia al documento activo.
     * Debe producir una lista nativa independiente.
     */
    bridgeCopy.setText('NATIVE NEW LIST TEST');

    const inserted = activeParent.insertListItem(
      sourceIndex + 1,
      bridgeCopy
    );

    inserted
      .setIndentFirstLine(18)
      .setIndentStart(36)
      .setIndentEnd(0);

    const insertedListId = inserted.getListId();

    result = {
      testId: 'TEST-018-INDEPENDENT-NATIVE-INCISO',
      ok: true,
      sourceListId: sourceListId,
      temporaryListId: temporaryListId,
      insertedListId: insertedListId,
      independentFromSource:
        insertedListId !== sourceListId,
      insertedIsNative:
        inserted.getType() ===
        DocumentApp.ElementType.LIST_ITEM,
      sourceGlyphType: sourceGlyphType,
      insertedGlyphType:
        String(inserted.getGlyphType()),
      elapsedMsBeforeCleanup:
        Date.now() - started
    };

    document.saveAndClose();

  } finally {
    if (temporaryDocumentId) {
      try {
        DriveApp
          .getFileById(temporaryDocumentId)
          .setTrashed(true);

        temporaryFileTrashed = true;

      } catch (error) {
        cleanupError = String(error);
      }
    }
  }

  result.temporaryFileTrashed =
    temporaryFileTrashed;

  result.cleanupError = cleanupError;
  result.elapsedMs = Date.now() - started;

  return result;
}
