/**
 * TEST LAB
 *
 * TEST-018 corregido:
 * Crear una lista nativa a) independiente mediante
 * un documento temporal.
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
  const sourceIndex =
    activeParent.getChildIndex(source);

  const sourceListId = source.getListId();
  const sourceGlyphType =
    String(source.getGlyphType());

  let temporaryDocument = null;
  let temporaryDocumentId = null;
  let temporaryDocumentClosed = false;
  let temporaryFileTrashed = false;
  let cleanupError = null;
  let result = null;

  try {
    temporaryDocument = DocumentApp.create(
      'DocsAssistant Native List Bridge ' +
      Date.now()
    );

    temporaryDocumentId =
      temporaryDocument.getId();

    const temporaryBody =
      temporaryDocument.getBody();

    /*
     * Primera transferencia:
     * documento activo → documento temporal.
     */
    const temporaryItem =
      temporaryBody.insertListItem(
        0,
        source.copy()
      );

    temporaryItem.setText(
      'NATIVE LIST TEMPLATE'
    );

    const temporaryListId =
      temporaryItem.getListId();

    /*
     * Segunda transferencia:
     * documento temporal → documento activo.
     *
     * Debe realizarse antes de cerrar
     * el documento temporal.
     */
    const bridgeCopy = temporaryItem.copy();

    bridgeCopy.setText(
      'NATIVE NEW LIST TEST'
    );

    const inserted =
      activeParent.insertListItem(
        sourceIndex + 1,
        bridgeCopy
      );

    inserted
      .setIndentFirstLine(18)
      .setIndentStart(36)
      .setIndentEnd(0);

    const insertedListId =
      inserted.getListId();

    result = {
      testId:
        'TEST-018-INDEPENDENT-NATIVE-INCISO',

      ok: true,

      sourceListId: sourceListId,
      temporaryListId: temporaryListId,
      insertedListId: insertedListId,

      independentFromSource:
        insertedListId !== sourceListId,

      insertedIsNative:
        inserted.getType() ===
        DocumentApp.ElementType.LIST_ITEM,

      sourceGlyphType:
        sourceGlyphType,

      insertedGlyphType:
        String(inserted.getGlyphType()),

      elapsedMsBeforeCleanup:
        Date.now() - started
    };

    /*
     * Cerramos ambos documentos solamente después
     * de terminar las inserciones.
     */
    temporaryDocument.saveAndClose();
    temporaryDocumentClosed = true;

    document.saveAndClose();

  } finally {
    if (
      temporaryDocument &&
      !temporaryDocumentClosed
    ) {
      try {
        temporaryDocument.saveAndClose();
        temporaryDocumentClosed = true;
      } catch (error) {
        cleanupError =
          'Close error: ' + String(error);
      }
    }

    if (temporaryDocumentId) {
      try {
        DriveApp
          .getFileById(temporaryDocumentId)
          .setTrashed(true);

        temporaryFileTrashed = true;

      } catch (error) {
        cleanupError =
          cleanupError ||
          'Trash error: ' + String(error);
      }
    }
  }

  result.temporaryDocumentClosed =
    temporaryDocumentClosed;

  result.temporaryFileTrashed =
    temporaryFileTrashed;

  result.cleanupError = cleanupError;
  result.elapsedMs = Date.now() - started;

  return result;
}
