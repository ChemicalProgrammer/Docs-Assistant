/**
 * TEST LAB
 *
 * TEST-018:
 * Comprueba si una copia trasladada mediante otro documento
 * conserva el formato nativo a), pero recibe un List ID nuevo.
 */

function runCurrentTest() {
  return testCreateIndependentNativeLetterList_();
}


function testCreateIndependentNativeLetterList_() {
  const started = Date.now();
  const document = DocumentApp.getActiveDocument();

  const source = testGetNativeListItemFromContext_(document);

  if (!source) {
    throw new Error(
      'Place the cursor inside a native a) list item.'
    );
  }

  const activeParent = source.getParent();

  if (
    !activeParent ||
    typeof activeParent.insertListItem !== 'function'
  ) {
    throw new Error(
      'The source list item is inside an unsupported container.'
    );
  }

  const sourceIndex = activeParent.getChildIndex(source);
  const sourceListId = source.getListId();
  const sourceGlyphType = String(source.getGlyphType());
  const sourceNestingLevel = source.getNestingLevel();

  let temporaryDocumentId = null;
  let temporaryFileTrashed = false;
  let cleanupError = null;
  let result = null;

  try {
    /*
     * El documento intermedio obliga a Google Docs a remapear
     * la definición y el ID de la lista.
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
     * Inserta nuevamente la copia en el documento activo.
     * Su List ID procede del documento temporal y no debería
     * coincidir con el List ID original.
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

      source: {
        text: source.getText(),
        listId: sourceListId,
        glyphType: sourceGlyphType,
        nestingLevel: sourceNestingLevel
      },

      temporary: {
        listId: temporaryListId
      },

      inserted: {
        text: inserted.getText(),
        listId: insertedListId,
        glyphType: String(inserted.getGlyphType()),
        nestingLevel: inserted.getNestingLevel(),
        indentFirstLine: inserted.getIndentFirstLine(),
        indentStart: inserted.getIndentStart(),
        indentEnd: inserted.getIndentEnd(),
        isNativeListItem:
          inserted.getType() ===
          DocumentApp.ElementType.LIST_ITEM
      },

      independentFromSource:
        insertedListId !== sourceListId,

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

  result.temporaryFileTrashed = temporaryFileTrashed;
  result.cleanupError = cleanupError;
  result.elapsedMs = Date.now() - started;

  return result;
}


function testGetNativeListItemFromContext_(document) {
  const cursor = document.getCursor();

  if (cursor) {
    const cursorItem = testFindListItemAncestor_(
      cursor.getElement()
    );

    if (cursorItem) {
      return cursorItem;
    }
  }

  const selection = document.getSelection();

  if (!selection) {
    return null;
  }

  const rangeElements = selection.getRangeElements();

  for (let i = 0; i < rangeElements.length; i++) {
    const listItem = testFindListItemAncestor_(
      rangeElements[i].getElement()
    );

    if (listItem) {
      return listItem;
    }
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
