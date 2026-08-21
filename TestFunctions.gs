/**
 * TEST LAB
 *
 * Prueba actual:
 * Verificar si un inciso nativo a) puede copiarse conservando:
 * - Numeración nativa de Google Docs.
 * - Sufijo ).
 * - List ID.
 * - Nivel de anidación.
 *
 * Admite cursor o selección dentro del inciso nativo.
 */

function runCurrentTest() {
  return testCopyNativeLetterListItem_();
}


/**
 * TEST-016
 *
 * Copia un inciso nativo e inserta la copia inmediatamente
 * después del elemento original.
 */
function testCopyNativeLetterListItem_() {
  const started = Date.now();
  const document = DocumentApp.getActiveDocument();

  const source = testGetNativeListItemFromContext_(document);

  if (!source) {
    throw new Error(
      'Place the cursor or selection inside a native a) list item.'
    );
  }

  const parent = source.getParent();

  if (
    !parent ||
    typeof parent.insertListItem !== 'function'
  ) {
    throw new Error(
      'The selected list item is inside an unsupported container.'
    );
  }

  const sourceIndex = parent.getChildIndex(source);

  const before = {
    text: source.getText(),
    listId: source.getListId(),
    glyphType: String(source.getGlyphType()),
    nestingLevel: source.getNestingLevel(),
    indentFirstLine: source.getIndentFirstLine(),
    indentStart: source.getIndentStart(),
    indentEnd: source.getIndentEnd()
  };

  const detachedCopy = source.copy();

  detachedCopy.setText('NATIVE INCISO COPY TEST');

  const inserted = parent.insertListItem(
    sourceIndex + 1,
    detachedCopy
  );

  // Left = 0.25 in
  // Hanging = 0.25 in
  // Right = 0
  inserted
    .setIndentFirstLine(18)
    .setIndentStart(36)
    .setIndentEnd(0);

  const after = {
    text: inserted.getText(),
    listId: inserted.getListId(),
    glyphType: String(inserted.getGlyphType()),
    nestingLevel: inserted.getNestingLevel(),
    indentFirstLine: inserted.getIndentFirstLine(),
    indentStart: inserted.getIndentStart(),
    indentEnd: inserted.getIndentEnd()
  };

  const result = {
    testId: 'TEST-016-NATIVE-INCISO-COPY',
    ok: true,
    before: before,
    after: after,
    sameListId: before.listId === after.listId,
    insertedIsNativeListItem:
      inserted.getType() ===
      DocumentApp.ElementType.LIST_ITEM,
    elapsedMs: Date.now() - started
  };

  document.saveAndClose();

  return result;
}


/**
 * Obtiene un ListItem nativo desde el cursor o la selección.
 */
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


/**
 * Recorre los padres de un elemento hasta encontrar
 * el ListItem nativo que lo contiene.
 */
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
