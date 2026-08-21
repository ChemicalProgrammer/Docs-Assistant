/**
 * TEST LAB
 *
 * TEST-017:
 * Convierte líneas seleccionadas dentro de un único ListItem
 * en elementos nativos independientes de la misma lista.
 *
 * Preparación esperada:
 *
 * a) Julio
 *    Agosto
 *    Septiembre
 *
 * Seleccionar únicamente:
 * Agosto
 * Septiembre
 */

function runCurrentTest() {
  return testSplitSelectedLinesIntoNativeListItems_();
}


function testSplitSelectedLinesIntoNativeListItems_() {
  const started = Date.now();
  const document = DocumentApp.getActiveDocument();
  const selection = document.getSelection();

  if (!selection) {
    throw new Error(
      'Select Agosto and Septiembre before running the test.'
    );
  }

  const rangeElements = selection.getRangeElements();

  if (rangeElements.length !== 1) {
    throw new Error(
      'For this test, the selection must be inside one native list item.'
    );
  }

  const rangeElement = rangeElements[0];
  const element = rangeElement.getElement();

  if (
    element.getType() !== DocumentApp.ElementType.TEXT ||
    !rangeElement.isPartial()
  ) {
    throw new Error(
      'Select only the text of Agosto and Septiembre.'
    );
  }

  const source = testFindListItemAncestor_(element);

  if (!source) {
    throw new Error(
      'The selected text is not inside a native list item.'
    );
  }

  const textElement = element.asText();
  const completeText = textElement.getText();

  const selectionStart = rangeElement.getStartOffset();
  const selectionEnd =
    rangeElement.getEndOffsetInclusive() + 1;

  if (
    selectionStart > 0 &&
    completeText.charAt(selectionStart - 1) !== '\r'
  ) {
    throw new Error(
      'The selection must begin at the start of Agosto.'
    );
  }

  if (
    selectionEnd < completeText.length &&
    completeText.charAt(selectionEnd) !== '\r'
  ) {
    throw new Error(
      'The selection must end at the end of Septiembre.'
    );
  }

  const selectedText = completeText.substring(
    selectionStart,
    selectionEnd
  );

  const selectedLines = selectedText
    .split(/\r|\n/)
    .map(function (line) {
      return line.trim();
    })
    .filter(function (line) {
      return line.length > 0;
    });

  if (selectedLines.length === 0) {
    throw new Error(
      'No non-empty lines were found in the selection.'
    );
  }

  let remainingBefore = completeText.substring(
    0,
    selectionStart
  );

  const remainingAfter = completeText.substring(
    selectionEnd
  );

  remainingBefore = remainingBefore.replace(/\r+$/, '');

  if (!remainingBefore.trim()) {
    throw new Error(
      'Leave the first list item, such as Julio, outside the selection.'
    );
  }

  if (remainingAfter.replace(/\r/g, '').trim()) {
    throw new Error(
      'For this test, select through the final logical line.'
    );
  }

  const parent = source.getParent();

  if (
    !parent ||
    typeof parent.insertListItem !== 'function'
  ) {
    throw new Error(
      'The native list item is inside an unsupported container.'
    );
  }

  const sourceIndex = parent.getChildIndex(source);
  const sourceListId = source.getListId();
  const sourceNestingLevel = source.getNestingLevel();
  const sourceGlyphType = String(source.getGlyphType());

  // La copia conserva la configuración a), b), c)…
  const nativeTemplate = source.copy();

  // El elemento original conserva únicamente Julio.
  source.setText(remainingBefore);

  source
    .setIndentFirstLine(18)
    .setIndentStart(36)
    .setIndentEnd(0);

  const insertedItems = [];

  for (let i = 0; i < selectedLines.length; i++) {
    const detachedCopy = nativeTemplate.copy();

    detachedCopy.setText(selectedLines[i]);

    const inserted = parent.insertListItem(
      sourceIndex + 1 + i,
      detachedCopy
    );

    // Garantiza que pertenece a la misma numeración nativa.
    inserted.setListId(source);

    inserted
      .setNestingLevel(sourceNestingLevel)
      .setIndentFirstLine(18)
      .setIndentStart(36)
      .setIndentEnd(0);

    insertedItems.push({
      text: inserted.getText(),
      listId: inserted.getListId(),
      sameListId:
        inserted.getListId() === sourceListId,
      glyphType: String(inserted.getGlyphType()),
      nestingLevel: inserted.getNestingLevel(),
      isNativeListItem:
        inserted.getType() ===
        DocumentApp.ElementType.LIST_ITEM
    });
  }

  const result = {
    testId: 'TEST-017-SPLIT-NATIVE-INCISOS',
    ok: true,
    sourceTextAfter: source.getText(),
    sourceListId: sourceListId,
    sourceGlyphType: sourceGlyphType,
    selectedLogicalLines: selectedLines.length,
    insertedItems: insertedItems,
    allSameListId: insertedItems.every(function (item) {
      return item.sameListId;
    }),
    allNative: insertedItems.every(function (item) {
      return item.isNativeListItem;
    }),
    elapsedMs: Date.now() - started
  };

  document.saveAndClose();

  return result;
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
