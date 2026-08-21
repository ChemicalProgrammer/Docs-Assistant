/**
 * TEST LAB
 *
 * Diagnóstico del cursor y de la selección actual.
 * No modifica el documento.
 */

function runCurrentTest() {
  return testInspectCurrentContext_();
}


function testInspectCurrentContext_() {
  const started = Date.now();
  const document = DocumentApp.getActiveDocument();

  const cursor = document.getCursor();
  const selection = document.getSelection();

  const result = {
    testId: 'TEST-019-CONTEXT-DIAGNOSTIC',
    ok: true,
    hasCursor: Boolean(cursor),
    hasSelection: Boolean(selection),
    cursor: null,
    selection: [],
    elapsedMs: 0
  };

  if (cursor) {
    result.cursor = testDescribeElement_(
      cursor.getElement()
    );
  }

  if (selection) {
    const rangeElements = selection.getRangeElements();

    for (let i = 0; i < rangeElements.length; i++) {
      const rangeElement = rangeElements[i];
      const element = rangeElement.getElement();

      const description = testDescribeElement_(element);

      description.rangeIndex = i;
      description.isPartial = rangeElement.isPartial();
      description.startOffset =
        rangeElement.getStartOffset();
      description.endOffsetInclusive =
        rangeElement.getEndOffsetInclusive();

      result.selection.push(description);
    }
  }

  result.elapsedMs = Date.now() - started;

  return result;
}


function testDescribeElement_(element) {
  const description = {
    elementType: null,
    text: null,
    ancestors: []
  };

  if (!element) {
    return description;
  }

  description.elementType = String(element.getType());

  try {
    description.text = element.getText();
  } catch (error) {
    description.text = null;
  }

  let current = element;
  let depth = 0;

  while (current && depth < 15) {
    const ancestor = {
      depth: depth,
      type: String(current.getType()),
      text: null,
      listId: null,
      glyphType: null,
      nestingLevel: null
    };

    try {
      ancestor.text = current.getText();
    } catch (error) {
      ancestor.text = null;
    }

    if (
      current.getType() ===
      DocumentApp.ElementType.LIST_ITEM
    ) {
      const listItem = current.asListItem();

      ancestor.listId = listItem.getListId();
      ancestor.glyphType =
        String(listItem.getGlyphType());
      ancestor.nestingLevel =
        listItem.getNestingLevel();
    }

    description.ancestors.push(ancestor);

    current = current.getParent();
    depth++;
  }

  return description;
}
