function runCurrentTest() {
  var started = Date.now();
  var doc = DocumentApp.getActiveDocument();
  var cursor = doc.getCursor();

  if (!cursor) {
    throw new Error(
      'Place only the cursor inside the line. Do not select text.'
    );
  }

  var source = findTestParagraph_(cursor.getElement());

  if (!source) {
    throw new Error('No paragraph was found at the cursor.');
  }

  /*
   * Si contiene Shift+Enter, extrae solamente la línea del cursor.
   */
  source = extractCursorLine_(source, cursor);

  var parent = source.getParent();
  var sourceIndex = parent.getChildIndex(source);
  var paragraphAttributes = source.getAttributes();
  var content = source.getText();

  var textAttributes = {};

  if (content.length) {
    try {
      textAttributes = source.editAsText().getAttributes();
    } catch (error) {}
  }

  /*
   * Aíslan la nueva lista de cualquier lista anterior o posterior.
   */
  var before = parent.insertParagraph(
    sourceIndex,
    '\uE510'
  );

  var after = parent.insertParagraph(
    parent.getChildIndex(source) + 1,
    '\uE511'
  );

  var item;

  try {
    item = parent.insertListItem(
      parent.getChildIndex(source),
      content
    );

    try {
      item.setAttributes(paragraphAttributes);
    } catch (error) {}

    if (content.length) {
      try {
        item.editAsText().setAttributes(textAttributes);
      } catch (error) {}
    }

    /*
     * Nueva lista automática.
     * No reutiliza listId y no busca listas anteriores.
     */
 /*
 * Primero establece el nivel y después el tipo de marcador.
 */
item.setNestingLevel(1);

item.setGlyphType(
  DocumentApp.GlyphType.LATIN_LOWER
);

    // Left 0.25", Hanging 0.25", Right 0".
    item.setIndentFirstLine(18);
    item.setIndentStart(36);
    item.setIndentEnd(0);

    source.removeFromParent();
  } finally {
    try {
      before.removeFromParent();
    } catch (error) {}

    try {
      after.removeFromParent();
    } catch (error) {}
  }

  return {
    testId: 'TEST-DIRECT-NATIVE-LETTER',
    glyphType: String(item.getGlyphType()),
    nestingLevel: item.getNestingLevel(),
    elapsedMs: Date.now() - started,
    ok: true
  };
}


function findTestParagraph_(element) {
  var current = element;

  while (current) {
    var type = current.getType();

    if (
      type === DocumentApp.ElementType.PARAGRAPH ||
      type === DocumentApp.ElementType.LIST_ITEM
    ) {
      return current;
    }

    current = current.getParent();
  }

  return null;
}


function extractCursorLine_(source, cursor) {
  var sourceText = source.getText();

  if (!/[\r\n]/.test(sourceText)) {
    return source;
  }

  var rawElement = cursor.getElement();
  var cursorOffset = Math.max(
    0,
    cursor.getSurroundingTextOffset()
  );

  if (rawElement.getType() === DocumentApp.ElementType.TEXT) {
    cursorOffset += getTextOffsetInTestParagraph_(
      source,
      rawElement
    );
  }

  var lines = [];
  var breakPattern = /\r\n|\r|\n/g;
  var start = 0;
  var match;

  while ((match = breakPattern.exec(sourceText)) !== null) {
    lines.push({
      text: sourceText.substring(start, match.index),
      start: start
    });

    start = match.index + match[0].length;
  }

  lines.push({
    text: sourceText.substring(start),
    start: start
  });

  var selectedLine = 0;

  for (var i = 1; i < lines.length; i++) {
    if (cursorOffset >= lines[i].start) {
      selectedLine = i;
    }
  }

  var parent = source.getParent();
  var sourceIndex = parent.getChildIndex(source);
  var paragraphAttributes = source.getAttributes();
  var created = [];

  lines.forEach(function(line, index) {
    var paragraph = parent.insertParagraph(
      sourceIndex + index,
      line.text
    );

    try {
      paragraph.setAttributes(paragraphAttributes);
    } catch (error) {}

    created.push(paragraph);
  });

  source.removeFromParent();

  return created[selectedLine];
}


function getTextOffsetInTestParagraph_(owner, target) {
  if (owner === target) return 0;

  var offset = 0;
  var found = false;

  function walk(element) {
    if (element === target) {
      found = true;
      return true;
    }

    if (element.getType() === DocumentApp.ElementType.TEXT) {
      offset += element.asText().getText().length;
      return false;
    }

    if (!element.getNumChildren) {
      return false;
    }

    for (var i = 0; i < element.getNumChildren(); i++) {
      if (walk(element.getChild(i))) {
        return true;
      }
    }

    return false;
  }

  walk(owner);

  return found ? offset : 0;
}
