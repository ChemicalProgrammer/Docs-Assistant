/*
 * Cambia solamente este valor:
 * LETTER | NUMBER | BULLET | ROMAN
 */
var TEST_NATIVE_LIST_TYPE = 'LETTER';


function runCurrentTest() {
  return testApplyNativeIncisoFromPrevious_();
}


/**
 * Aplica el formato nativo a), b), c) utilizando la definición
 * de la lista latina nativa anterior más cercana.
 *
 * No intenta reiniciar ni decidir si debe continuar.
 */
function testApplyNativeIncisoFromPrevious_() {
  var started = Date.now();

  var targets = getNativeListTestTargets_();

  if (!targets.length) {
    throw new Error(
      'Place the cursor in a line or select one or more lines.'
    );
  }

  /*
   * Helper que ya agregamos para convertir Shift+Enter
   * en líneas independientes.
   */
  targets = expandSoftBreakTargetsForNativeTest_(targets);

  if (!targets.length) {
    throw new Error('No target lines were found.');
  }

  var parent = targets[0].getParent();
  var donor = findPreviousNativeInciso_(parent, targets[0]);

  if (!donor) {
    throw new Error(
      'No previous native a) list was found in this container.'
    );
  }

  var created = [];

  targets.forEach(function(source) {
    var sourceIndex;

    try {
      sourceIndex = parent.getChildIndex(source);
    } catch (error) {
      throw new Error(
        'The selected lines must belong to the same document container.'
      );
    }

    if (sourceIndex < 0) {
      throw new Error(
        'The selected lines must belong to the same document container.'
      );
    }

    var content = source.getText();
    var paragraphAttributes = source.getAttributes();
    var textAttributes = {};

    if (content.length) {
      try {
        textAttributes = source.editAsText().getAttributes();
      } catch (error) {}
    }

    var item = parent.insertListItem(
      sourceIndex,
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
     * Esta es la operación que ya había funcionado:
     * copiar la definición nativa que muestra a).
     */
    item.setListId(donor);
    item.setNestingLevel(donor.getNestingLevel());

    // Left 0.25", Hanging 0.25", Right 0".
    item.setIndentFirstLine(18);
    item.setIndentStart(36);
    item.setIndentEnd(0);

    created.push(item);
    source.removeFromParent();
  });

  return {
    testId: 'TEST-NATIVE-INCISO-PROVEN-METHOD',
    itemsApplied: created.length,
    glyphType: String(created[0].getGlyphType()),
    listId: created[0].getListId(),
    nestingLevel: created[0].getNestingLevel(),
    elapsedMs: Date.now() - started,
    ok: true
  };
}


/**
 * Busca hacia arriba el inciso latino automático más cercano.
 */
function findPreviousNativeInciso_(parent, target) {
  var targetIndex = parent.getChildIndex(target);

  for (var i = targetIndex - 1; i >= 0; i--) {
    var child = parent.getChild(i);

    if (
      child.getType() !==
      DocumentApp.ElementType.LIST_ITEM
    ) {
      continue;
    }

    var item = child.asListItem();

    try {
      if (
        item.getGlyphType() ===
        DocumentApp.GlyphType.LATIN_LOWER
      ) {
        return item;
      }
    } catch (error) {}
  }

  return null;
}


function testNativeAutomaticList_(requestedType) {
  var started = Date.now();
  var type = String(requestedType || '').toUpperCase();

  var glyphs = {
    BULLET: DocumentApp.GlyphType.BULLET,
    NUMBER: DocumentApp.GlyphType.NUMBER,
    LETTER: DocumentApp.GlyphType.LATIN_LOWER,
    ROMAN: DocumentApp.GlyphType.ROMAN_UPPER
  };

  var glyph = glyphs[type];

  if (!glyph) {
    throw new Error(
      'Invalid test type. Use BULLET, NUMBER, LETTER or ROMAN.'
    );
  }

  var targets = getNativeListTestTargets_();

  if (!targets.length) {
    throw new Error(
      'Place the cursor in a paragraph or select several paragraphs.'
    );
  }

targets = expandSoftBreakTargetsForNativeTest_(targets);

  var parent = targets[0].getParent();

  targets.forEach(function(paragraph) {
    if (paragraph.getParent() !== parent) {
      throw new Error(
        'The test selection must be inside the same document container.'
      );
    }
  });

  for (var i = 1; i < targets.length; i++) {
    if (
      parent.getChildIndex(targets[i]) !==
      parent.getChildIndex(targets[i - 1]) + 1
    ) {
      throw new Error(
        'Select one contiguous block of paragraphs.'
      );
    }
  }

  /*
   * Los separadores temporales impiden que la prueba reutilice
   * accidentalmente una lista anterior.
   */
  var before = parent.insertParagraph(
    parent.getChildIndex(targets[0]),
    '\uE410'
  );

  var lastTarget = targets[targets.length - 1];

  var after = parent.insertParagraph(
    parent.getChildIndex(lastTarget) + 1,
    '\uE411'
  );

  var anchor = null;
  var created = [];

  try {
    targets.forEach(function(source) {
      var text = source.getText();
      var paragraphAttributes = source.getAttributes();
      var textAttributes = source.editAsText().getAttributes();

      var item = parent.insertListItem(
        parent.getChildIndex(source),
        text
      );

      /*
       * Conserva el formato común del texto.
       */
      try {
        item.setAttributes(paragraphAttributes);
      } catch (error) {}

      if (text.length) {
        try {
          item.editAsText().setAttributes(textAttributes);
        } catch (error) {}
      }

      if (!anchor) {
        /*
         * Primer elemento: crea la lista automática nativa.
         */
        item.setGlyphType(glyph);
        anchor = item;
      } else {
        /*
         * Los demás párrafos pertenecen a la misma lista automática.
         */
        item.setListId(anchor);
      }

      item.setNestingLevel(0);

      /*
       * Left = 0.25"
       * Hanging = 0.25"
       * Right = 0"
       */
      item.setIndentFirstLine(18);
      item.setIndentStart(36);
      item.setIndentEnd(0);

      created.push(item);
      source.removeFromParent();
    });
  } finally {
    try {
      before.removeFromParent();
    } catch (error) {}

    try {
      after.removeFromParent();
    } catch (error) {}
  }

  var sameListId = true;
  var anchorListId = anchor.getListId();

  created.forEach(function(item) {
    if (item.getListId() !== anchorListId) {
      sameListId = false;
    }
  });

  return {
    testId: 'TEST-NATIVE-LIST-GLYPH',
    requestedType: type,
    itemsApplied: created.length,
    glyphType: String(anchor.getGlyphType()),
    sameListId: sameListId,
    nestingLevel: anchor.getNestingLevel(),
    elapsedMs: Date.now() - started,
    ok: true
  };
}


function getNativeListTestTargets_() {
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  var targets = [];

  function addTarget(element) {
    var paragraph = findNativeListParagraph_(element);

    if (
      paragraph &&
      targets.indexOf(paragraph) === -1
    ) {
      targets.push(paragraph);
    }
  }

  if (selection) {
    selection.getRangeElements().forEach(function(rangeElement) {
      addTarget(rangeElement.getElement());
    });

    return targets;
  }

  var cursor = doc.getCursor();

  if (cursor) {
    addTarget(cursor.getElement());
  }

  return targets;
}


function findNativeListParagraph_(element) {
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
function expandSoftBreakTargetsForNativeTest_(targets) {
  var expanded = [];

  targets.forEach(function(source) {
    var sourceText = source.getText();
    var lines = sourceText.split(/\r\n|\r|\n/);

    if (lines.length === 1) {
      expanded.push(source);
      return;
    }

    var parent = source.getParent();
    var sourceIndex = parent.getChildIndex(source);
    var paragraphAttributes = source.getAttributes();

    /*
     * Para esta prueba conservamos los atributos comunes del texto.
     */
    var textAttributes = {};

    if (sourceText.length) {
      try {
        textAttributes = source.editAsText().getAttributes();
      } catch (error) {}
    }

    var created = [];

    lines.forEach(function(line, index) {
      var paragraph = parent.insertParagraph(
        sourceIndex + index,
        line
      );

      try {
        paragraph.setAttributes(paragraphAttributes);
      } catch (error) {}

      if (line.length) {
        try {
          paragraph.editAsText().setAttributes(textAttributes);
        } catch (error) {}
      }

      created.push(paragraph);
    });

    source.removeFromParent();

    created.forEach(function(paragraph) {
      expanded.push(paragraph);
    });
  });

  return expanded;
}
