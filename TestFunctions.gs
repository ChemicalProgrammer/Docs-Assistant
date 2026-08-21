/**
 * TEST LAB
 * El botón siempre llama esta función.
 * Para cada experimento cambiaremos únicamente la prueba ejecutada aquí.
 */
function runCurrentTest() {
  return testCopyNativeLetterListItem_();
}

function testCopyNativeLetterListItem_() {
  const started = Date.now();
  const document = DocumentApp.getActiveDocument();
  const cursor = document.getCursor();

  if (!cursor) {
    throw new Error(
      'Place the cursor inside a native a) list item.'
    );
  }

  let source = cursor.getElement();

  while (
    source &&
    source.getType() !== DocumentApp.ElementType.LIST_ITEM
  ) {
    source = source.getParent();
  }

  if (
    !source ||
    source.getType() !== DocumentApp.ElementType.LIST_ITEM
  ) {
    throw new Error(
      'The cursor is not inside a native Google Docs list item.'
    );
  }

  const parent = source.getParent();
  const sourceIndex = parent.getChildIndex(source);
  const sourceListId = source.getListId();
  const sourceGlyphType = String(source.getGlyphType());
  const sourceNestingLevel = source.getNestingLevel();

  const detachedCopy = source.copy();
  detachedCopy.setText('NATIVE INCISO COPY TEST');

  const inserted = parent.insertListItem(
    sourceIndex + 1,
    detachedCopy
  );

  inserted
    .setIndentFirstLine(18) // posición del marcador
    .setIndentStart(36)     // texto a 0.50 in
    .setIndentEnd(0);

  const insertedListId = inserted.getListId();

  document.saveAndClose();

  return {
    testId: 'TEST-016-NATIVE-INCISO-COPY',
    ok: true,
    sourceListId: sourceListId,
    insertedListId: insertedListId,
    sameListId: sourceListId === insertedListId,
    sourceGlyphType: sourceGlyphType,
    insertedGlyphType: String(inserted.getGlyphType()),
    sourceNestingLevel: sourceNestingLevel,
    insertedNestingLevel: inserted.getNestingLevel(),
    elapsedMs: Date.now() - started
  };
}

function testRestartExactLetterList_() {
  const started = Date.now();
  const document = DocumentApp.getActiveDocument();

  const targets = testGetParagraphsFromCurrentContext_(document);

  if (!targets || targets.length === 0) {
    throw new Error(
      'Place the cursor in a paragraph or select one or more lines.'
    );
  }

  // Checkbox desactivado:
  // no busca listas anteriores y siempre comienza desde a).
  const result = testCreateExactManualList_(
    targets,
    'LETTER',
    1
  );

  document.saveAndClose();

  let paragraphsApplied = null;

  if (Array.isArray(result)) {
    paragraphsApplied = result.length;
  } else if (
    result &&
    typeof result.paragraphsApplied === 'number'
  ) {
    paragraphsApplied = result.paragraphsApplied;
  }

  return {
    testId: 'TEST-015-RESTART-LETTER-LIST',
    ok: true,
    requestedType: 'LETTER',
    requestedContinue: false,
    initialOrdinal: 1,
    sourceElements: targets.length,
    paragraphsApplied: paragraphsApplied,
    elapsedMs: Date.now() - started
  };
}

/**
 * TEST-014
 *
 * Busca el inciso anterior y continúa desde el siguiente.
 *
 * Ejemplo:
 * c) Septiembre
 * d) Octubre
 * e) Noviembre
 */
function testContinueExactLetterList_() {
  const started = Date.now();

  const document =
    DocumentApp.getActiveDocument();

  const context =
    testGetParagraphsFromCurrentContext_(
      document
    );

  const targets = context.paragraphs;

  if (!targets.length) {
    throw new Error(
      'Selecciona las líneas nuevas que deben continuar la numeración.'
    );
  }

  const searchStarted = Date.now();

  const previous =
    testFindPreviousManualListOrdinal_(
      targets[0],
      'LETTER'
    );

  const searchMs =
    Date.now() - searchStarted;

  /*
   * Si encuentra c), comienza en d).
   * Si no encuentra una lista anterior, comienza en a).
   */
  const initialOrdinal =
    previous.ordinal > 0
      ? previous.ordinal + 1
      : 1;

  const updateStarted = Date.now();

  const result =
    testCreateExactManualList_(
      targets,
      'LETTER',
      initialOrdinal
    );

  const updateMs =
    Date.now() - updateStarted;

  document.saveAndClose();

  return {
    ok: true,
    testId:
      'TEST-014-CONTINUE-LETTER-LIST',

    requestedContinue: true,
    previousFound:
      previous.ordinal > 0,

    previousOrdinal:
      previous.ordinal,

    previousText:
      previous.text,

    siblingsScanned:
      previous.siblingsScanned,

    initialOrdinal:
      initialOrdinal,

    nextOrdinal:
      result.nextOrdinal,

    paragraphsCreated:
      result.paragraphs.length,

    items: result.paragraphs.map(
      function (paragraph, index) {
        return {
          index: index,
          text: paragraph.getText(),
          indentStart:
            paragraph.getIndentStart(),
          indentFirstLine:
            paragraph.getIndentFirstLine(),
          indentEnd:
            paragraph.getIndentEnd()
        };
      }
    ),

    searchMs: searchMs,
    updateMs: updateMs,
    apiReadMs: 0,
    apiWriteMs: 0,
    elapsedMs: Date.now() - started
  };
}

/**
 * Busca hacia atrás el elemento más cercano que tenga
 * un prefijo compatible.
 */
function testFindPreviousManualListOrdinal_(
  target,
  type
) {
  const parent = target.getParent();

  if (!parent) {
    return {
      ordinal: 0,
      text: '',
      siblingsScanned: 0
    };
  }

  const targetIndex =
    parent.getChildIndex(target);

  let siblingsScanned = 0;

  for (
    let index = targetIndex - 1;
    index >= 0;
    index--
  ) {
    const child = parent.getChild(index);

    const childType = child.getType();

    if (
      childType !==
        DocumentApp.ElementType.PARAGRAPH &&
      childType !==
        DocumentApp.ElementType.LIST_ITEM
    ) {
      continue;
    }

    siblingsScanned++;

    const logicalLines = String(
      child.getText() || ''
    ).split(/\r\n|\r|\n/);

    /*
     * Revisar desde la última línea hacia la primera.
     */
    for (
      let lineIndex =
        logicalLines.length - 1;
      lineIndex >= 0;
      lineIndex--
    ) {
      const line =
        logicalLines[lineIndex];

      const ordinal =
        testReadManualListOrdinal_(
          line,
          type
        );

      if (ordinal > 0) {
        return {
          ordinal: ordinal,
          text: line,
          siblingsScanned:
            siblingsScanned
        };
      }
    }
  }

  return {
    ordinal: 0,
    text: '',
    siblingsScanned:
      siblingsScanned
  };
}

/**
 * Lee el ordinal de un prefijo administrado
 * por el add-on.
 */
function testReadManualListOrdinal_(
  value,
  type
) {
  const text = String(value || '');

  if (type === 'NUMBER') {
    const match = text.match(
      /^\s*(\d+)\.\s+/
    );

    return match
      ? Number(match[1])
      : 0;
  }

  if (type === 'LETTER') {
    const match = text.match(
      /^\s*([A-Za-z]+)\)\s+/
    );

    return match
      ? testLettersToNumber_(
          match[1]
        )
      : 0;
  }

  if (type === 'ROMAN') {
    const match = text.match(
      /^\s*([ivxlcdm]+)\.\s+/i
    );

    return match
      ? testRomanToNumber_(
          match[1]
        )
      : 0;
  }

  return 0;
}

function testLettersToNumber_(letters) {
  const value = String(
    letters || ''
  ).toLowerCase();

  let result = 0;

  for (
    let index = 0;
    index < value.length;
    index++
  ) {
    const code =
      value.charCodeAt(index) - 96;

    if (code < 1 || code > 26) {
      return 0;
    }

    result =
      result * 26 +
      code;
  }

  return result;
}

function testRomanToNumber_(roman) {
  const value = String(
    roman || ''
  ).toUpperCase();

  const values = {
    I: 1,
    V: 5,
    X: 10,
    L: 50,
    C: 100,
    D: 500,
    M: 1000
  };

  let total = 0;
  let previous = 0;

  for (
    let index = value.length - 1;
    index >= 0;
    index--
  ) {
    const current =
      values[value[index]] || 0;

    if (!current) {
      return 0;
    }

    if (current < previous) {
      total -= current;
    } else {
      total += current;
      previous = current;
    }
  }

  return total;
}

/**
 * TEST-013
 *
 * Convierte párrafos y saltos manuales en incisos:
 *
 * a) Julio
 * b) Agosto
 * c) Septiembre
 *
 * Esta prueba inicia siempre desde a).
 */
function testApplyExactLetterList_() {
  const started = Date.now();

  const document =
    DocumentApp.getActiveDocument();

  const context =
    testGetParagraphsFromCurrentContext_(
      document
    );

  const targets = context.paragraphs;

  if (!targets.length) {
    throw new Error(
      'Selecciona uno o más párrafos o coloca el cursor en uno.'
    );
  }

  const updateStarted = Date.now();

  const result =
    testCreateExactManualList_(
      targets,
      'LETTER',
      1
    );

  const updateMs =
    Date.now() - updateStarted;

  document.saveAndClose();

  return {
    ok: true,
    testId:
      'TEST-013-EXACT-MANUAL-LETTER-LIST',

    contextType: context.contextType,
    sourceElements: targets.length,
    paragraphsCreated:
      result.paragraphs.length,

    firstOrdinal: 1,
    nextOrdinal: result.nextOrdinal,

    expectedPrefix: 'a)',
    expectedIndentInches: {
      left: 0.25,
      hanging: 0.25,
      right: 0
    },

    items: result.paragraphs.map(
      function (paragraph, index) {
        return {
          index: index,
          text: paragraph.getText(),
          heading: String(
            paragraph.getHeading()
          ),
          indentStart:
            paragraph.getIndentStart(),
          indentFirstLine:
            paragraph.getIndentFirstLine(),
          indentEnd:
            paragraph.getIndentEnd()
        };
      }
    ),

    updateMs: updateMs,
    apiReadMs: 0,
    apiWriteMs: 0,
    elapsedMs: Date.now() - started
  };
}

/**
 * Convierte los elementos recibidos en párrafos separados,
 * aplicando un prefijo visible controlado por el add-on.
 */
function testCreateExactManualList_(
  targets,
  type,
  initialOrdinal
) {
  let ordinal = initialOrdinal;
  const createdParagraphs = [];

  targets.forEach(function (target) {
    const parent = target.getParent();
    const targetIndex =
      parent.getChildIndex(target);

    /*
     * Google Docs representa los saltos manuales
     * internos mediante \r.
     */
    const logicalLines = String(
      target.getText() || ''
    )
      .split(/\r\n|\r|\n/)
      .map(function (line) {
        return testStripExistingListPrefix_(
          line
        ).trim();
      })
      .filter(function (line) {
        return line.length > 0;
      });

    /*
     * Insertar un párrafo independiente por cada línea.
     */
    logicalLines.forEach(function (
      content,
      lineIndex
    ) {
      const prefix =
        testCreateListPrefix_(
          type,
          ordinal
        );

      const paragraph =
        parent.insertParagraph(
          targetIndex + lineIndex,
          prefix + ' ' + content
        );

      /*
       * Aplicar el Normal Style actual del documento.
       */
      applyStyleToParagraph_(
        paragraph,
        'NORMAL'
      );

      testApplyExactListIndents_(
        paragraph
      );

      createdParagraphs.push(paragraph);
      ordinal++;
    });

    /*
     * Si había contenido, sustituir el elemento original.
     */
    if (logicalLines.length) {
      target.removeFromParent();
    } else {
      /*
       * Un párrafo vacío no se convierte en elemento de lista.
       */
      applyStyleToParagraph_(
        target,
        'NORMAL'
      );
    }
  });

  return {
    paragraphs: createdParagraphs,
    nextOrdinal: ordinal
  };
}

/**
 * Genera el prefijo requerido.
 */
function testCreateListPrefix_(
  type,
  ordinal
) {
  if (type === 'NUMBER') {
    return String(ordinal) + '.';
  }

  if (type === 'LETTER') {
    return (
      testNumberToLetters_(ordinal) +
      ')'
    );
  }

  if (type === 'ROMAN') {
    return (
      testNumberToRoman_(ordinal)
        .toLowerCase() +
      '.'
    );
  }

  throw new Error(
    'Unsupported manual list type: ' +
    type
  );
}

/**
 * Left: 0.25"
 * Hanging: 0.25"
 * Right: 0"
 */
function testApplyExactListIndents_(
  paragraph
) {
  const pointsPerInch = 72;
  const left = 0.25;
  const hanging = 0.25;

  paragraph.setIndentFirstLine(
    left * pointsPerInch
  );

  paragraph.setIndentStart(
    (left + hanging) *
      pointsPerInch
  );

  paragraph.setIndentEnd(0);
}

/**
 * Elimina un prefijo previo antes de crear el nuevo.
 */
function testStripExistingListPrefix_(
  value
) {
  let text = String(value || '');

  text = text.replace(
    /^\s*[•●○▪◦‣⁃-]\s+/,
    ''
  );

  text = text.replace(
    /^\s*\d+[.)]\s+/,
    ''
  );

  text = text.replace(
    /^\s*[A-Za-z]+[.)]\s+/,
    ''
  );

  text = text.replace(
    /^\s*[ivxlcdm]+[.)]\s+/i,
    ''
  );

  return text;
}

function testNumberToLetters_(number) {
  let value = Math.max(
    1,
    Number(number) || 1
  );

  let result = '';

  while (value > 0) {
    value--;

    result =
      String.fromCharCode(
        97 + (value % 26)
      ) +
      result;

    value = Math.floor(value / 26);
  }

  return result;
}

function testNumberToRoman_(number) {
  let value = Math.max(
    1,
    Number(number) || 1
  );

  const values = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I']
  ];

  let result = '';

  values.forEach(function (pair) {
    while (value >= pair[0]) {
      result += pair[1];
      value -= pair[0];
    }
  });

  return result;
}

/**
 * TEST-012
 *
 * Convierte los párrafos seleccionados en una lista nativa
 * de incisos LATIN_LOWER.
 *
 * Debe iniciar una lista nueva:
 * a), b), c) o el formato equivalente producido por Docs.
 */
function testCreateNewNativeLetterList_() {
  const started = Date.now();

  const document =
    DocumentApp.getActiveDocument();

  const context =
    testGetParagraphsFromCurrentContext_(
      document
    );

  const targets = context.paragraphs;

  if (!targets.length) {
    throw new Error(
      'Selecciona uno o más párrafos o coloca el cursor en uno.'
    );
  }

  const updateStarted = Date.now();

  /*
   * Crear una lista temporal aislada.
   * Su listId se asignará únicamente a los elementos seleccionados.
   */
  const temporary =
    testCreateIsolatedNativeListAnchor_(
      targets[0],
      DocumentApp.GlyphType.LATIN_LOWER
    );

  const items = [];

  try {
    targets.forEach(function (target) {
      const item =
        testConvertToNativeListItem_(
          target,
          temporary.anchor,
          DocumentApp.GlyphType.LATIN_LOWER
        );

      items.push(item);
    });
  } finally {
    /*
     * El listId permanece en los elementos seleccionados
     * aunque se elimine el elemento temporal.
     */
    try {
      temporary.anchor.removeFromParent();
    } catch (error) {}

    try {
      temporary.after.removeFromParent();
    } catch (error) {}

    try {
      temporary.before.removeFromParent();
    } catch (error) {}
  }

  const updateMs =
    Date.now() - updateStarted;

  const resultItems = items.map(
    function (item, index) {
      return {
        index: index,
        text: item.getText(),
        glyphType: String(
          item.getGlyphType()
        ),
        listId: item.getListId(),
        nestingLevel:
          item.getNestingLevel(),

        indentStart:
          item.getIndentStart(),

        indentFirstLine:
          item.getIndentFirstLine(),

        indentEnd:
          item.getIndentEnd()
      };
    }
  );

  const uniqueListIds = {};

  resultItems.forEach(function (item) {
    uniqueListIds[item.listId] = true;
  });

  document.saveAndClose();

  return {
    ok: true,
    testId:
      'TEST-012-NEW-NATIVE-LETTER-LIST',

    contextType: context.contextType,
    paragraphsApplied: items.length,

    requestedType: 'LETTER',
    requestedContinue: false,

    sameListId:
      Object.keys(uniqueListIds).length === 1,

    expectedIndentInches: {
      left: 0.25,
      hanging: 0.25,
      right: 0
    },

    items: resultItems,

    updateMs: updateMs,
    apiReadMs: 0,
    apiWriteMs: 0,
    elapsedMs: Date.now() - started
  };
}

/**
 * Crea un ListItem temporal con un listId nuevo y aislado.
 *
 * Los párrafos separadores impiden que Google Docs reutilice
 * accidentalmente la lista anterior o posterior.
 */
function testCreateIsolatedNativeListAnchor_(
  target,
  glyphType
) {
  const parent = target.getParent();
  const targetIndex =
    parent.getChildIndex(target);

  const before = parent.insertParagraph(
    targetIndex,
    '\uE210'
  );

  const anchor = parent.insertListItem(
    targetIndex + 1,
    '\uE211'
  );

  const after = parent.insertParagraph(
    targetIndex + 2,
    '\uE212'
  );

  anchor.setNestingLevel(0);
  anchor.setGlyphType(glyphType);

  return {
    before: before,
    anchor: anchor,
    after: after
  };
}

/**
 * Convierte o reutiliza el elemento seleccionado como ListItem.
 */
function testConvertToNativeListItem_(
  target,
  anchor,
  glyphType
) {
  let item;

  if (
    target.getType() ===
    DocumentApp.ElementType.LIST_ITEM
  ) {
    item = target.asListItem();
  } else {
    const parent = target.getParent();
    const index =
      parent.getChildIndex(target);

    const content =
      testStripVisibleListPrefix_(
        target.getText()
      ).trim();

    item = parent.insertListItem(
      index,
      content
    );

    target.removeFromParent();
  }

  /*
   * Todos los elementos reciben el mismo listId,
   * por lo que pertenecen a la misma lista.
   */
  item.setListId(anchor);
  item.setNestingLevel(0);
  item.setGlyphType(glyphType);

  /*
   * Aplicar Normal text usando la solución ya validada.
   */
  applyStyleToParagraph_(
    item,
    'NORMAL'
  );

  testApplyRequestedListIndents_(item);

  return item;
}

/**
 * Left: 0.25"
 * Hanging: 0.25"
 * Right: 0"
 *
 * En la geometría de Google Docs:
 * - indentFirstLine = Left
 * - indentStart = Left + Hanging
 */
function testApplyRequestedListIndents_(
  item
) {
  const pointsPerInch = 72;
  const left = 0.25;
  const hanging = 0.25;

  item.setIndentFirstLine(
    left * pointsPerInch
  );

  item.setIndentStart(
    (left + hanging) * pointsPerInch
  );

  item.setIndentEnd(0);
}

/**
 * Elimina prefijos visibles previos para evitar:
 * "1. a) Texto" o "• a) Texto".
 *
 * No afecta la numeración automática porque esa numeración
 * no forma parte de getText().
 */
function testStripVisibleListPrefix_(value) {
  let text = String(value || '');

  text = text.replace(
    /^\s*[•●○▪◦‣⁃-]\s+/,
    ''
  );

  text = text.replace(
    /^\s*\d+[.)]\s+/,
    ''
  );

  text = text.replace(
    /^\s*[A-Za-z]+[.)]\s+/,
    ''
  );

  text = text.replace(
    /^\s*[ivxlcdm]+[.)]\s+/i,
    ''
  );

  return text;
}

/**
 * TEST-011
 *
 * Aplica un estilo nombrado a todos los párrafos tocados
 * por la selección.
 *
 * Si no existe selección, utiliza el párrafo del cursor.
 */
function testApplyStyleToCurrentContext_(styleName) {
  const started = Date.now();
  const document = DocumentApp.getActiveDocument();

  const targetStyle =
    testResolveParagraphHeading_(styleName);

  const context =
    testGetParagraphsFromCurrentContext_(document);

  if (!context.paragraphs.length) {
    throw new Error(
      'No se encontraron párrafos en la selección o el cursor.'
    );
  }

  const updateStarted = Date.now();
  const results = [];

  context.paragraphs.forEach(function (
    paragraph,
    index
  ) {
    const result =
      testClearOverridesAndApplyStyle_(
        paragraph,
        targetStyle
      );

    result.index = index;
    results.push(result);
  });

  const updateMs = Date.now() - updateStarted;

  document.saveAndClose();

  return {
    ok: true,
    testId: 'TEST-011-MULTI-PARAGRAPH-STYLE',
    requestedStyle: styleName,
    contextType: context.contextType,
    paragraphsApplied: context.paragraphs.length,
    results: results,
    updateMs: updateMs,
    apiReadMs: 0,
    apiWriteMs: 0,
    elapsedMs: Date.now() - started
  };
}

/**
 * Elimina los overrides sin cambiar el contenido
 * y aplica el estilo nombrado solicitado.
 */
function testClearOverridesAndApplyStyle_(
  paragraph,
  targetStyle
) {
  const before = {
    elementType: String(paragraph.getType()),
    heading: String(paragraph.getHeading()),
    indentStart: paragraph.getIndentStart(),
    indentEnd: paragraph.getIndentEnd(),
    indentFirstLine:
      paragraph.getIndentFirstLine(),
    alignment: String(paragraph.getAlignment())
  };

  const resetAttributes =
    testBuildOverrideResetAttributes_();

  /*
   * No se modifica el texto del párrafo.
   * Los tabuladores y espacios reales se conservan.
   */
  paragraph.setAttributes(resetAttributes);
  paragraph.setHeading(targetStyle);

  const after = {
    elementType: String(paragraph.getType()),
    heading: String(paragraph.getHeading()),
    indentStart: paragraph.getIndentStart(),
    indentEnd: paragraph.getIndentEnd(),
    indentFirstLine:
      paragraph.getIndentFirstLine(),
    alignment: String(paragraph.getAlignment())
  };

  return {
    before: before,
    after: after,
    textLength: paragraph.getText().length
  };
}

/**
 * Atributos que se borran para recuperar la herencia
 * del estilo nombrado.
 *
 * LINK_URL no se incluye para conservar hipervínculos.
 * LIST_ID y GLYPH_TYPE tampoco se modifican.
 */
function testBuildOverrideResetAttributes_() {
  const resetAttributes = {};

  const attributesToReset = [
    DocumentApp.Attribute.FONT_FAMILY,
    DocumentApp.Attribute.FONT_SIZE,
    DocumentApp.Attribute.FOREGROUND_COLOR,
    DocumentApp.Attribute.BACKGROUND_COLOR,
    DocumentApp.Attribute.BOLD,
    DocumentApp.Attribute.ITALIC,
    DocumentApp.Attribute.UNDERLINE,
    DocumentApp.Attribute.STRIKETHROUGH,

    DocumentApp.Attribute.HORIZONTAL_ALIGNMENT,
    DocumentApp.Attribute.INDENT_START,
    DocumentApp.Attribute.INDENT_END,
    DocumentApp.Attribute.INDENT_FIRST_LINE,
    DocumentApp.Attribute.LINE_SPACING,
    DocumentApp.Attribute.SPACING_BEFORE,
    DocumentApp.Attribute.SPACING_AFTER
  ];

  attributesToReset.forEach(function (attribute) {
    resetAttributes[attribute] = null;
  });

  return resetAttributes;
}

/**
 * Obtiene una lista única de párrafos desde la selección.
 * Si no hay selección, obtiene el párrafo del cursor.
 */
function testGetParagraphsFromCurrentContext_(
  document
) {
  const selection = document.getSelection();

  if (selection) {
    const paragraphs = [];
    const existingKeys = {};

    selection
      .getRangeElements()
      .forEach(function (rangeElement) {
        const paragraph = testFindParagraph_(
          rangeElement.getElement()
        );

        if (!paragraph) {
          return;
        }

        const paragraphKey =
          testGetParagraphStructuralKey_(paragraph);

        if (existingKeys[paragraphKey]) {
          return;
        }

        existingKeys[paragraphKey] = true;
        paragraphs.push(paragraph);
      });

    if (paragraphs.length) {
      return {
        contextType: 'SELECTION',
        paragraphs: paragraphs
      };
    }
  }

  const cursor = document.getCursor();

  if (!cursor) {
    return {
      contextType: 'NONE',
      paragraphs: []
    };
  }

  const cursorParagraph = testFindParagraph_(
    cursor.getElement()
  );

  return {
    contextType: 'CURSOR',
    paragraphs: cursorParagraph
      ? [cursorParagraph]
      : []
  };
}

/**
 * Crea una clave estructural para evitar aplicar el
 * formato varias veces al mismo párrafo.
 */
function testGetParagraphStructuralKey_(paragraph) {
  const path = [];
  let current = paragraph;

  while (current) {
    const parent = current.getParent();

    if (
      !parent ||
      typeof parent.getChildIndex !== 'function'
    ) {
      break;
    }

    path.unshift(parent.getChildIndex(current));
    current = parent;
  }

  return String(paragraph.getType()) +
    ':' +
    path.join('.');
}

/**
 * Convierte el nombre del estilo al enum de DocumentApp.
 */
function testResolveParagraphHeading_(styleName) {
  const styles = {
    NORMAL:
      DocumentApp.ParagraphHeading.NORMAL,

    TITLE:
      DocumentApp.ParagraphHeading.TITLE,

    SUBTITLE:
      DocumentApp.ParagraphHeading.SUBTITLE,

    HEADING1:
      DocumentApp.ParagraphHeading.HEADING1,

    HEADING2:
      DocumentApp.ParagraphHeading.HEADING2,

    HEADING3:
      DocumentApp.ParagraphHeading.HEADING3,

    HEADING4:
      DocumentApp.ParagraphHeading.HEADING4,

    HEADING5:
      DocumentApp.ParagraphHeading.HEADING5,

    HEADING6:
      DocumentApp.ParagraphHeading.HEADING6
  };

  const normalizedName = String(styleName)
    .toUpperCase()
    .replace(/[\s_-]+/g, '');

  if (!styles[normalizedName]) {
    throw new Error(
      'Estilo no reconocido: ' + styleName
    );
  }

  return styles[normalizedName];
}

/**
 * TEST-009
 *
 * Confirma si basta con:
 *
 * borrar overrides con null → aplicar H4
 *
 * No utiliza una transición temporal por NORMAL.
 */
function testConfirmMinimalH4Reset_() {
  const started = Date.now();

  const document = DocumentApp.getActiveDocument();
  const cursor = document.getCursor();

  if (!cursor) {
    throw new Error(
      'Coloca el cursor dentro de un H4 con formato incorrecto.'
    );
  }

  const paragraph = testFindParagraph_(
    cursor.getElement()
  );

  if (!paragraph) {
    throw new Error('No se pudo localizar el párrafo.');
  }

  const before = {
    elementType: String(paragraph.getType()),
    heading: String(paragraph.getHeading()),
    indentStart: paragraph.getIndentStart(),
    indentEnd: paragraph.getIndentEnd(),
    indentFirstLine: paragraph.getIndentFirstLine(),
    alignment: String(paragraph.getAlignment())
  };

  /*
   * No incluimos LINK_URL para conservar hipervínculos.
   * No incluimos HEADING porque se aplicará después.
   */
  const resetAttributes = {};

  resetAttributes[
    DocumentApp.Attribute.FONT_FAMILY
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.FONT_SIZE
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.FOREGROUND_COLOR
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.BACKGROUND_COLOR
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.BOLD
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.ITALIC
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.UNDERLINE
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.STRIKETHROUGH
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.HORIZONTAL_ALIGNMENT
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.INDENT_START
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.INDENT_END
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.INDENT_FIRST_LINE
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.LINE_SPACING
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.SPACING_BEFORE
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.SPACING_AFTER
  ] = null;

  const updateStarted = Date.now();

  /*
   * Solución mínima:
   * limpiar overrides y aplicar el estilo.
   */
  paragraph.setAttributes(resetAttributes);

  paragraph.setHeading(
    DocumentApp.ParagraphHeading.HEADING4
  );

  const updateMs = Date.now() - updateStarted;

  const after = {
    elementType: String(paragraph.getType()),
    heading: String(paragraph.getHeading()),
    indentStart: paragraph.getIndentStart(),
    indentEnd: paragraph.getIndentEnd(),
    indentFirstLine: paragraph.getIndentFirstLine(),
    alignment: String(paragraph.getAlignment())
  };

  document.saveAndClose();

  return {
    ok: true,
    testId: 'TEST-009-MINIMAL-H4-RESET',
    before: before,
    after: after,
    updateMs: updateMs,
    apiReadMs: 0,
    apiWriteMs: 0,
    elapsedMs: Date.now() - started
  };
}

/**
 * TEST-008
 *
 * Fuerza la transición:
 *
 * NORMAL → limpiar overrides → HEADING4
 *
 * Esta prueba SÍ debe modificar visualmente el párrafo.
 */
function testForceHeading4Transition_() {
  const started = Date.now();

  const document = DocumentApp.getActiveDocument();
  const cursor = document.getCursor();

  if (!cursor) {
    throw new Error(
      'Coloca el cursor dentro de un párrafo, sin seleccionar texto.'
    );
  }

  const paragraph = testFindParagraph_(cursor.getElement());

  if (!paragraph) {
    throw new Error('No se pudo localizar el párrafo actual.');
  }

  const text = paragraph.editAsText();
  const textLength = text.getText().length;

  const before = {
    elementType: String(paragraph.getType()),
    heading: String(paragraph.getHeading()),
    indentStart: paragraph.getIndentStart(),
    indentEnd: paragraph.getIndentEnd(),
    indentFirstLine: paragraph.getIndentFirstLine(),
    alignment: String(paragraph.getAlignment())
  };

  /*
   * Atributos directos del texto que deben volver a heredarse.
   * LINK_URL se omite para conservar hipervínculos.
   */
  const textReset = {};

  textReset[DocumentApp.Attribute.FONT_FAMILY] = null;
  textReset[DocumentApp.Attribute.FONT_SIZE] = null;
  textReset[DocumentApp.Attribute.FOREGROUND_COLOR] = null;
  textReset[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
  textReset[DocumentApp.Attribute.BOLD] = null;
  textReset[DocumentApp.Attribute.ITALIC] = null;
  textReset[DocumentApp.Attribute.UNDERLINE] = null;
  textReset[DocumentApp.Attribute.STRIKETHROUGH] = null;

  /*
   * Overrides propios del párrafo.
   */
  const paragraphReset = {};

  paragraphReset[
    DocumentApp.Attribute.HORIZONTAL_ALIGNMENT
  ] = null;

  paragraphReset[
    DocumentApp.Attribute.INDENT_START
  ] = null;

  paragraphReset[
    DocumentApp.Attribute.INDENT_END
  ] = null;

  paragraphReset[
    DocumentApp.Attribute.INDENT_FIRST_LINE
  ] = null;

  paragraphReset[
    DocumentApp.Attribute.LINE_SPACING
  ] = null;

  paragraphReset[
    DocumentApp.Attribute.SPACING_BEFORE
  ] = null;

  paragraphReset[
    DocumentApp.Attribute.SPACING_AFTER
  ] = null;

  const updateStarted = Date.now();

  /*
   * Forzar un cambio real de estilo.
   */
  paragraph.setHeading(
    DocumentApp.ParagraphHeading.NORMAL
  );

  /*
   * Borrar overrides mientras el párrafo está en Normal.
   */
  paragraph.setAttributes(paragraphReset);

  if (textLength > 0) {
    text.setAttributes(
      0,
      textLength - 1,
      textReset
    );
  }

  /*
   * Aplicar finalmente el estilo configurado en el documento.
   */
  paragraph.setHeading(
    DocumentApp.ParagraphHeading.HEADING4
  );

  const updateMs = Date.now() - updateStarted;

  const after = {
    elementType: String(paragraph.getType()),
    heading: String(paragraph.getHeading()),
    indentStart: paragraph.getIndentStart(),
    indentEnd: paragraph.getIndentEnd(),
    indentFirstLine: paragraph.getIndentFirstLine(),
    alignment: String(paragraph.getAlignment())
  };

  document.saveAndClose();

  return {
    ok: true,
    testId: 'TEST-008-FORCE-H4-TRANSITION',
    before: before,
    after: after,
    textLength: textLength,
    updateMs: updateMs,
    apiReadMs: 0,
    apiWriteMs: 0,
    elapsedMs: Date.now() - started
  };
}

/**
 * TEST-007
 *
 * Intenta eliminar los overrides directamente con DocumentApp,
 * sin utilizar la API avanzada ni recorrer el documento.
 *
 * Esta prueba SÍ debe cambiar visualmente el párrafo.
 */
function testResetH4WithNullAttributes_() {
  const started = Date.now();

  const document = DocumentApp.getActiveDocument();
  const cursor = document.getCursor();

  if (!cursor) {
    throw new Error(
      'Coloca el cursor dentro de un párrafo, sin seleccionar texto.'
    );
  }

  const paragraph = testFindParagraph_(cursor.getElement());

  if (!paragraph) {
    throw new Error('No se pudo localizar el párrafo actual.');
  }

  const beforeHeading = String(paragraph.getHeading());
  const paragraphText = paragraph.getText();

  /*
   * No incluimos LINK_URL para conservar los hipervínculos.
   * Tampoco incluimos HEADING porque se aplicará H4 después.
   */
  const resetAttributes = {};

  resetAttributes[
    DocumentApp.Attribute.FONT_FAMILY
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.FONT_SIZE
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.FOREGROUND_COLOR
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.BACKGROUND_COLOR
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.BOLD
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.ITALIC
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.UNDERLINE
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.STRIKETHROUGH
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.HORIZONTAL_ALIGNMENT
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.INDENT_START
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.INDENT_END
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.INDENT_FIRST_LINE
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.LINE_SPACING
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.SPACING_BEFORE
  ] = null;

  resetAttributes[
    DocumentApp.Attribute.SPACING_AFTER
  ] = null;

  const updateStarted = Date.now();

  /*
   * Intentar eliminar los atributos directos.
   */
  paragraph.setAttributes(resetAttributes);

  /*
   * Aplicar posteriormente el estilo nombrado H4.
   */
  paragraph.setHeading(
    DocumentApp.ParagraphHeading.HEADING4
  );

  const updateMs = Date.now() - updateStarted;

  document.saveAndClose();

  return {
    ok: true,
    testId: 'TEST-007-NULL-ATTRIBUTE-RESET',
    beforeHeading: beforeHeading,
    afterHeading: 'HEADING4',
    textLength: paragraphText.length,
    updateMs: updateMs,
    apiReadMs: 0,
    apiWriteMs: 0,
    elapsedMs: Date.now() - started,
    message:
      'Null attributes were applied and the paragraph was set to H4.'
  };
}

/**
 * TEST-006
 *
 * Localiza los índices API del párrafo actual mediante un
 * NamedRange temporal, sin recorrer todo el documento.
 *
 * Esta prueba no cambia visualmente el formato.
 */
function testLocateWithTemporaryNamedRange_() {
  const started = Date.now();

  const document = DocumentApp.getActiveDocument();
  const cursor = document.getCursor();

  if (!cursor) {
    throw new Error(
      'Coloca el cursor dentro de un párrafo, sin seleccionar texto.'
    );
  }

  const paragraph = testFindParagraph_(cursor.getElement());

  if (!paragraph) {
    throw new Error('No se pudo localizar el párrafo actual.');
  }

  const documentId = document.getId();

  const markerName =
    'DOCS_ASSISTANT_TEST_' +
    Utilities.getUuid().replace(/-/g, '');

  /*
   * Crear el NamedRange temporal alrededor
   * del párrafo donde está el cursor.
   */
  const setupStarted = Date.now();

  const range = document
    .newRange()
    .addElement(paragraph)
    .build();

  const namedRange = document.addNamedRange(
    markerName,
    range
  );

  const documentAppNamedRangeId = namedRange.getId();
  const markerSetupMs = Date.now() - setupStarted;

  /*
   * Guardar los cambios para que el NamedRange
   * esté disponible desde la API de Docs.
   */
  const saveStarted = Date.now();

  document.saveAndClose();

  const saveMs = Date.now() - saveStarted;

  let apiReadMs = 0;
  let cleanupMs = 0;

  try {
    /*
     * Obtener solamente los NamedRanges.
     */
    const apiStarted = Date.now();

    const apiDocument = Docs.Documents.get(
      documentId,
      {
        fields: 'namedRanges'
      }
    );

    apiReadMs = Date.now() - apiStarted;

    /*
     * Encontrar el grupo mediante el nombre único.
     */
    const group =
      apiDocument.namedRanges &&
      apiDocument.namedRanges[markerName];

    const candidates =
      group && group.namedRanges
        ? group.namedRanges
        : [];

    const match = candidates.length
      ? candidates[0]
      : null;

    const apiRange =
      match &&
      match.ranges &&
      match.ranges.length
        ? match.ranges[0]
        : null;

    if (!apiRange) {
      throw new Error(
        'La API no devolvió el NamedRange temporal.'
      );
    }

    /*
     * Eliminar el NamedRange utilizando su nombre.
     *
     * La limpieza no debe impedir que la prueba
     * devuelva los índices encontrados.
     */
    const cleanupStarted = Date.now();
    let cleanupError = null;

    try {
      Docs.Documents.batchUpdate(
        {
          requests: [
            {
              deleteNamedRange: {
                name: markerName
              }
            }
          ]
        },
        documentId
      );
    } catch (error) {
      cleanupError =
        error && error.message
          ? error.message
          : String(error);
    }

    cleanupMs = Date.now() - cleanupStarted;

    return {
      ok: true,
      testId: 'TEST-006-TEMPORARY-NAMED-RANGE',

      startIndex: apiRange.startIndex,
      endIndex: apiRange.endIndex,
      targetLength: paragraph.getText().length,

      documentAppNamedRangeId:
        documentAppNamedRangeId,

      apiNamedRangeId:
        match.namedRangeId || null,

      markerSetupMs: markerSetupMs,
      saveMs: saveMs,
      apiReadMs: apiReadMs,
      cleanupMs: cleanupMs,

      cleanupOk: !cleanupError,
      cleanupError: cleanupError,

      elapsedMs: Date.now() - started
    };

  } catch (error) {
    /*
     * Si falla la lectura, intentar eliminar
     * igualmente el NamedRange temporal.
     */
    try {
      Docs.Documents.batchUpdate(
        {
          requests: [
            {
              deleteNamedRange: {
                name: markerName
              }
            }
          ]
        },
        documentId
      );
    } catch (cleanupError) {
      // No ocultar el error original.
    }

    throw error;
  }
}

/**
 * TEST 004
 * Aplica H4 y elimina overrides usando un rango API ya conocido.
 * No ejecuta Docs.Documents.get().
 */
function testResetH4AtKnownRange_() {
  const started = Date.now();
  const documentId = DocumentApp
    .getActiveDocument()
    .getId();

  // Rango obtenido en TEST-003.
  const paragraphRange = {
    startIndex: 29973,
    endIndex: 30003
  };

  /*
   * Excluimos el salto de línea final para que UpdateTextStyle
   * actúe únicamente sobre el contenido del párrafo.
   */
  const textRange = {
    startIndex: 29973,
    endIndex: 30002
  };

  const paragraphFields = [
    'namedStyleType',
    'alignment',
    'lineSpacing',
    'spacingMode',
    'spaceAbove',
    'spaceBelow',
    'borderBetween',
    'borderTop',
    'borderBottom',
    'borderLeft',
    'borderRight',
    'indentFirstLine',
    'indentStart',
    'indentEnd',
    'tabStops',
    'keepLinesTogether',
    'keepWithNext',
    'avoidWidowAndOrphan',
    'shading',
    'pageBreakBefore'
  ].join(',');

  const textFields = [
    'bold',
    'italic',
    'underline',
    'strikethrough',
    'smallCaps',
    'backgroundColor',
    'foregroundColor',
    'fontSize',
    'weightedFontFamily',
    'baselineOffset'
  ].join(',');

  const apiStarted = Date.now();

  Docs.Documents.batchUpdate(
    {
      requests: [
        {
          updateParagraphStyle: {
            range: paragraphRange,
            paragraphStyle: {
              namedStyleType: 'HEADING_4'
            },
            fields: paragraphFields
          }
        },
        {
          updateTextStyle: {
            range: textRange,
            textStyle: {},
            fields: textFields
          }
        }
      ]
    },
    documentId
  );

  const apiWriteMs = Date.now() - apiStarted;

  return {
    ok: true,
    testId: 'TEST-004-RESET-H4-KNOWN-RANGE',
    startIndex: paragraphRange.startIndex,
    endIndex: paragraphRange.endIndex,
    apiReadMs: 0,
    apiWriteMs: apiWriteMs,
    elapsedMs: Date.now() - started,
    message: 'H4 and override reset requests were completed.'
  };
}

/**
 * TEST 002
 * Obtiene la configuración actual de Heading 4 del documento y la aplica
 * explícitamente al párrafo.
 *
 * Es una prueba diagnóstica: sí crea formato directo.
 */
function testApplyHeading4Attributes_() {
  const started = Date.now();
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();

  if (!cursor) {
    throw new Error(
      'No cursor detected. Click inside one paragraph without selecting text.'
    );
  }

  const paragraph = testFindParagraph_(cursor.getElement());

  if (!paragraph) {
    throw new Error(
      'The cursor is not inside a Paragraph or ListItem.'
    );
  }

  let body = doc.getBody();

  try {
    const tab = doc.getActiveTab();

    if (tab && typeof tab.asDocumentTab === 'function') {
      body = tab.asDocumentTab().getBody();
    }
  } catch (error) {}

  const target = DocumentApp.ParagraphHeading.HEADING4;
  const styleAttributes = body.getHeadingAttributes(target);

  paragraph.setHeading(target);
  paragraph.setAttributes(styleAttributes);

  return {
    ok: paragraph.getHeading() === target,
    testId: 'TEST-002-H4-EFFECTIVE-ATTRIBUTES',
    operation: 'setHeading(H4) + setAttributes(document H4)',
    after: String(paragraph.getHeading()),
    attributeCount: Object.keys(styleAttributes).length,
    fontFamily:
      styleAttributes[DocumentApp.Attribute.FONT_FAMILY] || null,
    fontSize:
      styleAttributes[DocumentApp.Attribute.FONT_SIZE] || null,
    bold:
      styleAttributes[DocumentApp.Attribute.BOLD],
    foregroundColor:
      styleAttributes[DocumentApp.Attribute.FOREGROUND_COLOR] || null,
    elapsedMs: Date.now() - started,
    message:
      'The document H4 attributes were applied explicitly. Check the visual result.'
  };
}

/**
 * TEST 001
 * Ejecuta la operación H4 más pequeña posible.
 *
 * No utiliza:
 * - API avanzada de Google Docs.
 * - Segmentación.
 * - Limpieza de overrides.
 * - Funciones de Formatting.gs.
 */
function testNativeHeading4_() {
  const started = Date.now();
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();

  if (!cursor) {
    throw new Error(
      'No cursor detected. Click inside one paragraph without selecting text.'
    );
  }

  const paragraph = testFindParagraph_(cursor.getElement());

  if (!paragraph) {
    throw new Error(
      'The cursor is not inside a Paragraph or ListItem.'
    );
  }

  const before = paragraph.getHeading();
  const target = DocumentApp.ParagraphHeading.HEADING4;

  paragraph.setHeading(target);

  const after = paragraph.getHeading();
  const accepted = after === target;

  return {
    ok: accepted,
    testId: 'TEST-001-NATIVE-H4',
    operation: 'paragraph.setHeading(HEADING4)',
    targetType: String(paragraph.getType()),
    textPreview: String(paragraph.getText() || '').slice(0, 120),
    before: String(before),
    after: String(after),
    elapsedMs: Date.now() - started,
    message: accepted
      ? 'Google Docs accepted the native H4 style.'
      : 'Google Docs returned a different style after setHeading().'
  };
}

function testFindParagraph_(element) {
  let current = element;

  while (current) {
    const type = current.getType();

    if (
      type === DocumentApp.ElementType.PARAGRAPH ||
      type === DocumentApp.ElementType.LIST_ITEM
    ) {
      return current;
    }

    current = current.getParent
      ? current.getParent()
      : null;
  }

  return null;
}

/**
 * TEST 003
 * Localiza el rango API exacto del párrafo del cursor.
 * No cambia ningún formato.
 */
function testLocateParagraphWithApi_() {
  const started = Date.now();
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();

  if (!cursor) {
    throw new Error(
      'No cursor detected. Click inside one paragraph without selecting text.'
    );
  }

  const paragraph = testFindParagraph_(cursor.getElement());

  if (!paragraph) {
    throw new Error(
      'The cursor is not inside a Paragraph or ListItem.'
    );
  }

  let body = doc.getBody();
  let activeTabId = '';

  try {
    const activeTab = doc.getActiveTab();

    if (activeTab) {
      activeTabId = activeTab.getId();
      body = activeTab.asDocumentTab().getBody();
    }
  } catch (error) {}

  /*
   * Esta prueba usa la respuesta API simplificada del primer tab.
   * Evita descargar todo el documento con todos sus atributos y objetos.
   */
  try {
    const tabs = doc.getTabs();

    if (
      tabs.length &&
      activeTabId &&
      tabs[0].getId() !== activeTabId
    ) {
      throw new Error(
        'TEST-003 currently requires the first document tab to be active.'
      );
    }
  } catch (error) {
    if (String(error.message || error).indexOf('TEST-003') >= 0) {
      throw error;
    }
  }

  const childIndex = body.getChildIndex(paragraph);
  let targetOrdinal = -1;
  let paragraphOrdinal = 0;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const element = body.getChild(i);
    const type = element.getType();

    const styleable =
      type === DocumentApp.ElementType.PARAGRAPH ||
      type === DocumentApp.ElementType.LIST_ITEM;

    if (!styleable) continue;

    if (i === childIndex) {
      targetOrdinal = paragraphOrdinal;
    }

    paragraphOrdinal++;
  }

  if (targetOrdinal < 0) {
    throw new Error(
      'The cursor paragraph could not be mapped inside the document body.'
    );
  }

  const apiStarted = Date.now();

  const apiDocument = Docs.Documents.get(doc.getId(), {
    fields: 'revisionId,body(content(startIndex,endIndex,paragraph))'
  });

  const apiReadMs = Date.now() - apiStarted;

  const apiParagraphs = (apiDocument.body.content || [])
    .filter(element => element.paragraph);

  const apiParagraph = apiParagraphs[targetOrdinal];

  if (!apiParagraph) {
    throw new Error(
      'The corresponding API paragraph was not found.'
    );
  }

  return {
    ok: true,
    testId: 'TEST-003-LOCATE-API-RANGE',
    childIndex: childIndex,
    paragraphOrdinal: targetOrdinal,
    documentParagraphs: paragraphOrdinal,
    apiParagraphs: apiParagraphs.length,
    startIndex: apiParagraph.startIndex,
    endIndex: apiParagraph.endIndex,
    apiReadMs: apiReadMs,
    elapsedMs: Date.now() - started,
    message: 'The paragraph API range was located successfully.'
  };
}
/**
 * TEST 005
 * Calcula localmente los índices de la API.
 * No lee la Docs API y no modifica el documento.
 */
function testCalculateApiRangeLocally_() {
  const started = Date.now();
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();

  if (!cursor) {
    throw new Error(
      'No cursor detected. Click inside one paragraph without selecting text.'
    );
  }

  const paragraph = testFindParagraph_(cursor.getElement());

  if (!paragraph) {
    throw new Error(
      'The cursor is not inside a Paragraph or ListItem.'
    );
  }

  let body = doc.getBody();

  try {
    const tab = doc.getActiveTab();

    if (tab && typeof tab.asDocumentTab === 'function') {
      body = tab.asDocumentTab().getBody();
    }
  } catch (error) {}

  const childIndex = body.getChildIndex(paragraph);

  /*
   * La API reserva el índice 0 para el section break inicial.
   * El primer elemento visible comienza en el índice 1.
   */
  let predictedStart = 1;

  const stats = {
    paragraphs: 0,
    tables: 0,
    cells: 0,
    specialElements: 0
  };

  for (let i = 0; i < childIndex; i++) {
    predictedStart += testApiStructuralLength_(
      body.getChild(i),
      stats
    );
  }

  const targetLength = testApiStructuralLength_(
    paragraph,
    stats
  );

  const predictedEnd = predictedStart + targetLength;

  // Valores reales obtenidos en TEST-003.
  const knownStart = 29973;
  const knownEnd = 30003;

  const startMatches = predictedStart === knownStart;
  const endMatches = predictedEnd === knownEnd;

  return {
    ok: startMatches && endMatches,
    testId: 'TEST-005-LOCAL-RANGE-CALCULATION',
    childIndex: childIndex,
    predictedStart: predictedStart,
    predictedEnd: predictedEnd,
    knownStart: knownStart,
    knownEnd: knownEnd,
    startDifference: predictedStart - knownStart,
    endDifference: predictedEnd - knownEnd,
    targetLength: targetLength,
    paragraphsScanned: stats.paragraphs,
    tablesScanned: stats.tables,
    tableCellsScanned: stats.cells,
    specialElements: stats.specialElements,
    apiReadMs: 0,
    elapsedMs: Date.now() - started,
    message: startMatches && endMatches
      ? 'The local API range matches exactly.'
      : 'The local range requires an index adjustment.'
  };
}

function testApiStructuralLength_(element, stats) {
  const type = String(element.getType());

  if (type === 'PARAGRAPH' || type === 'LIST_ITEM') {
    stats.paragraphs++;
    return testApiParagraphLength_(element, stats);
  }

  if (type === 'TABLE') {
    stats.tables++;
    return testApiTableLength_(element, stats);
  }

  if (type === 'TABLE_OF_CONTENTS') {
    let length = 1;

    for (let i = 0; i < element.getNumChildren(); i++) {
      length += testApiStructuralLength_(
        element.getChild(i),
        stats
      );
    }

    return length;
  }

  /*
   * Otros elementos estructurales ocupan normalmente una unidad
   * dentro del modelo de índices.
   */
  stats.specialElements++;
  return 1;
}

function testApiParagraphLength_(paragraph, stats) {
  // Un carácter adicional corresponde al salto de línea del párrafo.
  let length = 1;

  const oneUnitTypes = [
    'INLINE_IMAGE',
    'PAGE_BREAK',
    'COLUMN_BREAK',
    'HORIZONTAL_RULE',
    'FOOTNOTE',
    'EQUATION',
    'PERSON',
    'RICH_LINK',
    'DATE',
    'DATE_ELEMENT',
    'AUTO_TEXT'
  ];

  for (let i = 0; i < paragraph.getNumChildren(); i++) {
    const child = paragraph.getChild(i);
    const type = String(child.getType());

    if (type === 'TEXT') {
      length += String(child.getText() || '').length;
      continue;
    }

    if (oneUnitTypes.indexOf(type) >= 0) {
      length++;
      stats.specialElements++;
      continue;
    }

    if (typeof child.getText === 'function') {
      length += String(child.getText() || '').length;
    } else {
      length++;
      stats.specialElements++;
    }
  }

  return length;
}

function testApiTableLength_(table, stats) {
  /*
   * Modelo de índices:
   * 1 unidad para la tabla;
   * 1 unidad por fila;
   * 1 unidad por celda;
   * más el contenido estructural de cada celda.
   */
  let length = 1;

  for (let rowIndex = 0; rowIndex < table.getNumRows(); rowIndex++) {
    const row = table.getRow(rowIndex);

    length++;

    for (let cellIndex = 0; cellIndex < row.getNumCells(); cellIndex++) {
      const cell = row.getCell(cellIndex);

      stats.cells++;
      length++;

      for (
        let childIndex = 0;
        childIndex < cell.getNumChildren();
        childIndex++
      ) {
        length += testApiStructuralLength_(
          cell.getChild(childIndex),
          stats
        );
      }
    }
  }

  return length;
}


