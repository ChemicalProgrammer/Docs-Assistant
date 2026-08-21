/**
 * TEST LAB
 * El botón siempre llama esta función.
 * Para cada experimento cambiaremos únicamente la prueba ejecutada aquí.
 */
function runCurrentTest() {
  return testForceHeading4Transition_();
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


