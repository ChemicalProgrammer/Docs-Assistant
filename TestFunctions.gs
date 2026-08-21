/**
 * TEST LAB
 * El botón siempre llama esta función.
 * Para cada experimento cambiaremos únicamente la prueba ejecutada aquí.
 */
function runCurrentTest() {
  return testResetH4AtKnownRange_();
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



