/**
 * TEST LAB
 * El botón siempre llama esta función.
 * Para cada experimento cambiaremos únicamente la prueba ejecutada aquí.
 */
function runCurrentTest() {
  return testApplyHeading4Attributes_();
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
