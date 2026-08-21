/**
 * TEST LAB
 * El botón siempre llama esta función.
 * Para cada experimento cambiaremos únicamente la prueba ejecutada aquí.
 */
function runCurrentTest() {
  return testNativeHeading4_();
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
