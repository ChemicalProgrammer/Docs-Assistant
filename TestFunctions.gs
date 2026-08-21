/**
 * TEST LAB
 *
 * Prueba actual:
 * aplica una lista automática a., b., c.
 */
function runCurrentTest() {
  return applyFastNativeList_('LETTER');
}

/**
 * Aplica una lista automática nativa mediante DocumentApp.
 *
 * Tipos admitidos:
 * - LETTER: a., b., c.
 * - NUMBER: 1., 2., 3.
 * - BULLET: •
 * - ROMAN:  I., II., III.
 *
 * No utiliza Google Docs API.
 */
function applyFastNativeList_(requestedType) {
  const started = Date.now();

  const doc = DocumentApp.getActiveDocument();
  const activeTab = doc.getActiveTab();

  if (!activeTab) {
    throw new Error('No se pudo obtener el tab activo.');
  }

  const documentTab = activeTab.asDocumentTab();
  const body = documentTab.getBody();

  const paragraphs =
    fastListGetTargetParagraphs_(doc, body);

  if (!paragraphs.length) {
    throw new Error(
      'Selecciona párrafos o coloca el cursor dentro de un párrafo.'
    );
  }

  const glyphType =
    fastListGetGlyphType_(requestedType);

  const listItems = [];

  /*
   * Convertir cada Paragraph en ListItem.
   *
   * Si ya es un ListItem, se reutiliza.
   */
  paragraphs.forEach(function (paragraph) {
    let listItem;

    if (
      paragraph.getType() ===
      DocumentApp.ElementType.LIST_ITEM
    ) {
      listItem = paragraph.asListItem();
    } else {
      listItem = fastListConvertParagraph_(
        body,
        paragraph
      );
    }

    listItems.push(listItem);
  });

  /*
   * El primer elemento será la referencia de listId.
   */
  const firstListItem = listItems[0];

  firstListItem
    .setGlyphType(glyphType)
    .setNestingLevel(0)
    .setHeading(
      DocumentApp.ParagraphHeading.NORMAL
    )
    .setIndentStart(36)
    .setIndentFirstLine(18)
    .setIndentEnd(0);

  /*
   * Todos los demás elementos reciben el mismo listId para formar
   * una sola secuencia automática.
   */
  for (
    let index = 1;
    index < listItems.length;
    index++
  ) {
    listItems[index]
      .setListId(firstListItem)
      .setGlyphType(glyphType)
      .setNestingLevel(0)
      .setHeading(
        DocumentApp.ParagraphHeading.NORMAL
      )
      .setIndentStart(36)
      .setIndentFirstLine(18)
      .setIndentEnd(0);
  }

  return {
    ok: true,
    testId: 'TEST-FAST-NATIVE-LIST',
    requestedType:
      String(requestedType).toUpperCase(),
    expectedGlyph:
      fastListExpectedGlyph_(requestedType),
    automaticList: true,
    paragraphsApplied: listItems.length,
    usesDocsApi: false,
    expectedIndentInches: {
      left: 0.25,
      hanging: 0.25,
      right: 0
    },
    elapsedMs: Date.now() - started
  };
}

/**
 * Convierte un Paragraph normal en ListItem.
 *
 * Conserva:
 * - Texto.
 * - Formato de caracteres.
 * - Links.
 *
 * Después elimina el Paragraph original.
 */
function fastListConvertParagraph_(
  body,
  paragraph
) {
  const childIndex =
    body.getChildIndex(paragraph);

  if (childIndex < 0) {
    throw new Error(
      'El párrafo no pertenece al cuerpo principal.'
    );
  }

  const sourceText =
    paragraph.editAsText();

  const textContent =
    sourceText.getText();

  /*
   * Crear el ListItem antes del párrafo original.
   */
  const listItem = body.insertListItem(
    childIndex,
    textContent
  );

  /*
   * Copiar el formato directo del texto.
   */
  fastListCopyTextFormatting_(
    sourceText,
    listItem.editAsText()
  );

  /*
   * El ListItem ya contiene una copia del texto.
   * Ahora puede eliminarse el Paragraph original.
   */
  paragraph.removeFromParent();

  return listItem;
}

/**
 * Copia los diferentes segmentos de formato de un Text a otro.
 */
function fastListCopyTextFormatting_(
  sourceText,
  destinationText
) {
  const textLength =
    sourceText.getText().length;

  if (!textLength) {
    return;
  }

  const attributeIndexes =
    sourceText.getTextAttributeIndices();

  for (
    let index = 0;
    index < attributeIndexes.length;
    index++
  ) {
    const startOffset =
      attributeIndexes[index];

    const endOffset =
      index + 1 < attributeIndexes.length
        ? attributeIndexes[index + 1] - 1
        : textLength - 1;

    const attributes =
      sourceText.getAttributes(startOffset);

    destinationText.setAttributes(
      startOffset,
      endOffset,
      attributes
    );
  }
}

/**
 * Devuelve el GlyphType nativo solicitado.
 */
function fastListGetGlyphType_(requestedType) {
  const normalizedType =
    String(requestedType || '')
      .toUpperCase();

  const glyphTypes = {
    LETTER:
      DocumentApp.GlyphType.LATIN_LOWER,

    NUMBER:
      DocumentApp.GlyphType.NUMBER,

    BULLET:
      DocumentApp.GlyphType.BULLET,

    ROMAN:
      DocumentApp.GlyphType.ROMAN_UPPER
  };

  const glyphType =
    glyphTypes[normalizedType];

  if (!glyphType) {
    throw new Error(
      'Tipo de lista desconocido: ' +
      requestedType
    );
  }

  return glyphType;
}

/**
 * Texto esperado para el reporte de la prueba.
 */
function fastListExpectedGlyph_(requestedType) {
  const expected = {
    LETTER: 'a.',
    NUMBER: '1.',
    BULLET: '•',
    ROMAN: 'I.'
  };

  return expected[
    String(requestedType || '')
      .toUpperCase()
  ] || '';
}

/**
 * Obtiene los párrafos completos seleccionados.
 *
 * Si no hay selección, utiliza el párrafo donde está el cursor.
 */
function fastListGetTargetParagraphs_(
  doc,
  body
) {
  const selection = doc.getSelection();
  const indexes = {};

  if (selection) {
    selection
      .getRangeElements()
      .forEach(function (rangeElement) {
        const paragraph =
          fastListFindParagraph_(
            rangeElement.getElement()
          );

        if (!paragraph) {
          return;
        }

        const childIndex =
          fastListBodyChildIndex_(
            body,
            paragraph
          );

        if (childIndex >= 0) {
          indexes[childIndex] = true;
        }
      });
  } else {
    const cursor = doc.getCursor();

    if (!cursor) {
      throw new Error(
        'No se detectó una selección ni un cursor.'
      );
    }

    const paragraph =
      fastListFindParagraph_(
        cursor.getElement()
      );

    if (!paragraph) {
      throw new Error(
        'El cursor no está dentro de un párrafo.'
      );
    }

    const childIndex =
      fastListBodyChildIndex_(
        body,
        paragraph
      );

    if (childIndex < 0) {
      throw new Error(
        'Esta función solo admite párrafos del cuerpo principal.'
      );
    }

    indexes[childIndex] = true;
  }

  const selectedIndexes =
    Object.keys(indexes)
      .map(Number)
      .sort(function (a, b) {
        return a - b;
      });

  if (!selectedIndexes.length) {
    return [];
  }

  const firstIndex =
    selectedIndexes[0];

  const lastIndex =
    selectedIndexes[
      selectedIndexes.length - 1
    ];

  const paragraphs = [];

  /*
   * Incluye párrafos vacíos entre el primero y el último.
   */
  for (
    let childIndex = firstIndex;
    childIndex <= lastIndex;
    childIndex++
  ) {
    const element =
      body.getChild(childIndex);

    const elementType =
      element.getType();

    if (
      elementType ===
        DocumentApp.ElementType.PARAGRAPH ||
      elementType ===
        DocumentApp.ElementType.LIST_ITEM
    ) {
      paragraphs.push(element);
      continue;
    }

    throw new Error(
      'La selección contiene un elemento que no es un párrafo.'
    );
  }

  return paragraphs;
}

/**
 * Busca el Paragraph o ListItem que contiene un elemento.
 */
function fastListFindParagraph_(element) {
  let current = element;

  while (current) {
    const elementType =
      current.getType();

    if (
      elementType ===
        DocumentApp.ElementType.PARAGRAPH ||
      elementType ===
        DocumentApp.ElementType.LIST_ITEM
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
 * Obtiene el índice del párrafo en el Body.
 */
function fastListBodyChildIndex_(
  body,
  paragraph
) {
  try {
    return body.getChildIndex(paragraph);
  } catch (error) {
    return -1;
  }
}
