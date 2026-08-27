/**
 * Aplica el Named Style configurado actualmente en el documento.
 *
 * Primero elimina los overrides directos mediante null.
 * Después aplica el estilo solicitado.
 *
 * No modifica:
 * - contenido;
 * - hipervínculos;
 * - listId;
 * - tipo de numeración;
 * - nivel de anidación.
 */
function applyStyleToParagraph_(
  paragraph,
  styleName
) {
  if (!paragraph) {
    return null;
  }

  const targetHeading =
    getParagraphHeadingEnum_(styleName);

  paragraph.setAttributes(
    getNamedStyleResetAttributes_()
  );

  paragraph.setHeading(targetHeading);

  resetTextColorToAutomatic_(
    paragraph.editAsText()
  );

  return paragraph;
}

/**
 * Removes a direct foreground-color override so Google Docs can use the
 * inherited/Automatic text color. This intentionally does not force #000000.
 */
function resetTextColorToAutomatic_(textElement) {
  if (!textElement) return textElement;

  const attributes = {};
  attributes[DocumentApp.Attribute.FOREGROUND_COLOR] = null;
  textElement.setAttributes(attributes);

  return textElement;
}

/**
 * Atributos que deben volver a heredarse del Named Style.
 */
function getNamedStyleResetAttributes_() {
  const attributes = {};

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

  attributesToReset.forEach(function (
    attribute
  ) {
    attributes[attribute] = null;
  });

  return attributes;
}

/**
 * Alias utilizado por otras funciones existentes.
 */
function applyNamedStyleToParagraph_(
  paragraph,
  styleName
) {
  return applyStyleToParagraph_(
    paragraph,
    styleName
  );
}

/**
 * Convierte el nombre utilizado por la interfaz
 * al enum de DocumentApp.
 */
function getParagraphHeadingEnum_(styleName) {
  const map = {

    NORMAL:
      DocumentApp.ParagraphHeading.NORMAL,
TITLE:
  DocumentApp.ParagraphHeading.TITLE,

SUBTITLE:
  DocumentApp.ParagraphHeading.SUBTITLE,
    H1:
      DocumentApp.ParagraphHeading.HEADING1,

    H2:
      DocumentApp.ParagraphHeading.HEADING2,

    H3:
      DocumentApp.ParagraphHeading.HEADING3,

    H4:
      DocumentApp.ParagraphHeading.HEADING4,

    H5:
      DocumentApp.ParagraphHeading.HEADING5,

    H6:
      DocumentApp.ParagraphHeading.HEADING6
  };

  const normalizedStyle = String(
    styleName || ''
  ).toUpperCase();

  const heading = map[normalizedStyle];

  if (!heading) {
    throw new Error(
      'Unknown style: ' + styleName
    );
  }

  return heading;
}


/**
 * Reaplica los estilos existentes dentro de una selección.
 *
 * Reglas:
 * - H1–H6 conservan su nivel.
 * - Title y Subtitle se conservan.
 * - Texto común recibe NORMAL.
 * - Párrafos vacíos reciben NORMAL.
 * - Numeración, listas, contenido e hipervínculos se conservan.
 */
function formatSelectedNamedStyles() {
  return formatSelectedSection_();
}

/**
 * Devuelve el nombre del estilo actual del párrafo.
 * Cualquier estilo desconocido se trata como NORMAL.
 */
function getExistingNamedStyleName_(
  paragraph
) {
  const heading = String(
    paragraph.getHeading()
  );

  const map = {
    NORMAL: 'NORMAL',
    TITLE: 'TITLE',
    SUBTITLE: 'SUBTITLE',
    HEADING1: 'H1',
    HEADING2: 'H2',
    HEADING3: 'H3',
    HEADING4: 'H4',
    HEADING5: 'H5',
    HEADING6: 'H6'
  };

  return map[heading] || 'NORMAL';
}


/**
 * B) RETURN SEGMENTS
 * Arguments: selection + comparison.
 *
 * Supported comparisons:
 * ALL, STYLEABLE_TEXT, NORMAL_PARAGRAPH, HEADING, H1..H6,
 * BLANK, TABLE, FIGURE, FIGURE_CAPTION, TABLE_CAPTION, NOTE,
 * LIST, BULLET, NUMBER, LETTER, ROMAN, EQUATION.
 */
function getSegments_(selection, comparison) {
  const segments = collectSegments_(selection);
  return filterSegments_(segments, comparison || 'ALL');
}

/**
 * Public diagnostic wrapper. Returns serializable metadata only.
 */
function getSegmentSummary(comparison) {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  const segments = getSegments_(selection, comparison || 'ALL');

  return segments.map((segment, index) => ({
    index: index,
    type: segment.type,
    subtype: segment.subtype || '',
    text: segment.text || ''
  }));
}

function collectSegments_(selection) {
  const doc = DocumentApp.getActiveDocument();
  const body = getActiveBody_();

  if (selection) {
    const ranges = selection.getRangeElements();
    if (!ranges.length) return [];

    const firstTop = getTopLevelBodyElement_(ranges[0].getElement(), body);
    const lastTop = getTopLevelBodyElement_(
      ranges[ranges.length - 1].getElement(),
      body
    );

    if (firstTop && lastTop) {
      try {
        const firstIndex = body.getChildIndex(firstTop);
        const lastIndex = body.getChildIndex(lastTop);
        const minIndex = Math.min(firstIndex, lastIndex);
        const maxIndex = Math.max(firstIndex, lastIndex);
        const result = [];

        for (let i = minIndex; i <= maxIndex; i++) {
          const segment = classifySegment_(body.getChild(i));
          if (segment) result.push(segment);
        }

        return result;
      } catch (e) {}
    }

    return collectSegmentsFromRangeElements_(ranges);
  }

  const cursor = doc.getCursor();
  if (!cursor) return [];

  const top = getTopLevelBodyElement_(cursor.getElement(), body);

  if (top && top.getType() === DocumentApp.ElementType.TABLE) {
    const owner = getOwningParagraph_(cursor.getElement());
    if (owner) {
      return [{
        element: owner,
        type: 'TABLE_CONTENT',
        subtype: classifyParagraphSubtype_(owner),
        text: owner.getText()
      }];
    }
  }

  const segment = top ? classifySegment_(top) : null;
  return segment ? [segment] : [];
}

function collectSegmentsFromRangeElements_(ranges) {
  const result = [];
  const seen = {};

  ranges.forEach(re => {
    const owner = getOwningParagraph_(re.getElement());
    if (!owner) return;

    const key = buildElementPathKey_(owner);
    if (seen[key]) return;
    seen[key] = true;

    const segment = classifySegment_(owner);
    if (segment) result.push(segment);
  });

  return result;
}

function getTopLevelBodyElement_(element, body) {
  let current = element;

  while (current) {
    let parent = null;
    try { parent = current.getParent(); } catch (e) { return null; }
    if (!parent) return null;

    try {
      if (parent.getType() === DocumentApp.ElementType.BODY_SECTION) {
        body.getChildIndex(current);
        return current;
      }
    } catch (e) {}

    current = parent;
  }

  return null;
}

function getOwningParagraph_(element) {
  let current = element;

  while (current) {
    const type = current.getType();

    if (
      type === DocumentApp.ElementType.PARAGRAPH ||
      type === DocumentApp.ElementType.LIST_ITEM
    ) {
      return current;
    }

    current = current.getParent ? current.getParent() : null;
  }

  return null;
}

function buildElementPathKey_(element) {
  const parts = [];
  let current = element;

  while (current && current.getParent) {
    const parent = current.getParent();
    if (!parent || !parent.getChildIndex) break;

    try {
      parts.unshift(parent.getChildIndex(current));
    } catch (e) {
      break;
    }

    try {
      if (parent.getType() === DocumentApp.ElementType.BODY_SECTION) break;
    } catch (e) {}

    current = parent;
  }

  return parts.join('.');
}

function classifySegment_(element) {
  if (!element || !element.getType) return null;

  const type = element.getType();

  if (type === DocumentApp.ElementType.TABLE) {
    return {
      element: element,
      type: 'TABLE',
      subtype: '',
      text: ''
    };
  }

  if (
    type !== DocumentApp.ElementType.PARAGRAPH &&
    type !== DocumentApp.ElementType.LIST_ITEM
  ) {
    return {
      element: element,
      type: 'OTHER',
      subtype: String(type),
      text: ''
    };
  }

  const text = element.getText() || '';
  const subtype = classifyParagraphSubtype_(element);

  return {
    element: element,
    type: subtype,
    subtype: subtype,
    text: text
  };
}

function classifyParagraphSubtype_(paragraph) {
  const text = String(paragraph.getText() || '');
  const trimmed = text.trim();

  if (!trimmed) return 'BLANK';

  // IMPORTANT:
  // A numbered Heading in Google Docs is usually a ListItem AND a Heading.
  // Heading identity must therefore be checked before list identity.
  try {
    const heading = paragraph.getHeading();

    if (heading === DocumentApp.ParagraphHeading.HEADING1) return 'H1';
    if (heading === DocumentApp.ParagraphHeading.HEADING2) return 'H2';
    if (heading === DocumentApp.ParagraphHeading.HEADING3) return 'H3';
    if (heading === DocumentApp.ParagraphHeading.HEADING4) return 'H4';
    if (heading === DocumentApp.ParagraphHeading.HEADING5) return 'H5';
    if (heading === DocumentApp.ParagraphHeading.HEADING6) return 'H6';
  } catch (e) {}

  if (/^\s*figure\b/i.test(trimmed)) return 'FIGURE_CAPTION';
  if (/^\s*table\b/i.test(trimmed)) return 'TABLE_CAPTION';
  if (/^\s*(?:notes?|notas?)\b/i.test(trimmed)) return 'NOTE';

  if (paragraphContainsFigure_(paragraph)) return 'FIGURE';
  if (paragraphContainsEquation_(paragraph)) return 'EQUATION';

  if (paragraph.getType() === DocumentApp.ElementType.LIST_ITEM) {
    return classifyListItem_(paragraph.asListItem());
  }

  return 'NORMAL_PARAGRAPH';
}

function classifyListItem_(item) {
  try {
    const glyph = item.getGlyphType();

    if (glyph === DocumentApp.GlyphType.BULLET) return 'BULLET';
    if (glyph === DocumentApp.GlyphType.NUMBER) return 'NUMBER';
    if (glyph === DocumentApp.GlyphType.LATIN_LOWER) return 'LETTER';
    if (glyph === DocumentApp.GlyphType.LATIN_UPPER) return 'LETTER';
    if (glyph === DocumentApp.GlyphType.ROMAN_LOWER) return 'ROMAN';
    if (glyph === DocumentApp.GlyphType.ROMAN_UPPER) return 'ROMAN';
  } catch (e) {}

  return 'LIST';
}

function paragraphContainsFigure_(paragraph) {
  if (paragraphContainsInlineFigure_(paragraph)) return true;

  try {
    const positioned = paragraph.getPositionedImages();
    if (positioned && positioned.length) return true;
  } catch (e) {}

  return false;
}

function paragraphContainsInlineFigure_(paragraph) {
  try {
    for (let i = 0; i < paragraph.getNumChildren(); i++) {
      const childType = paragraph.getChild(i).getType();

      if (
        childType === DocumentApp.ElementType.INLINE_IMAGE ||
        childType === DocumentApp.ElementType.INLINE_DRAWING
      ) {
        return true;
      }
    }
  } catch (e) {}

  return false;
}

function paragraphContainsEquation_(paragraph) {
  try {
    for (let i = 0; i < paragraph.getNumChildren(); i++) {
      if (paragraph.getChild(i).getType() === DocumentApp.ElementType.EQUATION) {
        return true;
      }
    }
  } catch (e) {}

  return false;
}

function filterSegments_(segments, comparison) {
  const filters = Array.isArray(comparison)
    ? comparison.map(v => String(v).toUpperCase())
    : [String(comparison || 'ALL').toUpperCase()];

  if (filters.indexOf('ALL') >= 0) return segments;

  return segments.filter(segment =>
    filters.some(filter => segmentMatches_(segment, filter))
  );
}

function segmentMatches_(segment, filter) {
  const type = String(segment.type || '').toUpperCase();

  if (filter === type) return true;

  if (filter === 'HEADING') {
    return /^H[1-6]$/.test(type);
  }

  if (filter === 'LIST') {
    return ['LIST', 'BULLET', 'NUMBER', 'LETTER', 'ROMAN'].indexOf(type) >= 0;
  }

  if (filter === 'STYLEABLE_TEXT') {
    return (
      type === 'NORMAL_PARAGRAPH' ||
      type === 'BLANK' ||
      /^H[1-6]$/.test(type)
    );
  }

  if (filter === 'CAPTION') {
    return type === 'FIGURE_CAPTION' || type === 'TABLE_CAPTION';
  }

  return false;
}

/**
 * Aplica el estilo a todos los párrafos o ListItems
 * encontrados en la selección.
 *
 * Se utiliza ALL porque un encabezado numerado puede
 * seguir siendo un ListItem con estilo NORMAL.
 */
function applyStyleToSelection_(
  selection,
  styleName
) {
  const segments = getSegments_(
    selection,
    'ALL'
  );

  let count = 0;

  segments.forEach(function (segment) {
    const element = segment.element;

    if (
      !element ||
      typeof element.getType !== 'function'
    ) {
      return;
    }

    const elementType = element.getType();

    const isStyleable =
      elementType ===
        DocumentApp.ElementType.PARAGRAPH ||
      elementType ===
        DocumentApp.ElementType.LIST_ITEM;

    if (!isStyleable) {
      return;
    }

    applyStyleToParagraph_(
      element,
      styleName
    );

    count++;
  });

  return count;
}

/**
 * Entrada pública utilizada por los botones del sidebar.
 */
function applyStyleToCurrentContext(styleName) {
  const started = Date.now();
  const document =
    DocumentApp.getActiveDocument();

  const selection = document.getSelection();

  let count = applyStyleToSelection_(
    selection,
    styleName
  );

  /*
   * Fallback para el cursor si la segmentación
   * no devolvió ningún elemento.
   */
  if (!count) {
    const cursor = document.getCursor();

    if (cursor) {
      const owner = getOwningParagraph_(
        cursor.getElement()
      );

      if (owner) {
        applyStyleToParagraph_(
          owner,
          styleName
        );

        count = 1;
      }
    }
  }

  if (!count) {
    throw new Error(
      'No paragraph or heading was found ' +
      'at the cursor/selection.'
    );
  }

  /*
   * Fuerza el guardado antes de responder
   * a la interfaz.
   */
  document.saveAndClose();

  return {
    ok: true,
    paragraphs: count,
    elapsedMs: Date.now() - started
  };
}

/**
 * Nombre público anterior, conservado por compatibilidad.
 */
function applyNamedStyle(styleName) {
  return applyStyleToCurrentContext(styleName);
}
function sentenceCaseHeading_(value) {
  value = String(value || '');
  if (!value) return value;

  const lower = value.toLowerCase();

  // Capitalize the first alphabetic character while preserving a leading
  // number, section number, punctuation, etc.
  // "2. PROCESS CONDITIONS" -> "2. Process conditions"
  return lower.replace(
    /[A-Za-zÁÉÍÓÚÜÑÀÈÌÒÙÂÊÎÔÛÄËÏÖÜÇ]/,
    ch => ch.toUpperCase()
  );
}

function getCurrentIndentation() {
  const targets = getStyleTargetParagraphs_();
  const target = targets.length ? targets[0] : null;

  if (!target) {
    throw new Error('Place the cursor in a paragraph or select text first.');
  }

  const PT_PER_IN = 72;
  const start = Number(target.getIndentStart() || 0) / PT_PER_IN;
  const end = Number(target.getIndentEnd() || 0) / PT_PER_IN;
  const first = Number(target.getIndentFirstLine() || 0) / PT_PER_IN;

  let special = 'NONE';
  let by = 0;

  if (first > start + 0.001) {
    special = 'FIRST_LINE';
    by = first - start;
  } else if (start > first + 0.001) {
    special = 'HANGING';
    by = start - first;
  }

  return {
    left: roundIndentValue_(special === 'HANGING' ? first : start),
    right: roundIndentValue_(end),
    special: special,
    by: roundIndentValue_(by)
  };
}

function applyIndentation(left, right, special, by) {
  const started = Date.now();
  const paragraphs = getStyleTargetParagraphs_();

  if (!paragraphs.length) {
    throw new Error('Place the cursor in a paragraph or select one or more paragraphs.');
  }

  const leftIn = Number(left);
  const rightIn = Number(right);
  const byIn = Number(by || 0);
  const mode = String(special || 'NONE').toUpperCase();

  if (!Number.isFinite(leftIn) || !Number.isFinite(rightIn) || !Number.isFinite(byIn)) {
    throw new Error('Indentation values must be valid numbers.');
  }

  if (leftIn < 0 || rightIn < 0 || byIn < 0) {
    throw new Error('Indentation values cannot be negative.');
  }

  if (['NONE', 'FIRST_LINE', 'HANGING'].indexOf(mode) === -1) {
    throw new Error('Unknown special indent.');
  }

  const PT_PER_IN = 72;
  const leftPt = leftIn * PT_PER_IN;
  const rightPt = rightIn * PT_PER_IN;
  const byPt = byIn * PT_PER_IN;
  const attribute = DocumentApp.Attribute;
  const attributes = {};

  if (mode === 'HANGING') {
    // First line starts at Left; wrapped lines at Left + Hanging.
    attributes[attribute.INDENT_FIRST_LINE] = leftPt;
    attributes[attribute.INDENT_START] = leftPt + byPt;
  } else if (mode === 'FIRST_LINE') {
    // Wrapped lines start at Left; first line starts at Left + By.
    attributes[attribute.INDENT_START] = leftPt;
    attributes[attribute.INDENT_FIRST_LINE] = leftPt + byPt;
  } else {
    attributes[attribute.INDENT_START] = leftPt;
    attributes[attribute.INDENT_FIRST_LINE] = leftPt;
  }

  attributes[attribute.INDENT_END] = rightPt;

  paragraphs.forEach(p => {
    p.setAttributes(attributes);
  });

  return {
    ok: true,
    paragraphs: paragraphs.length,
    left: leftIn,
    right: rightIn,
    special: mode,
    by: byIn,
    elapsedMs: Date.now() - started
  };
}

function roundIndentValue_(value) {
  return Math.round(Number(value || 0) * 1000) / 1000;
}

function setParagraphSpacing(before, after, lineSpacing) {
  const started = Date.now();
  const paragraphs = getStyleTargetParagraphs_();

  if (!paragraphs.length) {
    throw new Error('Place the cursor in a paragraph or select one or more paragraphs.');
  }

  const attribute = DocumentApp.Attribute;
  const attributes = {};

  if (before !== null && before !== undefined) {
    const value = Number(before);
    if (!Number.isFinite(value) || value < 0) {
      throw new Error('Spacing before must be zero or greater.');
    }
    attributes[attribute.SPACING_BEFORE] = value;
  }

  if (after !== null && after !== undefined) {
    const value = Number(after);
    if (!Number.isFinite(value) || value < 0) {
      throw new Error('Spacing after must be zero or greater.');
    }
    attributes[attribute.SPACING_AFTER] = value;
  }

  if (lineSpacing !== null && lineSpacing !== undefined) {
    const value = Number(lineSpacing);
    if (!Number.isFinite(value) || value <= 0) {
      throw new Error('Line spacing must be greater than zero.');
    }
    attributes[attribute.LINE_SPACING] = value;
  }

  paragraphs.forEach(p => {
    p.setAttributes(attributes);
  });

  return {
    ok: true,
    paragraphs: paragraphs.length,
    elapsedMs: Date.now() - started
  };
}

function setKeepWithNext(value) {
  const started = Date.now();
  const paragraphs = getStyleTargetParagraphs_();

  if (!paragraphs.length) {
    throw new Error('Place the cursor in a paragraph or select one or more paragraphs.');
  }

  paragraphs.forEach(p => p.setKeepWithNext(Boolean(value)));

  return {
    ok: true,
    paragraphs: paragraphs.length,
    elapsedMs: Date.now() - started
  };
}

/**
 * Applies a real, automatic Google Docs list to the current paragraph or to
 * every paragraph in the selection. This path intentionally uses only
 * DocumentApp so the four list buttons remain fast and predictable.
 *
 * The previous-list/continue behavior is deliberately not part of this UI
 * command: the selected paragraphs always become one native list sequence.
 */
function applyListPreset(type) {
  return applyFastNativeListPreset_(type);
}

function applyFastNativeListPreset_(requestedType) {
  const started = Date.now();
  const doc = DocumentApp.getActiveDocument();
  const activeTab = doc.getActiveTab();

  if (!activeTab) {
    throw new Error('Could not read the active document tab.');
  }

  const body = activeTab.asDocumentTab().getBody();
  const paragraphs = uiListGetTargetParagraphs_(doc, body);

  if (!paragraphs.length) {
    throw new Error('Place the cursor in a paragraph or select one or more paragraphs.');
  }

  const result = applyFastNativeListToParagraphs_(
    body,
    paragraphs,
    requestedType,
    null
  );

  return {
    ok: true,
    requestedType: result.requestedType,
    automaticList: true,
    paragraphsApplied: result.listItems.length,
    glyph: result.glyph,
    usesDocsApi: false,
    elapsedMs: Date.now() - started
  };
}

/**
 * Shared native-list engine used by both the four list buttons and
 * Full Smart Format. It uses DocumentApp only.
 *
 * When anchorListItem is supplied, the new items continue that native list.
 */
function applyFastNativeListToParagraphs_(
  body,
  paragraphs,
  requestedType,
  anchorListItem
) {
  const normalizedType = String(requestedType || '').toUpperCase();
  const glyphType = uiListGetGlyphType_(normalizedType);
  const listItems = paragraphs.map(function(paragraph) {
    if (paragraph.getType() === DocumentApp.ElementType.LIST_ITEM) {
      return paragraph.asListItem();
    }

    return uiListConvertParagraph_(body, paragraph);
  });

  if (!listItems.length) {
    throw new Error('No paragraphs were supplied to the native-list engine.');
  }

  const firstListItem = anchorListItem || listItems[0];
  const firstNewItemIndex = anchorListItem ? 0 : 1;

  if (!anchorListItem) {
    uiListFormatItem_(firstListItem, glyphType);
  }

  for (let i = firstNewItemIndex; i < listItems.length; i++) {
    listItems[i].setListId(firstListItem);
    uiListFormatItem_(listItems[i], glyphType);
  }

  return {
    requestedType: normalizedType,
    glyph: uiListExpectedGlyph_(normalizedType),
    listItems: listItems,
    firstListItem: firstListItem
  };
}

function uiListFormatItem_(listItem, glyphType) {
  listItem
    .setGlyphType(glyphType)
    .setNestingLevel(0)
    .setHeading(DocumentApp.ParagraphHeading.NORMAL)
    .setLineSpacing(1.5)
    .setIndentStart(36)
    .setIndentFirstLine(18)
    .setIndentEnd(0);

  resetTextColorToAutomatic_(listItem.editAsText());
}

function uiListConvertParagraph_(body, paragraph) {
  const childIndex = body.getChildIndex(paragraph);

  if (childIndex < 0) {
    throw new Error('The paragraph does not belong to the active document body.');
  }

  const sourceText = paragraph.editAsText();
  const textContent = sourceText.getText();
  const listItem = body.insertListItem(childIndex, textContent);

  uiListCopyTextFormatting_(sourceText, listItem.editAsText());
  paragraph.removeFromParent();

  return listItem;
}

function uiListCopyTextFormatting_(sourceText, destinationText) {
  const textLength = sourceText.getText().length;
  if (!textLength) return;

  const indices = sourceText.getTextAttributeIndices();

  for (let i = 0; i < indices.length; i++) {
    const start = indices[i];
    const end = i + 1 < indices.length ? indices[i + 1] - 1 : textLength - 1;
    destinationText.setAttributes(start, end, sourceText.getAttributes(start));
  }
}

function uiListGetGlyphType_(type) {
  const glyphTypes = {
    BULLET: DocumentApp.GlyphType.BULLET,
    NUMBER: DocumentApp.GlyphType.NUMBER,
    LETTER: DocumentApp.GlyphType.LATIN_LOWER,
    ROMAN: DocumentApp.GlyphType.ROMAN_UPPER
  };

  if (!glyphTypes[type]) {
    throw new Error('Unknown list type: ' + type);
  }

  return glyphTypes[type];
}

function uiListExpectedGlyph_(type) {
  return {
    BULLET: '•',
    NUMBER: '1.',
    LETTER: 'a.',
    ROMAN: 'I.'
  }[type] || '';
}

function uiListGetTargetParagraphs_(doc, body) {
  const selection = doc.getSelection();
  const foundIndexes = {};

  if (selection) {
    selection.getRangeElements().forEach(function(rangeElement) {
      const paragraph = uiListFindParagraph_(rangeElement.getElement());
      if (!paragraph) return;

      const index = uiListBodyChildIndex_(body, paragraph);
      if (index >= 0) foundIndexes[index] = true;
    });
  } else {
    const cursor = doc.getCursor();

    if (!cursor) {
      throw new Error('No selection or cursor was detected.');
    }

    const paragraph = uiListFindParagraph_(cursor.getElement());
    if (!paragraph) {
      throw new Error('The cursor is not inside a paragraph.');
    }

    const index = uiListBodyChildIndex_(body, paragraph);
    if (index < 0) {
      throw new Error('Lists can currently be applied only in the main document body.');
    }

    foundIndexes[index] = true;
  }

  const selectedIndexes = Object.keys(foundIndexes)
    .map(Number)
    .sort(function(a, b) { return a - b; });

  if (!selectedIndexes.length) return [];

  const first = selectedIndexes[0];
  const last = selectedIndexes[selectedIndexes.length - 1];
  const paragraphs = [];

  for (let i = first; i <= last; i++) {
    const element = body.getChild(i);
    const elementType = element.getType();

    if (
      elementType === DocumentApp.ElementType.PARAGRAPH ||
      elementType === DocumentApp.ElementType.LIST_ITEM
    ) {
      paragraphs.push(element);
      continue;
    }

    throw new Error('The selection contains an element that is not a paragraph.');
  }

  return paragraphs;
}

function uiListFindParagraph_(element) {
  let current = element;

  while (current) {
    const elementType = current.getType();

    if (
      elementType === DocumentApp.ElementType.PARAGRAPH ||
      elementType === DocumentApp.ElementType.LIST_ITEM
    ) {
      return current;
    }

    current = current.getParent ? current.getParent() : null;
  }

  return null;
}

function uiListBodyChildIndex_(body, paragraph) {
  try {
    return body.getChildIndex(paragraph);
  } catch (error) {
    return -1;
  }
}
function insertBreak(kind) {
  const started = Date.now();
  const doc = DocumentApp.getActiveDocument();
  const requested = String(kind || 'PAGE').toUpperCase();

  if (requested !== 'PAGE') {
    throw new Error('Only local page breaks are supported.');
  }

  const selection = doc.getSelection();

  if (selection) {
    const ranges = selection.getRangeElements();

    if (ranges.length) {
      const body = getActiveBody_();
      const first = getTopLevelElementForParent_(
        ranges[0].getElement(),
        body
      );

      if (!first) {
        throw new Error('The selection must be inside the active document body.');
      }

      body.insertPageBreak(body.getChildIndex(first));

      return {
        ok: true,
        kind: 'PAGE',
        location: 'SELECTION_START',
        elapsedMs: Date.now() - started
      };
    }
  }

  const cursor = doc.getCursor();
  if (!cursor) {
    throw new Error('Place the cursor or select the block before which the page break should be inserted.');
  }

  cursor.insertPageBreak();

  return {
    ok: true,
    kind: 'PAGE',
    location: 'CURSOR',
    elapsedMs: Date.now() - started
  };
}

function getSelectedParagraphs_() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) return [];

  const paragraphs = [];
  const seen = {};

  selection.getRangeElements().forEach(re => {
    let el = re.getElement();

    // RangeElements can point to Text or other nested elements.
    // Walk upward until reaching the owning paragraph/list item.
    while (
      el &&
      el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
      el.getType() !== DocumentApp.ElementType.LIST_ITEM
    ) {
      el = el.getParent();
    }

    const key = el ? buildElementPathKey_(el) : '';

    if (el && key && !seen[key]) {
      seen[key] = true;
      paragraphs.push(el);
    }
  });

  return paragraphs;
}

/**
 * Returns every body-level paragraph/ListItem between the first and last
 * selected elements. This includes blank lines that Google Docs may omit from
 * Selection.getRangeElements(). A selection contained inside one table falls
 * back to the table-cell paragraphs returned by getSelectedParagraphs_().
 */
function getContiguousSelectedParagraphs_(selection) {
  const body = getActiveBody_();
  const ranges = selection ? selection.getRangeElements() : [];
  const paragraphs = [];

  if (!ranges.length) return paragraphs;

  const first = getTopLevelBodyElement_(ranges[0].getElement(), body);
  const last = getTopLevelBodyElement_(
    ranges[ranges.length - 1].getElement(),
    body
  );

  if (first && last) {
    const firstIndex = body.getChildIndex(first);
    const lastIndex = body.getChildIndex(last);
    const minIndex = Math.min(firstIndex, lastIndex);
    const maxIndex = Math.max(firstIndex, lastIndex);

    for (let index = minIndex; index <= maxIndex; index++) {
      const element = body.getChild(index);
      const type = element.getType();

      if (
        type === DocumentApp.ElementType.PARAGRAPH ||
        type === DocumentApp.ElementType.LIST_ITEM
      ) {
        paragraphs.push(element);
      }
    }
  }

  return paragraphs.length
    ? paragraphs
    : getSelectedParagraphs_();
}

/**
 * Returns every paragraph/ListItem targeted by the current selection, or the
 * paragraph that owns the cursor when there is no selection.
 *
 * This shared selector is intentionally based only on DocumentApp. It is used
 * by indentation, spacing, equations, captions and notes.
 */
function getStyleTargetParagraphs_() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  if (selection) {
    return getContiguousSelectedParagraphs_(selection);
  }

  const cursor = doc.getCursor();
  if (!cursor) return [];

  const paragraph = getOwningParagraph_(cursor.getElement());
  return paragraph ? [paragraph] : [];
}

function formatEquationLine() {
  const existingEquation = getActiveEquationTableContext_();

  if (existingEquation) {
    const equationMarkers = readEquationMarkerIndices_(
      existingEquation.body
    );
    const isKnownEquation =
      isEquationLayoutTable_(existingEquation.table) ||
      Boolean(equationMarkers[existingEquation.bodyIndex]);

    if (
      !isKnownEquation ||
      !hasEquationLayoutStructure_(existingEquation.table)
    ) {
      throw new Error(
        'The selected table is not an equation row created by Docs Assistant.'
      );
    }

    return refreshExistingEquationTable_(existingEquation);
  }

  const targets = getStyleTargetParagraphs_();

  if (targets.length !== 1) {
    throw new Error('Place the cursor in one equation line or select only that line.');
  }

  const source = targets[0];

  if (isInsideTable_(source)) {
    throw new Error('The equation line is already inside a table.');
  }

  if (source.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    throw new Error('Equation formatting requires a normal paragraph, not a list item.');
  }

  const body = getActiveBody_();
  const top = getTopLevelElementForParent_(source, body);

  if (!top || top.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    throw new Error('The equation line must be a body-level paragraph.');
  }

  const sourceParagraph = top.asParagraph();
  const sourceIndex = body.getChildIndex(sourceParagraph);
  const equationNumber = getNextEquationNumberBeforeIndex_(body, sourceIndex);

  // Paragraph.copy() is a deep copy, so a real embedded Equation remains
  // an Equation rather than being flattened to plain text.
  const sourceCopy = sourceParagraph.copy();

  const table = body.insertTable(sourceIndex, [['', '', '']]);
  table.setBorderWidth(0);

  applyEquationTableColumnWidths_(table, body, equationNumber);

  const row = table.getRow(0);
  row.setMinimumHeight(24);

  const leftCell = row.getCell(0);
  const centerCell = row.getCell(1);
  const rightCell = row.getCell(2);

  [leftCell, centerCell, rightCell].forEach(cell => {
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    cell.setPaddingTop(0);
    cell.setPaddingBottom(0);
    cell.setPaddingLeft(0);
    cell.setPaddingRight(0);
  });

  leftCell.setText('');

  // Center: preserve rich content/equation.
  centerCell.clear();
  const centerParagraph = centerCell.appendParagraph(sourceCopy);
  removeEmptyCellParagraphsExcept_(centerCell, centerParagraph);
  centerParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  centerParagraph.setIndentStart(0);
  centerParagraph.setIndentEnd(0);
  centerParagraph.setIndentFirstLine(0);
  centerParagraph.setSpacingBefore(0);
  centerParagraph.setSpacingAfter(0);

  // Right: dotted leader + Equation N.
  const label = buildEquationLabelText_(equationNumber);
  rightCell.setText(label);

  const rightParagraph = rightCell.getChild(0).asParagraph();
  rightParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  rightParagraph.setIndentStart(0);
  rightParagraph.setIndentEnd(0);
  rightParagraph.setIndentFirstLine(0);
  rightParagraph.setSpacingBefore(0);
  rightParagraph.setSpacingAfter(0);

  formatEquationLabel_(rightCell, equationNumber);

  sourceParagraph.removeFromParent();

  const renumberResult = renumberEquationTablesFromBodyIndex_(
    body,
    sourceIndex,
    [{bodyIndex: sourceIndex, element: table}]
  );

  return {
    ok: true,
    mode: 'CREATED',
    equationNumber:
      renumberResult.numberByBodyIndex[sourceIndex] || equationNumber,
    equationsRenumbered: renumberResult.renumbered,
    equationsIndexed: renumberResult.equationsIndexed,
    equationMarkersRecovered: renumberResult.equationMarkersRecovered,
    equationIndexMs: renumberResult.equationIndexMs,
    followingEquationsRenumbered:
      Math.max(0, renumberResult.renumbered - 1)
  };
}

/**
 * Returns the one body-level table targeted entirely by the current cursor or
 * selection. Body indices are compared instead of Apps Script wrapper object
 * identity so selecting several cells in the same equation remains reliable.
 */
function getActiveEquationTableContext_() {
  const doc = DocumentApp.getActiveDocument();
  const body = getActiveBody_();
  const selection = doc.getSelection();
  const elements = [];

  if (selection) {
    const ranges = selection.getRangeElements();
    for (let index = 0; index < ranges.length; index++) {
      elements.push(ranges[index].getElement());
    }
  } else {
    const cursor = doc.getCursor();
    if (cursor) elements.push(cursor.getElement());
  }

  if (!elements.length) return null;

  let activeBodyIndex = null;

  for (let index = 0; index < elements.length; index++) {
    const table = findAncestorTable_(elements[index]);
    if (!table) return null;

    const top = getTopLevelElementForParent_(table, body);
    if (!top || top.getType() !== DocumentApp.ElementType.TABLE) return null;

    const bodyIndex = body.getChildIndex(top);

    if (activeBodyIndex === null) {
      activeBodyIndex = bodyIndex;
    } else if (bodyIndex !== activeBodyIndex) {
      return null;
    }
  }

  return {
    body: body,
    bodyIndex: activeBodyIndex,
    table: body.getChild(activeBodyIndex).asTable()
  };
}

/**
 * Updates an existing equation row in place and renumbers it together with
 * every later equation. No layout table is recreated.
 */
function refreshExistingEquationTable_(context) {
  const renumberResult = renumberEquationTablesFromBodyIndex_(
    context.body,
    context.bodyIndex,
    [{bodyIndex: context.bodyIndex, element: context.table}]
  );
  const equationNumber =
    renumberResult.numberByBodyIndex[context.bodyIndex];

  if (!equationNumber) {
    throw new Error('The equation row could not be added to the document index.');
  }

  return {
    ok: true,
    mode: 'UPDATED',
    equationNumber: equationNumber,
    equationsRenumbered: renumberResult.renumbered,
    equationsIndexed: renumberResult.equationsIndexed,
    equationMarkersRecovered: renumberResult.equationMarkersRecovered,
    equationIndexMs: renumberResult.equationIndexMs,
    followingEquationsRenumbered:
      Math.max(0, renumberResult.renumbered - 1)
  };
}

/**
 * Applies one continuous bold run to "Equation N" and explicitly keeps it
 * non-italic. The dotted leader remains regular.
 */
function formatEquationLabel_(rightCell, equationNumber) {
  const expected = buildEquationLabelText_(equationNumber);
  const current = String(rightCell.getText() || '').trim();
  const numberCorrected = current !== expected;
  const paragraph = rightCell.getChild(0).asParagraph();
  const text = paragraph.editAsText();

  if (numberCorrected) {
    // Keep the existing paragraph object. TableCell.setText() may replace the
    // paragraph and invalidate the NamedRange used as the equation marker.
    text.setText(expected);
  }

  paragraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  paragraph.setIndentStart(0);
  paragraph.setIndentEnd(0);
  paragraph.setIndentFirstLine(0);
  paragraph.setSpacingBefore(0);
  paragraph.setSpacingAfter(0);

  text.setFontFamily('Arial').setFontSize(9).setBold(false).setItalic(false);
  resetTextColorToAutomatic_(text);

  const boldLabel = buildEquationBoldLabelText_(equationNumber);
  const labelStart = expected.indexOf(boldLabel);

  if (labelStart >= 0) {
    text.setBold(labelStart, labelStart + boldLabel.length - 1, true);
    text.setItalic(labelStart, labelStart + boldLabel.length - 1, false);
  }

  return {numberCorrected: numberCorrected};
}

function buildEquationLabelText_(equationNumber) {
  return '.................... ' +
    buildEquationBoldLabelText_(equationNumber);
}

function buildEquationBoldLabelText_(equationNumber) {
  // A non-breaking space looks identical but prevents the number from being
  // orphaned on a second line in Google Docs.
  return 'Equation\u00A0' + equationNumber;
}

/**
 * Rebuilds the complete equation index from the direct table collection and
 * updates only the selected/new equation and the equations that follow it.
 */
function renumberEquationTablesFromBodyIndex_(
  body,
  firstBodyIndex,
  selectedEquationEntries
) {
  const objectIndex = buildCompleteEquationIndex_(
    body,
    selectedEquationEntries || []
  );
  const numberByBodyIndex = {};
  let renumbered = 0;

  objectIndex.equationIndices.forEach(function(bodyIndex, position) {
    const equationNumber = position + 1;
    numberByBodyIndex[bodyIndex] = equationNumber;

    if (bodyIndex < firstBodyIndex) return;

    const element = body.getChild(bodyIndex);
    if (!hasEquationLayoutStructure_(element.asTable())) return;

    const table = element.asTable();
    applyEquationTableColumnWidths_(table, body, equationNumber);
    formatEquationLabel_(table.getRow(0).getCell(2), equationNumber);
    renumbered++;
  });

  return {
    renumbered: renumbered,
    numberByBodyIndex: numberByBodyIndex,
    equationsIndexed: objectIndex.equationIndices.length,
    equationMarkersRecovered: objectIndex.markersRecovered,
    equationIndexMs: objectIndex.elapsedMs
  };
}

/**
 * Reserves enough room for the full dotted "Equation N" label without
 * sacrificing geometric centering. Both side columns always receive the same
 * width; the middle column remains centered on the page.
 */
function applyEquationTableColumnWidths_(table, body, equationNumber) {
  const calculatedWidth =
    Number(body.getPageWidth()) -
    Number(body.getMarginLeft()) -
    Number(body.getMarginRight());
  const usableWidth =
    Number.isFinite(calculatedWidth) && calculatedWidth > 0
      ? calculatedWidth
      : 468;
  const digitCount = String(Math.abs(Number(equationNumber) || 0)).length;

  // 132 pt safely accommodates the 20-dot leader plus "Equation 999" at
  // Arial 9. Add a little room only if the counter grows beyond three digits.
  const labelMinimumWidth =
    132 + Math.max(0, digitCount - 3) * 6;
  const preferredSideWidth = Math.max(
    labelMinimumWidth,
    usableWidth * 0.28
  );
  const minimumCenterWidth = Math.max(
    144,
    usableWidth * 0.38
  );
  const maximumSideWidth = Math.max(
    72,
    (usableWidth - minimumCenterWidth) / 2
  );
  const sideWidth = Math.min(
    preferredSideWidth,
    maximumSideWidth
  );
  const centerWidth = Math.max(
    72,
    usableWidth - sideWidth * 2
  );

  table.setColumnWidth(0, sideWidth);
  table.setColumnWidth(1, centerWidth);
  table.setColumnWidth(2, sideWidth);

  return {
    usableWidth: usableWidth,
    sideWidth: sideWidth,
    centerWidth: centerWidth,
    digitCount: digitCount
  };
}

/**
 * Discovers every body-level equation table on each equation-button run.
 * Valid labels are sufficient for discovery; existing markers also allow a
 * damaged label to be repaired. Missing markers are restored for later fast
 * section-formatting runs.
 */
function buildCompleteEquationIndex_(body, selectedEquationEntries) {
  const started = Date.now();
  const markerIndexSet = readEquationMarkerIndices_(body);
  const selectedIndexSet = {};
  let markersRecovered = 0;

  (selectedEquationEntries || []).forEach(function(entry) {
    selectedIndexSet[entry.bodyIndex] = true;
  });

  const equationIndices = [];

  getTopLevelTableEntries_(body).forEach(function(entry) {
    const structurallyValid = hasEquationLayoutStructure_(entry.table);
    const recognized = structurallyValid && (
      isEquationLayoutTable_(entry.table) ||
      Boolean(markerIndexSet[entry.bodyIndex]) ||
      Boolean(selectedIndexSet[entry.bodyIndex])
    );

    if (!recognized) return;

    equationIndices.push(entry.bodyIndex);

    if (!markerIndexSet[entry.bodyIndex]) {
      try {
        addEquationTableMarker_(entry.table);
        markerIndexSet[entry.bodyIndex] = true;
        markersRecovered++;
      } catch (error) {}
    }
  });

  equationIndices.sort(function(a, b) { return a - b; });

  return {
    equationIndices: equationIndices,
    markersRecovered: markersRecovered,
    elapsedMs: Date.now() - started
  };
}

/**
 * Returns each unique body-level physical table without walking every body
 * child. Nested tables are collapsed into their top-level physical table.
 */
function getTopLevelTableEntries_(body) {
  const entries = [];
  const seenBodyIndices = {};
  const tables = body.getTables() || [];

  tables.forEach(function(table) {
    const top = getTopLevelElementForParent_(table, body);
    if (!top || top.getType() !== DocumentApp.ElementType.TABLE) return;

    const bodyIndex = body.getChildIndex(top);
    if (seenBodyIndices[bodyIndex]) return;

    seenBodyIndices[bodyIndex] = true;
    entries.push({
      bodyIndex: bodyIndex,
      table: top.asTable()
    });
  });

  entries.sort(function(a, b) {
    return a.bodyIndex - b.bodyIndex;
  });

  return entries;
}

function removeEmptyCellParagraphsExcept_(cell, keepParagraph) {
  if (cell.getNumChildren() <= 1) return;

  for (let i = cell.getNumChildren() - 1; i >= 0; i--) {
    const child = cell.getChild(i);

    // Avoid wrapper identity assumptions: the kept paragraph has content
    // copied from the source; only delete empty paragraphs.
    if (
      child.getType() === DocumentApp.ElementType.PARAGRAPH &&
      child.asParagraph().getText() === '' &&
      cell.getNumChildren() > 1
    ) {
      try { child.removeFromParent(); } catch (e) {}
    }
  }
}

function getNextEquationNumberBeforeIndex_(body, targetIndex) {
  let maxNumber = 0;

  for (let i = 0; i < targetIndex; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TABLE) continue;

    const table = child.asTable();
    if (table.getNumRows() < 1) continue;

    const row = table.getRow(0);
    if (row.getNumCells() !== 3) continue;

    const rightText = row.getCell(2).getText();
    const match = String(rightText || '').match(/\bEquation\s+(\d+)\s*$/i);

    if (match) {
      maxNumber = Math.max(maxNumber, Number(match[1]) || 0);
    }
  }

  return maxNumber + 1;
}

/**
 * Identifies the one-row, three-cell layout table created by
 * formatEquationLine(). The label pattern is used instead of border width so
 * the table remains identifiable even if its borders were changed later.
 */
function isEquationLayoutTable_(table) {
  try {
    if (!hasEquationLayoutStructure_(table)) return false;

    const row = table.getRow(0);
    const rightText = String(row.getCell(2).getText() || '')
      .replace(/\u00A0/g, ' ')
      .trim();

    return /^\.{10,}\s*Equation\s+\d+\s*$/i.test(rightText);
  } catch (error) {
    return false;
  }
}

/**
 * Checks the stable physical shape of an equation row independently of its
 * label. This lets a persistent marker repair a missing or damaged number.
 */
function hasEquationLayoutStructure_(table) {
  try {
    if (!table || table.getNumRows() !== 1) return false;

    const row = table.getRow(0);
    if (row.getNumCells() !== 3) return false;

    const leftText = String(row.getCell(0).getText() || '')
      .replace(/\u00A0/g, ' ')
      .trim();

    return leftText === '';
  } catch (error) {
    return false;
  }
}

/**
 * Returns true only for physical tables that must advance Table caption
 * numbering. Equation layout tables are intentionally excluded.
 */
function isCountableCaptionTable_(element) {
  if (
    !element ||
    !element.getType ||
    element.getType() !== DocumentApp.ElementType.TABLE
  ) {
    return false;
  }

  return !isEquationLayoutTable_(element.asTable());
}

function formatSelectedTable() {
  const started = Date.now();
  const doc = DocumentApp.getActiveDocument();
  const body = getActiveBody_();
  const table = getActiveTable_();

  if (!table) {
    throw new Error(
      'Table formatting is only allowed when the cursor/selection is entirely inside one table.'
    );
  }

  const top = getTopLevelElementForParent_(table, body);
  if (!top || top.getType() !== DocumentApp.ElementType.TABLE) {
    throw new Error('Only body-level tables can be formatted.');
  }

  const bodyIndex = body.getChildIndex(top);
  const tableResult = formatTableElement_(table);
  const pinPlan = buildTableHeaderPinPlan_(body, [bodyIndex]);

  doc.saveAndClose();
  const pinResult = executeTableHeaderPinPlan_(pinPlan);

  return {
    ok: true,
    rows: tableResult.rows,
    cells: tableResult.cells,
    paragraphs: tableResult.paragraphs,
    listItems: tableResult.listItems,
    listIndentLeftInches: 0,
    listHangingInches: 0.10,
    headerRowsPinned: pinResult.pinnedTables,
    tableAlignment: 'UNCHANGED_NOT_EXPOSED_BY_GOOGLE_DOCS_API',
    usesDocsApi: true,
    elapsedMs: Date.now() - started
  };
}

/**
 * Resolves body-level table positions before DocumentApp is closed. The Docs
 * API identifies a table by its structural start index, so table ordinals are
 * used to bridge the two document models without scanning cell contents.
 */
function buildTableHeaderPinPlan_(body, bodyIndices) {
  const requested = {};
  (bodyIndices || []).forEach(function(index) {
    requested[Number(index)] = true;
  });

  const tableOrdinals = [];
  let tableOrdinal = 0;

  for (let bodyIndex = 0; bodyIndex < body.getNumChildren(); bodyIndex++) {
    const element = body.getChild(bodyIndex);
    if (element.getType() !== DocumentApp.ElementType.TABLE) continue;

    if (requested[bodyIndex] && !isEquationLayoutTable_(element.asTable())) {
      tableOrdinals.push(tableOrdinal);
    }

    tableOrdinal++;
  }

  const doc = DocumentApp.getActiveDocument();

  return {
    documentId: doc.getId(),
    tabId: getActiveDocumentTabId_(),
    tableOrdinals: tableOrdinals
  };
}

/**
 * Pins the first row of all planned tables in one Docs API batch. This is the
 * only supported public operation that makes a header row repeat across pages.
 */
function executeTableHeaderPinPlan_(plan) {
  const tableOrdinals = plan && plan.tableOrdinals
    ? plan.tableOrdinals
    : [];

  if (!tableOrdinals.length) {
    return {pinnedTables: 0, usesDocsApi: false};
  }

  if (typeof Docs === 'undefined' || !Docs.Documents) {
    throw new Error(
      'The Advanced Google Docs service is required to pin table header rows.'
    );
  }

  const apiDocument = Docs.Documents.get(plan.documentId, {
    includeTabsContent: true
  });
  const apiTab = getApiDocumentTab_(apiDocument, plan.tabId);
  const apiTables = (apiTab.body.content || []).filter(function(element) {
    return element && element.table;
  });
  const requests = [];

  tableOrdinals.forEach(function(tableOrdinal) {
    const structural = apiTables[tableOrdinal];

    if (!structural || !Number.isFinite(Number(structural.startIndex))) {
      throw new Error(
        'Could not resolve a selected table in the Google Docs API structure.'
      );
    }

    const tableStartLocation = {
      index: Number(structural.startIndex)
    };

    if (plan.tabId) {
      tableStartLocation.tabId = plan.tabId;
    }

    requests.push({
      pinTableHeaderRows: {
        tableStartLocation: tableStartLocation,
        pinnedHeaderRowsCount: 1
      }
    });
  });

  Docs.Documents.batchUpdate(
    {requests: requests},
    plan.documentId
  );

  return {
    pinnedTables: requests.length,
    usesDocsApi: true
  };
}

function getActiveDocumentTabId_() {
  try {
    const tab = DocumentApp.getActiveDocument().getActiveTab();
    return tab && tab.getId ? tab.getId() : '';
  } catch (error) {
    return '';
  }
}

function getApiDocumentTab_(apiDocument, activeTabId) {
  const found = findApiTabById_(apiDocument.tabs || [], activeTabId);
  if (found && found.documentTab) return found.documentTab;

  if (apiDocument.body) return apiDocument;

  throw new Error(
    'Could not read the active Google Docs tab through the Docs API.'
  );
}

function findApiTabById_(tabs, tabId) {
  for (let index = 0; index < (tabs || []).length; index++) {
    const tab = tabs[index];
    const properties = tab.tabProperties || {};

    if (!tabId || String(properties.tabId || '') === String(tabId)) {
      return tab;
    }

    const childMatch = findApiTabById_(tab.childTabs || [], tabId);
    if (childMatch) return childMatch;
  }

  return null;
}

function getActiveTable_() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  if (selection) {
    const ranges = selection.getRangeElements();
    let foundTable = null;

    for (let i = 0; i < ranges.length; i++) {
      const table = findAncestorTable_(ranges[i].getElement());

      if (!table) return null;

      if (!foundTable) {
        foundTable = table;
      } else if (table !== foundTable) {
        return null;
      }
    }

    return foundTable;
  }

  const cursor = doc.getCursor();
  return cursor ? findAncestorTable_(cursor.getElement()) : null;
}

function findAncestorTable_(element) {
  let current = element;

  while (current) {
    if (
      current.getType &&
      current.getType() === DocumentApp.ElementType.TABLE
    ) {
      return current.asTable();
    }

    current = current.getParent ? current.getParent() : null;
  }

  return null;
}

function buildTableParagraphAttributes_(
  isHeader,
  isListItem,
  pointsPerInch
) {
  const attribute = DocumentApp.Attribute;
  const attributes = {};

  attributes[attribute.HORIZONTAL_ALIGNMENT] = isHeader
    ? DocumentApp.HorizontalAlignment.CENTER
    : DocumentApp.HorizontalAlignment.LEFT;
  attributes[attribute.LINE_SPACING] = isListItem ? 1.5 : 1;
  attributes[attribute.SPACING_BEFORE] = 0;
  attributes[attribute.SPACING_AFTER] = 0;
  attributes[attribute.INDENT_END] = 0;

  if (isListItem) {
    // Left = 0 in; Hanging = 0.10 in.
    attributes[attribute.INDENT_FIRST_LINE] = 0;
    attributes[attribute.INDENT_START] = 0.10 * pointsPerInch;
  } else {
    // Left = 0.05 in; Special indent = None.
    const paragraphIndent = 0.05 * pointsPerInch;
    attributes[attribute.INDENT_FIRST_LINE] = paragraphIndent;
    attributes[attribute.INDENT_START] = paragraphIndent;
  }

  attributes[attribute.FONT_FAMILY] = 'Arial';
  attributes[attribute.FONT_SIZE] = 9;
  attributes[attribute.BOLD] = Boolean(isHeader);
  attributes[attribute.FOREGROUND_COLOR] = null;

  return attributes;
}


/**
 * Formats the current paragraph as a Figure/Table caption.
 * Inputs such as "Table 5.1. Overview" are treated as one old caption
 * identifier; ".1" is NOT left behind as part of the description.
 */
function formatCaptionLine(captionType) {
  captionType = String(captionType || '').toLowerCase();
  const forcedType = captionType === 'figure' ? 'Figure'
                   : captionType === 'table' ? 'Table'
                   : null;

  if (!forcedType) throw new Error('Unknown caption type.');

  const targets = getStyleTargetParagraphs_();
  if (!targets.length) {
    throw new Error('Place the cursor in the caption line or select it.');
  }
  if (targets.length !== 1) {
    throw new Error('Format one caption at a time.');
  }

  const p = targets[0];
  if (isInsideTable_(p)) {
    throw new Error('Figure/Table captions cannot be applied inside a table. Use Format selected table instead.');
  }
  if (p.getType() === DocumentApp.ElementType.LIST_ITEM) {
    throw new Error('Captions cannot be list items.');
  }
  if (forcedType === 'Figure' && paragraphContainsInlineFigure_(p)) {
    throw new Error(
      'Place the cursor in a separate caption line above the inline image. The image paragraph will not be modified.'
    );
  }

  const original = p.getText();

  // If the line already starts with Figure/Table, strip that existing prefix.
  // Otherwise treat the current line as the description, but also remove a
  // leading old numeric caption identifier such as "5.1. Overview".
  const parsed = parseCaptionLine_(original);
  const content = parsed
    ? {description: parsed.description}
    : parseCaptionDescriptionOnly_(original);

  const renumberResult = formatCaptionAndFollowing_(
    p,
    forcedType,
    content.description
  );

  return {
    ok: true,
    type: forcedType,
    number: renumberResult.number,
    text: p.getText(),
    followingCaptionsRenumbered:
      renumberResult.followingCaptionsRenumbered,
    objectsIndexed: renumberResult.objectsIndexed,
    anchorBodyIndex: renumberResult.anchorBodyIndex,
    positionPreserved: renumberResult.positionPreserved,
    elapsedMs: renumberResult.elapsedMs
  };
}

/**
 * Formats the selected caption and reenumerates only later captions of the
 * same type. The table/figure index is built once for the whole operation.
 */
function formatCaptionAndFollowing_(targetParagraph, type, description) {
  const started = Date.now();
  const body = getActiveBody_();
  const targetIndex = body.getChildIndex(targetParagraph);
  let localFigureMarker = {
    bodyIndex: -1,
    markerAdded: false
  };

  // A newly inserted positioned image or drawing is not returned by
  // Body.getImages(). Discover the Figure associated with the active caption
  // locally and persist its marker before building the complete index.
  if (type === 'Figure') {
    const associatedFigureIndex = findCaptionObjectBodyIndex_(
      'Figure',
      body,
      targetIndex
    );

    if (associatedFigureIndex >= 0) {
      localFigureMarker = ensurePersistentFigureBlockMarker_(
        body,
        body.getChild(associatedFigureIndex)
      );
    }
  }

  const requirements = {
    tableIndex: type === 'Table',
    figureIndex: type === 'Figure',
    equationIndex: false
  };
  const objectIndex = buildNeededDocumentIndex_(
    body,
    requirements,
    []
  );

  let number = getIndexedCaptionOrdinal_(
    targetParagraph,
    type,
    targetIndex,
    body,
    objectIndex
  );

  // Preserve the established manual behavior for a caption before a custom
  // anchor, while bulk updates continue to leave pre-anchor captions alone.
  if (number === null || number === undefined) {
    number = getCaptionOrdinal_(targetParagraph, type);
  }

  formatCaptionParagraph_(targetParagraph, type, description, number);

  let followingCaptionsRenumbered = 0;

  for (let bodyIndex = targetIndex + 1; bodyIndex < body.getNumChildren(); bodyIndex++) {
    const element = body.getChild(bodyIndex);
    if (element.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const paragraph = element.asParagraph();
    const parsed = parseCaptionLine_(paragraph.getText());
    if (!parsed || parsed.type !== type) continue;

    const followingNumber = getIndexedCaptionOrdinal_(
      paragraph,
      type,
      bodyIndex,
      body,
      objectIndex
    );

    if (followingNumber === null || followingNumber === undefined) continue;

    formatCaptionParagraph_(
      paragraph,
      type,
      parsed.description,
      followingNumber
    );
    followingCaptionsRenumbered++;
  }

  const finalTargetIndex = body.getChildIndex(targetParagraph);
  const objectsIndexed = type === 'Table'
    ? objectIndex.tableIndices.length
    : objectIndex.figureIndices.length;

  return {
    number: number,
    followingCaptionsRenumbered: followingCaptionsRenumbered,
    objectsIndexed: objectsIndexed,
    anchorBodyIndex: getCaptionCounterAnchorIndex_(type, body),
    localFigureMarkerAdded: localFigureMarker.markerAdded,
    positionPreserved: finalTargetIndex === targetIndex,
    originalBodyIndex: targetIndex,
    finalBodyIndex: finalTargetIndex,
    elapsedMs: Date.now() - started
  };
}

function parseCaptionDescriptionOnly_(value) {
  let remainder = String(value || '').trim();

  // When the caption keyword is missing, allow old identifiers such as:
  // "5.1. Overview" -> "Overview"
  // "3. Process diagram" -> "Process diagram"
  // "X. Overview" -> "Overview"
  const numberMatch = remainder.match(
    /^((?:X)|(?:\d+(?:\s*\.\s*\d+)*))(?=\s|[.:–—-]|$)/i
  );

  if (numberMatch) {
    remainder = remainder.substring(numberMatch[0].length);
    remainder = remainder.replace(/^\s*[.:–—-]?\s*/, '');
  }

  return {description: remainder.trim()};
}


function formatNoteLine() {
  const targets = getStyleTargetParagraphs_();
  if (!targets.length) {
    throw new Error('Place the cursor in the note line or select it.');
  }
  if (targets.length !== 1) {
    throw new Error('Format one note at a time.');
  }

  const p = targets[0];
  if (isInsideTable_(p)) {
    throw new Error('Note formatting cannot be applied inside a table. Use Format selected table instead.');
  }
  if (p.getType() === DocumentApp.ElementType.LIST_ITEM) {
    throw new Error('Notes cannot be list items.');
  }

  const parsed = parseNoteLine_(p.getText());
  formatNoteParagraph_(p, parsed.description);

  return {ok:true, text:p.getText()};
}

function parseNoteLine_(value) {
  const original = String(value || '').trim();
  let source = original;

  // Detect an existing Note/Notes/Nota/Notas marker after optional opening
  // wrappers. This intentionally handles cases such as:
  //   (Note: Abcdefg...)
  //   ((NOTA - Abcdefg...))
  //   [Notes. Abcdefg...]
  //   Note (Abcdefg...)
  const prefix = source.match(
    /^\s*[\(\[\{]*\s*(?:Notes?|Notas?)\b\s*[\)\]\}]*\s*[.:–—-]?\s*/i
  );

  // No existing Note/Nota marker: do NOT clean or reinterpret the line.
  // The formatter will simply prepend "Note. ".
  if (!prefix) {
    return {
      description: original,
      hadNotePrefix: false
    };
  }

  source = source.substring(prefix[0].length).trim();

  // If the existing note was wrapped, remove trailing closing wrappers
  // without deleting the description's punctuation.
  source = source.replace(/\s*[\)\]\}]+\s*$/, '').trim();

  // A second wrapper around only the description is also normalized.
  while (
    source.length >= 2 &&
    (
      (source.startsWith('(') && source.endsWith(')')) ||
      (source.startsWith('[') && source.endsWith(']')) ||
      (source.startsWith('{') && source.endsWith('}'))
    )
  ) {
    source = source.substring(1, source.length - 1).trim();
  }

  return {
    description: source,
    hadNotePrefix: true
  };
}

function formatNoteParagraph_(p, description) {
  if (isInsideTable_(p)) {
    throw new Error('Note formatting is not allowed inside a table.');
  }
  const noteText = 'Note. ' + String(description || '').trim();

  applyNamedStyleToParagraph_(p, 'NORMAL');

  const t = p.editAsText();
  t.setText(noteText);
  t.setFontFamily('Arial');
  t.setFontSize(9);
  t.setBold(false);
  resetTextColorToAutomatic_(t);

  const prefix = 'Note.';
  t.setBold(0, prefix.length - 1, true);

  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function parseCaptionLine_(value) {
  const source = String(value || '');

  // Caption keyword is case-insensitive:
  // FIGURE..., Figure..., figure..., TABLE..., etc.
  const head = source.match(/^\s*(Figure|Table)\b\s*(.*)$/i);
  if (!head) return null;

  const type = head[1].toLowerCase() === 'figure' ? 'Figure' : 'Table';
  let remainder = String(head[2] || '').trim();

  // Allow punctuation immediately after the keyword:
  // "FIGURE: Process flow"
  remainder = remainder.replace(/^[\s:–—-]+/, '');

  let oldNumber = '';

  // Consume the COMPLETE old numbering prefix in one operation.
  // Spaces around decimal dots are accepted intentionally, so even a
  // previously malformed value such as "Table 54. 1. Overview" is read
  // as one old identifier and becomes "Table N. Overview".
  //
  // Examples consumed:
  // X
  // 5
  // 5.1
  // 5.1.3
  // 54. 1
  const numberMatch = remainder.match(
    /^((?:X)|(?:\d+(?:\s*\.\s*\d+)*))(?=\s|[.:–—-]|$)/i
  );

  if (numberMatch) {
    oldNumber = numberMatch[1].replace(/\s+/g, '');
    remainder = remainder.substring(numberMatch[0].length);

    // Remove the punctuation that terminated the old identifier.
    // Handles both "5.1. Overview" and "5.1 Overview".
    remainder = remainder.replace(/^\s*[.:–—-]?\s*/, '');
  }

  // If no old number exists, the line is still accepted:
  // "FIGURE Process flow" -> "Figure N. Process flow"
  return {
    type: type,
    oldNumber: oldNumber,
    description: remainder.trim()
  };
}

function formatCaptionParagraph_(p, type, description, number) {
  if (isInsideTable_(p)) {
    throw new Error('Caption formatting is not allowed inside a table.');
  }
  if (type === 'Figure' && paragraphContainsInlineFigure_(p)) {
    throw new Error(
      'A Figure caption must use a separate paragraph above the inline image.'
    );
  }
  const captionText = type + ' ' + number + '. ' + String(description || '').trim();
  const verticalLayout = type === 'Figure'
    ? captureCaptionVerticalLayout_(p)
    : null;

  // Start from CURRENT Normal text style, then apply caption overrides.
  applyNamedStyleToParagraph_(p, 'NORMAL');

  const t = p.editAsText();
  t.setText(captionText);
  t.setFontFamily('Arial');
  t.setFontSize(9);
  t.setBold(false);
  t.setItalic(false);
  resetTextColorToAutomatic_(t);

  // One continuous rich-text range produces **Table 1**, not two adjacent
  // runs such as **Table** **1** when exported to Markdown-like notation.
  const prefix = type + ' ' + number;
  if (prefix.length) {
    t.setBold(0, prefix.length - 1, true);
  }

  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  restoreCaptionVerticalLayout_(p, verticalLayout);
}

/**
 * Caption formatting must not change the paragraph's vertical placement
 * relative to its image/table. Named Style reapplication can otherwise alter
 * spacing and create the impression that the caption moved.
 */
function captureCaptionVerticalLayout_(paragraph) {
  const layout = {};

  try { layout.spacingBefore = paragraph.getSpacingBefore(); } catch (error) {}
  try { layout.spacingAfter = paragraph.getSpacingAfter(); } catch (error) {}
  try { layout.lineSpacing = paragraph.getLineSpacing(); } catch (error) {}

  return layout;
}

function restoreCaptionVerticalLayout_(paragraph, layout) {
  if (!layout) return;

  if (Number.isFinite(layout.spacingBefore)) {
    paragraph.setSpacingBefore(layout.spacingBefore);
  }
  if (Number.isFinite(layout.spacingAfter)) {
    paragraph.setSpacingAfter(layout.spacingAfter);
  }
  if (Number.isFinite(layout.lineSpacing) && layout.lineSpacing > 0) {
    paragraph.setLineSpacing(layout.lineSpacing);
  }
}

function getCaptionOrdinal_(targetParagraph, type) {
  if (type === 'Table') {
    return getTableCaptionOrdinal_(targetParagraph);
  }

  if (type === 'Figure') {
    return getFigureCaptionOrdinal_(targetParagraph);
  }

  return getCaptionOrdinalByPreviousCaptions_(targetParagraph, type);
}

const CAPTION_COUNTER_CONFIG_ = {
  Table: {
    anchorName: 'DOCS_ASSISTANT_TABLE_COUNTER_ANCHOR',
    startProperty: 'DOCS_ASSISTANT_TABLE_COUNTER_START'
  },
  Figure: {
    anchorName: 'DOCS_ASSISTANT_FIGURE_COUNTER_ANCHOR',
    startProperty: 'DOCS_ASSISTANT_FIGURE_COUNTER_START'
  }
};

function normalizeCaptionCounterType_(type) {
  const value = String(type || '').toLowerCase();
  if (value === 'table') return 'Table';
  if (value === 'figure') return 'Figure';
  throw new Error('Counter type must be Table or Figure.');
}

function getCaptionCounterStart_(type) {
  const normalized = normalizeCaptionCounterType_(type);
  const config = CAPTION_COUNTER_CONFIG_[normalized];
  const props = PropertiesService.getDocumentProperties();
  const raw = props ? props.getProperty(config.startProperty) : null;
  const value = parseInt(raw || '1', 10);
  return Number.isFinite(value) && value >= 1 ? value : 1;
}

function getCaptionNumberingSettings() {
  return {
    table: getCaptionCounterSettings_('Table'),
    figure: getCaptionCounterSettings_('Figure')
  };
}

function getCaptionCounterSettings_(type) {
  const normalized = normalizeCaptionCounterType_(type);
  const body = getActiveBody_();
  const anchorIndex = getCaptionCounterAnchorIndex_(normalized, body);

  return {
    startAt: getCaptionCounterStart_(normalized),
    anchorSet: anchorIndex >= 0,
    anchorIndex: anchorIndex
  };
}

function setCaptionCounterAnchor(type, startAt) {
  const normalized = normalizeCaptionCounterType_(type);
  const start = parseInt(startAt, 10);

  if (!Number.isFinite(start) || start < 1) {
    throw new Error('Start at must be an integer of 1 or greater.');
  }

  const doc = DocumentApp.getActiveDocument();
  const body = getActiveBody_();
  const reference = getCurrentReferenceElement_();

  if (!reference) {
    throw new Error('Place the cursor on the first caption line you want to number.');
  }

  const top = getTopLevelElementForParent_(reference, body);
  if (!top) {
    throw new Error('The cursor must be in the main document body.');
  }

  const topType = top.getType();
  const isCaptionParagraph =
    topType === DocumentApp.ElementType.PARAGRAPH ||
    topType === DocumentApp.ElementType.LIST_ITEM;
  const isTableObject =
    normalized === 'Table' &&
    topType === DocumentApp.ElementType.TABLE &&
    isCountableCaptionTable_(top);
  const isFigureObject =
    normalized === 'Figure' &&
    isStandaloneFigureBlock_(top);

  if (!isCaptionParagraph && !isTableObject && !isFigureObject) {
    throw new Error(
      normalized === 'Table'
        ? 'Place the cursor on the Table caption or inside the table, then press Set here.'
        : 'Place the cursor on the Figure caption or image paragraph, then press Set here.'
    );
  }

  const config = CAPTION_COUNTER_CONFIG_[normalized];
  removeNamedRangesByName_(config.anchorName);

  // A NamedRange is most reliable on text. When the cursor is inside a table,
  // anchor its owning cell paragraph; resolving the range later returns the
  // body-level table index.
  const rangeElement = isTableObject
    ? (getOwningParagraph_(reference) || reference)
    : top;
  addActiveTabNamedRange_(config.anchorName, rangeElement);

  const props = PropertiesService.getDocumentProperties();
  if (props) props.setProperty(config.startProperty, String(start));

  const anchorBodyIndex = body.getChildIndex(top);
  const anchorObjectBodyIndex = findCaptionObjectBodyIndex_(
    normalized,
    body,
    anchorBodyIndex
  );
  let figureMarkerAdded = false;

  // Direct inline images are collected on every run. This marker additionally
  // covers a newly inserted positioned image or InlineDrawing at the anchor.
  if (normalized === 'Figure' && anchorObjectBodyIndex >= 0) {
    figureMarkerAdded = ensurePersistentFigureBlockMarker_(
      body,
      body.getChild(anchorObjectBodyIndex)
    ).markerAdded;
  }

  return {
    type: normalized,
    startAt: start,
    anchorSet: true,
    anchorBodyIndex: anchorBodyIndex,
    anchorObjectBodyIndex: anchorObjectBodyIndex,
    figureMarkerAdded: figureMarkerAdded,
    anchorTarget: isTableObject
      ? 'TABLE'
      : isFigureObject
        ? 'FIGURE'
        : 'CAPTION'
  };
}

function clearCaptionCounterAnchor(type) {
  const normalized = normalizeCaptionCounterType_(type);
  const config = CAPTION_COUNTER_CONFIG_[normalized];

  removeNamedRangesByName_(config.anchorName);

  const props = PropertiesService.getDocumentProperties();
  if (props) props.deleteProperty(config.startProperty);

  return {
    type: normalized,
    startAt: 1,
    anchorSet: false
  };
}

function getActiveTabNamedRanges_(name) {
  const doc = DocumentApp.getActiveDocument();
  const tab = getActiveDocumentTab_();

  try {
    return tab ? (tab.getNamedRanges(name) || []) : (doc.getNamedRanges(name) || []);
  } catch (e) {
    return [];
  }
}

function addActiveTabNamedRange_(name, element) {
  const doc = DocumentApp.getActiveDocument();
  const tab = getActiveDocumentTab_();

  if (tab) {
    const range = tab.newRange().addElement(element).build();
    return tab.addNamedRange(name, range);
  }

  const range = doc.newRange().addElement(element).build();
  return doc.addNamedRange(name, range);
}

function removeNamedRangesByName_(name) {
  const named = getActiveTabNamedRanges_(name);

  named.forEach(n => {
    try { n.remove(); } catch (e) {}
  });
}

function getCaptionCounterAnchorIndex_(type, parent) {
  const normalized = normalizeCaptionCounterType_(type);
  const config = CAPTION_COUNTER_CONFIG_[normalized];
  const named = getActiveTabNamedRanges_(config.anchorName);

  if (!named.length) return -1;

  for (let n = 0; n < named.length; n++) {
    try {
      const rangeElements = named[n].getRange().getRangeElements();

      for (let i = 0; i < rangeElements.length; i++) {
        const top = getTopLevelElementForParent_(rangeElements[i].getElement(), parent);
        if (!top) continue;

        return parent.getChildIndex(top);
      }
    } catch (e) {}
  }

  return -1;
}

/**
 * Resolve the stored caption-line anchor to the ACTUAL first Table/Figure
 * object that belongs to that caption. This is done at count time, so setup
 * does not fail just because the object is not "near enough" in the DOM.
 */
function getCaptionCounterAnchorObjectIndex_(type, parent) {
  const normalized = normalizeCaptionCounterType_(type);
  const anchorLineIndex = getCaptionCounterAnchorIndex_(normalized, parent);
  if (anchorLineIndex < 0) return -1;

  return findCaptionObjectBodyIndex_(
    normalized,
    parent,
    anchorLineIndex
  );
}

function findCaptionObjectBodyIndex_(type, parent, referenceIndex) {
  const normalized = normalizeCaptionCounterType_(type);

  // Both Table and Figure captions belong above their object. Equality is
  // accepted for an anchor set inside a table or on a positioned image.
  for (let i = referenceIndex; i < parent.getNumChildren(); i++) {
    const child = parent.getChild(i);
    const matches =
      normalized === 'Table'
        ? isCountableCaptionTable_(child)
        : isStandaloneFigureBlock_(child);

    if (matches) return i;
  }

  // Fallback for legacy captions that remain below their object.
  for (let i = referenceIndex - 1; i >= 0; i--) {
    const child = parent.getChild(i);
    const matches =
      normalized === 'Table'
        ? isCountableCaptionTable_(child)
        : isStandaloneFigureBlock_(child);

    if (matches) return i;
  }

  return -1;
}

function getTopLevelElementForParent_(element, parent) {
  let current = element;

  while (current) {
    let currentParent = null;

    try {
      currentParent = current.getParent();
    } catch (e) {
      return null;
    }

    if (!currentParent) return null;

    // Robust body detection:
    // do not compare Apps Script element wrapper objects with ===.
    // Instead detect the actual Body container by its ElementType.
    try {
      if (currentParent.getType() === DocumentApp.ElementType.BODY_SECTION) {
        // If a specific active body was supplied, confirm that this top-level
        // element really belongs to it using getChildIndex(), not object identity.
        if (parent && parent.getChildIndex) {
          try {
            parent.getChildIndex(current);
          } catch (e) {
            return null;
          }
        }

        return current;
      }
    } catch (e) {}

    current = currentParent;
  }

  return null;
}

function getCurrentReferenceElement_() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  if (selection) {
    const ranges = selection.getRangeElements();
    if (ranges.length) return ranges[0].getElement();
  }

  const cursor = doc.getCursor();
  if (cursor) return cursor.getElement();

  return null;
}


function getAnchoredObjectOrdinal_(type, parent, objectIndex) {
  const normalized = normalizeCaptionCounterType_(type);
  const anchorLineIndex = getCaptionCounterAnchorIndex_(normalized, parent);

  // No custom anchor: legacy behavior, count from document start.
  if (anchorLineIndex < 0) {
    let count = 0;

    for (let i = 0; i <= objectIndex; i++) {
      const child = parent.getChild(i);
      const matches =
        normalized === 'Table'
          ? isCountableCaptionTable_(child)
          : isStandaloneFigureBlock_(child);

      if (matches) count++;
    }

    return Math.max(1, count);
  }

  const anchorObjectIndex = getCaptionCounterAnchorObjectIndex_(normalized, parent);
  if (anchorObjectIndex < 0) return null;

  // Anything physically before the first anchored object is outside
  // the custom numbering sequence.
  if (objectIndex < anchorObjectIndex) return null;

  let relativeCount = 0;

  for (let i = anchorObjectIndex; i <= objectIndex; i++) {
    const child = parent.getChild(i);
    const matches =
      normalized === 'Table'
        ? isCountableCaptionTable_(child)
        : isStandaloneFigureBlock_(child);

    if (matches) relativeCount++;
  }

  if (relativeCount < 1) return null;

  return getCaptionCounterStart_(normalized) + relativeCount - 1;
}

function renumberAllCaptions() {
  const started = Date.now();
  const body = getActiveBody_();
  const objectIndex = buildNeededDocumentIndex_(
    body,
    {
      tableIndex: true,
      figureIndex: true,
      equationIndex: false
    },
    []
  );
  let tables = 0;
  let figures = 0;
  let skippedBeforeAnchor = 0;

  // Body-level captions only; captions inside table cells remain protected.
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);

    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const p = child.asParagraph();
    const parsed = parseCaptionLine_(p.getText());

    if (!parsed || (parsed.type !== 'Table' && parsed.type !== 'Figure')) continue;

    const ordinal = getIndexedCaptionOrdinal_(
      p,
      parsed.type,
      i,
      body,
      objectIndex
    );

    if (ordinal === null || ordinal === undefined) {
      skippedBeforeAnchor++;
      continue;
    }

    formatCaptionParagraph_(p, parsed.type, parsed.description, ordinal);

    if (parsed.type === 'Table') tables++;
    else figures++;
  }

  return {
    tables: tables,
    figures: figures,
    skippedBeforeAnchor: skippedBeforeAnchor,
    tableIndexMs: objectIndex.tableIndexMs,
    figureIndexMs: objectIndex.figureIndexMs,
    elapsedMs: Date.now() - started
  };
}


/**
 * Table numbering is based on ACTUAL Google Docs table elements, not on
 * the number already written in caption text.
 *
 * The caption belongs to the first actual TABLE element at or after it. A
 * previous-table fallback is retained for legacy captions that remain below.
 *
 * If no actual table can be found, we safely fall back to counting prior
 * Table caption paragraphs.
 */
function getTableCaptionOrdinal_(targetParagraph) {
  const parent = targetParagraph.getParent();
  if (!parent || !parent.getChildIndex) {
    return getCaptionOrdinalByPreviousCaptions_(targetParagraph, 'Table');
  }

  let targetIndex;
  try {
    targetIndex = parent.getChildIndex(targetParagraph);
  } catch (e) {
    return getCaptionOrdinalByPreviousCaptions_(targetParagraph, 'Table');
  }

  let nearestTableIndex = -1;

  for (let i = targetIndex; i < parent.getNumChildren(); i++) {
    if (isCountableCaptionTable_(parent.getChild(i))) {
      nearestTableIndex = i;
      break;
    }
  }

  if (nearestTableIndex < 0) {
    for (let i = targetIndex - 1; i >= 0; i--) {
      if (isCountableCaptionTable_(parent.getChild(i))) {
        nearestTableIndex = i;
        break;
      }
    }
  }

  if (nearestTableIndex < 0) {
    return getCaptionOrdinalByPreviousCaptions_(targetParagraph, 'Table');
  }

  const anchored = getAnchoredObjectOrdinal_('Table', parent, nearestTableIndex);

  // A direct click on a caption before the custom anchor still receives the
  // legacy document-start number. Bulk Renumber leaves pre-anchor captions
  // untouched, but manual formatting remains predictable.
  if (anchored === null) {
    let count = 0;
    for (let i = 0; i <= nearestTableIndex; i++) {
      if (isCountableCaptionTable_(parent.getChild(i))) count++;
    }
    return Math.max(1, count);
  }

  return anchored;
}

/**
 * Figure numbering is based on actual visual objects in the document body.
 *
 * A "figure block" is a body-level Paragraph/ListItem containing one or more:
 * - InlineImage
 * - InlineDrawing
 * - PositionedImage anchored to the paragraph
 *
 * Multiple visual objects in the same paragraph are treated as ONE figure
 * block, which is useful for composite figures.
 *
 * Images/drawings inside tables are intentionally not counted here because
 * table-cell graphics should not normally advance standalone Figure numbering.
 */
function getFigureCaptionOrdinal_(targetParagraph) {
  const parent = targetParagraph.getParent();
  if (!parent || !parent.getChildIndex) {
    return getCaptionOrdinalByPreviousCaptions_(targetParagraph, 'Figure');
  }

  let targetIndex;
  try {
    targetIndex = parent.getChildIndex(targetParagraph);
  } catch (e) {
    return getCaptionOrdinalByPreviousCaptions_(targetParagraph, 'Figure');
  }

  let nearestFigureIndex = -1;

  // Figure captions are above their image. Include the same paragraph for a
  // positioned image anchored to the caption line.
  for (let i = targetIndex; i < parent.getNumChildren(); i++) {
    if (isStandaloneFigureBlock_(parent.getChild(i))) {
      nearestFigureIndex = i;
      break;
    }
  }

  // Fallback for legacy captions that remain below their image.
  if (nearestFigureIndex < 0) {
    for (let i = targetIndex - 1; i >= 0; i--) {
      if (isStandaloneFigureBlock_(parent.getChild(i))) {
        nearestFigureIndex = i;
        break;
      }
    }
  }

  if (nearestFigureIndex < 0) {
    return getCaptionOrdinalByPreviousCaptions_(targetParagraph, 'Figure');
  }

  const anchored = getAnchoredObjectOrdinal_('Figure', parent, nearestFigureIndex);

  if (anchored === null) {
    let count = 0;
    for (let i = 0; i <= nearestFigureIndex; i++) {
      if (isStandaloneFigureBlock_(parent.getChild(i))) count++;
    }
    return Math.max(1, count);
  }

  return anchored;
}

function isStandaloneFigureBlock_(element) {
  if (!element || !element.getType) return false;

  const type = element.getType();
  if (
    type !== DocumentApp.ElementType.PARAGRAPH &&
    type !== DocumentApp.ElementType.LIST_ITEM
  ) {
    return false;
  }

  // Inline images / drawings.
  try {
    const count = element.getNumChildren();
    for (let i = 0; i < count; i++) {
      const childType = element.getChild(i).getType();

      if (
        childType === DocumentApp.ElementType.INLINE_IMAGE ||
        childType === DocumentApp.ElementType.INLINE_DRAWING
      ) {
        return true;
      }
    }
  } catch (e) {}

  // Positioned images are not child Elements; they are anchored to a paragraph.
  try {
    const positioned = element.getPositionedImages();
    if (positioned && positioned.length > 0) return true;
  } catch (e) {}

  return false;
}

/**
 * Caption fallback when no actual Table/Figure object can be associated.
 *
 * IMPORTANT: compare by child index, not JavaScript object identity.
 * Apps Script can return different wrapper objects for the same document
 * element, so `child === targetParagraph` is not a reliable position test.
 */
function getCaptionOrdinalByPreviousCaptions_(targetParagraph, type) {
  const parent = targetParagraph.getParent();
  if (!parent || !parent.getChildIndex) return 1;

  let targetIndex;
  try {
    targetIndex = parent.getChildIndex(targetParagraph);
  } catch (e) {
    return 1;
  }

  let count = 0;

  for (let i = 0; i < targetIndex; i++) {
    const child = parent.getChild(i);

    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const parsed = parseCaptionLine_(child.asParagraph().getText());
    if (parsed && parsed.type === type) count++;
  }

  return count + 1;
}

function getCurrentListType_(p) {
  if (p.getType() !== DocumentApp.ElementType.LIST_ITEM) return '';

  try {
    const glyph = String(p.asListItem().getGlyphType()).toUpperCase();

    if (glyph.indexOf('ROMAN') >= 0) return 'roman';
    if (glyph.indexOf('LATIN') >= 0 || glyph.indexOf('LETTER') >= 0) return 'letter';
    if (glyph.indexOf('NUMBER') >= 0 || glyph.indexOf('DECIMAL') >= 0) return 'number';
    return 'bullet';
  } catch (e) {
    return 'bullet';
  }
}

function stripManualListPrefix_(p, type) {
  // Native ListItems do not contain their rendered glyph in getText(),
  // so no prefix should be stripped from them.
  if (p.getType() === DocumentApp.ElementType.LIST_ITEM) return;

  const original = p.getText();
  let cleaned = original;

  switch (type) {
    case 'BULLET':
      cleaned = original.replace(/^\s*[•●○▪◦‣⁃-]\s+/, '');
      break;
    case 'NUMBER':
      cleaned = original.replace(/^\s*\d+[.)]\s+/, '');
      break;
    case 'LETTER':
      cleaned = original.replace(/^\s*[A-Za-zÁÉÍÓÚÜÑáéíóúüñ][.)]\s+/, '');
      break;
    case 'ROMAN':
      cleaned = original.replace(/^\s*[ivxlcdm]+[.)]\s+/i, '');
      break;
  }

  if (cleaned !== original) {
    p.editAsText().setText(cleaned);
  }
}

function isInsideTable_(el) {
  let current = el;
  while (current) {
    if (current.getType && current.getType() === DocumentApp.ElementType.TABLE_CELL) {
      return true;
    }
    current = current.getParent ? current.getParent() : null;
  }
  return false;
}
