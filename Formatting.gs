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

  return paragraph;
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
  const started = Date.now();

  const document =
    DocumentApp.getActiveDocument();

  const selection = document.getSelection();

  if (!selection) {
    throw new Error(
      'Select the paragraphs you want to format.'
    );
  }

  /*
   * ALL permite incluir encabezados, texto normal,
   * ListItems y párrafos vacíos.
   */
  const segments = getSegments_(
    selection,
    'ALL'
  );

  let formatted = 0;
  let headings = 0;
  let normal = 0;
  let blank = 0;
  let skipped = 0;

  segments.forEach(function (segment) {
    const paragraph = segment.element;

    if (
      !paragraph ||
      typeof paragraph.getType !== 'function'
    ) {
      skipped++;
      return;
    }

    const elementType = paragraph.getType();

    const isParagraph =
      elementType ===
        DocumentApp.ElementType.PARAGRAPH ||
      elementType ===
        DocumentApp.ElementType.LIST_ITEM;

    if (!isParagraph) {
      skipped++;
      return;
    }

    const text = String(
      paragraph.getText() || ''
    );

    const isBlank = !text.trim();

    let styleName = 'NORMAL';

    /*
     * Los párrafos vacíos siempre reciben NORMAL.
     * Los demás conservan su estilo nombrado actual.
     */
    if (!isBlank) {
      styleName =
        getExistingNamedStyleName_(paragraph);
    }

    applyStyleToParagraph_(
      paragraph,
      styleName
    );

    formatted++;

    if (isBlank) {
      blank++;
    } else if (
      /^H[1-6]$/.test(styleName) ||
      styleName === 'TITLE' ||
      styleName === 'SUBTITLE'
    ) {
      headings++;
    } else {
      normal++;
    }
  });

  if (!formatted) {
    throw new Error(
      'The selection contains no paragraphs to format.'
    );
  }

  document.saveAndClose();

  return {
    ok: true,
    paragraphs: formatted,
    headings: headings,
    normal: normal,
    blank: blank,
    skipped: skipped,
    elapsedMs: Date.now() - started
  };
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

  if (paragraphContainsFigure_(paragraph)) return 'FIGURE';
  if (paragraphContainsEquation_(paragraph)) return 'EQUATION';

  if (/^\s*figure\b/i.test(trimmed)) return 'FIGURE_CAPTION';
  if (/^\s*table\b/i.test(trimmed)) return 'TABLE_CAPTION';
  if (/^\s*(?:notes?|notas?)\b/i.test(trimmed)) return 'NOTE';

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

  try {
    const positioned = paragraph.getPositionedImages();
    if (positioned && positioned.length) return true;
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
  const target = getCurrentParagraph_() || getFirstSelectedParagraph_();

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

  paragraphs.forEach(p => {
    if (mode === 'HANGING') {
      // Google Docs geometry:
      // first line starts at Left;
      // wrapped lines start at Left + Hanging.
      p.setIndentFirstLine(leftPt);
      p.setIndentStart(leftPt + byPt);
    } else if (mode === 'FIRST_LINE') {
      // Wrapped lines start at Left;
      // first line starts at Left + By.
      p.setIndentStart(leftPt);
      p.setIndentFirstLine(leftPt + byPt);
    } else {
      p.setIndentStart(leftPt);
      p.setIndentFirstLine(leftPt);
    }

    p.setIndentEnd(rightPt);
  });

  return {
    ok: true,
    paragraphs: paragraphs.length,
    left: leftIn,
    right: rightIn,
    special: mode,
    by: byIn
  };
}

function getFirstSelectedParagraph_() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) return null;

  const ranges = selection.getRangeElements();
  if (!ranges.length) return null;

  let el = ranges[0].getElement();

  while (
    el &&
    el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
    el.getType() !== DocumentApp.ElementType.LIST_ITEM
  ) {
    el = el.getParent();
  }

  return el || null;
}

function roundIndentValue_(value) {
  return Math.round(Number(value || 0) * 1000) / 1000;
}

function setParagraphSpacing(before, after, lineSpacing) {
  eachSelectedParagraph_(p => {
    if (before !== null && before !== undefined) p.setSpacingBefore(Number(before));
    if (after !== null && after !== undefined) p.setSpacingAfter(Number(after));
    if (lineSpacing) p.setLineSpacing(Number(lineSpacing));
  });
  return true;
}

function setKeepWithNext(value) {
  eachSelectedParagraph_(p => p.setKeepWithNext(Boolean(value)));
  return true;
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
    .setIndentStart(36)
    .setIndentFirstLine(18)
    .setIndentEnd(0);
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
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  if (!cursor) throw new Error('Place the cursor where the break should be inserted.');

  if (kind === 'PAGE') {
    cursor.insertPageBreak();
    return true;
  }

  throw new Error('Section breaks require the advanced Google Docs API. The UI is prepared; implementation is the next step.');
}

function getSelectedParagraphs_() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) return [];

  const paragraphs = [];

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

    // Compare element objects directly instead of String(el).
    // String(el) is not a reliable unique identifier and could cause
    // multiple selected paragraphs to be treated as the same paragraph.
    if (el && paragraphs.indexOf(el) === -1) {
      paragraphs.push(el);
    }
  });

  return paragraphs;
}

function eachSelectedParagraph_(fn) {
  const paragraphs = getStyleTargetParagraphs_();
  if (!paragraphs.length) {
    throw new Error('Place the cursor in a paragraph or select one or more paragraphs.');
  }
  paragraphs.forEach(fn);
}


function formatEquationLine() {
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

  // Symmetric side columns keep the center cell geometrically centered.
  const usableWidth = Math.max(
    360,
    Number(body.getPageWidth()) - Number(body.getMarginLeft()) - Number(body.getMarginRight())
  );
  const sideWidth = usableWidth * 0.22;
  const centerWidth = usableWidth * 0.56;

  table.setColumnWidth(0, sideWidth);
  table.setColumnWidth(1, centerWidth);
  table.setColumnWidth(2, sideWidth);

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
  const label = '.................... Equation ' + equationNumber;
  rightCell.setText(label);

  const rightParagraph = rightCell.getChild(0).asParagraph();
  rightParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  rightParagraph.setIndentStart(0);
  rightParagraph.setIndentEnd(0);
  rightParagraph.setIndentFirstLine(0);
  rightParagraph.setSpacingBefore(0);
  rightParagraph.setSpacingAfter(0);

  const rightText = rightParagraph.editAsText();
  rightText.setFontFamily('Arial').setFontSize(9).setBold(false).setItalic(false);

  const labelStart = label.indexOf('Equation ');
  if (labelStart >= 0) {
    rightText.setBold(labelStart, label.length - 1, true);
    rightText.setItalic(labelStart, label.length - 1, true);
  }

  sourceParagraph.removeFromParent();

  return {
    ok: true,
    equationNumber: equationNumber
  };
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

function formatSelectedTable() {
  const started = Date.now();
  const table = getActiveTable_();

  if (!table) {
    throw new Error(
      'Table formatting is only allowed when the cursor/selection is entirely inside one table.'
    );
  }

  const PT_PER_IN = 72;
  const attribute = DocumentApp.Attribute;

  const cellAttributes = {};
  cellAttributes[attribute.VERTICAL_ALIGNMENT] =
    DocumentApp.VerticalAlignment.CENTER;
  cellAttributes[attribute.PADDING_TOP] = 0.028 * PT_PER_IN;
  cellAttributes[attribute.PADDING_BOTTOM] = 0.028 * PT_PER_IN;
  cellAttributes[attribute.PADDING_LEFT] = 0.028 * PT_PER_IN;
  cellAttributes[attribute.PADDING_RIGHT] = 0.028 * PT_PER_IN;

  const headerParagraphAttributes =
    buildTableParagraphAttributes_(true, false, PT_PER_IN);
  const bodyParagraphAttributes =
    buildTableParagraphAttributes_(false, false, PT_PER_IN);
  const headerListAttributes =
    buildTableParagraphAttributes_(true, true, PT_PER_IN);
  const bodyListAttributes =
    buildTableParagraphAttributes_(false, true, PT_PER_IN);

  table.setBorderColor('#000000');
  table.setBorderWidth(1);

  const rows = table.getNumRows();
  let cells = 0;
  let paragraphs = 0;
  let listItems = 0;

  for (let rowIndex = 0; rowIndex < rows; rowIndex++) {
    const row = table.getRow(rowIndex);
    const isHeader = rowIndex === 0;

    row.setMinimumHeight(0.49 * PT_PER_IN);

    const paragraphAttributes = isHeader
      ? headerParagraphAttributes
      : bodyParagraphAttributes;
    const listAttributes = isHeader
      ? headerListAttributes
      : bodyListAttributes;

    const cellCount = row.getNumCells();
    cells += cellCount;

    for (let cellIndex = 0; cellIndex < cellCount; cellIndex++) {
      const cell = row.getCell(cellIndex);

      cell.setAttributes(cellAttributes);
      cell.setBackgroundColor(null);

      for (
        let childIndex = 0;
        childIndex < cell.getNumChildren();
        childIndex++
      ) {
        const child = cell.getChild(childIndex);
        const type = child.getType();

        if (type === DocumentApp.ElementType.PARAGRAPH) {
          child.asParagraph().setAttributes(paragraphAttributes);
          paragraphs++;
        } else if (type === DocumentApp.ElementType.LIST_ITEM) {
          child.asListItem().setAttributes(listAttributes);
          listItems++;
        }
      }
    }
  }

  return {
    ok: true,
    rows: rows,
    cells: cells,
    paragraphs: paragraphs,
    listItems: listItems,
    listIndentLeftInches: 0,
    listHangingInches: 0.10,
    usesDocsApi: false,
    elapsedMs: Date.now() - started
  };
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
  attributes[attribute.LINE_SPACING] = 1;
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

  const original = p.getText();

  // If the line already starts with Figure/Table, strip that existing prefix.
  // Otherwise treat the current line as the description, but also remove a
  // leading old numeric caption identifier such as "5.1. Overview".
  const parsed = parseCaptionLine_(original);
  const content = parsed
    ? {description: parsed.description}
    : parseCaptionDescriptionOnly_(original);

  const nextNumber = getCaptionOrdinal_(p, forcedType);
  formatCaptionParagraph_(p, forcedType, content.description, nextNumber);

  return {
    ok: true,
    type: forcedType,
    number: nextNumber,
    text: p.getText()
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
  const captionText = type + ' ' + number + '. ' + String(description || '').trim();

  // Start from CURRENT Normal text style, then apply caption overrides.
  applyNamedStyleToParagraph_(p, 'NORMAL');

  const t = p.editAsText();
  t.setText(captionText);
  t.setFontFamily('Arial');
  t.setFontSize(9);
  t.setBold(false);

  const prefix = type + ' ' + number + '.';
  if (prefix.length) {
    t.setBold(0, prefix.length - 1, true);
  }

  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
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

  // The anchor is the CURRENT BODY-LEVEL PARAGRAPH/LIST ITEM.
  // We deliberately do not require the physical table/image to be found
  // at setup time. That was too brittle for real Docs layouts.
  const top = getTopLevelElementForParent_(reference, body);
  if (!top) {
    throw new Error('The cursor must be in the main document body.');
  }

  const topType = top.getType();
  if (
    topType !== DocumentApp.ElementType.PARAGRAPH &&
    topType !== DocumentApp.ElementType.LIST_ITEM
  ) {
    throw new Error('Place the cursor on the caption text line, then press Set here.');
  }

  const config = CAPTION_COUNTER_CONFIG_[normalized];
  removeNamedRangesByName_(config.anchorName);

  addActiveTabNamedRange_(config.anchorName, top);

  const props = PropertiesService.getDocumentProperties();
  if (props) props.setProperty(config.startProperty, String(start));

  return {
    type: normalized,
    startAt: start,
    anchorSet: true
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

  let bestIndex = -1;
  let bestDistance = Number.MAX_SAFE_INTEGER;

  for (let i = 0; i < parent.getNumChildren(); i++) {
    const child = parent.getChild(i);
    const matches =
      normalized === 'Table'
        ? child.getType() === DocumentApp.ElementType.TABLE
        : isStandaloneFigureBlock_(child);

    if (!matches) continue;

    const distance = Math.abs(i - anchorLineIndex);

    // Table captions are usually above the table -> prefer object after.
    // Figure captions are usually below the figure -> prefer object before.
    const preferredTie =
      normalized === 'Table'
        ? i > anchorLineIndex
        : i < anchorLineIndex;

    const currentBestPreferred =
      bestIndex >= 0 &&
      (normalized === 'Table'
        ? bestIndex > anchorLineIndex
        : bestIndex < anchorLineIndex);

    if (
      distance < bestDistance ||
      (distance === bestDistance && preferredTie && !currentBestPreferred)
    ) {
      bestIndex = i;
      bestDistance = distance;
    }
  }

  return bestIndex;
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
          ? child.getType() === DocumentApp.ElementType.TABLE
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
        ? child.getType() === DocumentApp.ElementType.TABLE
        : isStandaloneFigureBlock_(child);

    if (matches) relativeCount++;
  }

  if (relativeCount < 1) return null;

  return getCaptionCounterStart_(normalized) + relativeCount - 1;
}

function renumberAllCaptions() {
  const body = getActiveBody_();
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

    const ordinal = getCaptionOrdinal_(p, parsed.type);

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
    skippedBeforeAnchor: skippedBeforeAnchor
  };
}


/**
 * Table numbering is based on ACTUAL Google Docs table elements, not on
 * the number already written in caption text.
 *
 * The caption may be immediately above or below its table. We find the
 * nearest actual TABLE element and return that table's 1-based position
 * in the current document container.
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
  let nearestDistance = Number.MAX_SAFE_INTEGER;

  for (let i = 0; i < parent.getNumChildren(); i++) {
    const child = parent.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TABLE) continue;

    const distance = Math.abs(i - targetIndex);

    if (
      distance < nearestDistance ||
      (distance === nearestDistance && i > targetIndex)
    ) {
      nearestDistance = distance;
      nearestTableIndex = i;
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
      if (parent.getChild(i).getType() === DocumentApp.ElementType.TABLE) count++;
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
  let nearestDistance = Number.MAX_SAFE_INTEGER;

  for (let i = 0; i < parent.getNumChildren(); i++) {
    const child = parent.getChild(i);
    if (!isStandaloneFigureBlock_(child)) continue;

    const distance = Math.abs(i - targetIndex);

    if (
      distance < nearestDistance ||
      (distance === nearestDistance && i < targetIndex)
    ) {
      nearestDistance = distance;
      nearestFigureIndex = i;
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

/**
 * Full Smart Format
 * -----------------
 * One Gemini classification request for the entire selected text, followed
 * by deterministic local formatting. Actual table objects and image objects
 * are intentionally not reformatted here for speed/safety.
 */
function smartFormatSelection() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) {
    throw new Error('Select the text you want to format completely.');
  }

  const allTargets = getSelectedParagraphs_();
  if (!allTargets.length) {
    throw new Error('No paragraphs were found in the selection.');
  }

  // Actual table contents keep their dedicated formatter.
  const targets = allTargets.filter(p => !isInsideTable_(p));
  const skippedTableParagraphs = allTargets.length - targets.length;

  const items = [];
  const targetById = {};
  const ambiguous = [];

  targets.forEach((p, index) => {
    const value = p.getText();
    if (!value.trim()) return;

    const id = 'p' + index;
    const detected = detectFormattingType_(p, value);

    const item = {
      id: id,
      text: value,
      fixedType: detected || '',
      existingHeading: getCurrentHeadingName_(p),
      existingList: getCurrentListType_(p)
    };

    items.push(item);
    targetById[id] = p;

    if (!detected) ambiguous.push(item);
  });

  if (!items.length) {
    throw new Error('The selection contains no text paragraphs to format.');
  }

  // Gemini is used only where document structure is genuinely ambiguous.
  // This is more reliable for long selections than asking it to classify
  // obvious bullets, captions and numbered subsections too.
  const aiPlan = ambiguous.length
    ? classifyFormattingPlanWithGemini_(addFormattingContext_(items, ambiguous))
    : [];

  const typeById = {};
  items.forEach(item => {
    if (item.fixedType) typeById[item.id] = item.fixedType;
  });
  aiPlan.forEach(x => {
    if (!typeById[x.id]) typeById[x.id] = x.type;
  });

  // Any item omitted by Gemini safely falls back to Normal text.
  items.forEach(item => {
    if (!typeById[item.id]) typeById[item.id] = 'normal';
  });

  // Captions already recognizable from their text are numbered in one scan.
  const captionNumbers = buildCaptionNumberPlan_(items, targetById);

  let formatted = 0;
  let headings = 0;
  let lists = 0;
  let captions = 0;
  let notes = 0;
  let normal = 0;
  const styleAssignments = [];
  const body = getActiveBody_();
  let activeListAnchor = null;
  let activeListType = '';
  let previousItemIndex = -1;

  items.forEach(item => {
    let p = targetById[item.id];
    if (!p) return;

    const type = typeById[item.id];
    const isListType = ['bullet', 'number', 'letter', 'roman'].indexOf(type) >= 0;
    const itemIndex = parseInt(String(item.id).substring(1), 10);
    const isContiguous = previousItemIndex >= 0 && itemIndex === previousItemIndex + 1;

    if (!isListType || !isContiguous) {
      activeListAnchor = null;
      activeListType = '';
    }

    switch (type) {
      case 'heading1':
        styleAssignments.push({paragraph: p, styleName: 'H1'}); headings++; break;
      case 'heading2':
        styleAssignments.push({paragraph: p, styleName: 'H2'}); headings++; break;
      case 'heading3':
        styleAssignments.push({paragraph: p, styleName: 'H3'}); headings++; break;
      case 'heading4':
        styleAssignments.push({paragraph: p, styleName: 'H4'}); headings++; break;
      case 'heading5':
        styleAssignments.push({paragraph: p, styleName: 'H5'}); headings++; break;
      case 'heading6':
        styleAssignments.push({paragraph: p, styleName: 'H6'}); headings++; break;

      case 'bullet':
      case 'number':
      case 'letter':
      case 'roman': {
        const listType = type.toUpperCase();
        const continuationAnchor = activeListType === listType
          ? activeListAnchor
          : null;

        stripManualListPrefix_(p, listType);

        const listResult = applyFastNativeListToParagraphs_(
          body,
          [p],
          listType,
          continuationAnchor
        );

        p = listResult.listItems[0];
        targetById[item.id] = p;
        activeListAnchor = listResult.firstListItem;
        activeListType = listType;
        lists++;
        break;
      }

      case 'figure_caption':
      case 'table_caption': {
        const parsed = parseCaptionLine_(p.getText());
        if (parsed) {
          const captionType = type === 'figure_caption' ? 'Figure' : 'Table';
          const n = captionNumbers[item.id] || getCaptionOrdinal_(p, captionType);
          formatCaptionParagraph_(p, captionType, parsed.description, n);
          captions++;
        } else {
          // If Gemini inferred a caption without an explicit prefix, do not
          // invent whether it is Figure/Table here; keep the content safe.
          styleAssignments.push({paragraph: p, styleName: 'NORMAL'});
          normal++;
        }
        break;
      }

      case 'note':
        formatNoteParagraph_(p, parseNoteLine_(p.getText()).description);
        notes++;
        break;

      case 'normal':
      default:
        styleAssignments.push({paragraph: p, styleName: 'NORMAL'});
        normal++;
        break;
    }

    formatted++;
    previousItemIndex = itemIndex;
  });

  styleAssignments.forEach(assignment => {
    applyStyleToParagraph_(assignment.paragraph, assignment.styleName);
  });

  DocumentApp.getActiveDocument().saveAndClose();

  return {
    ok: true,
    formatted: formatted,
    headings: headings,
    lists: lists,
    captions: captions,
    notes: notes,
    normal: normal,
    aiClassified: ambiguous.length,
    skippedTableParagraphs: skippedTableParagraphs
  };
}

function detectFormattingType_(p, value) {
  const text = String(value || '').trim();
  if (!text) return 'normal';

  // Preserve an existing native heading assignment.
  const existingHeading = getCurrentHeadingName_(p);
  if (/^heading[1-6]$/.test(existingHeading)) return existingHeading;

  // Existing native list items are highly reliable evidence.
  const existingList = getCurrentListType_(p);
  if (existingList) return existingList;

  const caption = parseCaptionLine_(text);
  if (caption) {
    return caption.type === 'Figure' ? 'figure_caption' : 'table_caption';
  }

  if (/^\s*(?:Notes?|Notas?)\b(?:\s*[.:–—-]|\s+)/i.test(text)) {
    return 'note';
  }

  // Explicit textual list markers.
  if (/^\s*[•●○▪◦‣⁃-]\s+/.test(text)) return 'bullet';

  // Roman must be tested before letters because "i)" is also a letter.
  if (/^\s*(?:[ivxlcdm]+)[.)]\s+/i.test(text)) return 'roman';
  if (/^\s*[A-Za-zÁÉÍÓÚÜÑáéíóúüñ][.)]\s+/.test(text)) return 'letter';

  // Decimal section numbering is deterministic:
  // 1.1 -> H2, 1.1.1 -> H3, etc.
  const section = text.match(/^\s*(\d+(?:\.\d+)+)\.?\s+\S/);
  if (section) {
    const depth = section[1].split('.').length;
    return 'heading' + Math.min(depth, 6);
  }

  // A plain integer prefix such as "1. ..." is ambiguous: it may be H1
  // or an ordered list. Leave it to Gemini with surrounding context.
  // Unnumbered short title-like lines are also left to Gemini.

  // Long prose ending in normal sentence punctuation is safe to classify
  // locally, reducing Gemini load on large documents.
  const words = text.split(/\s+/).length;
  if (words >= 18 && /[.!?]$/.test(text)) return 'normal';

  return '';
}

function getCurrentHeadingName_(p) {
  try {
    const h = String(p.getHeading());
    const map = {
      HEADING1:'heading1', HEADING2:'heading2', HEADING3:'heading3',
      HEADING4:'heading4', HEADING5:'heading5', HEADING6:'heading6'
    };
    return map[h] || '';
  } catch (e) {
    return '';
  }
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

function addFormattingContext_(allItems, ambiguousItems) {
  const indexById = {};
  allItems.forEach((x, i) => indexById[x.id] = i);

  return ambiguousItems.map(item => {
    const i = indexById[item.id];
    return {
      id: item.id,
      text: item.text,
      existingHeading: item.existingHeading || '',
      existingList: item.existingList || '',
      previous: i > 0 ? allItems[i - 1].text : '',
      next: i < allItems.length - 1 ? allItems[i + 1].text : ''
    };
  });
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

function buildCaptionNumberPlan_(items, targetById) {
  const result = {};

  items.forEach(item => {
    if (
      item.fixedType !== 'figure_caption' &&
      item.fixedType !== 'table_caption'
    ) {
      return;
    }

    const p = targetById[item.id];
    if (!p) return;

    const type = item.fixedType === 'figure_caption' ? 'Figure' : 'Table';
    result[item.id] = getCaptionOrdinal_(p, type);
  });

  return result;
}
