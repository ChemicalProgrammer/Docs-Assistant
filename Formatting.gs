function applyNamedStyle(styleName) {
  const targets = getStyleTargetParagraphs_();
  if (!targets.length) {
    throw new Error('Select text or place the cursor in the paragraph you want to format.');
  }

  targets.forEach(p => applyNamedStyleToParagraph_(p, styleName));
  return true;
}

function applyNamedStyleToParagraph_(p, styleName) {
  const map = {
    NORMAL: DocumentApp.ParagraphHeading.NORMAL,
    H1: DocumentApp.ParagraphHeading.HEADING1,
    H2: DocumentApp.ParagraphHeading.HEADING2,
    H3: DocumentApp.ParagraphHeading.HEADING3,
    H4: DocumentApp.ParagraphHeading.HEADING4,
    H5: DocumentApp.ParagraphHeading.HEADING5,
    H6: DocumentApp.ParagraphHeading.HEADING6
  };

  const heading = map[styleName];
  if (!heading) throw new Error('Unknown style.');

  const body = getActiveBody_();
  const styleAttributes = body.getHeadingAttributes(heading);

  if (styleName !== 'NORMAL') {
    const currentText = p.getText();
    const converted = sentenceCaseHeading_(currentText);
    if (converted !== currentText) {
      p.editAsText().setText(converted);
    }
  }

  p.setHeading(heading);
  p.setAttributes(styleAttributes);
  p.setHeading(heading);

  applyHeadingIndentation_(p, styleName);

  if (styleName !== 'NORMAL') {
    normalizeHeadingNumberSpacing_(p);
  }

  applyNamedTextAttributes_(p, styleAttributes);
}

function applyHeadingIndentation_(paragraph, styleName) {
  const PT_PER_IN = 72;

  const indents = {
    H1: {left: -0.12, right: 0},
    H2: {left:  0.00, right: 0},
    H3: {left:  0.19, right: 0}
  };

  const cfg = indents[styleName];
  if (!cfg) return;

  const leftPt = cfg.left * PT_PER_IN;
  const rightPt = cfg.right * PT_PER_IN;

  // "Special indent: None" means the first line begins at the same
  // position as the paragraph's left/start indent.
  paragraph.setIndentStart(leftPt);
  paragraph.setIndentEnd(rightPt);
  paragraph.setIndentFirstLine(leftPt);
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
}

function normalizeHeadingNumberSpacing_(paragraph) {
  const text = paragraph.getText();
  if (!text) return;

  // Examples:
  // "1.    Introduction"    -> "1. Introduction"
  // "1.2\tScope"            -> "1.2 Scope"
  // "1.2.3   Methodology"   -> "1.2.3 Methodology"
  //
  // Only affects headings that begin with a numeric section identifier.
  const normalized = text.replace(
    /^(\s*\d+(?:\.\d+)*\.?)[\t ]+/,
    '$1 '
  );

  if (normalized !== text) {
    paragraph.editAsText().setText(normalized);
  }
}

function getStyleTargetParagraphs_() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  // If there is a selection, apply the style to every paragraph touched
  // by the selection, exactly as before.
  if (selection) {
    return getSelectedParagraphs_();
  }

  // If there is no selection, the paragraph/list item containing the cursor
  // is the target. The user does not need to select the heading text.
  const cursor = doc.getCursor();
  if (!cursor) return [];

  let el = cursor.getElement();
  while (
    el &&
    el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
    el.getType() !== DocumentApp.ElementType.LIST_ITEM
  ) {
    el = el.getParent();
  }

  return el ? [el] : [];
}

function getActiveBody_() {
  const doc = DocumentApp.getActiveDocument();
  try {
    return doc.getActiveTab().asDocumentTab().getBody();
  } catch (e) {
    return doc.getBody();
  }
}

function applyNamedTextAttributes_(paragraph, attrs) {
  const text = paragraph.editAsText();
  if (!text || text.getText().length === 0) return;

  const textAttrs = {};
  const supported = [
    DocumentApp.Attribute.FONT_FAMILY,
    DocumentApp.Attribute.FONT_SIZE,
    DocumentApp.Attribute.BOLD,
    DocumentApp.Attribute.ITALIC,
    DocumentApp.Attribute.UNDERLINE,
    DocumentApp.Attribute.STRIKETHROUGH,
    DocumentApp.Attribute.FOREGROUND_COLOR,
    DocumentApp.Attribute.BACKGROUND_COLOR
  ];

  supported.forEach(attr => {
    if (attrs[attr] !== undefined && attrs[attr] !== null) {
      textAttrs[attr] = attrs[attr];
    }
  });

  if (Object.keys(textAttrs).length) {
    text.setAttributes(textAttrs);
  }
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

function applyListPreset(type, continuePrevious) {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) throw new Error('Select paragraphs first.');

  const paragraphs = getSelectedParagraphs_();
  paragraphs.forEach(p => applyListFormatToParagraph_(p, type));

  return {
    ok: true,
    continuePrevious: Boolean(continuePrevious),
    leftIndentInches: 0.06,
    hangingIndentInches: 0.25
  };
}

function applyListFormatToParagraph_(p, type) {
  const glyphs = {
    BULLET: DocumentApp.GlyphType.BULLET,
    NUMBER: DocumentApp.GlyphType.NUMBER,
    LETTER: DocumentApp.GlyphType.LATIN_LOWER,
    ROMAN: DocumentApp.GlyphType.ROMAN_LOWER
  };

  const glyph = glyphs[type];
  if (!glyph) throw new Error('Unknown list type.');

  const LEFT_IN = 0.06;
  const HANGING_IN = 0.25;
  const PT_PER_IN = 72;
  const firstLinePt = LEFT_IN * PT_PER_IN;
  const startPt = (LEFT_IN + HANGING_IN) * PT_PER_IN;

  // Always start from the CURRENT Normal text style.
  applyNamedStyleToParagraph_(p, 'NORMAL');

  let item = p;

  // Reuse an existing list item; otherwise convert the paragraph.
  if (p.getType() !== DocumentApp.ElementType.LIST_ITEM) {
    const parent = p.getParent();
    const idx = parent.getChildIndex(p);
    item = parent.insertListItem(idx, p.getText());
    p.removeFromParent();

    // Converting to ListItem creates a new element, so reapply Normal text.
    applyNamedStyleToParagraph_(item, 'NORMAL');
  }

  item.setGlyphType(glyph);
  item.setIndentStart(startPt);
  item.setIndentFirstLine(firstLinePt);

  return item;
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
  const paragraphs = getSelectedParagraphs_();
  if (!paragraphs.length) throw new Error('Select one or more paragraphs first.');
  paragraphs.forEach(fn);
}


function formatSelectedTable() {
  const table = getActiveTable_();
  if (!table) throw new Error('Place the cursor inside a table or select content inside a table.');

  table.setBorderColor('#000000');
  table.setBorderWidth(1);

  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    row.setMinimumHeight(0.49 * 72);

    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      cell.setPaddingTop(0.028 * 72);
      cell.setPaddingBottom(0.028 * 72);
      cell.setPaddingLeft(0.028 * 72);
      cell.setPaddingRight(0.028 * 72);
      cell.setBackgroundColor(null);
      formatTableCellContent_(cell, r === 0);
    }
  }
  return {ok:true, rows:table.getNumRows()};
}

function getActiveTable_() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  if (selection) {
    const ranges = selection.getRangeElements();
    for (let i = 0; i < ranges.length; i++) {
      const table = findAncestorTable_(ranges[i].getElement());
      if (table) return table;
    }
  }
  const cursor = doc.getCursor();
  if (cursor) {
    const table = findAncestorTable_(cursor.getElement());
    if (table) return table;
  }
  return null;
}

function findAncestorTable_(el) {
  while (el) {
    if (el.getType && el.getType() === DocumentApp.ElementType.TABLE) return el.asTable();
    el = el.getParent ? el.getParent() : null;
  }
  return null;
}

function formatTableCellContent_(cell, isHeader) {
  for (let i = 0; i < cell.getNumChildren(); i++) {
    const child = cell.getChild(i);
    const type = child.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      p.setAlignment(isHeader ? DocumentApp.HorizontalAlignment.CENTER : DocumentApp.HorizontalAlignment.LEFT);
      p.setLineSpacing(1.5);
      const t = p.editAsText();
      t.setFontFamily('Arial').setFontSize(9).setBold(isHeader);
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      const p = child.asListItem();
      p.setAlignment(isHeader ? DocumentApp.HorizontalAlignment.CENTER : DocumentApp.HorizontalAlignment.LEFT);
      p.setLineSpacing(1.5);
      const t = p.editAsText();
      t.setFontFamily('Arial').setFontSize(9).setBold(isHeader);
    }
  }
}


/**
 * Formats the current paragraph as a Figure/Table caption.
 * Inputs such as "Table 5.1. Overview" are treated as one old caption
 * identifier; ".1" is NOT left behind as part of the description.
 */
function formatCaptionLine() {
  const targets = getStyleTargetParagraphs_();
  if (!targets.length) {
    throw new Error('Place the cursor in a Figure/Table caption or select the caption line.');
  }
  if (targets.length !== 1) {
    throw new Error('Format one Figure/Table caption at a time.');
  }

  const p = targets[0];
  if (p.getType() === DocumentApp.ElementType.LIST_ITEM) {
    throw new Error('Figure/Table captions cannot be list items.');
  }

  const parsed = parseCaptionLine_(p.getText());
  if (!parsed) {
    throw new Error('The line must begin with "Figure X." or "Table X."');
  }

  const nextNumber = getCaptionOrdinal_(p, parsed.type);
  formatCaptionParagraph_(p, parsed.type, parsed.description, nextNumber);

  return {
    ok: true,
    type: parsed.type,
    number: nextNumber,
    text: p.getText()
  };
}

function parseCaptionLine_(value) {
  const text = String(value || '');

  // Important: consume the COMPLETE old identifier:
  // Table 5.1. Overview  -> id "5.1", description "Overview"
  // Table 5.1.2. Title   -> id "5.1.2", description "Title"
  // Table 7. Title       -> id "7", description "Title"
  // Table X. Title       -> id "X", description "Title"
  const match = text.match(
    /^\s*(Figure|Table)\s+(X|\d+(?:\.\d+)*)\.\s*(.*)$/i
  );
  if (!match) return null;

  return {
    type: match[1].toLowerCase() === 'figure' ? 'Figure' : 'Table',
    oldNumber: match[2],
    description: match[3] || ''
  };
}

function formatCaptionParagraph_(p, type, description, number) {
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
  const body = getActiveBody_();
  let count = 0;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);

    if (child === targetParagraph) {
      return count + 1;
    }

    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const parsed = parseCaptionLine_(child.asParagraph().getText());
      if (parsed && parsed.type === type) count++;
    }
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
  if (!allTargets.length) throw new Error('No paragraphs were found in the selection.');

  // Skip paragraphs inside actual table cells. Tables keep their dedicated button.
  const targets = allTargets.filter(p => !isInsideTable_(p));
  const skippedTableParagraphs = allTargets.length - targets.length;

  const items = [];
  const targetById = {};

  targets.forEach((p, index) => {
    const value = p.getText();
    if (!value.trim()) return;

    const id = 'p' + index;
    const parsedCaption = parseCaptionLine_(value);
    let currentList = '';

    if (p.getType() === DocumentApp.ElementType.LIST_ITEM) {
      try {
        currentList = String(p.asListItem().getGlyphType());
      } catch (e) {}
    }

    items.push({
      id: id,
      text: value,
      fixedType: parsedCaption
        ? (parsedCaption.type === 'Figure' ? 'figure_caption' : 'table_caption')
        : '',
      existingList: currentList
    });
    targetById[id] = p;
  });

  if (!items.length) throw new Error('The selection contains no text paragraphs to format.');

  const plan = classifyFormattingPlanWithGemini_(items);
  const planById = {};
  plan.forEach(x => planById[x.id] = x.type);

  // Compute caption numbers BEFORE changing any selected caption text.
  // This avoids rescanning the document once per caption.
  const captionNumbers = buildCaptionNumberPlan_(items, targetById);

  let formatted = 0;
  let captions = 0;
  let lists = 0;
  let headings = 0;

  items.forEach(item => {
    let p = targetById[item.id];
    if (!p) return;

    let type = item.fixedType || planById[item.id] || 'normal';

    switch (type) {
      case 'heading1':
        applyNamedStyleToParagraph_(p, 'H1'); headings++; break;
      case 'heading2':
        applyNamedStyleToParagraph_(p, 'H2'); headings++; break;
      case 'heading3':
        applyNamedStyleToParagraph_(p, 'H3'); headings++; break;
      case 'heading4':
        applyNamedStyleToParagraph_(p, 'H4'); headings++; break;
      case 'heading5':
        applyNamedStyleToParagraph_(p, 'H5'); headings++; break;
      case 'heading6':
        applyNamedStyleToParagraph_(p, 'H6'); headings++; break;

      case 'bullet':
        applyListFormatToParagraph_(p, 'BULLET'); lists++; break;
      case 'number':
        applyListFormatToParagraph_(p, 'NUMBER'); lists++; break;
      case 'letter':
        applyListFormatToParagraph_(p, 'LETTER'); lists++; break;
      case 'roman':
        applyListFormatToParagraph_(p, 'ROMAN'); lists++; break;

      case 'figure_caption':
      case 'table_caption': {
        const parsed = parseCaptionLine_(p.getText());
        if (parsed) {
          const n = captionNumbers[item.id] || getCaptionOrdinal_(p, parsed.type);
          formatCaptionParagraph_(p, parsed.type, parsed.description, n);
          captions++;
        }
        break;
      }

      case 'normal':
      default:
        applyNamedStyleToParagraph_(p, 'NORMAL');
        break;
    }

    formatted++;
  });

  return {
    ok: true,
    formatted: formatted,
    headings: headings,
    lists: lists,
    captions: captions,
    skippedTableParagraphs: skippedTableParagraphs
  };
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
  const wanted = new Map();

  items.forEach(item => {
    if (item.fixedType === 'figure_caption' || item.fixedType === 'table_caption') {
      const p = targetById[item.id];
      if (p) wanted.set(p, item.id);
    }
  });

  if (!wanted.size) return result;

  const body = getActiveBody_();
  let figureCount = 0;
  let tableCount = 0;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const parsed = parseCaptionLine_(child.asParagraph().getText());
    if (!parsed) continue;

    const next = parsed.type === 'Figure' ? ++figureCount : ++tableCount;

    if (wanted.has(child)) {
      result[wanted.get(child)] = next;
    }
  }

  return result;
}
