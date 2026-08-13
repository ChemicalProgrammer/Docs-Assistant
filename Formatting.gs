function applyNamedStyle(styleName) {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) throw new Error('Select one or more paragraphs first.');

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

  eachSelectedParagraph_(p => p.setHeading(heading));
  return true;
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

  const glyphs = {
    BULLET: DocumentApp.GlyphType.BULLET,
    NUMBER: DocumentApp.GlyphType.NUMBER,
    LETTER: DocumentApp.GlyphType.LATIN_LOWER,
    ROMAN: DocumentApp.GlyphType.ROMAN_LOWER
  };

  const glyph = glyphs[type];
  if (!glyph) throw new Error('Unknown list type.');

  const paragraphs = getSelectedParagraphs_();
  paragraphs.forEach((p, i) => {
    const parent = p.getParent();
    const idx = parent.getChildIndex(p);
    const text = p.getText();
    const listItem = parent.insertListItem(idx, text);
    listItem.setGlyphType(glyph);
    p.removeFromParent();
  });

  // Google Docs list continuation has limitations in DocumentApp.
  // The flag is retained in the API/UI so a Docs API implementation can
  // be added without changing the interface.
  return { ok: true, continuePrevious: Boolean(continuePrevious) };
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
