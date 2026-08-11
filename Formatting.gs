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
  const out = [];
  const seen = new Set();

  selection.getRangeElements().forEach(re => {
    let el = re.getElement();
    while (el && el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
           el.getType() !== DocumentApp.ElementType.LIST_ITEM) {
      el = el.getParent();
    }
    if (el) {
      const key = String(el);
      if (!seen.has(key)) {
        seen.add(key);
        out.push(el);
      }
    }
  });
  return out;
}

function eachSelectedParagraph_(fn) {
  const paragraphs = getSelectedParagraphs_();
  if (!paragraphs.length) throw new Error('Select one or more paragraphs first.');
  paragraphs.forEach(fn);
}
