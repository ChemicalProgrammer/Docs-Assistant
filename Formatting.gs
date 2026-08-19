function applyNamedStyle(styleName) {
  const started = Date.now();
  const targets = getStyleTargetParagraphs_();

  if (!targets.length) {
    throw new Error('Select text or place the cursor in the paragraph you want to format.');
  }

  const context = getNamedStyleContext_(styleName);

  targets.forEach(p => {
    formatSingleParagraph_(p, styleName, context);
  });

  return {
    ok: true,
    paragraphs: targets.length,
    mode: targets.length === 1 ? 'single' : 'range',
    elapsedMs: Date.now() - started
  };
}

function applyNamedStyleToCurrentParagraph(styleName) {
  const started = Date.now();
  const target = getCurrentParagraph_();

  if (!target) {
    throw new Error('Place the cursor in the paragraph you want to format.');
  }

  const context = getNamedStyleContext_(styleName);
  formatSingleParagraph_(target, styleName, context);

  return {
    ok: true,
    paragraphs: 1,
    mode: 'single',
    elapsedMs: Date.now() - started
  };
}

function applyNamedStyleToSelectedParagraphs(styleName) {
  const started = Date.now();
  const targets = getSelectedParagraphsIncludingGaps_();

  if (!targets.length) {
    throw new Error('Select one or more paragraphs first.');
  }

  const context = getNamedStyleContext_(styleName);

  targets.forEach(p => {
    formatSingleParagraph_(p, styleName, context);
  });

  return {
    ok: true,
    paragraphs: targets.length,
    mode: 'range',
    elapsedMs: Date.now() - started
  };
}

/**
 * Reliable single-paragraph formatter.
 *
 * Important:
 * - Paragraph.setAttributes() is intentionally NOT used here.
 *   Mixing paragraph and character attributes in one attribute map proved
 *   unreliable for Normal/Heading formatting.
 * - setHeading() applies the document's real named style.
 * - left/right/special indentation are then normalized explicitly.
 * - character formatting is applied only through the Text element.
 */
function formatSingleParagraph_(paragraph, styleName, context) {
  const ctx = context || getNamedStyleContext_(styleName);

  // Do not let the style buttons operate inside a table cell. Tables have
  // their own formatting rules and this prevents accidental cross-formatting.
  if (isInsideTable_(paragraph)) {
    throw new Error('Normal/Heading styles cannot be applied inside a table cell.');
  }

  // Heading text normalization happens in memory and is written at most once.
  if (styleName !== 'NORMAL') {
    const text = paragraph.editAsText();
    const original = text.getText();
    const normalized = normalizeHeadingTextValue_(original);

    if (normalized !== original) {
      text.setText(normalized);
    }
  }

  // Apply the actual Google Docs named style.
  paragraph.setHeading(ctx.heading);

  // Project rule: Normal + H1-H6 all align to the exact same left axis.
  paragraph.setIndentStart(0);
  paragraph.setIndentEnd(0);
  paragraph.setIndentFirstLine(0);
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

  const text = paragraph.editAsText();
  if (!text || !text.getText().length) return;

  // Reapply the named style's character attributes so direct formatting
  // does not survive unexpectedly (e.g. Normal remaining bold).
  if (ctx.textAttributes && Object.keys(ctx.textAttributes).length) {
    text.setAttributes(ctx.textAttributes);
  }

  // Project rule: every Heading 1-6 is fully Bold and 10 pt, including
  // manually typed numbering and the title text.
  if (styleName !== 'NORMAL') {
    text.setBold(true);
    text.setFontSize(10);
  }
}

function getNamedStyleContext_(styleName) {
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
  if (!heading) {
    throw new Error('Unknown style: ' + styleName);
  }

  // One style read for the entire operation.
  const styleAttributes = getActiveBody_().getHeadingAttributes(heading);

  return {
    heading: heading,
    textAttributes: getNamedTextAttributes_(styleAttributes)
  };
}

function getNamedTextAttributes_(attrs) {
  const result = {};
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
      result[attr] = attrs[attr];
    }
  });

  return result;
}

function normalizeHeadingTextValue_(value) {
  const source = String(value || '');
  if (!source) return source;

  let normalized = source.toLowerCase().replace(
    /[A-Za-zÁÉÍÓÚÜÑÀÈÌÒÙÂÊÎÔÛÄËÏÖÜÇ]/,
    ch => ch.toUpperCase()
  );

  normalized = normalized.replace(
    /^(\s*\d+(?:\.\d+)*\.?)[\t ]+/,
    '$1 '
  );

  return normalized;
}

function getCurrentParagraph_() {
  const doc = DocumentApp.getActiveDocument();
  const cursor = doc.getCursor();
  if (!cursor) return null;

  let el = cursor.getElement();

  while (
    el &&
    el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
    el.getType() !== DocumentApp.ElementType.LIST_ITEM
  ) {
    el = el.getParent();
  }

  return el || null;
}

function getStyleTargetParagraphs_() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  if (selection) {
    return getSelectedParagraphsIncludingGaps_();
  }

  const current = getCurrentParagraph_();
  return current ? [current] : [];
}

function getSelectedParagraphsIncludingGaps_() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) return [];

  const ranges = selection.getRangeElements();
  if (!ranges.length) return [];

  const touched = [];

  ranges.forEach(re => {
    let el = re.getElement();

    while (
      el &&
      el.getType() !== DocumentApp.ElementType.PARAGRAPH &&
      el.getType() !== DocumentApp.ElementType.LIST_ITEM
    ) {
      el = el.getParent();
    }

    if (el) touched.push(el);
  });

  if (!touched.length) return [];

  const first = touched[0];
  const last = touched[touched.length - 1];
  const parent = first.getParent();

  try {
    const firstIndex = parent.getChildIndex(first);
    const lastIndex = parent.getChildIndex(last);
    const minIndex = Math.min(firstIndex, lastIndex);
    const maxIndex = Math.max(firstIndex, lastIndex);
    const result = [];

    for (let i = minIndex; i <= maxIndex; i++) {
      const child = parent.getChild(i);
      const type = child.getType();

      if (
        type === DocumentApp.ElementType.PARAGRAPH ||
        type === DocumentApp.ElementType.LIST_ITEM
      ) {
        result.push(child);
      }
    }

    return result;
  } catch (e) {
    return getSelectedParagraphs_();
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
  // Selection is optional. With only the cursor, format the whole current line.
  const paragraphs = getStyleTargetParagraphs_();
  if (!paragraphs.length) {
    throw new Error('Place the cursor in a paragraph or select one or more paragraphs.');
  }

  // Bullets must be REAL Google Docs bullets (native ListItems).
  if (type === 'BULLET') {
    return applyNativeBulletPreset_(paragraphs, Boolean(continuePrevious));
  }

  return applyManualListPreset_(paragraphs, type, Boolean(continuePrevious));
}

/**
 * SAFE LIST ENGINE
 * ----------------
 * Number/Letter/Roman use the explicit-prefix fallback.
 * BULLET uses the separate native-bullet engine above so bullets remain
 * true Google Docs ListItems.
 */
function applyNativeBulletPreset_(paragraphs, continuePrevious) {
  let anchor = continuePrevious
    ? findPreviousSafeNativeBullet_(paragraphs[0])
    : null;

  let tableBullets = 0;
  let bodyBullets = 0;

  paragraphs.forEach(p => {
    const targetInsideTable = isInsideTable_(p);

    if (isExistingNativeListOfType_(p, 'BULLET')) {
      normalizeExistingNativeListParagraph_(p, 'BULLET');
      if (!anchor) anchor = p.asListItem();

      if (targetInsideTable) tableBullets++;
      else bodyBullets++;
      return;
    }

    const item = convertToSafeNativeBullet_(p, anchor);
    normalizeExistingNativeListParagraph_(item, 'BULLET');

    if (!anchor) anchor = item;

    if (isInsideTable_(item)) tableBullets++;
    else bodyBullets++;
  });

  return {
    ok: true,
    requestedContinue: continuePrevious,
    nativeBullets: true,
    bodyBullets: bodyBullets,
    tableBullets: tableBullets,
    bodyIndent: {left: 0.06, hanging: 0.25},
    tableIndent: {left: 0.00, hanging: 0.20}
  };
}

function convertToSafeNativeBullet_(p, anchor) {
  const parent = p.getParent();
  const idx = parent.getChildIndex(p);
  const content = stripAnyListPrefixText_(p.getText()).trim();

  if (anchor) {
    const item = parent.insertListItem(idx, content);

    // Critical: assign the known bullet list before removing the original.
    // Do not call setGlyphType(); the anchor's list definition supplies it.
    item.setListId(anchor);

    p.removeFromParent();
    return item;
  }

  // Create an isolated native bullet list.
  const before = parent.insertParagraph(idx, '\uE210');
  const after = parent.insertParagraph(idx + 1, '\uE211');
  const item = parent.insertListItem(idx + 1, content);

  // Item is surrounded by normal paragraphs, so this glyph change cannot
  // mutate a neighboring numbered-heading list.
  item.setGlyphType(DocumentApp.GlyphType.BULLET);

  // Remove original target and temporary separators.
  p.removeFromParent();
  try { after.removeFromParent(); } catch (e) {}
  try { before.removeFromParent(); } catch (e) {}

  return item;
}

function findPreviousSafeNativeBullet_(target) {
  const parent = target.getParent();
  if (!parent) return null;

  const targetIndex = parent.getChildIndex(target);

  for (let i = targetIndex - 1; i >= 0; i--) {
    const child = parent.getChild(i);

    if (!isExistingNativeListOfType_(child, 'BULLET')) continue;

    // Prefer a bullet already using this Add-on's expected geometry.
    if (matchesAddonListIndent_(child.asListItem(), 'BULLET')) {
      return child.asListItem();
    }
  }

  return null;
}

function applyManualListPreset_(paragraphs, type, continuePrevious) {
  const allowed = ['BULLET', 'NUMBER', 'LETTER', 'ROMAN'];
  if (allowed.indexOf(type) === -1) throw new Error('Unknown list type.');

  let ordinal = 1;

  if (continuePrevious && type !== 'BULLET') {
    const previous = findPreviousListOrdinal_(paragraphs[0], type);
    if (previous > 0) ordinal = previous + 1;
  }

  paragraphs.forEach(p => {
    // Surgical preservation rule:
    // if the line is already a native Google Docs list of the requested type,
    // do NOT recreate it, do NOT touch listId, and do NOT touch glyph type.
    // That keeps automatic a)->b), 1.->2., etc. intact and cannot alter
    // neighboring heading-numbering definitions.
    if (isExistingNativeListOfType_(p, type)) {
      normalizeExistingNativeListParagraph_(p, type);
      return;
    }

    const original = p.getText();
    const paragraph = recreateAsParagraph_(p);

    // Stable v0.7.5 behavior for plain paragraphs / incompatible lists.
    applyNamedStyleToParagraph_(paragraph, 'NORMAL');

    const content = stripAnyListPrefixText_(original).trim();

    switch (type) {
      case 'BULLET':
        throw new Error('Bullet formatting must use the native bullet engine.');

      case 'NUMBER':
        paragraph.setText(String(ordinal) + '. ' + content);
        ordinal++;
        break;

      case 'LETTER':
        paragraph.setText(numberToLetters_(ordinal) + ') ' + content);
        ordinal++;
        break;

      case 'ROMAN':
        paragraph.setText(numberToRoman_(ordinal).toLowerCase() + '. ' + content);
        ordinal++;
        break;
    }

    normalizeManualListParagraph_(paragraph);
    applyListIndents_(paragraph);
  });

  return {
    ok: true,
    requestedContinue: continuePrevious,
    leftIndentInches: 0.06,
    hangingIndentInches: 0.25,
    safeManualList: true
  };
}

function isExistingNativeListOfType_(p, type) {
  if (!p || p.getType() !== DocumentApp.ElementType.LIST_ITEM) return false;

  try {
    const glyph = p.asListItem().getGlyphType();

    if (type === 'BULLET') return glyph === DocumentApp.GlyphType.BULLET;
    if (type === 'NUMBER') return glyph === DocumentApp.GlyphType.NUMBER;
    if (type === 'LETTER') return glyph === DocumentApp.GlyphType.LATIN_LOWER;
    if (type === 'ROMAN')  return glyph === DocumentApp.GlyphType.ROMAN_LOWER;
  } catch (e) {}

  return false;
}

function normalizeExistingNativeListParagraph_(item, listType) {
  // Preserve native Google Docs list structure:
  // - listId unchanged
  // - glyph type unchanged
  //
  // Only style/indent the selected ListItem.

  try {
    item.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  } catch (e) {}

  const cfg = getNativeListIndentConfig_(item, listType);
  const PT_PER_IN = 72;

  // In Docs geometry:
  // first-line position = Left
  // wrapped-line position = Left + Hanging
  item.setIndentFirstLine(cfg.left * PT_PER_IN);
  item.setIndentStart((cfg.left + cfg.hanging) * PT_PER_IN);
  item.setIndentEnd(0);

  // Prevent accidental whole-line bold inherited from a preceding heading,
  // while leaving mixed rich text alone.
  const text = item.editAsText();
  const value = text.getText();
  if (!value) return;

  const normalAttrs = getActiveBody_().getHeadingAttributes(DocumentApp.ParagraphHeading.NORMAL);
  const normalBold = normalAttrs[DocumentApp.Attribute.BOLD];

  if (normalBold !== true) {
    const boldState = text.isBold();
    if (boldState === true) {
      text.setBold(false);
    }
  }
}

function getNativeListIndentConfig_(item, listType) {
  // Special rule requested for BULLETS inside a table cell.
  if (listType === 'BULLET' && isInsideTable_(item)) {
    return {
      left: 0.00,
      hanging: 0.20
    };
  }

  return {
    left: 0.06,
    hanging: 0.25
  };
}

function recreateAsParagraph_(p) {
  if (p.getType() === DocumentApp.ElementType.PARAGRAPH) {
    return p.asParagraph();
  }

  // If the source is a native ListItem, convert only the selected item
  // into a plain paragraph. No glyph/list definition is modified.
  const parent = p.getParent();
  const idx = parent.getChildIndex(p);
  const paragraph = parent.insertParagraph(idx, p.getText());
  p.removeFromParent();
  return paragraph;
}

function normalizeManualListParagraph_(paragraph) {
  // Reapply the CURRENT Normal text style after setText().
  // Google Docs can inherit direct character formatting (notably bold)
  // from a neighboring paragraph when a paragraph/list item is recreated.
  applyNamedStyleToParagraph_(paragraph, 'NORMAL');

  // Some documents retain direct BOLD on newly created text even after the
  // named style is reassigned. A manual list is expected to begin as Normal
  // text, so explicitly clear whole-paragraph bold only when Normal text is
  // not defined as bold.
  const body = getActiveBody_();
  const normalAttrs = body.getHeadingAttributes(DocumentApp.ParagraphHeading.NORMAL);
  const normalBold = normalAttrs[DocumentApp.Attribute.BOLD];

  if (normalBold !== true) {
    const t = paragraph.editAsText();
    if (t.getText().length) t.setBold(false);
  }
}

function applyListIndents_(item) {
  const PT_PER_IN = 72;

  // Requested list geometry:
  // Left = 0.06 in
  // Hanging = 0.25 in
  item.setIndentFirstLine(0.06 * PT_PER_IN);
  item.setIndentStart((0.06 + 0.25) * PT_PER_IN);
  item.setIndentEnd(0);
}

function findPreviousListOrdinal_(target, type) {
  const parent = target.getParent();
  if (!parent) return 0;

  const targetIndex = parent.getChildIndex(target);

  for (let i = targetIndex - 1; i >= 0; i--) {
    const child = parent.getChild(i);

    if (
      child.getType() !== DocumentApp.ElementType.PARAGRAPH &&
      child.getType() !== DocumentApp.ElementType.LIST_ITEM
    ) {
      continue;
    }

    const value = child.getText();

    // First prefer lists already formatted by this Add-on.
    const manual = getManualListOrdinal_(value, type);
    if (manual > 0) return manual;

    // Backward compatibility: older Add-on versions used native ListItems.
    // Read them only; never modify their glyph/listId.
    if (child.getType() === DocumentApp.ElementType.LIST_ITEM) {
      const item = child.asListItem();

      if (!matchesAddonListIndent_(item, type)) continue;

      try {
        const glyph = item.getGlyphType();

        if (
          (type === 'NUMBER' && glyph === DocumentApp.GlyphType.NUMBER) ||
          (type === 'LETTER' && glyph === DocumentApp.GlyphType.LATIN_LOWER) ||
          (type === 'ROMAN' && glyph === DocumentApp.GlyphType.ROMAN_LOWER)
        ) {
          const nativeOrdinal = getNativeListOrdinalReadOnly_(item);
          if (nativeOrdinal > 0) return nativeOrdinal;
        }
      } catch (e) {}
    }
  }

  return 0;
}

function getManualListOrdinal_(value, type) {
  const text = String(value || '');

  if (type === 'NUMBER') {
    const m = text.match(/^\s*(\d+)\.\s+/);
    return m ? Number(m[1]) : 0;
  }

  if (type === 'LETTER') {
    const m = text.match(/^\s*([A-Za-z]+)\)\s+/);
    return m ? lettersToNumber_(m[1]) : 0;
  }

  if (type === 'ROMAN') {
    const m = text.match(/^\s*([ivxlcdm]+)\.\s+/i);
    return m ? romanToNumber_(m[1]) : 0;
  }

  return 0;
}

function matchesAddonListIndent_(item, listType) {
  const cfg = getNativeListIndentConfig_(item, listType);
  const EXPECTED_START = (cfg.left + cfg.hanging) * 72;
  const EXPECTED_FIRST = cfg.left * 72;
  const TOLERANCE = 1.5;

  try {
    return (
      Math.abs(Number(item.getIndentStart()) - EXPECTED_START) <= TOLERANCE &&
      Math.abs(Number(item.getIndentFirstLine()) - EXPECTED_FIRST) <= TOLERANCE
    );
  } catch (e) {
    return false;
  }
}

function getNativeListOrdinalReadOnly_(targetItem) {
  const body = getActiveBody_();
  let count = 0;
  let targetListId = '';

  try {
    targetListId = targetItem.getListId();
  } catch (e) {
    return 0;
  }

  if (!targetListId) return 0;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

    const item = child.asListItem();

    try {
      if (item.getListId() === targetListId) count++;
      if (item === targetItem) return count;
    } catch (e) {}
  }

  return count;
}

function numberToLetters_(number) {
  let n = Math.max(1, Number(number) || 1);
  let result = '';

  while (n > 0) {
    n--;
    result = String.fromCharCode(97 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }

  return result;
}

function lettersToNumber_(letters) {
  const value = String(letters || '').toLowerCase();
  let result = 0;

  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i) - 96;
    if (code < 1 || code > 26) return 0;
    result = result * 26 + code;
  }

  return result;
}

function numberToRoman_(number) {
  let n = Math.max(1, Number(number) || 1);
  const map = [
    [1000,'M'],[900,'CM'],[500,'D'],[400,'CD'],
    [100,'C'],[90,'XC'],[50,'L'],[40,'XL'],
    [10,'X'],[9,'IX'],[5,'V'],[4,'IV'],[1,'I']
  ];

  let result = '';

  map.forEach(pair => {
    while (n >= pair[0]) {
      result += pair[1];
      n -= pair[0];
    }
  });

  return result;
}

function romanToNumber_(roman) {
  const value = String(roman || '').toUpperCase();
  const values = {I:1,V:5,X:10,L:50,C:100,D:500,M:1000};
  let total = 0;
  let previous = 0;

  for (let i = value.length - 1; i >= 0; i--) {
    const current = values[value[i]] || 0;
    if (!current) return 0;

    if (current < previous) total -= current;
    else {
      total += current;
      previous = current;
    }
  }

  return total;
}

function stripAnyListPrefixText_(value) {
  let text = String(value || '');

  text = text.replace(/^\s*[•●○▪◦‣⁃-]\s+/, '');
  text = text.replace(/^\s*\d+[.)]\s+/, '');
  text = text.replace(/^\s*[A-Za-z]+[.)]\s+/, '');
  text = text.replace(/^\s*[ivxlcdm]+[.)]\s+/i, '');

  return text;
}

/**
 * Used by Full Smart Format.
 * Preserves an existing visible ordinal when one is already present.
 */
function applyListFormatToParagraph_(p, type) {
  if (isExistingNativeListOfType_(p, type)) {
    normalizeExistingNativeListParagraph_(p, type);
    return p;
  }

  const original = p.getText();
  const paragraph = recreateAsParagraph_(p);

  applyNamedStyleToParagraph_(paragraph, 'NORMAL');

  const content = stripAnyListPrefixText_(original).trim();

  if (type === 'BULLET') {
    // Full Smart Format must also create a REAL bullet, not a text character.
    const item = convertToSafeNativeBullet_(paragraph, null);
    normalizeExistingNativeListParagraph_(item, 'BULLET');
    return item;
  } else if (type === 'NUMBER') {
    const n = getManualListOrdinal_(original, 'NUMBER') || 1;
    paragraph.setText(String(n) + '. ' + content);
  } else if (type === 'LETTER') {
    const n = getManualListOrdinal_(original, 'LETTER') || 1;
    paragraph.setText(numberToLetters_(n) + ') ' + content);
  } else if (type === 'ROMAN') {
    const n = getManualListOrdinal_(original, 'ROMAN') || 1;
    paragraph.setText(numberToRoman_(n).toLowerCase() + '. ' + content);
  }

  normalizeManualListParagraph_(paragraph);
  applyListIndents_(paragraph);
  return paragraph;
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
  const table = getActiveTable_();
  if (!table) throw new Error('Table formatting is only allowed when the cursor/selection is entirely inside one table.');

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
    let foundTable = null;

    for (let i = 0; i < ranges.length; i++) {
      const table = findAncestorTable_(ranges[i].getElement());

      // Safety: if even one selected element is outside a table,
      // do not guess which table the user intended.
      if (!table) return null;

      if (!foundTable) {
        foundTable = table;
      } else if (table !== foundTable) {
        // Safety: selection spans more than one table.
        return null;
      }
    }

    return foundTable;
  }

  const cursor = doc.getCursor();
  if (cursor) {
    return findAncestorTable_(cursor.getElement());
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
      p.setLineSpacing(1);
      p.setSpacingBefore(0);
      p.setSpacingAfter(0);

      // Table cell paragraph indentation:
      // Left 0.05 in, Right 0 in, Special indent None.
      const tableLeftIndentPt = 0.05 * 72;
      p.setIndentStart(tableLeftIndentPt);
      p.setIndentEnd(0);
      p.setIndentFirstLine(tableLeftIndentPt);

      const t = p.editAsText();
      t.setFontFamily('Arial').setFontSize(9).setBold(isHeader);
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      const p = child.asListItem();
      p.setAlignment(isHeader ? DocumentApp.HorizontalAlignment.CENTER : DocumentApp.HorizontalAlignment.LEFT);
      p.setLineSpacing(1);
      p.setSpacingBefore(0);
      p.setSpacingAfter(0);

      // Table cell paragraph indentation:
      // Left 0.05 in, Right 0 in, Special indent None.
      const tableLeftIndentPt = 0.05 * 72;
      p.setIndentStart(tableLeftIndentPt);
      p.setIndentEnd(0);
      p.setIndentFirstLine(tableLeftIndentPt);

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

  items.forEach(item => {
    let p = targetById[item.id];
    if (!p) return;

    const type = typeById[item.id];

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
        stripManualListPrefix_(p, 'BULLET');
        applyListFormatToParagraph_(p, 'BULLET');
        lists++;
        break;

      case 'number':
        stripManualListPrefix_(p, 'NUMBER');
        applyListFormatToParagraph_(p, 'NUMBER');
        lists++;
        break;

      case 'letter':
        applyListFormatToParagraph_(p, 'LETTER');
        lists++;
        break;

      case 'roman':
        stripManualListPrefix_(p, 'ROMAN');
        applyListFormatToParagraph_(p, 'ROMAN');
        lists++;
        break;

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
          applyNamedStyleToParagraph_(p, 'NORMAL');
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
        applyNamedStyleToParagraph_(p, 'NORMAL');
        normal++;
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
