/**
 * TEST LAB V6.8
 *
 * Non-destructive audit for the production formatting rules. Place the cursor
 * in a caption, equation table, real table, or select list paragraphs.
 */
function runCurrentTest() {
  const started = Date.now();
  const doc = DocumentApp.getActiveDocument();
  const body = getActiveBody_();
  const targets = getStyleTargetParagraphs_();
  const activeTableContext = getActiveEquationTableContext_();
  const activeTable = activeTableContext
    ? activeTableContext.table
    : getActiveTable_();

  const report = {
    ok: true,
    testId: 'FORMATTING-COMPLIANCE-V6.8',
    visibleContentMutations: false,
    mayCreateInternalIndexMarkers: true,
    equationRowUpdateMode: {
      supported: true,
      behavior:
        'Format Equation Row updates the active equation and every later equation.'
    },
    automaticTextColor: {
      policy: 'FOREGROUND_COLOR = null',
      hardCodedBlackText: false,
      explanation:
        'Automatic inherits the Google Docs/Named Style color; it is not forced to #000000.'
    },
    selectedLists: auditSelectedListSpacing_(targets),
    selectedCaption: auditSelectedCaption_(targets),
    activeTable: activeTable
      ? auditActiveTable_(activeTable, body, doc)
      : null,
    sequences: auditCaptionAndEquationSequences_(body)
  };

  const checks = [
    report.selectedLists.ok,
    report.selectedCaption ? report.selectedCaption.ok : true,
    report.activeTable ? report.activeTable.ok : true,
    report.sequences.ok
  ];

  report.ok = checks.every(function(value) { return value !== false; });
  report.elapsedMs = Date.now() - started;
  return report;
}

function auditSelectedListSpacing_(paragraphs) {
  const listItems = (paragraphs || []).filter(function(paragraph) {
    return paragraph.getType() === DocumentApp.ElementType.LIST_ITEM;
  });
  const details = listItems.map(function(item) {
    const spacing = Number(item.asListItem().getLineSpacing());
    return {
      preview: String(item.getText() || '').substring(0, 60),
      lineSpacing: spacing,
      ok: Math.abs(spacing - 1.5) < 0.001
    };
  });

  return {
    ok: details.every(function(item) { return item.ok; }),
    selectedListItems: details.length,
    expectedLineSpacing: 1.5,
    details: details
  };
}

function auditSelectedCaption_(paragraphs) {
  if (!paragraphs || paragraphs.length !== 1) return null;

  const paragraph = paragraphs[0];
  if (paragraph.getType() !== DocumentApp.ElementType.PARAGRAPH) return null;

  const parsed = parseCaptionLine_(paragraph.getText());
  if (!parsed || !/^\d+$/.test(parsed.oldNumber)) return null;

  const prefix = parsed.type + ' ' + parsed.oldNumber;
  const text = paragraph.editAsText();
  const start = String(paragraph.getText()).indexOf(prefix);
  const end = start + prefix.length - 1;
  const bold = testRangeState_(text, start, end, 'isBold', true);
  const italic = testRangeState_(text, start, end, 'isItalic', false);
  const boldRunStarts = testFormattingRunStarts_(text, start, end, 'isBold', true);

  return {
    ok: start >= 0 && bold && italic && boldRunStarts.length === 1,
    type: parsed.type,
    number: Number(parsed.oldNumber),
    prefix: prefix,
    continuousBoldPrefix: bold && boldRunStarts.length === 1,
    boldRunStarts: boldRunStarts,
    prefixNotItalic: italic
  };
}

function auditActiveTable_(table, body, doc) {
  const top = getTopLevelElementForParent_(table, body);
  const bodyIndex = top ? body.getChildIndex(top) : -1;
  const equationMarkers = readEquationMarkerIndices_(body);

  if (
    isEquationLayoutTable_(table) ||
    Boolean(equationMarkers[bodyIndex])
  ) {
    return auditEquationTable_(table, bodyIndex);
  }

  const header = auditFirstTableRow_(table);
  const tableLists = auditTableListSpacing_(table);
  const pinned = auditPinnedHeaderRow_(body, bodyIndex, doc);

  return {
    ok: header.ok && tableLists.ok && pinned,
    kind: 'REAL_TABLE',
    bodyIndex: bodyIndex,
    firstRow: header,
    lists: tableLists,
    pinnedHeaderRow: pinned,
    tableObjectCentering: {
      supportedByPublicGoogleDocsApi: false,
      changedByAddon: false,
      note: 'Cell content can be centered; table-object alignment is not exposed.'
    }
  };
}

function auditFirstTableRow_(table) {
  if (!table.getNumRows()) {
    return {ok: false, cells: 0, allTextBold: false};
  }

  const row = table.getRow(0);
  let textCharacters = 0;
  let boldCharacters = 0;

  for (let cellIndex = 0; cellIndex < row.getNumCells(); cellIndex++) {
    const text = row.getCell(cellIndex).editAsText();
    const value = text.getText();

    for (let offset = 0; offset < value.length; offset++) {
      if (!value.charAt(offset).trim()) continue;
      textCharacters++;
      if (text.isBold(offset) === true) boldCharacters++;
    }
  }

  return {
    ok: textCharacters === 0 || boldCharacters === textCharacters,
    cells: row.getNumCells(),
    textCharacters: textCharacters,
    boldCharacters: boldCharacters,
    allTextBold: textCharacters === 0 || boldCharacters === textCharacters
  };
}

function auditTableListSpacing_(table) {
  let listItems = 0;
  let correct = 0;

  for (let rowIndex = 0; rowIndex < table.getNumRows(); rowIndex++) {
    const row = table.getRow(rowIndex);

    for (let cellIndex = 0; cellIndex < row.getNumCells(); cellIndex++) {
      const cell = row.getCell(cellIndex);

      for (let childIndex = 0; childIndex < cell.getNumChildren(); childIndex++) {
        const child = cell.getChild(childIndex);
        if (child.getType() !== DocumentApp.ElementType.LIST_ITEM) continue;

        listItems++;
        if (Math.abs(Number(child.asListItem().getLineSpacing()) - 1.5) < 0.001) {
          correct++;
        }
      }
    }
  }

  return {
    ok: correct === listItems,
    listItems: listItems,
    correctLineSpacing: correct,
    expectedLineSpacing: 1.5
  };
}

function auditEquationTable_(table, bodyIndex) {
  const rightCell = table.getRow(0).getCell(2);
  const value = String(rightCell.getText() || '').trim();
  const match = value.match(/\bEquation\s+(\d+)\s*$/i);

  if (!match) {
    return {
      ok: false,
      kind: 'EQUATION_TABLE',
      bodyIndex: bodyIndex,
      labelRecognized: false
    };
  }

  const label = 'Equation\u00A0' + match[1];
  const text = rightCell.getChild(0).asParagraph().editAsText();
  const start = value.indexOf(label);
  const end = start + label.length - 1;
  const bold = testRangeState_(text, start, end, 'isBold', true);
  const notItalic = testRangeState_(text, start, end, 'isItalic', false);
  const boldRunStarts = testFormattingRunStarts_(text, start, end, 'isBold', true);

  return {
    ok: bold && notItalic && boldRunStarts.length === 1,
    kind: 'EQUATION_TABLE',
    bodyIndex: bodyIndex,
    number: Number(match[1]),
    continuousBoldLabel: bold && boldRunStarts.length === 1,
    boldRunStarts: boldRunStarts,
    labelNotItalic: notItalic
  };
}

function auditPinnedHeaderRow_(body, bodyIndex, doc) {
  const plan = buildTableHeaderPinPlan_(body, [bodyIndex]);
  if (!plan.tableOrdinals.length) return false;

  const apiDocument = Docs.Documents.get(doc.getId(), {
    includeTabsContent: true
  });
  const apiTab = getApiDocumentTab_(apiDocument, plan.tabId);
  const tables = (apiTab.body.content || []).filter(function(element) {
    return element && element.table;
  });
  const structural = tables[plan.tableOrdinals[0]];
  const firstRow = structural && structural.table &&
    structural.table.tableRows && structural.table.tableRows[0];

  return Boolean(
    firstRow &&
    firstRow.tableRowStyle &&
    firstRow.tableRowStyle.tableHeader === true
  );
}

function auditCaptionAndEquationSequences_(body) {
  const objectIndex = buildNeededDocumentIndex_(
    body,
    {
      tableIndex: true,
      figureIndex: true,
      equationIndex: true
    },
    []
  );
  const mismatches = [];
  let captions = 0;

  for (let bodyIndex = 0; bodyIndex < body.getNumChildren(); bodyIndex++) {
    const element = body.getChild(bodyIndex);
    if (element.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const paragraph = element.asParagraph();
    const parsed = parseCaptionLine_(paragraph.getText());
    if (!parsed || (parsed.type !== 'Table' && parsed.type !== 'Figure')) continue;

    captions++;
    const expected = getIndexedCaptionOrdinal_(
      paragraph,
      parsed.type,
      bodyIndex,
      body,
      objectIndex
    );
    const actual = /^\d+$/.test(parsed.oldNumber)
      ? Number(parsed.oldNumber)
      : null;

    if (expected !== null && expected !== undefined && actual !== expected) {
      if (mismatches.length < 20) {
        mismatches.push({
          bodyIndex: bodyIndex,
          type: parsed.type,
          actual: actual,
          expected: expected
        });
      }
    }
  }

  objectIndex.equationIndices.forEach(function(bodyIndex, position) {
    const value = body.getChild(bodyIndex).asTable().getRow(0).getCell(2).getText();
    const match = String(value || '').match(/\bEquation\s+(\d+)\s*$/i);
    const actual = match ? Number(match[1]) : null;
    const expected = position + 1;

    if (actual !== expected && mismatches.length < 20) {
      mismatches.push({
        bodyIndex: bodyIndex,
        type: 'Equation',
        actual: actual,
        expected: expected
      });
    }
  });

  return {
    ok: mismatches.length === 0,
    captionsAudited: captions,
    equationsAudited: objectIndex.equationIndices.length,
    figuresIndexed: objectIndex.figureIndices.length,
    inlineImagesCollected: objectIndex.figureImageCollectionCount,
    directInlineFigureBlocks: objectIndex.figureDirectBlockCount,
    mismatches: mismatches,
    tableIndexMs: objectIndex.tableIndexMs,
    figureIndexMs: objectIndex.figureIndexMs
  };
}

function testRangeState_(text, start, end, methodName, expected) {
  if (start < 0 || end < start) return false;

  for (let offset = start; offset <= end; offset++) {
    if (text[methodName](offset) !== expected) return false;
  }

  return true;
}

function testFormattingRunStarts_(text, start, end, methodName, expected) {
  if (start < 0 || end < start) return [];

  return text.getTextAttributeIndices().filter(function(offset) {
    return (
      offset >= start &&
      offset <= end &&
      text[methodName](offset) === expected
    );
  });
}
