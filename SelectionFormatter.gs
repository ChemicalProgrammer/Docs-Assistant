/**
 * Deterministic selected-section formatter.
 *
 * Production engine used by the “Format selected styles” button.
 * Uses the fast local formatters. The only advanced-API operation is one
 * optional batch that pins the first row of selected real tables.
 */
function formatSelectedSection_() {
  const started = Date.now();
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  if (!selection) {
    throw new Error('Select the section you want to format.');
  }

  const body = getActiveBody_();

  const selectionAnalysisStarted = Date.now();
  const selectedRange = getSelectedBodyRange_(selection, body);
  const selectionPlan = buildSectionSelectionPlan_(body, selectedRange);
  const selectionAnalysisMs = Date.now() - selectionAnalysisStarted;

  const indexStarted = Date.now();
  const objectIndex = buildNeededDocumentIndex_(
    body,
    selectionPlan.requirements,
    selectionPlan.selectedEquationEntries
  );
  const indexMs = Date.now() - indexStarted;

  const report = {
    ok: true,
    engineId: 'DETERMINISTIC-SECTION-FORMAT-V6.8',
    normal: 0,
    blank: 0,
    headings: 0,
    titleOrSubtitle: 0,
    tables: 0,
    tableCells: 0,
    equations: 0,
    equationNumbersCorrected: 0,
    lists: 0,
    listRuns: 0,
    notes: 0,
    figures: 0,
    tableCaptions: 0,
    captionsSkippedBeforeAnchor: 0,
    skipped: 0,
    usesGemini: false,
    usesAdvancedDocsApi: false
  };

  const listActions = [];
  const selectedTableBodyIndices = [];
  const mainFormattingStarted = Date.now();

  selectionPlan.entries.forEach(function(entry) {
    const element = entry.element;
    const bodyIndex = entry.bodyIndex;

    if (entry.kind === 'EQUATION_TABLE') {
      report.equations++;
      return;
    }

    if (entry.kind === 'TABLE') {
      const tableResult = formatTableElement_(element.asTable());
      selectedTableBodyIndices.push(bodyIndex);
      report.tables++;
      report.tableCells += tableResult.cells;
      return;
    }

    if (entry.kind === 'SKIP') {
      report.skipped++;
      return;
    }

    const paragraph = element;
    const classification = entry.classification;

    if (classification.kind === 'LIST') {
      listActions.push({
        bodyIndex: bodyIndex,
        paragraph: paragraph,
        listType: classification.listType
      });
      return;
    }

    if (classification.kind === 'NOTE') {
      formatNoteParagraph_(
        paragraph,
        parseNoteLine_(paragraph.getText()).description
      );
      report.notes++;
      return;
    }

    if (
      classification.kind === 'FIGURE_CAPTION' ||
      classification.kind === 'TABLE_CAPTION'
    ) {
      const captionType = classification.kind === 'FIGURE_CAPTION'
        ? 'Figure'
        : 'Table';
      const parsed = parseCaptionLine_(paragraph.getText());
      const number = getIndexedCaptionOrdinal_(
        paragraph,
        captionType,
        bodyIndex,
        body,
        objectIndex
      );

      if (number === null || number === undefined) {
        report.captionsSkippedBeforeAnchor++;
        return;
      }

      formatCaptionParagraph_(
        paragraph,
        captionType,
        parsed.description,
        number
      );

      if (captionType === 'Figure') report.figures++;
      else report.tableCaptions++;

      return;
    }

    applyStyleToParagraph_(paragraph, classification.styleName);

    if (classification.blank) {
      report.blank++;
    } else if (/^H[1-6]$/.test(classification.styleName)) {
      report.headings++;
    } else if (
      classification.styleName === 'TITLE' ||
      classification.styleName === 'SUBTITLE'
    ) {
      report.titleOrSubtitle++;
    } else {
      report.normal++;
    }
  });

  if (selectionPlan.selectedEquationEntries.length) {
    const firstSelectedEquationIndex = Math.min.apply(
      null,
      selectionPlan.selectedEquationEntries.map(function(entry) {
        return entry.bodyIndex;
      })
    );

    objectIndex.equationIndices.forEach(function(bodyIndex, position) {
      if (bodyIndex < firstSelectedEquationIndex) return;

      const equationTable = body.getChild(bodyIndex).asTable();
      applyEquationTableColumnWidths_(
        equationTable,
        body,
        position + 1
      );
      const equationResult = formatEquationLabel_(
        equationTable.getRow(0).getCell(2),
        position + 1
      );

      if (equationResult.numberCorrected) {
        report.equationNumbersCorrected++;
      }
    });
  }

  const mainFormattingMs = Date.now() - mainFormattingStarted;
  const listFormattingStarted = Date.now();

  const listRuns = groupContiguousListActions_(listActions);

  listRuns.forEach(function(run) {
    run.actions.forEach(function(action) {
      stripManualListPrefix_(action.paragraph, run.listType);
    });

    applyFastNativeListToParagraphs_(
      body,
      run.actions.map(function(action) {
        return action.paragraph;
      }),
      run.listType,
      null
    );

    report.lists += run.actions.length;
    report.listRuns++;
  });

  const listFormattingMs = Date.now() - listFormattingStarted;
  const headerPinPlan = buildTableHeaderPinPlan_(
    body,
    selectedTableBodyIndices
  );
  const saveStarted = Date.now();
  doc.saveAndClose();
  const saveMs = Date.now() - saveStarted;
  const headerPinStarted = Date.now();
  const headerPinResult = executeTableHeaderPinPlan_(headerPinPlan);
  const headerPinMs = Date.now() - headerPinStarted;

  report.usesAdvancedDocsApi = headerPinResult.usesDocsApi;
  report.tableHeaderRowsPinned = headerPinResult.pinnedTables;
  report.tableAlignment = 'UNCHANGED_NOT_EXPOSED_BY_GOOGLE_DOCS_API';

  report.selectedBodyElements = selectionPlan.entries.length;
  report.tableIndexBuilt = selectionPlan.requirements.tableIndex;
  report.figureIndexBuilt = selectionPlan.requirements.figureIndex;
  report.equationIndexBuilt = selectionPlan.requirements.equationIndex;
  report.documentPhysicalTables = objectIndex.physicalTables;
  report.documentTableCollectionSize = objectIndex.tableCollectionCount;
  report.documentRealTables = objectIndex.tableIndices.length;
  report.documentEquationTables = objectIndex.equationIndices.length;
  report.documentFigures = objectIndex.figureIndices.length;
  report.documentInlineImages = objectIndex.figureImageCollectionCount;
  report.documentDirectInlineFigureBlocks =
    objectIndex.figureDirectBlockCount;
  report.equationMarkerMigrationPerformed =
    objectIndex.markerMigrationPerformed;
  report.equationMarkerCount = objectIndex.equationIndices.length;
  report.figureMarkerMigrationPerformed =
    objectIndex.figureMarkerMigrationPerformed;
  report.figureMarkerCount = objectIndex.figurePersistentMarkerCount;
  report.selectionAnalysisMs = selectionAnalysisMs;
  report.equationMarkerMigrationMs = objectIndex.markerMigrationMs;
  report.figureMarkerMigrationMs = objectIndex.figureMarkerMigrationMs;
  report.tableIndexMs = objectIndex.tableIndexMs;
  report.figureIndexMs = objectIndex.figureIndexMs;
  report.indexMs = indexMs;
  report.mainFormattingMs = mainFormattingMs;
  report.listFormattingMs = listFormattingMs;
  report.saveMs = saveMs;
  report.tableHeaderPinMs = headerPinMs;
  report.elapsedMs = Date.now() - started;

  return report;
}

function buildSectionSelectionPlan_(body, selectedRange) {
  const entries = [];
  const selectedEquationEntries = [];
  const requirements = {
    tableIndex: false,
    figureIndex: false,
    equationIndex: false
  };

  for (
    let bodyIndex = selectedRange.firstIndex;
    bodyIndex <= selectedRange.lastIndex;
    bodyIndex++
  ) {
    const element = body.getChild(bodyIndex);
    const elementType = element.getType();

    if (elementType === DocumentApp.ElementType.TABLE) {
      const isEquation = isEquationLayoutTable_(element.asTable());
      const entry = {
        kind: isEquation ? 'EQUATION_TABLE' : 'TABLE',
        bodyIndex: bodyIndex,
        element: element
      };

      entries.push(entry);

      if (isEquation) {
        selectedEquationEntries.push(entry);
        requirements.equationIndex = true;
      }

      continue;
    }

    if (
      elementType !== DocumentApp.ElementType.PARAGRAPH &&
      elementType !== DocumentApp.ElementType.LIST_ITEM
    ) {
      entries.push({
        kind: 'SKIP',
        bodyIndex: bodyIndex,
        element: element
      });
      continue;
    }

    const classification = classifySectionParagraph_(element);
    entries.push({
      kind: 'PARAGRAPH',
      bodyIndex: bodyIndex,
      element: element,
      classification: classification
    });

    if (classification.kind === 'TABLE_CAPTION') {
      requirements.tableIndex = true;
    } else if (classification.kind === 'FIGURE_CAPTION') {
      requirements.figureIndex = true;
    }
  }

  return {
    entries: entries,
    selectedEquationEntries: selectedEquationEntries,
    requirements: requirements
  };
}

function getSelectedBodyRange_(selection, body) {
  const ranges = selection.getRangeElements();
  if (!ranges.length) {
    throw new Error('The selection is empty.');
  }

  const firstElement = getTopLevelElementForParent_(
    ranges[0].getElement(),
    body
  );
  const lastElement = getTopLevelElementForParent_(
    ranges[ranges.length - 1].getElement(),
    body
  );

  if (!firstElement || !lastElement) {
    throw new Error('The selection must be inside the active document body.');
  }

  const first = body.getChildIndex(firstElement);
  const last = body.getChildIndex(lastElement);

  return {
    firstIndex: Math.min(first, last),
    lastIndex: Math.max(first, last)
  };
}

/**
 * Construye únicamente los índices que necesita la selección actual.
 */
function buildNeededDocumentIndex_(
  body,
  requirements,
  selectedEquationEntries
) {
  const tableIndices = [];
  const equationIndices = [];
  const equationNumberByBodyIndex = {};
  const figureIndices = [];
  let physicalTables = null;
  let tableCollectionCount = null;
  let markerMigrationPerformed = false;
  let markerMigrationMs = 0;
  let figureMarkerMigrationPerformed = false;
  let figureMarkerMigrationMs = 0;
  let figureImageCollectionCount = null;
  let figureDirectBlockCount = null;
  let figurePersistentMarkerCount = null;
  let tableIndexMs = 0;
  let figureIndexMs = 0;

  const needsTableObjects =
    requirements.tableIndex || requirements.equationIndex;

  if (needsTableObjects) {
    const tableIndexStarted = Date.now();
    const markerResult = ensureEquationTableMarkers_(
      body,
      selectedEquationEntries
    );

    markerMigrationPerformed = markerResult.migrationPerformed;
    markerMigrationMs = markerResult.elapsedMs;
    const tables = body.getTables() || [];
    const seenTableIndices = {};
    tableCollectionCount = tables.length;
    physicalTables = 0;

    /*
     * getTables() obtiene directamente la colección de tablas de la sección.
     * Evita recorrer uno por uno todos los elementos del documento.
     * Si existiera una tabla anidada, se conserva únicamente su tabla física
     * de nivel superior para no alterar el ordinal de las leyendas.
     */
    tables.forEach(function(table) {
      const top = getTopLevelElementForParent_(table, body);

      if (!top || top.getType() !== DocumentApp.ElementType.TABLE) return;

      const bodyIndex = body.getChildIndex(top);
      if (seenTableIndices[bodyIndex]) return;

      seenTableIndices[bodyIndex] = true;
      physicalTables++;

      if (markerResult.markerIndexSet[bodyIndex]) {
        equationIndices.push(bodyIndex);
      } else {
        tableIndices.push(bodyIndex);
      }
    });

    tableIndices.sort(function(a, b) { return a - b; });
    equationIndices.sort(function(a, b) { return a - b; });

    equationIndices.forEach(function(bodyIndex, position) {
      equationNumberByBodyIndex[bodyIndex] = position + 1;
    });

    tableIndexMs = Date.now() - tableIndexStarted;
  }

  if (requirements.figureIndex) {
    const figureIndexStarted = Date.now();
    const figureMarkerResult = ensureFigureBlockMarkers_(body);

    figureMarkerMigrationPerformed =
      figureMarkerResult.migrationPerformed;
    figureMarkerMigrationMs = figureMarkerResult.elapsedMs;
    figureImageCollectionCount =
      figureMarkerResult.imageCollectionCount;
    figureDirectBlockCount =
      figureMarkerResult.directBlockCount;
    figurePersistentMarkerCount =
      figureMarkerResult.persistentMarkerCount;

    Object.keys(figureMarkerResult.markerIndexSet)
      .map(Number)
      .sort(function(a, b) { return a - b; })
      .forEach(function(bodyIndex) {
        figureIndices.push(bodyIndex);
      });

    figureIndexMs = Date.now() - figureIndexStarted;
  }

  return {
    physicalTables: physicalTables,
    tableCollectionCount: tableCollectionCount,
    tableIndices: tableIndices,
    equationIndices: equationIndices,
    equationNumberByBodyIndex: equationNumberByBodyIndex,
    figureIndices: figureIndices,
    markerMigrationPerformed: markerMigrationPerformed,
    markerMigrationMs: markerMigrationMs,
    figureMarkerMigrationPerformed: figureMarkerMigrationPerformed,
    figureMarkerMigrationMs: figureMarkerMigrationMs,
    figureImageCollectionCount: figureImageCollectionCount,
    figureDirectBlockCount: figureDirectBlockCount,
    figurePersistentMarkerCount: figurePersistentMarkerCount,
    tableIndexMs: tableIndexMs,
    figureIndexMs: figureIndexMs
  };
}

const SECTION_EQUATION_MARKER_NAME_ =
  'DOCS_ASSISTANT_EQUATION_LAYOUT_MARKER';

/**
 * Las tablas de ecuación se detectan estructuralmente una sola vez. Después
 * quedan identificadas mediante NamedRanges que se desplazan con el contenido.
 */
function ensureEquationTableMarkers_(body, selectedEquationEntries) {
  const started = Date.now();
  const markerIndexSet = readEquationMarkerIndices_(body);
  const migrationProperty = getEquationMarkerMigrationProperty_();
  const properties = PropertiesService.getDocumentProperties();
  const migrationAlreadyDone = properties &&
    properties.getProperty(migrationProperty) === '1';
  let migrationPerformed = false;

  /*
   * Una ecuación seleccionada se marca inmediatamente. Esto también protege
   * ecuaciones nuevas creadas durante las pruebas.
   */
  (selectedEquationEntries || []).forEach(function(entry) {
    if (markerIndexSet[entry.bodyIndex]) return;

    addEquationTableMarker_(entry.element.asTable());
    markerIndexSet[entry.bodyIndex] = true;
  });

  if (!migrationAlreadyDone) {
    migrationPerformed = true;

    getTopLevelTableEntries_(body).forEach(function(entry) {
      if (markerIndexSet[entry.bodyIndex]) return;

      if (isEquationLayoutTable_(entry.table)) {
        addEquationTableMarker_(entry.table);
        markerIndexSet[entry.bodyIndex] = true;
      }
    });

    if (properties) {
      properties.setProperty(migrationProperty, '1');
    }
  }

  return {
    markerIndexSet: markerIndexSet,
    migrationPerformed: migrationPerformed,
    elapsedMs: Date.now() - started
  };
}

function readEquationMarkerIndices_(body) {
  const indexSet = {};
  const namedRanges = getActiveTabNamedRanges_(SECTION_EQUATION_MARKER_NAME_);

  namedRanges.forEach(function(namedRange) {
    try {
      const ranges = namedRange.getRange().getRangeElements();

      ranges.forEach(function(rangeElement) {
        const top = getTopLevelElementForParent_(
          rangeElement.getElement(),
          body
        );

        if (!top || top.getType() !== DocumentApp.ElementType.TABLE) return;
        indexSet[body.getChildIndex(top)] = true;
      });
    } catch (error) {}
  });

  return indexSet;
}

function addEquationTableMarker_(table) {
  const rightCell = table.getRow(0).getCell(2);
  const paragraph = rightCell.getChild(0).asParagraph();
  addActiveTabNamedRange_(SECTION_EQUATION_MARKER_NAME_, paragraph);
}

function getEquationMarkerMigrationProperty_() {
  let tabId = 'LEGACY';

  try {
    const tab = DocumentApp.getActiveDocument().getActiveTab();
    if (tab && tab.getId) tabId = String(tab.getId());
  } catch (error) {}

  return (
    'DOCS_ASSISTANT_EQUATION_MARKERS_MIGRATED_V2_' +
    tabId.replace(/[^A-Za-z0-9_-]/g, '_')
  );
}

const SECTION_FIGURE_MARKER_NAME_ =
  'DOCS_ASSISTANT_FIGURE_BLOCK_MARKER';

/**
 * Detecta todos los bloques de figura una sola vez y los identifica mediante
 * NamedRanges. Después, la numeración obtiene sus posiciones desde esos
 * marcadores sin volver a abrir cada párrafo del documento.
 */
function ensureFigureBlockMarkers_(body) {
  const started = Date.now();
  const markerIndexSet = readFigureMarkerIndices_(body);
  let persistentMarkerCount = Object.keys(markerIndexSet).length;
  const directResult = mergeDirectInlineFigureIndices_(
    body,
    markerIndexSet
  );
  const migrationProperty = getFigureMarkerMigrationProperty_();
  const properties = PropertiesService.getDocumentProperties();
  const migrationAlreadyDone = properties &&
    properties.getProperty(migrationProperty) === '1';
  let migrationPerformed = false;

  if (!migrationAlreadyDone) {
    migrationPerformed = true;

    for (let bodyIndex = 0; bodyIndex < body.getNumChildren(); bodyIndex++) {
      if (markerIndexSet[bodyIndex]) continue;

      const element = body.getChild(bodyIndex);
      if (!isStandaloneFigureBlock_(element)) continue;

      addActiveTabNamedRange_(SECTION_FIGURE_MARKER_NAME_, element);
      markerIndexSet[bodyIndex] = true;
      persistentMarkerCount++;
    }

    if (properties) {
      properties.setProperty(migrationProperty, '1');
    }
  }

  return {
    markerIndexSet: markerIndexSet,
    migrationPerformed: migrationPerformed,
    imageCollectionCount: directResult.imageCollectionCount,
    directBlockCount: directResult.directBlockCount,
    persistentMarkerCount: persistentMarkerCount,
    elapsedMs: Date.now() - started
  };
}

/**
 * Adds every current body-level InlineImage to the Figure index.
 *
 * Unlike the one-time marker migration, Body.getImages() reflects images
 * inserted after the migration. It also avoids reopening every paragraph in
 * a large document. Images inside tables are excluded because their top-level
 * element is a TABLE, not a standalone figure paragraph.
 */
function mergeDirectInlineFigureIndices_(body, indexSet) {
  const targetIndexSet = indexSet || {};
  const directBlockSet = {};
  let images = [];

  try {
    images = body.getImages() || [];
  } catch (error) {
    images = [];
  }

  images.forEach(function(image) {
    const top = getTopLevelElementForParent_(image, body);
    if (!top || !isStandaloneFigureBlock_(top)) return;

    const bodyIndex = body.getChildIndex(top);
    targetIndexSet[bodyIndex] = true;
    directBlockSet[bodyIndex] = true;
  });

  return {
    imageCollectionCount: images.length,
    directBlockCount: Object.keys(directBlockSet).length
  };
}

/**
 * Persists a marker for a positioned image or drawing discovered locally at
 * the active Figure anchor. Inline images are also accepted, but their direct
 * Body collection already guarantees that they participate in every index.
 */
function ensurePersistentFigureBlockMarker_(body, element) {
  const top = getTopLevelElementForParent_(element, body);

  if (!top || !isStandaloneFigureBlock_(top)) {
    return {
      bodyIndex: -1,
      markerAdded: false
    };
  }

  const bodyIndex = body.getChildIndex(top);

  // Body.getImages() already refreshes ordinary InlineImages on every index
  // build, including images inserted after the original marker migration.
  // Avoid a second NamedRange read for this common and fast path.
  if (figureBlockContainsInlineImage_(top)) {
    return {
      bodyIndex: bodyIndex,
      markerAdded: false,
      coveredByDirectCollection: true
    };
  }

  const existing = readFigureMarkerIndices_(body);

  if (!existing[bodyIndex]) {
    addActiveTabNamedRange_(SECTION_FIGURE_MARKER_NAME_, top);
    return {
      bodyIndex: bodyIndex,
      markerAdded: true,
      coveredByDirectCollection: false
    };
  }

  return {
    bodyIndex: bodyIndex,
    markerAdded: false,
    coveredByDirectCollection: false
  };
}

function figureBlockContainsInlineImage_(element) {
  try {
    for (let index = 0; index < element.getNumChildren(); index++) {
      if (
        element.getChild(index).getType() ===
        DocumentApp.ElementType.INLINE_IMAGE
      ) {
        return true;
      }
    }
  } catch (error) {}

  return false;
}

function readFigureMarkerIndices_(body) {
  const indexSet = {};
  const namedRanges = getActiveTabNamedRanges_(SECTION_FIGURE_MARKER_NAME_);

  namedRanges.forEach(function(namedRange) {
    try {
      const ranges = namedRange.getRange().getRangeElements();

      ranges.forEach(function(rangeElement) {
        const top = getTopLevelElementForParent_(
          rangeElement.getElement(),
          body
        );

        if (!top) return;

        /*
         * Un NamedRange de párrafo puede ampliar sus límites cuando se edita
         * la leyenda contigua. Por eso no basta con aceptar PARAGRAPH: se
         * vuelve a comprobar que el elemento realmente contenga una imagen o
         * dibujo independiente. Esto impide que las leyendas se cuenten como
         * figuras y que el índice crezca de 46 a 56 tras formatearlas.
         */
        if (!isStandaloneFigureBlock_(top)) return;

        indexSet[body.getChildIndex(top)] = true;
      });
    } catch (error) {}
  });

  return indexSet;
}

function getFigureMarkerMigrationProperty_() {
  let tabId = 'LEGACY';

  try {
    const tab = DocumentApp.getActiveDocument().getActiveTab();
    if (tab && tab.getId) tabId = String(tab.getId());
  } catch (error) {}

  return (
    'DOCS_ASSISTANT_FIGURE_MARKERS_MIGRATED_V2_' +
    tabId.replace(/[^A-Za-z0-9_-]/g, '_')
  );
}

function classifySectionParagraph_(paragraph) {
  const text = String(paragraph.getText() || '');
  const trimmed = text.trim();

  if (!trimmed) {
    return {
      kind: 'STYLE',
      styleName: 'NORMAL',
      blank: true
    };
  }

  const caption = parseCaptionLine_(trimmed);
  if (caption) {
    return {
      kind: caption.type === 'Figure'
        ? 'FIGURE_CAPTION'
        : 'TABLE_CAPTION'
    };
  }

  if (/^\s*(?:Notes?|Notas?)\b(?:\s*[.:–—-]|\s+)/i.test(trimmed)) {
    return {kind: 'NOTE'};
  }

  /*
   * La numeración jerárquica tiene prioridad sobre un Heading incorrecto.
   * 2.3 -> H2; 2.3.1 -> H3.
   */
  const section = trimmed.match(/^\s*(\d+(?:\.\d+)+)\.?\s+\S/);
  if (section) {
    const depth = Math.min(section[1].split('.').length, 6);
    return {
      kind: 'STYLE',
      styleName: 'H' + depth,
      blank: false
    };
  }

  const existingStyle = getExistingNamedStyleName_(paragraph);
  if (
    /^H[1-6]$/.test(existingStyle) ||
    existingStyle === 'TITLE' ||
    existingStyle === 'SUBTITLE'
  ) {
    return {
      kind: 'STYLE',
      styleName: existingStyle,
      blank: false
    };
  }

  const existingList = getCurrentListType_(paragraph);
  if (existingList) {
    return {
      kind: 'LIST',
      listType: existingList.toUpperCase()
    };
  }

  if (/^\s*[•●○▪◦‣⁃-]\s+/.test(trimmed)) {
    return {kind: 'LIST', listType: 'BULLET'};
  }

  if (/^\s*[ivxlcdm]+[.)]\s+/i.test(trimmed)) {
    return {kind: 'LIST', listType: 'ROMAN'};
  }

  if (/^\s*[A-Za-zÁÉÍÓÚÜÑáéíóúüñ][.)]\s+/.test(trimmed)) {
    return {kind: 'LIST', listType: 'LETTER'};
  }

  if (/^\s*\d+[.)]\s+/.test(trimmed)) {
    return {kind: 'LIST', listType: 'NUMBER'};
  }

  return {
    kind: 'STYLE',
    styleName: 'NORMAL',
    blank: false
  };
}

function groupContiguousListActions_(actions) {
  const runs = [];

  actions.forEach(function(action) {
    const lastRun = runs.length ? runs[runs.length - 1] : null;
    const previousAction = lastRun && lastRun.actions.length
      ? lastRun.actions[lastRun.actions.length - 1]
      : null;

    if (
      lastRun &&
      lastRun.listType === action.listType &&
      previousAction.bodyIndex + 1 === action.bodyIndex
    ) {
      lastRun.actions.push(action);
      return;
    }

    runs.push({
      listType: action.listType,
      actions: [action]
    });
  });

  return runs;
}

function formatTableElement_(table) {
  const PT_PER_IN = 72;
  const attribute = DocumentApp.Attribute;
  const cellAttributes = {};

  cellAttributes[attribute.VERTICAL_ALIGNMENT] =
    DocumentApp.VerticalAlignment.CENTER;
  cellAttributes[attribute.PADDING_TOP] = 0.028 * PT_PER_IN;
  cellAttributes[attribute.PADDING_BOTTOM] = 0.028 * PT_PER_IN;
  cellAttributes[attribute.PADDING_LEFT] = 0.028 * PT_PER_IN;
  cellAttributes[attribute.PADDING_RIGHT] = 0.028 * PT_PER_IN;

  const headerParagraph = buildTableParagraphAttributes_(true, false, PT_PER_IN);
  const bodyParagraph = buildTableParagraphAttributes_(false, false, PT_PER_IN);
  const headerList = buildTableParagraphAttributes_(true, true, PT_PER_IN);
  const bodyList = buildTableParagraphAttributes_(false, true, PT_PER_IN);

  table.setBorderColor('#000000');
  table.setBorderWidth(1);

  let cells = 0;
  let paragraphs = 0;
  let listItems = 0;

  for (let rowIndex = 0; rowIndex < table.getNumRows(); rowIndex++) {
    const row = table.getRow(rowIndex);
    const isHeader = rowIndex === 0;
    row.setMinimumHeight(0.49 * PT_PER_IN);

    for (let cellIndex = 0; cellIndex < row.getNumCells(); cellIndex++) {
      const cell = row.getCell(cellIndex);
      cell.setAttributes(cellAttributes);
      cell.setBackgroundColor(null);
      cells++;

      for (let childIndex = 0; childIndex < cell.getNumChildren(); childIndex++) {
        const child = cell.getChild(childIndex);

        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
          child.asParagraph().setAttributes(
            isHeader ? headerParagraph : bodyParagraph
          );
          paragraphs++;
        } else if (child.getType() === DocumentApp.ElementType.LIST_ITEM) {
          child.asListItem().setAttributes(
            isHeader ? headerList : bodyList
          );
          listItems++;
        }
      }
    }
  }

  return {
    rows: table.getNumRows(),
    cells: cells,
    paragraphs: paragraphs,
    listItems: listItems
  };
}

function formatEquationLayoutNumber_(table, equationNumber) {
  return formatEquationLabel_(
    table.getRow(0).getCell(2),
    equationNumber
  );
}

function getIndexedCaptionOrdinal_(
  paragraph,
  type,
  captionBodyIndex,
  body,
  objectIndex
) {
  const anchorLineIndex = getCaptionCounterAnchorIndex_(type, body);

  /*
   * Once a Figure anchor is explicitly set, the user's captions define the
   * numbering sequence. Count valid Figure captions from that anchor and
   * always count the active line as the next caption, even when it is still
   * plain unnumbered text. Extra image blocks inside a composite figure do not
   * advance the caption counter.
   */
  if (type === 'Figure' && anchorLineIndex >= 0) {
    const anchorCaptionIndex = getFigureAnchorCaptionBodyIndex_(
      body,
      anchorLineIndex
    );

    if (anchorCaptionIndex >= 0) {
      if (captionBodyIndex < anchorCaptionIndex) return null;

      let captionCount = 0;

      for (
        let bodyIndex = anchorCaptionIndex;
        bodyIndex <= captionBodyIndex;
        bodyIndex++
      ) {
        const element = body.getChild(bodyIndex);
        const elementType = element.getType();
        const canBeCaption =
          elementType === DocumentApp.ElementType.PARAGRAPH ||
          elementType === DocumentApp.ElementType.LIST_ITEM;
        const parsed = canBeCaption
          ? parseCaptionLine_(element.getText())
          : null;
        const isExistingFigureCaption =
          parsed && parsed.type === 'Figure';

        if (
          isExistingFigureCaption ||
          (bodyIndex === captionBodyIndex && !isExistingFigureCaption)
        ) {
          captionCount++;
        }
      }

      if (captionCount > 0) {
        return getCaptionCounterStart_('Figure') + captionCount - 1;
      }
    }
  }

  const indices = type === 'Table'
    ? objectIndex.tableIndices
    : objectIndex.figureIndices;
  const objectPosition = findCaptionObjectPosition_(
    indices,
    captionBodyIndex,
    type
  );

  if (objectPosition < 0) {
    return getCaptionOrdinalByPreviousCaptions_(paragraph, type);
  }

  if (anchorLineIndex < 0) {
    return objectPosition + 1;
  }

  const anchorObjectPosition = findCaptionObjectPosition_(
    indices,
    anchorLineIndex,
    type
  );

  if (anchorObjectPosition < 0 || objectPosition < anchorObjectPosition) {
    return null;
  }

  return (
    getCaptionCounterStart_(type) +
    objectPosition -
    anchorObjectPosition
  );
}

/**
 * Resolves a Figure counter anchor to its caption line. Usually the NamedRange
 * already points at "Figure N.". When Set here was pressed on the image
 * paragraph instead, locate the nearest preceding Figure caption associated
 * with that same image block.
 */
function getFigureAnchorCaptionBodyIndex_(body, anchorLineIndex) {
  const anchorElement = body.getChild(anchorLineIndex);
  const anchorType = anchorElement.getType();

  if (
    anchorType === DocumentApp.ElementType.PARAGRAPH ||
    anchorType === DocumentApp.ElementType.LIST_ITEM
  ) {
    const parsedAnchor = parseCaptionLine_(anchorElement.getText());
    if (parsedAnchor && parsedAnchor.type === 'Figure') {
      return anchorLineIndex;
    }
  }

  const anchorObjectIndex = findCaptionObjectBodyIndex_(
    'Figure',
    body,
    anchorLineIndex
  );

  if (anchorObjectIndex < 0) return -1;

  for (let bodyIndex = anchorObjectIndex; bodyIndex >= 0; bodyIndex--) {
    const element = body.getChild(bodyIndex);
    const elementType = element.getType();

    if (
      elementType !== DocumentApp.ElementType.PARAGRAPH &&
      elementType !== DocumentApp.ElementType.LIST_ITEM
    ) {
      continue;
    }

    const parsed = parseCaptionLine_(element.getText());
    if (!parsed || parsed.type !== 'Figure') continue;

    if (
      findCaptionObjectBodyIndex_('Figure', body, bodyIndex) ===
      anchorObjectIndex
    ) {
      return bodyIndex;
    }
  }

  return -1;
}

function findCaptionObjectPosition_(indices, referenceIndex, type) {
  if (!indices || !indices.length) return -1;

  if (type === 'Figure') {
    // Figure captions belong to the first image at or after the caption.
    // Equality supports positioned images anchored to the caption paragraph.
    for (let position = 0; position < indices.length; position++) {
      if (indices[position] >= referenceIndex) return position;
    }

    // Fallback for legacy documents whose captions remain below their images.
    for (let position = indices.length - 1; position >= 0; position--) {
      if (indices[position] < referenceIndex) return position;
    }
    return -1;
  }

  // Table captions belong to the first table at or after the caption. Equality
  // supports an anchor set directly from inside the table.
  for (let position = 0; position < indices.length; position++) {
    if (indices[position] >= referenceIndex) return position;
  }

  // Fallback for legacy captions that remain below their table.
  for (let position = indices.length - 1; position >= 0; position--) {
    if (indices[position] < referenceIndex) return position;
  }

  return -1;
}
