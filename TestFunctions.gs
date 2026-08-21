function runCurrentTest() {
  var started = Date.now();
  var setupStarted = Date.now();

  var doc = DocumentApp.getActiveDocument();
  var documentId = doc.getId();
  var cursor = doc.getCursor();

  if (!cursor) {
    throw new Error(
      'Place only the cursor inside one line.'
    );
  }

  var source = findTestParagraph_(
    cursor.getElement()
  );

  if (!source) {
    throw new Error(
      'No paragraph was found at the cursor.'
    );
  }

  /*
   * Si la línea pertenece a un párrafo con Shift+Enter,
   * extrae únicamente la línea del cursor.
   */
  source = extractCursorLine_(source, cursor);

  var activeTab = doc.getActiveTab();
  var tabId = activeTab.getId();
  var documentTab = activeTab.asDocumentTab();

  var parent = source.getParent();
  var token = Utilities
    .getUuid()
    .replace(/-/g, '');

  var beforeText = '\uE610' + token;
  var afterText = '\uE611' + token;
  var rangeName = 'DA_INCISO_' + token;

  /*
   * Evitan que Google una el inciso con una lista vecina.
   */
  parent.insertParagraph(
    parent.getChildIndex(source),
    beforeText
  );

  parent.insertParagraph(
    parent.getChildIndex(source) + 1,
    afterText
  );

  /*
   * Creamos un rango nombrado localmente.
   */
  var range = documentTab
    .newRange()
    .addElement(source)
    .build();

  var namedRange = documentTab.addNamedRange(
    rangeName,
    range
  );

  var namedRangeId = namedRange.getId();

  /*
   * La API debe ver el rango nombrado.
   */
  doc.saveAndClose();
  Utilities.sleep(250);

  var setupMs = Date.now() - setupStarted;

  /*
   * Solicitamos únicamente metadatos de rangos nombrados.
   * No pedimos el cuerpo del documento.
   */
  var readStarted = Date.now();

  var apiDocument = Docs.Documents.get(
    documentId,
    {
      includeTabsContent: true,
      fields: namedRangeFieldsMask_()
    }
  );

  var apiReadMs = Date.now() - readStarted;

  var metadata = findNamedRangeMetadata_(
    apiDocument.tabs || [],
    tabId,
    rangeName,
    namedRangeId
  );

  if (
    !metadata ||
    !metadata.ranges ||
    !metadata.ranges.length
  ) {
    cleanupNamedRangeTest_(
      documentId,
      tabId,
      namedRangeId,
      beforeText,
      afterText
    );

    throw new Error(
      'The optimized metadata read did not return the named range.'
    );
  }

  var apiRange = metadata.ranges[0];

  /*
   * Crea una lista nativa cuyo segundo nivel es a).
   */
  var writeStarted = Date.now();

  Docs.Documents.batchUpdate(
    {
      requests: [
        {
          createParagraphBullets: {
            range: {
              startIndex: apiRange.startIndex,
              endIndex: apiRange.endIndex,
              tabId: apiRange.tabId || tabId
            },
            bulletPreset:
              'NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS'
          }
        }
      ]
    },
    documentId
  );

  var apiWriteMs = Date.now() - writeStarted;
  var finishStarted = Date.now();

  /*
   * Volvemos a abrir el documento porque saveAndClose()
   * invalida los objetos DocumentApp anteriores.
   */
  var reopened = DocumentApp.openById(
    documentId
  );

  var reopenedTab = reopened
    .getTab(tabId)
    .asDocumentTab();

  var localNamedRange =
    reopenedTab.getNamedRangeById(
      namedRangeId
    );

  if (!localNamedRange) {
    cleanupNamedRangeTest_(
      documentId,
      tabId,
      namedRangeId,
      beforeText,
      afterText
    );

    throw new Error(
      'The local named range was not found after the API update.'
    );
  }

  var item = findListItemInRange_(
    localNamedRange.getRange()
  );

  if (!item) {
    cleanupMarkersNearRange_(
      localNamedRange,
      beforeText,
      afterText
    );

    try {
      localNamedRange.remove();
    } catch (error) {}

    reopened.saveAndClose();

    throw new Error(
      'The Docs API did not create a native list item.'
    );
  }

  /*
   * Nivel 1 del preset:
   * nivel 0 = 1)
   * nivel 1 = a)
   * nivel 2 = i)
   */
  item.setNestingLevel(1);

  /*
   * Left 0.25"
   * Hanging 0.25"
   * Right 0"
   */
  item.setIndentFirstLine(18);
  item.setIndentStart(36);
  item.setIndentEnd(0);

  /*
   * Elimina los separadores temporales.
   */
  removeMarkerSibling_(
    item.getPreviousSibling(),
    beforeText
  );

  removeMarkerSibling_(
    item.getNextSibling(),
    afterText
  );

  try {
    localNamedRange.remove();
  } catch (error) {}

  var glyphType = String(
    item.getGlyphType()
  );

  var nestingLevel =
    item.getNestingLevel();

  reopened.saveAndClose();

  return {
    testId:
      'TEST-OPTIMIZED-NAMED-RANGE-A-PAREN',

    setupMs: setupMs,
    apiReadMs: apiReadMs,
    apiWriteMs: apiWriteMs,

    localFinishMs:
      Date.now() - finishStarted,

    elapsedMs:
      Date.now() - started,

    glyphType: glyphType,
    nestingLevel: nestingLevel,
    ok: true
  };
}


/**
 * Field mask limitado a:
 * - ID de las pestañas
 * - rangos nombrados
 *
 * Incluye hasta tres niveles de pestañas anidadas.
 */
function namedRangeFieldsMask_() {
  var level3 =
    'childTabs(' +
      'tabProperties(tabId),' +
      'documentTab(namedRanges)' +
    ')';

  var level2 =
    'childTabs(' +
      'tabProperties(tabId),' +
      'documentTab(namedRanges),' +
      level3 +
    ')';

  var level1 =
    'childTabs(' +
      'tabProperties(tabId),' +
      'documentTab(namedRanges),' +
      level2 +
    ')';

  return (
    'tabs(' +
      'tabProperties(tabId),' +
      'documentTab(namedRanges),' +
      level1 +
    ')'
  );
}


/**
 * Busca el rango nombrado dentro de las pestañas
 * devueltas por Docs API.
 */
function findNamedRangeMetadata_(
  tabs,
  wantedTabId,
  rangeName,
  rangeId
) {
  for (var i = 0; i < tabs.length; i++) {
    var tab = tabs[i];

    var currentTabId =
      tab.tabProperties &&
      tab.tabProperties.tabId;

    var namedRanges =
      tab.documentTab &&
      tab.documentTab.namedRanges;

    if (
      (!wantedTabId ||
       currentTabId === wantedTabId) &&
      namedRanges &&
      namedRanges[rangeName]
    ) {
      var matches =
        namedRanges[rangeName]
          .namedRanges || [];

      for (
        var j = 0;
        j < matches.length;
        j++
      ) {
        if (
          !rangeId ||
          matches[j].namedRangeId ===
            rangeId
        ) {
          return matches[j];
        }
      }
    }

    var nested =
      findNamedRangeMetadata_(
        tab.childTabs || [],
        wantedTabId,
        rangeName,
        rangeId
      );

    if (nested) {
      return nested;
    }
  }

  return null;
}


/**
 * Obtiene el ListItem creado dentro del rango.
 */
function findListItemInRange_(range) {
  var elements =
    range.getRangeElements();

  for (
    var i = 0;
    i < elements.length;
    i++
  ) {
    var element =
      elements[i].getElement();

    while (element) {
      if (
        element.getType() ===
        DocumentApp.ElementType.LIST_ITEM
      ) {
        return element.asListItem();
      }

      if (
        element.getType() ===
        DocumentApp.ElementType.PARAGRAPH
      ) {
        break;
      }

      element = element.getParent();
    }
  }

  return null;
}


/**
 * Elimina un separador temporal únicamente
 * cuando coincide exactamente con su token.
 */
function removeMarkerSibling_(
  element,
  expectedText
) {
  if (!element) return;

  var type = element.getType();

  if (
    type !==
      DocumentApp.ElementType.PARAGRAPH &&
    type !==
      DocumentApp.ElementType.LIST_ITEM
  ) {
    return;
  }

  if (element.getText() === expectedText) {
    try {
      element.removeFromParent();
    } catch (error) {}
  }
}


/**
 * Limpieza utilizando el rango nombrado.
 */
function cleanupMarkersNearRange_(
  namedRange,
  beforeText,
  afterText
) {
  if (!namedRange) return;

  var elements = namedRange
    .getRange()
    .getRangeElements();

  if (!elements.length) return;

  var owner = findTestParagraph_(
    elements[0].getElement()
  );

  if (!owner) return;

  removeMarkerSibling_(
    owner.getPreviousSibling(),
    beforeText
  );

  removeMarkerSibling_(
    owner.getNextSibling(),
    afterText
  );
}


/**
 * Limpieza de emergencia cuando ocurre un error
 * después de cerrar el documento original.
 */
function cleanupNamedRangeTest_(
  documentId,
  tabId,
  namedRangeId,
  beforeText,
  afterText
) {
  try {
    var doc = DocumentApp.openById(
      documentId
    );

    var tab = doc
      .getTab(tabId)
      .asDocumentTab();

    var namedRange =
      tab.getNamedRangeById(
        namedRangeId
      );

    if (namedRange) {
      cleanupMarkersNearRange_(
        namedRange,
        beforeText,
        afterText
      );

      try {
        namedRange.remove();
      } catch (error) {}
    }

    doc.saveAndClose();
  } catch (error) {}
}


/**
 * Encuentra el Paragraph o ListItem del cursor.
 */
function findTestParagraph_(element) {
  var current = element;

  while (current) {
    var type = current.getType();

    if (
      type ===
        DocumentApp.ElementType.PARAGRAPH ||
      type ===
        DocumentApp.ElementType.LIST_ITEM
    ) {
      return current;
    }

    current = current.getParent();
  }

  return null;
}


/**
 * Si el cursor está dentro de líneas separadas
 * mediante Shift+Enter, convierte esas líneas
 * en párrafos reales y devuelve únicamente
 * la línea del cursor.
 */
function extractCursorLine_(
  source,
  cursor
) {
  var sourceText = source.getText();

  if (!/[\r\n]/.test(sourceText)) {
    return source;
  }

  var rawElement = cursor.getElement();

  var cursorOffset = Math.max(
    0,
    cursor.getSurroundingTextOffset()
  );

  if (
    rawElement.getType() ===
    DocumentApp.ElementType.TEXT
  ) {
    cursorOffset +=
      getTextOffsetInTestParagraph_(
        source,
        rawElement
      );
  }

  var lines = [];
  var breakPattern = /\r\n|\r|\n/g;
  var start = 0;
  var match;

  while (
    (match =
      breakPattern.exec(sourceText)) !==
    null
  ) {
    lines.push({
      text: sourceText.substring(
        start,
        match.index
      ),
      start: start
    });

    start =
      match.index +
      match[0].length;
  }

  lines.push({
    text: sourceText.substring(start),
    start: start
  });

  var selectedLine = 0;

  for (
    var i = 1;
    i < lines.length;
    i++
  ) {
    if (
      cursorOffset >= lines[i].start
    ) {
      selectedLine = i;
    }
  }

  var parent = source.getParent();

  var sourceIndex =
    parent.getChildIndex(source);

  var paragraphAttributes =
    source.getAttributes();

  var textAttributes = {};

  if (sourceText.length) {
    try {
      textAttributes =
        source.editAsText()
          .getAttributes();
    } catch (error) {}
  }

  var created = [];

  lines.forEach(
    function(line, index) {
      var paragraph =
        parent.insertParagraph(
          sourceIndex + index,
          line.text
        );

      try {
        paragraph.setAttributes(
          paragraphAttributes
        );
      } catch (error) {}

      if (line.text.length) {
        try {
          paragraph
            .editAsText()
            .setAttributes(
              textAttributes
            );
        } catch (error) {}
      }

      created.push(paragraph);
    }
  );

  source.removeFromParent();

  return created[selectedLine];
}


/**
 * Calcula el offset del Text dentro
 * de su Paragraph/ListItem.
 */
function getTextOffsetInTestParagraph_(
  owner,
  target
) {
  if (owner === target) return 0;

  var offset = 0;
  var found = false;

  function walk(element) {
    if (element === target) {
      found = true;
      return true;
    }

    if (
      element.getType() ===
      DocumentApp.ElementType.TEXT
    ) {
      offset += element
        .asText()
        .getText()
        .length;

      return false;
    }

    if (!element.getNumChildren) {
      return false;
    }

    for (
      var i = 0;
      i < element.getNumChildren();
      i++
    ) {
      if (
        walk(element.getChild(i))
      ) {
        return true;
      }
    }

    return false;
  }

  walk(owner);

  return found ? offset : 0;
}
