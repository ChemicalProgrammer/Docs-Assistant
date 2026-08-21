function runCurrentTest() {
  return testApplyNativeIncisos_();
}

/**
 * Aplica una lista automática:
 *   a)
 *   b)
 *   c)
 *
 * A la selección o al párrafo donde está el cursor.
 *
 * Sangrías:
 * - Left:    0.25"
 * - Hanging: 0.25"
 * - Right:   0"
 */
function testApplyNativeIncisos_() {
  const started = Date.now();
  const doc = DocumentApp.getActiveDocument();
  const documentId = doc.getId();

  if (typeof Docs === 'undefined' || !Docs.Documents) {
    throw new Error(
      'Activa el servicio avanzado Google Docs API en Apps Script.'
    );
  }

  const activeTab = doc.getActiveTab();
  const tabId = activeTab.getId();
  const documentTab = activeTab.asDocumentTab();
  const body = documentTab.getBody();

  const paragraphs = incisoGetTargetParagraphs_(doc, body);

  if (!paragraphs.length) {
    throw new Error(
      'Selecciona uno o varios párrafos, o coloca el cursor en un párrafo.'
    );
  }

  /*
   * El preset requerido usa:
   * nivel 0 = 1)
   * nivel 1 = a)
   * nivel 2 = i)
   *
   * Una tabulación inicial obliga al nivel 1. La API elimina esa
   * tabulación al crear la lista.
   */
  const markers = [];
  const markerPrefix =
    'DA_INCISO_' +
    Utilities.getUuid().replace(/-/g, '') +
    '_';

  paragraphs.forEach(function (paragraph, index) {
    const text = paragraph.editAsText();

    // Dejar exactamente una tabulación inicial.
    while (text.getText().charAt(0) === '\t') {
      text.deleteText(0, 0);
    }

    text.insertText(0, '\t');

    const name = markerPrefix + index;
    const range = documentTab
      .newRange()
      .addElement(paragraph)
      .build();

    const namedRange = documentTab.addNamedRange(name, range);

    markers.push({
      name: name,
      id: namedRange.getId()
    });
  });

  /*
   * Publica los cambios de DocumentApp antes de usar Docs API.
   * No se vuelve a utilizar el objeto doc después de esta línea.
   */
  doc.saveAndClose();

  const apiStarted = Date.now();

  /*
   * Lee únicamente los named ranges, no las 295 páginas del documento.
   */
  const apiDocument = Docs.Documents.get(documentId, {
    includeTabsContent: true,
    fields:
      'tabs(tabProperties(tabId),documentTab(namedRanges))'
  });

  const apiTab = incisoFindApiTab_(apiDocument.tabs || [], tabId);

  if (!apiTab || !apiTab.documentTab) {
    throw new Error(
      'La API no devolvió el tab activo del documento.'
    );
  }

  const namedRanges =
    apiTab.documentTab.namedRanges || {};

  const locatedRanges = markers.map(function (marker) {
    const group = namedRanges[marker.name];

    if (
      !group ||
      !group.namedRanges ||
      !group.namedRanges.length
    ) {
      throw new Error(
        'No se encontró el marcador temporal: ' + marker.name
      );
    }

    const matching =
      group.namedRanges.find(function (item) {
        return String(item.namedRangeId) === String(marker.id);
      }) || group.namedRanges[0];

    if (!matching.ranges || !matching.ranges.length) {
      throw new Error(
        'El marcador temporal no contiene un rango.'
      );
    }

    return matching.ranges[0];
  });

  const startIndex = Math.min.apply(
    null,
    locatedRanges.map(function (range) {
      return Number(range.startIndex);
    })
  );

  const endIndex = Math.max.apply(
    null,
    locatedRanges.map(function (range) {
      return Number(range.endIndex);
    })
  );

  const originalRange = {
    startIndex: startIndex,
    endIndex: endIndex,
    tabId: tabId
  };

  /*
   * CreateParagraphBullets elimina una tabulación por párrafo.
   * Por eso el rango final es más corto.
   */
  const formattedRange = {
    startIndex: startIndex,
    endIndex: endIndex - paragraphs.length,
    tabId: tabId
  };

  const requests = [
    {
      deleteParagraphBullets: {
        range: originalRange
      }
    },
    {
      createParagraphBullets: {
        range: originalRange,
        bulletPreset:
          'NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS'
      }
    },
    {
      updateParagraphStyle: {
        range: formattedRange,
        paragraphStyle: {
          indentStart: {
            magnitude: 36,
            unit: 'PT'
          },
          indentFirstLine: {
            magnitude: 18,
            unit: 'PT'
          },
          indentEnd: {
            magnitude: 0,
            unit: 'PT'
          }
        },
        fields:
          'indentStart,indentFirstLine,indentEnd'
      }
    }
  ];

  // Eliminar los marcadores temporales.
  markers.forEach(function (marker) {
    requests.push({
      deleteNamedRange: {
        namedRangeId: marker.id,
        tabsCriteria: {
          tabIds: [tabId]
        }
      }
    });
  });

  Docs.Documents.batchUpdate(
    {
      requests: requests
    },
    documentId
  );

  return {
    ok: true,
    testId: 'TEST-NATIVE-A-PAREN',
    paragraphsApplied: paragraphs.length,
    expectedGlyph: 'a)',
    automaticList: true,
    expectedIndentInches: {
      left: 0.25,
      hanging: 0.25,
      right: 0
    },
    apiMs: Date.now() - apiStarted,
    elapsedMs: Date.now() - started
  };
}

/**
 * Obtiene los párrafos completos comprendidos por la selección.
 * Si no hay selección, usa el párrafo donde está el cursor.
 */
function incisoGetTargetParagraphs_(doc, body) {
  const selection = doc.getSelection();
  const found = {};

  if (selection) {
    selection.getRangeElements().forEach(function (rangeElement) {
      const paragraph =
        incisoFindParagraph_(rangeElement.getElement());

      if (!paragraph) return;

      const index = incisoBodyChildIndex_(body, paragraph);

      if (index >= 0) {
        found[index] = true;
      }
    });
  } else {
    const cursor = doc.getCursor();

    if (!cursor) {
      throw new Error(
        'No hay selección ni cursor dentro del documento.'
      );
    }

    const paragraph =
      incisoFindParagraph_(cursor.getElement());

    if (!paragraph) {
      throw new Error(
        'El cursor no está dentro de un párrafo.'
      );
    }

    const index = incisoBodyChildIndex_(body, paragraph);

    if (index < 0) {
      throw new Error(
        'Esta prueba solo admite párrafos del cuerpo principal.'
      );
    }

    found[index] = true;
  }

  const indexes = Object.keys(found)
    .map(Number)
    .sort(function (a, b) {
      return a - b;
    });

  if (!indexes.length) return [];

  /*
   * Incluye los párrafos vacíos que existan entre el primero
   * y el último párrafo seleccionado.
   */
  const first = indexes[0];
  const last = indexes[indexes.length - 1];
  const result = [];

  for (let index = first; index <= last; index++) {
    const element = body.getChild(index);
    const type = element.getType();

    if (
      type === DocumentApp.ElementType.PARAGRAPH ||
      type === DocumentApp.ElementType.LIST_ITEM
    ) {
      result.push(element);
      continue;
    }

    throw new Error(
      'La selección contiene un elemento que no es un párrafo.'
    );
  }

  return result;
}

function incisoFindParagraph_(element) {
  let current = element;

  while (current) {
    const type = current.getType();

    if (
      type === DocumentApp.ElementType.PARAGRAPH ||
      type === DocumentApp.ElementType.LIST_ITEM
    ) {
      return current;
    }

    current = current.getParent
      ? current.getParent()
      : null;
  }

  return null;
}

function incisoBodyChildIndex_(body, paragraph) {
  try {
    return body.getChildIndex(paragraph);
  } catch (error) {
    return -1;
  }
}

function incisoFindApiTab_(tabs, wantedTabId) {
  for (let index = 0; index < tabs.length; index++) {
    const tab = tabs[index];
    const properties = tab.tabProperties || {};

    if (
      String(properties.tabId || '') ===
      String(wantedTabId || '')
    ) {
      return tab;
    }
  }

  return null;
}
