/**
 * TEST LAB
 *
 * El botón de pruebas ejecuta siempre esta función.
 */
function runCurrentTest() {
  return testApplyNativeIncisos_();
}

/**
 * Aplica una lista automática nativa:
 *
 * a) Primer párrafo
 * b) Segundo párrafo
 * c) Tercer párrafo
 *
 * Funciona con:
 * - Uno o varios párrafos seleccionados.
 * - El párrafo donde está colocado el cursor.
 *
 * Sangrías:
 * - Left: 0.25"
 * - Hanging: 0.25"
 * - Right: 0"
 */
function testApplyNativeIncisos_() {
  const started = Date.now();

  if (
    typeof Docs === 'undefined' ||
    !Docs.Documents
  ) {
    throw new Error(
      'Activa el servicio avanzado Google Docs API.'
    );
  }

  const doc = DocumentApp.getActiveDocument();
  const documentId = doc.getId();

  const activeTab = doc.getActiveTab();

  if (!activeTab) {
    throw new Error(
      'No se pudo obtener el tab activo.'
    );
  }

  const tabId = activeTab.getId();
  const documentTab = activeTab.asDocumentTab();
  const body = documentTab.getBody();

  const paragraphs =
    incisoGetTargetParagraphs_(doc, body);

  const paragraphCount = paragraphs.length;

  if (!paragraphCount) {
    throw new Error(
      'Selecciona párrafos o coloca el cursor dentro de un párrafo.'
    );
  }

  /*
   * Crear un named range para cada párrafo.
   *
   * Los marcadores permitirán recuperar los mismos párrafos después
   * de utilizar Docs API y volver a abrir el documento.
   */
  const markers = [];

  const markerPrefix =
    'DA_INCISO_' +
    Utilities.getUuid().replace(/-/g, '') +
    '_';

  paragraphs.forEach(function (paragraph, index) {
    const markerName = markerPrefix + index;

    const range = documentTab
      .newRange()
      .addElement(paragraph)
      .build();

    documentTab.addNamedRange(
      markerName,
      range
    );

    markers.push({
      name: markerName
    });
  });

  /*
   * Publicar los named ranges antes de llamar Docs API.
   *
   * La instancia doc queda cerrada después de esta instrucción.
   */
  doc.saveAndClose();

  const apiStarted = Date.now();

  /*
   * Obtener solamente los named ranges.
   *
   * No se solicita el contenido completo del documento.
   */
  const apiDocument = Docs.Documents.get(
    documentId,
    {
      includeTabsContent: true,
      fields:
        'tabs(tabProperties(tabId),documentTab(namedRanges))'
    }
  );

  const apiTab = incisoFindApiTab_(
    apiDocument.tabs || [],
    tabId
  );

  if (!apiTab || !apiTab.documentTab) {
    throw new Error(
      'Google Docs API no devolvió el tab activo.'
    );
  }

  const apiNamedRanges =
    apiTab.documentTab.namedRanges || {};

  /*
   * Obtener el startIndex y endIndex de cada párrafo.
   */
  const locatedRanges = markers.map(function (marker) {
    const namedRangeGroup =
      apiNamedRanges[marker.name];

    if (
      !namedRangeGroup ||
      !namedRangeGroup.namedRanges ||
      !namedRangeGroup.namedRanges.length
    ) {
      throw new Error(
        'No se encontró el marcador temporal: ' +
        marker.name
      );
    }

    const apiNamedRange =
      namedRangeGroup.namedRanges[0];

    if (
      !apiNamedRange.ranges ||
      !apiNamedRange.ranges.length
    ) {
      throw new Error(
        'El marcador temporal no contiene un rango.'
      );
    }

    const range = apiNamedRange.ranges[0];

    const startIndex =
      Number(range.startIndex);

    const endIndex =
      Number(range.endIndex);

    if (
      !Number.isFinite(startIndex) ||
      !Number.isFinite(endIndex)
    ) {
      throw new Error(
        'El marcador temporal no contiene índices válidos.'
      );
    }

    return {
      startIndex: startIndex,
      endIndex: endIndex,
      tabId: tabId
    };
  });

  /*
   * Crear un solo rango que incluya todos los párrafos.
   */
  const startIndex = Math.min.apply(
    null,
    locatedRanges.map(function (range) {
      return range.startIndex;
    })
  );

  const endIndex = Math.max.apply(
    null,
    locatedRanges.map(function (range) {
      return range.endIndex;
    })
  );

  const targetRange = {
    startIndex: startIndex,
    endIndex: endIndex,
    tabId: tabId
  };

  /*
   * Crear una lista automática con el preset:
   *
   * Nivel 0 = 1)
   * Nivel 1 = a)
   * Nivel 2 = i)
   *
   * Docs API crea inicialmente los elementos en el nivel 0.
   */
  Docs.Documents.batchUpdate(
    {
      requests: [
        {
          deleteParagraphBullets: {
            range: targetRange
          }
        },
        {
          createParagraphBullets: {
            range: targetRange,
            bulletPreset:
              'NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS'
          }
        }
      ]
    },
    documentId
  );

  const apiMs = Date.now() - apiStarted;

  /*
   * Volver a abrir el documento.
   *
   * No se reutiliza la instancia cerrada anteriormente.
   */
  const reopenedDocument =
    DocumentApp.openById(documentId);

  const reopenedTabObject =
    reopenedDocument.getTab(tabId);

  if (!reopenedTabObject) {
    throw new Error(
      'No se pudo volver a abrir el tab activo.'
    );
  }

  const reopenedTab =
    reopenedTabObject.asDocumentTab();

  let paragraphsApplied = 0;

  /*
   * Recuperar cada ListItem utilizando su named range.
   */
  markers.forEach(function (marker) {
    const namedRanges =
      reopenedTab.getNamedRanges(marker.name);

    if (
      !namedRanges ||
      !namedRanges.length
    ) {
      throw new Error(
        'No se pudo recuperar el marcador: ' +
        marker.name
      );
    }

    const namedRange = namedRanges[0];

    const rangeElements =
      namedRange
        .getRange()
        .getRangeElements();

    let listItem = null;

    for (
      let index = 0;
      index < rangeElements.length;
      index++
    ) {
      const paragraph = incisoFindParagraph_(
        rangeElements[index].getElement()
      );

      if (
        paragraph &&
        paragraph.getType() ===
          DocumentApp.ElementType.LIST_ITEM
      ) {
        listItem = paragraph.asListItem();
        break;
      }
    }

    if (!listItem) {
      throw new Error(
        'El párrafo no fue convertido en un ListItem.'
      );
    }

    /*
     * Mover el elemento al nivel 1 del preset.
     *
     * Nivel 0 = 1)
     * Nivel 1 = a)
     */
    listItem.setNestingLevel(1);

    /*
     * Aplicar las sangrías solicitadas.
     *
     * 18 puntos = 0.25 pulgadas.
     *
     * indentStart:     36 pt
     * indentFirstLine: 18 pt
     * indentEnd:        0 pt
     */
    listItem.setIndentStart(36);
    listItem.setIndentFirstLine(18);
    listItem.setIndentEnd(0);

    /*
     * Eliminar el marcador temporal.
     */
    namedRange.remove();

    paragraphsApplied++;
  });

  reopenedDocument.saveAndClose();

  return {
    ok: true,
    testId:
      'TEST-NATIVE-A-PAREN-LEVEL-1',
    automaticList: true,
    expectedGlyph: 'a)',
    nestingLevel: 1,
    paragraphsApplied: paragraphsApplied,
    expectedIndentInches: {
      left: 0.25,
      hanging: 0.25,
      right: 0
    },
    apiMs: apiMs,
    elapsedMs: Date.now() - started
  };
}

/**
 * Obtiene los párrafos completos incluidos en la selección.
 *
 * Si no existe una selección, devuelve el párrafo donde está
 * colocado el cursor.
 *
 * También incluye los párrafos vacíos ubicados entre el primer
 * y el último párrafo seleccionado.
 */
function incisoGetTargetParagraphs_(doc, body) {
  const selection = doc.getSelection();
  const foundIndexes = {};

  if (selection) {
    const rangeElements =
      selection.getRangeElements();

    rangeElements.forEach(function (rangeElement) {
      const paragraph = incisoFindParagraph_(
        rangeElement.getElement()
      );

      if (!paragraph) {
        return;
      }

      const childIndex =
        incisoBodyChildIndex_(
          body,
          paragraph
        );

      if (childIndex >= 0) {
        foundIndexes[childIndex] = true;
      }
    });
  } else {
    const cursor = doc.getCursor();

    if (!cursor) {
      throw new Error(
        'No se detectó una selección ni un cursor.'
      );
    }

    const paragraph = incisoFindParagraph_(
      cursor.getElement()
    );

    if (!paragraph) {
      throw new Error(
        'El cursor no está dentro de un párrafo.'
      );
    }

    const childIndex =
      incisoBodyChildIndex_(
        body,
        paragraph
      );

    if (childIndex < 0) {
      throw new Error(
        'Esta prueba solamente admite párrafos del cuerpo principal.'
      );
    }

    foundIndexes[childIndex] = true;
  }

  const selectedIndexes =
    Object.keys(foundIndexes)
      .map(function (value) {
        return Number(value);
      })
      .sort(function (a, b) {
        return a - b;
      });

  if (!selectedIndexes.length) {
    return [];
  }

  const firstIndex =
    selectedIndexes[0];

  const lastIndex =
    selectedIndexes[
      selectedIndexes.length - 1
    ];

  const paragraphs = [];

  /*
   * Recorrer todos los elementos entre el primero y el último.
   *
   * Esto permite incluir párrafos vacíos dentro de la selección.
   */
  for (
    let childIndex = firstIndex;
    childIndex <= lastIndex;
    childIndex++
  ) {
    const element =
      body.getChild(childIndex);

    const elementType =
      element.getType();

    if (
      elementType ===
        DocumentApp.ElementType.PARAGRAPH ||
      elementType ===
        DocumentApp.ElementType.LIST_ITEM
    ) {
      paragraphs.push(element);
      continue;
    }

    throw new Error(
      'La selección contiene un elemento que no es un párrafo.'
    );
  }

  return paragraphs;
}

/**
 * Busca el Paragraph o ListItem que contiene un elemento.
 */
function incisoFindParagraph_(element) {
  let current = element;

  while (current) {
    const elementType =
      current.getType();

    if (
      elementType ===
        DocumentApp.ElementType.PARAGRAPH ||
      elementType ===
        DocumentApp.ElementType.LIST_ITEM
    ) {
      return current;
    }

    current = current.getParent
      ? current.getParent()
      : null;
  }

  return null;
}

/**
 * Obtiene el índice del párrafo dentro del Body.
 */
function incisoBodyChildIndex_(
  body,
  paragraph
) {
  try {
    return body.getChildIndex(
      paragraph
    );
  } catch (error) {
    return -1;
  }
}

/**
 * Localiza el tab activo en la respuesta de Docs API.
 */
function incisoFindApiTab_(
  tabs,
  wantedTabId
) {
  for (
    let index = 0;
    index < tabs.length;
    index++
  ) {
    const tab = tabs[index];

    const properties =
      tab.tabProperties || {};

    if (
      String(properties.tabId || '') ===
      String(wantedTabId || '')
    ) {
      return tab;
    }
  }

  return null;
}
