/**
 * TEST LAB
 *
 * Aplica una lista automática nativa:
 *
 * a) Primer párrafo
 * b) Segundo párrafo
 * c) Tercer párrafo
 *
 * Puede utilizar una selección o el párrafo donde está el cursor.
 */
function runCurrentTest() {
  return testApplyNativeIncisos_();
}

/**
 * Aplica incisos automáticos a), b), c).
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
  const tabId = activeTab.getId();
  const documentTab = activeTab.asDocumentTab();
  const body = documentTab.getBody();

  const paragraphs =
    incisoGetTargetParagraphs_(doc, body);

  const paragraphCount = paragraphs.length;

  if (!paragraphCount) {
    throw new Error(
      'Selecciona uno o varios párrafos, o coloca el cursor dentro de un párrafo.'
    );
  }

  /*
   * Crear un marcador temporal independiente para cada párrafo.
   *
   * Estos marcadores permiten obtener los índices exactos requeridos
   * por Google Docs API sin leer todo el contenido del documento.
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

    const namedRange = documentTab.addNamedRange(
      markerName,
      range
    );

    markers.push({
      name: markerName,
      documentAppId: namedRange.getId()
    });
  });

  /*
   * Publicar los named ranges antes de consultar Docs API.
   *
   * No utilizar el objeto doc después de saveAndClose().
   */
  doc.saveAndClose();

  const apiStarted = Date.now();

  /*
   * Solicitar únicamente los named ranges.
   *
   * No se descarga el contenido completo del documento.
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
   * Localizar el rango numérico de cada párrafo.
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

    /*
     * El ID de DocumentApp puede ser diferente del ID presentado
     * por Docs API. El nombre es único, por lo que podemos tomar
     * el primer resultado de este grupo.
     */
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

    if (
      !Number.isFinite(Number(range.startIndex)) ||
      !Number.isFinite(Number(range.endIndex))
    ) {
      throw new Error(
        'El marcador temporal no contiene índices válidos.'
      );
    }

    return {
      startIndex: Number(range.startIndex),
      endIndex: Number(range.endIndex),
      tabId: tabId
    };
  });

  /*
   * Rango completo desde el primer párrafo hasta el último.
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

  const originalRange = {
    startIndex: startIndex,
    endIndex: endIndex,
    tabId: tabId
  };

  /*
   * Para obtener a), la lista debe crearse en el nivel 1.
   *
   * Docs API determina el nivel contando las tabulaciones que
   * existen al inicio de cada párrafo:
   *
   * Nivel 0: 1)
   * Nivel 1: a)
   * Nivel 2: i)
   *
   * Insertaremos una tabulación en cada párrafo.
   */
  const paragraphStarts = locatedRanges
    .map(function (range) {
      return range.startIndex;
    })
    .sort(function (a, b) {
      /*
       * Insertar desde el último párrafo hacia el primero evita
       * desplazar los índices que todavía no se han utilizado.
       */
      return b - a;
    });

  /*
   * Después de insertar una tabulación en cada párrafo, el rango
   * será temporalmente más largo.
   */
  const rangeWithTabs = {
    startIndex: startIndex,
    endIndex: endIndex + paragraphCount,
    tabId: tabId
  };

  const requests = [];

  /*
   * Quitar cualquier lista automática existente en los párrafos.
   */
  requests.push({
    deleteParagraphBullets: {
      range: originalRange
    }
  });

  /*
   * Insertar una tabulación al inicio de cada párrafo.
   */
  paragraphStarts.forEach(function (paragraphStart) {
    requests.push({
      insertText: {
        location: {
          index: paragraphStart,
          tabId: tabId
        },
        text: '\t'
      }
    });
  });

  /*
   * Crear la lista nativa.
   *
   * NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS produce:
   *
   * Nivel 0: 1)
   * Nivel 1: a)
   * Nivel 2: i)
   *
   * CreateParagraphBullets elimina automáticamente las
   * tabulaciones iniciales después de determinar el nivel.
   */
  requests.push({
    createParagraphBullets: {
      range: rangeWithTabs,
      bulletPreset:
        'NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS'
    }
  });

  /*
   * Aplicar:
   *
   * Left:    0.25"
   * Hanging: 0.25"
   * Right:   0"
   *
   * Equivalencias:
   *
   * 0.25" = 18 pt
   * indentStart = 36 pt
   * indentFirstLine = 18 pt
   */
  requests.push({
    updateParagraphStyle: {
      range: originalRange,
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
  });

  /*
   * Eliminar los marcadores temporales usando sus nombres únicos.
   */
  markers.forEach(function (marker) {
    requests.push({
      deleteNamedRange: {
        name: marker.name,
        tabsCriteria: {
          tabIds: [tabId]
        }
      }
    });
  });

  /*
   * Ejecutar todos los cambios en un solo batchUpdate.
   */
  Docs.Documents.batchUpdate(
    {
      requests: requests
    },
    documentId
  );

  return {
    ok: true,
    testId: 'TEST-NATIVE-A-PAREN',
    automaticList: true,
    expectedGlyph: 'a)',
    paragraphsApplied: paragraphCount,
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
 * Obtiene los párrafos completos incluidos en la selección.
 *
 * Si no hay selección, utiliza el párrafo donde está el cursor.
 * También incluye los párrafos vacíos que existan dentro de la
 * selección.
 */
function incisoGetTargetParagraphs_(doc, body) {
  const selection = doc.getSelection();
  const foundIndexes = {};

  if (selection) {
    selection
      .getRangeElements()
      .forEach(function (rangeElement) {
        const paragraph = incisoFindParagraph_(
          rangeElement.getElement()
        );

        if (!paragraph) {
          return;
        }

        const childIndex = incisoBodyChildIndex_(
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

    const childIndex = incisoBodyChildIndex_(
      body,
      paragraph
    );

    if (childIndex < 0) {
      throw new Error(
        'Esta prueba solo admite párrafos del cuerpo principal.'
      );
    }

    foundIndexes[childIndex] = true;
  }

  const selectedIndexes = Object.keys(foundIndexes)
    .map(function (value) {
      return Number(value);
    })
    .sort(function (a, b) {
      return a - b;
    });

  if (!selectedIndexes.length) {
    return [];
  }

  const firstIndex = selectedIndexes[0];
  const lastIndex =
    selectedIndexes[selectedIndexes.length - 1];

  const paragraphs = [];

  /*
   * Recorrer todos los elementos entre el primer y último párrafo.
   * Esto permite incluir párrafos vacíos.
   */
  for (
    let childIndex = firstIndex;
    childIndex <= lastIndex;
    childIndex++
  ) {
    const element = body.getChild(childIndex);
    const elementType = element.getType();

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
    const elementType = current.getType();

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
function incisoBodyChildIndex_(body, paragraph) {
  try {
    return body.getChildIndex(paragraph);
  } catch (error) {
    return -1;
  }
}

/**
 * Localiza el tab activo dentro de la respuesta de Docs API.
 */
function incisoFindApiTab_(tabs, wantedTabId) {
  for (
    let index = 0;
    index < tabs.length;
    index++
  ) {
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
