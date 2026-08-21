/**
 * TEST-022
 *
 * Convierte UN párrafo normal en el primer elemento de una lista nativa:
 *
 * a) Octubre
 *
 * No utiliza:
 * - rangos con nombre
 * - documentos temporales
 * - DriveApp
 * - prefijos escritos como texto
 */
function runCurrentTest() {
  return testCreateIndependentNativeInciso_();
}


function testCreateIndependentNativeInciso_() {
  const started = Date.now();

  const activeDoc = DocumentApp.getActiveDocument();
  const documentId = activeDoc.getId();
  const activeTabId = activeDoc.getActiveTab().getId();

  const uniqueId = Utilities.getUuid().replace(/-/g, '');
  const targetMarker = 'DASTART' + uniqueId;
  const separatorMarker = 'DASEPARATOR' + uniqueId;

  let apiReadMs = 0;
  let apiWriteMs = 0;

  try {
    const target = getSingleSelectedParagraph_(activeDoc);

    if (target.getType() !== DocumentApp.ElementType.PARAGRAPH) {
      throw new Error(
        'Seleccione solamente un párrafo normal, por ejemplo Octubre.'
      );
    }

    if (/[\r\n]/.test(target.getText())) {
      throw new Error(
        'Esta prueba acepta un solo párrafo sin saltos de línea internos.'
      );
    }

    const parent = target.getParent();

    if (
      !parent ||
      parent.getType() !== DocumentApp.ElementType.BODY_SECTION
    ) {
      throw new Error(
        'Para esta prueba, el párrafo debe estar directamente en el cuerpo del documento.'
      );
    }

    /*
     * El separador impide que Docs una el nuevo inciso con una lista
     * anterior que tenga el mismo preset.
     */
    const targetIndex = parent.getChildIndex(target);
    parent.insertParagraph(targetIndex, separatorMarker);

    /*
     * Un tabulador inicial hace que el preset use el segundo nivel:
     * nivel 0 = 1)
     * nivel 1 = a)
     */
    target.editAsText().insertText(0, '\t' + targetMarker);

    /*
     * A partir de aquí no se vuelve a utilizar ningún objeto obtenido
     * desde activeDoc, porque el documento quedará cerrado.
     */
    activeDoc.saveAndClose();

    const readStarted = Date.now();

    const apiDocument = Docs.Documents.get(documentId, {
      includeTabsContent: true,
      fields: buildTest022Fields_()
    });

    apiReadMs = Date.now() - readStarted;

    const located = findApiParagraphWithMarker_(
      apiDocument.tabs || [],
      targetMarker
    );

    if (!located) {
      throw new Error(
        'No se encontró el marcador temporal en la respuesta reducida de Docs API.'
      );
    }

    const writeStarted = Date.now();

    Docs.Documents.batchUpdate(
      {
        requests: [
          {
            createParagraphBullets: {
              range: {
                startIndex: located.startIndex,
                endIndex: located.endIndex,
                tabId: located.tabId
              },
              bulletPreset:
                'NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS'
            }
          }
        ]
      },
      documentId
    );

    apiWriteMs = Date.now() - writeStarted;

    const reopenedDoc = DocumentApp.openById(documentId);
    const reopenedTab = reopenedDoc.getTab(located.tabId);

    if (!reopenedTab) {
      throw new Error(
        'No se pudo volver a abrir la pestaña donde está el párrafo.'
      );
    }

    const body = reopenedTab.asDocumentTab().getBody();
    const markerResult = body.findText(targetMarker);

    if (!markerResult) {
      throw new Error(
        'Docs API aplicó la operación, pero no se encontró el párrafo resultante.'
      );
    }

    const listItem = findParagraphAncestor_(
      markerResult.getElement()
    );

    if (
      !listItem ||
      listItem.getType() !== DocumentApp.ElementType.LIST_ITEM
    ) {
      throw new Error(
        'El párrafo resultante no es un elemento de lista nativo.'
      );
    }

    /*
     * El marcador se elimina; nunca queda como parte del texto real.
     */
    listItem.replaceText(targetMarker, '');

    /*
     * Left 0.25", hanging 0.25", right 0".
     */
    listItem
      .setIndentFirstLine(18)
      .setIndentStart(36)
      .setIndentEnd(0);

    const separatorRemoved = removeMarkerParagraph_(
      body,
      separatorMarker
    );

    const result = {
      testId: 'TEST-022-NATIVE-INDEPENDENT-INCISO',
      ok: true,
      isNative: true,
      glyphType: String(listItem.getGlyphType()),
      nestingLevel: listItem.getNestingLevel(),
      listId: listItem.getListId(),
      separatorRemoved: separatorRemoved,
      apiReadMs: apiReadMs,
      apiWriteMs: apiWriteMs,
      elapsedMs: Date.now() - started
    };

    reopenedDoc.saveAndClose();
    return result;

  } catch (error) {
    cleanupTest022Markers_(
      documentId,
      activeTabId,
      targetMarker,
      separatorMarker
    );

    throw new Error(
      'TEST-022: ' + (error && error.message
        ? error.message
        : String(error))
    );
  }
}


/**
 * Obtiene un único párrafo desde una selección o desde el cursor.
 */
function getSingleSelectedParagraph_(doc) {
  const selection = doc.getSelection();

  if (selection) {
    const rangeElements = selection.getRangeElements();
    let selectedParagraph = null;

    for (let i = 0; i < rangeElements.length; i++) {
      const paragraph = findParagraphAncestor_(
        rangeElements[i].getElement()
      );

      if (!paragraph) {
        continue;
      }

      if (!selectedParagraph) {
        selectedParagraph = paragraph;
        continue;
      }

      if (!sameDocumentElement_(selectedParagraph, paragraph)) {
        throw new Error(
          'Seleccione solamente un párrafo para esta prueba.'
        );
      }
    }

    if (selectedParagraph) {
      return selectedParagraph;
    }
  }

  const cursor = doc.getCursor();

  if (!cursor) {
    throw new Error(
      'Seleccione un párrafo normal o coloque el cursor dentro de él.'
    );
  }

  const cursorParagraph = findParagraphAncestor_(
    cursor.getElement()
  );

  if (!cursorParagraph) {
    throw new Error(
      'No se pudo identificar el párrafo que contiene el cursor.'
    );
  }

  return cursorParagraph;
}


/**
 * Sube desde Text hasta Paragraph o ListItem.
 */
function findParagraphAncestor_(element) {
  let current = element;

  while (current) {
    const type = current.getType();

    if (
      type === DocumentApp.ElementType.PARAGRAPH ||
      type === DocumentApp.ElementType.LIST_ITEM
    ) {
      return current;
    }

    current = current.getParent();
  }

  return null;
}


/**
 * Compara dos referencias sin depender de la igualdad de proxies.
 */
function sameDocumentElement_(first, second) {
  if (first === second) {
    return true;
  }

  const firstParent = first.getParent();
  const secondParent = second.getParent();

  if (!firstParent || firstParent !== secondParent) {
    return false;
  }

  return firstParent.getChildIndex(first) ===
    secondParent.getChildIndex(second);
}


/**
 * Field mask reducido.
 *
 * Incluye solamente:
 * - identificador de pestaña
 * - índices de párrafos
 * - contenido de textRun
 *
 * Se contemplan hasta cuatro niveles de pestañas anidadas.
 */
function buildTest022Fields_() {
  return 'tabs(' + buildTest022TabFields_(4) + ')';
}


function buildTest022TabFields_(remainingDepth) {
  let fields =
    'tabProperties(tabId),' +
    'documentTab(' +
      'body(' +
        'content(' +
          'startIndex,' +
          'endIndex,' +
          'paragraph(' +
            'elements(' +
              'textRun(content)' +
            ')' +
          ')' +
        ')' +
      ')' +
    ')';

  if (remainingDepth > 0) {
    fields +=
      ',childTabs(' +
        buildTest022TabFields_(remainingDepth - 1) +
      ')';
  }

  return fields;
}


/**
 * Busca el marcador dentro de todas las pestañas devueltas por Docs API.
 */
function findApiParagraphWithMarker_(tabs, marker) {
  for (let i = 0; i < tabs.length; i++) {
    const tab = tabs[i];
    const tabId = tab.tabProperties
      ? tab.tabProperties.tabId
      : null;

    const content =
      tab.documentTab &&
      tab.documentTab.body &&
      tab.documentTab.body.content
        ? tab.documentTab.body.content
        : [];

    for (let j = 0; j < content.length; j++) {
      const structuralElement = content[j];

      if (!structuralElement.paragraph) {
        continue;
      }

      const paragraphText = getApiParagraphText_(
        structuralElement.paragraph
      );

      if (paragraphText.indexOf(marker) !== -1) {
        return {
          tabId: tabId,
          startIndex: structuralElement.startIndex,
          endIndex: structuralElement.endIndex
        };
      }
    }

    const childResult = findApiParagraphWithMarker_(
      tab.childTabs || [],
      marker
    );

    if (childResult) {
      return childResult;
    }
  }

  return null;
}


function getApiParagraphText_(paragraph) {
  const elements = paragraph.elements || [];
  let text = '';

  for (let i = 0; i < elements.length; i++) {
    if (elements[i].textRun) {
      text += elements[i].textRun.content || '';
    }
  }

  return text;
}


/**
 * Elimina el párrafo separador temporal.
 */
function removeMarkerParagraph_(body, marker) {
  const result = body.findText(marker);

  if (!result) {
    return false;
  }

  const paragraph = findParagraphAncestor_(
    result.getElement()
  );

  if (!paragraph) {
    return false;
  }

  paragraph.removeFromParent();
  return true;
}


/**
 * Limpieza de emergencia si cualquier fase falla.
 */
function cleanupTest022Markers_(
  documentId,
  tabId,
  targetMarker,
  separatorMarker
) {
  try {
    const doc = DocumentApp.openById(documentId);
    const tab = doc.getTab(tabId);

    if (!tab) {
      return;
    }

    const body = tab.asDocumentTab().getBody();

    const targetResult = body.findText(targetMarker);

    if (targetResult) {
      const paragraph = findParagraphAncestor_(
        targetResult.getElement()
      );

      if (paragraph) {
        paragraph.replaceText(targetMarker, '');

        const text = paragraph.editAsText();

        if (
          text.getText().length > 0 &&
          text.getText().charAt(0) === '\t'
        ) {
          text.deleteText(0, 0);
        }
      }
    }

    removeMarkerParagraph_(body, separatorMarker);
    doc.saveAndClose();

  } catch (cleanupError) {
    // No ocultar el error original.
  }
}
