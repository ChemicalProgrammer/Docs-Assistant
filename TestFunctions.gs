/**
 * TEST-023
 * Crea una lista nativa a) independiente sin leer el documento completo.
 */
function runCurrentTest() {
  return testNativeIncisoWithNamedRange_();
}


function testNativeIncisoWithNamedRange_() {
  const started = Date.now();
  const doc = DocumentApp.getActiveDocument();
  const documentId = doc.getId();
  const tabId = doc.getActiveTab().getId();
  const documentTab = doc.getActiveTab().asDocumentTab();

  const target = getSingleParagraph023_(doc);

  if (target.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    throw new Error(
      'TEST-024: coloque el cursor dentro de un párrafo normal.'
    );
  }

  if (/[\r\n]/.test(target.getText())) {
    throw new Error(
      'TEST-024: use un párrafo independiente sin Shift+Enter.'
    );
  }

  const parent = target.getParent();

  if (
    !parent ||
    parent.getType() !== DocumentApp.ElementType.BODY_SECTION
  ) {
    throw new Error(
      'TEST-024: el párrafo debe estar en el cuerpo del documento.'
    );
  }

  const uniqueId = Utilities.getUuid().replace(/-/g, '');
  const rangeName = 'DOCSASSISTANT_TEST_024_' + uniqueId;
  const separatorMarker = 'DASEP024' + uniqueId;

  /*
   * Impide que Docs continúe automáticamente una lista anterior.
   */
  const targetIndex = parent.getChildIndex(target);
  parent.insertParagraph(targetIndex, separatorMarker);

  /*
   * Crear el rango explícitamente en la pestaña activa.
   */
  const rangeBuilder = documentTab.newRange();
  rangeBuilder.addElement(target);

  const namedRange = documentTab.addNamedRange(
    rangeName,
    rangeBuilder.build()
  );

  const namedRangeId = namedRange.getId();

  let apiReadMs = 0;
  let apiWriteMs = 0;

  doc.saveAndClose();

  try {
    /*
     * Pausa corta solamente para que Docs API vea el rango recién guardado.
     */
    Utilities.sleep(250);

    const readStarted = Date.now();

    const apiDocument = Docs.Documents.get(documentId, {
      includeTabsContent: true,
      fields: buildNamedRangeFields024_(4)
    });

    apiReadMs = Date.now() - readStarted;

    const locatedRange = findNamedRange024_(
      apiDocument.tabs || [],
      rangeName,
      namedRangeId
    );

    if (!locatedRange) {
      throw new Error(
        'Docs API no devolvió el rango con nombre.'
      );
    }

    const writeStarted = Date.now();

    Docs.Documents.batchUpdate(
      {
        requests: [
          {
            createParagraphBullets: {
              range: {
                startIndex: locatedRange.startIndex,
                endIndex: locatedRange.endIndex,
                tabId: locatedRange.tabId
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
    const reopenedTabObject =
      reopenedDoc.getTab(locatedRange.tabId);

    if (!reopenedTabObject) {
      throw new Error(
        'No se pudo volver a abrir la pestaña.'
      );
    }

    const reopenedTab =
      reopenedTabObject.asDocumentTab();

    const localNamedRange =
      reopenedTab.getNamedRangeById(namedRangeId);

    if (!localNamedRange) {
      throw new Error(
        'No se encontró localmente el rango después de aplicar la lista.'
      );
    }

    const rangeElements =
      localNamedRange.getRange().getRangeElements();

    let listItem = null;

    for (let i = 0; i < rangeElements.length; i++) {
      const paragraph = findParagraph023_(
        rangeElements[i].getElement()
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
        'El elemento resultante no es una lista nativa.'
      );
    }

    /*
     * El preset se crea en nivel 0 = 1).
     * Forzamos nivel 1 = a).
     */
    listItem
      .setNestingLevel(1)
      .setIndentFirstLine(18)
      .setIndentStart(36)
      .setIndentEnd(0);

    localNamedRange.remove();

    const separatorRemoved = removeSeparator024_(
      reopenedTab.getBody(),
      separatorMarker
    );

    const result = {
      testId: 'TEST-024-NAMED-RANGE-NATIVE-INCISO',
      ok: true,
      isNative: true,
      glyphType: String(listItem.getGlyphType()),
      nestingLevel: listItem.getNestingLevel(),
      listId: listItem.getListId(),
      separatorRemoved: separatorRemoved,
      namedRangeRemoved: true,
      apiReadMs: apiReadMs,
      apiWriteMs: apiWriteMs,
      elapsedMs: Date.now() - started
    };

    reopenedDoc.saveAndClose();
    return result;

  } catch (error) {
    cleanupTest024_(
      documentId,
      tabId,
      namedRangeId,
      separatorMarker
    );

    throw new Error(
      'TEST-024: ' +
      (error && error.message
        ? error.message
        : String(error))
    );
  }
}


function buildNamedRangeFields024_(depth) {
  return 'tabs(' + buildTabFields024_(depth) + ')';
}


function buildTabFields024_(depth) {
  let fields =
    'tabProperties(tabId),' +
    'documentTab(namedRanges)';

  if (depth > 0) {
    fields +=
      ',childTabs(' +
      buildTabFields024_(depth - 1) +
      ')';
  }

  return fields;
}


/**
 * documentTab.namedRanges es un mapa indexado por nombre:
 *
 * namedRanges[nombre].namedRanges[]
 */
function findNamedRange024_(tabs, rangeName, rangeId) {
  for (let i = 0; i < tabs.length; i++) {
    const tab = tabs[i];
    const currentTabId = tab.tabProperties
      ? tab.tabProperties.tabId
      : null;

    const map =
      tab.documentTab &&
      tab.documentTab.namedRanges
        ? tab.documentTab.namedRanges
        : {};

    const group = map[rangeName];

    if (group && group.namedRanges) {
      for (let j = 0; j < group.namedRanges.length; j++) {
        const candidate = group.namedRanges[j];

        if (
          !rangeId ||
          candidate.namedRangeId === rangeId
        ) {
          const ranges = candidate.ranges || [];

          if (ranges.length > 0) {
            return {
              startIndex: ranges[0].startIndex,
              endIndex: ranges[0].endIndex,
              tabId: ranges[0].tabId || currentTabId
            };
          }
        }
      }
    }

    const childResult = findNamedRange024_(
      tab.childTabs || [],
      rangeName,
      rangeId
    );

    if (childResult) {
      return childResult;
    }
  }

  return null;
}


function removeSeparator024_(body, marker) {
  const result = body.findText(marker);

  if (!result) {
    return false;
  }

  const paragraph = findParagraph023_(
    result.getElement()
  );

  if (!paragraph) {
    return false;
  }

  paragraph.removeFromParent();
  return true;
}


function cleanupTest024_(
  documentId,
  tabId,
  namedRangeId,
  separatorMarker
) {
  try {
    const cleanupDoc =
      DocumentApp.openById(documentId);

    const tabObject = cleanupDoc.getTab(tabId);

    if (!tabObject) {
      return;
    }

    const cleanupTab = tabObject.asDocumentTab();

    const namedRange =
      cleanupTab.getNamedRangeById(namedRangeId);

    if (namedRange) {
      namedRange.remove();
    }

    removeSeparator024_(
      cleanupTab.getBody(),
      separatorMarker
    );

    cleanupDoc.saveAndClose();

  } catch (cleanupError) {
    // No ocultar el error original.
  }
}


function testFastIndependentNativeInciso_() {
  const started = Date.now();
  const activeDoc = DocumentApp.getActiveDocument();
  const documentId = activeDoc.getId();
  const activeTabId = activeDoc.getActiveTab().getId();

  const target = getSingleParagraph023_(activeDoc);

  if (target.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    throw new Error(
      'TEST-023: coloque el cursor dentro de un párrafo normal.'
    );
  }

  if (/[\r\n]/.test(target.getText())) {
    throw new Error(
      'TEST-023: use un párrafo independiente sin Shift+Enter.'
    );
  }

  const targetParent = target.getParent();

  if (
    !targetParent ||
    targetParent.getType() !==
      DocumentApp.ElementType.BODY_SECTION
  ) {
    throw new Error(
      'TEST-023: el párrafo debe estar en el cuerpo del documento.'
    );
  }

  const targetIndex = targetParent.getChildIndex(target);
  const originalText = target.getText();
  const originalAttributes = target.getAttributes();

  let temporaryTabId = null;
  let apiWriteMs = 0;

  activeDoc.saveAndClose();

  try {
    /*
     * 1. Crear una pestaña temporal.
     */
    let apiStarted = Date.now();

    const addResponse = Docs.Documents.batchUpdate(
      {
        requests: [
          {
            addDocumentTab: {
              tabProperties: {
                title: 'DocsAssistant temporary list'
              }
            }
          }
        ]
      },
      documentId
    );

    apiWriteMs += Date.now() - apiStarted;

    temporaryTabId =
      addResponse &&
      addResponse.replies &&
      addResponse.replies[0] &&
      addResponse.replies[0].addDocumentTab &&
      addResponse.replies[0].addDocumentTab.tabProperties
        ? addResponse.replies[0]
            .addDocumentTab.tabProperties.tabId
        : null;

    if (!temporaryTabId) {
      throw new Error(
        'Docs API no devolvió el ID de la pestaña temporal.'
      );
    }

    /*
     * 2. En una pestaña nueva conocemos los índices:
     *    el primer párrafo comienza en el índice 1.
     */
    const templateText = 'DOCSASSISTANT_NATIVE_INCISO_TEMPLATE';

    apiStarted = Date.now();

    Docs.Documents.batchUpdate(
      {
        requests: [
          {
            insertText: {
              location: {
                index: 1,
                tabId: temporaryTabId
              },
              text: templateText
            }
          },
          {
            createParagraphBullets: {
              range: {
                startIndex: 1,
                endIndex: 1 + templateText.length,
                tabId: temporaryTabId
              },
              bulletPreset:
                'NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS'
            }
          }
        ]
      },
      documentId
    );

    apiWriteMs += Date.now() - apiStarted;

    /*
     * 3. Copiar localmente la lista nativa al párrafo seleccionado.
     */
    const workingDoc = DocumentApp.openById(documentId);

    const temporaryTab = workingDoc.getTab(temporaryTabId);
    const destinationTab = workingDoc.getTab(activeTabId);

    if (!temporaryTab || !destinationTab) {
      throw new Error(
        'No se pudieron abrir las pestañas necesarias.'
      );
    }

    const temporaryBody =
      temporaryTab.asDocumentTab().getBody();

    const templateItems = temporaryBody.getListItems();

    if (!templateItems || templateItems.length === 0) {
      throw new Error(
        'No se creó el elemento de lista nativo temporal.'
      );
    }

    const templateItem = templateItems[0];

    /*
     * Preset:
     * nivel 0 = 1)
     * nivel 1 = a)
     * nivel 2 = i)
     */
    templateItem.setNestingLevel(1);

    const destinationBody =
      destinationTab.asDocumentTab().getBody();

    const originalParagraph =
      destinationBody.getChild(targetIndex);

    if (
      !originalParagraph ||
      originalParagraph.getType() !==
        DocumentApp.ElementType.PARAGRAPH
    ) {
      throw new Error(
        'El párrafo seleccionado cambió durante la prueba.'
      );
    }

    if (originalParagraph.getText() !== originalText) {
      throw new Error(
        'El contenido del párrafo cambió durante la prueba.'
      );
    }

    const insertedItem = destinationBody.insertListItem(
      targetIndex,
      templateItem.copy()
    );

    /*
     * Reemplazar únicamente el contenido de la plantilla.
     * La configuración nativa a) y el listId permanecen.
     */
    insertedItem.setText(originalText);

    insertedItem
      .setAttributes(originalAttributes)
      .setHeading(DocumentApp.ParagraphHeading.NORMAL)
      .setNestingLevel(1)
      .setIndentFirstLine(18)
      .setIndentStart(36)
      .setIndentEnd(0);

    /*
     * El elemento nuevo está antes del párrafo original.
     */
    originalParagraph.removeFromParent();

    const createdListId = insertedItem.getListId();

    workingDoc.saveAndClose();

    /*
     * 4. Eliminar la pestaña temporal.
     */
    apiStarted = Date.now();

    Docs.Documents.batchUpdate(
      {
        requests: [
          {
            deleteTab: {
              tabId: temporaryTabId
            }
          }
        ]
      },
      documentId
    );

    apiWriteMs += Date.now() - apiStarted;
    temporaryTabId = null;

    /*
     * 5. Verificación final después de eliminar la pestaña.
     */
    const verificationDoc = DocumentApp.openById(documentId);
    const verificationTab =
      verificationDoc.getTab(activeTabId);

    if (!verificationTab) {
      throw new Error(
        'No se pudo abrir la pestaña para verificar.'
      );
    }

    const verificationBody =
      verificationTab.asDocumentTab().getBody();

    const verifiedElement =
      verificationBody.getChild(targetIndex);

    if (
      !verifiedElement ||
      verifiedElement.getType() !==
        DocumentApp.ElementType.LIST_ITEM
    ) {
      throw new Error(
        'El elemento final no es una lista nativa.'
      );
    }

    const verifiedItem = verifiedElement.asListItem();

    let sharesEarlierListId = false;

    for (let i = 0; i < targetIndex; i++) {
      const earlier = verificationBody.getChild(i);

      if (
        earlier.getType() ===
        DocumentApp.ElementType.LIST_ITEM
      ) {
        const earlierItem = earlier.asListItem();

        if (
          earlierItem.getListId() ===
          verifiedItem.getListId()
        ) {
          sharesEarlierListId = true;
          break;
        }
      }
    }

    const result = {
      testId: 'TEST-023-FAST-NATIVE-INDEPENDENT-INCISO',
      ok: true,
      isNative: true,
      glyphType: String(verifiedItem.getGlyphType()),
      nestingLevel: verifiedItem.getNestingLevel(),
      listId: verifiedItem.getListId(),
      originalListId: createdListId,
      independentFromEarlierLists: !sharesEarlierListId,
      temporaryTabRemoved: true,
      apiReadMs: 0,
      apiWriteMs: apiWriteMs,
      elapsedMs: Date.now() - started
    };

    verificationDoc.saveAndClose();
    return result;

  } catch (error) {
    if (temporaryTabId) {
      deleteTemporaryTab023_(
        documentId,
        temporaryTabId
      );
    }

    throw new Error(
      'TEST-023: ' +
      (error && error.message
        ? error.message
        : String(error))
    );
  }
}


function getSingleParagraph023_(doc) {
  const selection = doc.getSelection();

  if (selection) {
    const rangeElements = selection.getRangeElements();
    let paragraph = null;

    for (let i = 0; i < rangeElements.length; i++) {
      const current = findParagraph023_(
        rangeElements[i].getElement()
      );

      if (!current) {
        continue;
      }

      if (!paragraph) {
        paragraph = current;
        continue;
      }

      if (!sameElement023_(paragraph, current)) {
        throw new Error(
          'TEST-023: seleccione solamente un párrafo.'
        );
      }
    }

    if (paragraph) {
      return paragraph;
    }
  }

  const cursor = doc.getCursor();

  if (!cursor) {
    throw new Error(
      'TEST-023: coloque el cursor dentro de un párrafo.'
    );
  }

  const paragraph = findParagraph023_(
    cursor.getElement()
  );

  if (!paragraph) {
    throw new Error(
      'TEST-023: no se encontró el párrafo.'
    );
  }

  return paragraph;
}


function findParagraph023_(element) {
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


function sameElement023_(first, second) {
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


function deleteTemporaryTab023_(documentId, tabId) {
  try {
    Docs.Documents.batchUpdate(
      {
        requests: [
          {
            deleteTab: {
              tabId: tabId
            }
          }
        ]
      },
      documentId
    );
  } catch (cleanupError) {
    // No ocultar el error original.
  }
}
