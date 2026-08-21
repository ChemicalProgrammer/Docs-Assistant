/**
 * TEST-023
 * Crea una lista nativa a) independiente sin leer el documento completo.
 */
function runCurrentTest() {
  return testFastIndependentNativeInciso_();
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
