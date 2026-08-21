/**
 * ============================================
 * SOLUCIÓN SIMPLE (RECOMENDADA PARA LA MAYORÍA)
 * ============================================
 */

/**
 * Convierte las líneas seleccionadas en incisos a), b), c)
 * Usa la API estándar de DocumentApp - No requiere configuraciones adicionales
 */
function aplicarIncisosSimple() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (!selection) {
    DocumentApp.getUi().alert(
      'Por favor, selecciona el texto que deseas convertir en incisos.'
    );
    return;
  }
  
  const elements = selection.getRangeElements();
  let textoCompleto = '';
  
  // Extraer todo el texto de la selección
  for (let element of elements) {
    const el = element.getElement();
    if (el.getType() === DocumentApp.ElementType.TEXT) {
      textoCompleto += el.asText().getText();
    } else if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      textoCompleto += el.asParagraph().getText();
    } else if (el.getType() === DocumentApp.ElementType.LIST_ITEM) {
      textoCompleto += el.asListItem().getText();
    }
  }
  
  // Dividir en líneas y limpiar
  const lineas = textoCompleto
    .split('\n')
    .map(linea => linea.trim())
    .filter(linea => linea.length > 0);
  
  if (lineas.length === 0) {
    DocumentApp.getUi().alert('No se encontraron líneas para convertir.');
    return;
  }
  
  // Generar texto con incisos
  const letras = 'abcdefghijklmnopqrstuvwxyz';
  let nuevoTexto = '';
  
  for (let i = 0; i < lineas.length; i++) {
    let letra;
    if (i < letras.length) {
      letra = letras[i];
    } else {
      // Para más de 26 incisos
      const primera = letras[Math.floor(i / 26) - 1] || '';
      const segunda = letras[i % 26];
      letra = primera + segunda;
    }
    nuevoTexto += `${letra}) ${lineas[i]}\n`;
  }
  
  // Reemplazar el texto seleccionado
  try {
    const firstElement = elements[0].getElement();
    const lastElement = elements[elements.length - 1].getElement();
    
    // Si es un solo elemento de texto
    if (firstElement === lastElement && firstElement.getType() === DocumentApp.ElementType.TEXT) {
      const text = firstElement.asText();
      const startOffset = elements[0].getStartOffset();
      const endOffset = elements[elements.length - 1].getEndOffsetInclusive();
      text.deleteText(startOffset, endOffset);
      text.insertText(startOffset, nuevoTexto);
    } else {
      // Para selecciones múltiples, reemplazar todo el contenido
      // Primero, obtener el rango completo
      const body = doc.getBody();
      const firstIndex = body.getChildIndex(firstElement);
      const lastIndex = body.getChildIndex(lastElement);
      
      // Eliminar todos los elementos en el rango
      for (let i = lastIndex; i >= firstIndex; i--) {
        const child = body.getChild(i);
        if (child) {
          child.removeFromParent();
        }
      }
      
      // Insertar el nuevo texto como un solo párrafo
      body.insertParagraph(firstIndex, nuevoTexto);
    }
    
    DocumentApp.getUi().alert(`✅ ${lineas.length} líneas convertidas a incisos.`);
  } catch (error) {
    DocumentApp.getUi().alert('Error al reemplazar el texto: ' + error.toString());
  }
}

/**
 * Versión mejorada que preserva el formato de los párrafos
 * Convierte cada párrafo en un LIST_ITEM con formato de inciso
 */
function aplicarIncisosConFormato() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (!selection) {
    DocumentApp.getUi().alert('Por favor, selecciona los párrafos a convertir.');
    return;
  }
  
  const elements = selection.getRangeElements();
  const paragraphs = [];
  
  // Obtener todos los párrafos en la selección
  for (let element of elements) {
    let el = element.getElement();
    let paragraph = null;
    
    // Buscar el párrafo padre
    while (el) {
      const type = el.getType();
      if (type === DocumentApp.ElementType.PARAGRAPH || 
          type === DocumentApp.ElementType.LIST_ITEM) {
        paragraph = el;
        break;
      }
      el = el.getParent();
    }
    
    if (paragraph) {
      const index = doc.getBody().getChildIndex(paragraph);
      if (index !== -1 && !paragraphs.includes(paragraph)) {
        paragraphs.push(paragraph);
      }
    }
  }
  
  if (paragraphs.length === 0) {
    DocumentApp.getUi().alert('No se encontraron párrafos para convertir.');
    return;
  }
  
  // Ordenar por posición
  paragraphs.sort((a, b) => {
    return doc.getBody().getChildIndex(a) - doc.getBody().getChildIndex(b);
  });
  
  // Convertir cada párrafo a LIST_ITEM con inciso
  const letras = 'abcdefghijklmnopqrstuvwxyz';
  
  for (let i = 0; i < paragraphs.length; i++) {
    const paragraph = paragraphs[i];
    const texto = paragraph.getText();
    
    // Obtener la letra correspondiente
    let letra;
    if (i < letras.length) {
      letra = letras[i];
    } else {
      const primera = letras[Math.floor(i / 26) - 1] || '';
      const segunda = letras[i % 26];
      letra = primera + segunda;
    }
    
    // Crear un nuevo LIST_ITEM
    const body = doc.getBody();
    const index = body.getChildIndex(paragraph);
    
    // Crear el nuevo elemento LIST_ITEM
    const listItem = body.insertListItem(index, `${letra}) ${texto}`);
    
    // Configurar el formato de lista
    listItem.setGlyphType(DocumentApp.GlyphType.LOWER_ALPHA);
    listItem.setNestingLevel(0);
    
    // Eliminar el párrafo original
    paragraph.removeFromParent();
  }
  
  DocumentApp.getUi().alert(`✅ ${paragraphs.length} párrafos convertidos a incisos.`);
}


/**
 * ============================================
 * SOLUCIÓN AVANZADA (CON GOOGLE DOCS API)
 * ============================================
 * NOTA: Requiere activar el servicio Google Docs API
 * Extensiones → Apps Script → Servicios → Google Docs API
 */

/**
 * Aplica incisos usando la Google Docs API
 * Mayor control sobre el formato y sangrías
 */
function aplicarIncisosAvanzado() {
  // Verificar que el servicio esté activo
  if (typeof Docs === 'undefined' || !Docs.Documents) {
    DocumentApp.getUi().alert(
      '❌ Error: Activa el servicio "Google Docs API" en:\n' +
      'Extensiones → Apps Script → Servicios'
    );
    return;
  }
  
  try {
    const doc = DocumentApp.getActiveDocument();
    const documentId = doc.getId();
    const body = doc.getBody();
    const selection = doc.getSelection();
    
    if (!selection) {
      DocumentApp.getUi().alert('Por favor, selecciona los párrafos a convertir.');
      return;
    }
    
    // Obtener los párrafos seleccionados
    const elements = selection.getRangeElements();
    const paragraphs = [];
    
    for (let element of elements) {
      let el = element.getElement();
      while (el) {
        const type = el.getType();
        if (type === DocumentApp.ElementType.PARAGRAPH || 
            type === DocumentApp.ElementType.LIST_ITEM) {
          const index = body.getChildIndex(el);
          if (index !== -1 && !paragraphs.includes(el)) {
            paragraphs.push({
              element: el,
              index: index
            });
          }
          break;
        }
        el = el.getParent();
      }
    }
    
    if (paragraphs.length === 0) {
      DocumentApp.getUi().alert('No se encontraron párrafos.');
      return;
    }
    
    // Ordenar por índice
    paragraphs.sort((a, b) => a.index - b.index);
    
    // Crear marcadores temporales
    const markers = [];
    const markerPrefix = 'INCISO_' + Utilities.getUuid().replace(/-/g, '') + '_';
    
    for (let i = 0; i < paragraphs.length; i++) {
      const markerName = markerPrefix + i;
      const range = doc.newRange()
        .addElement(paragraphs[i].element)
        .build();
      
      doc.addNamedRange(markerName, range);
      markers.push(markerName);
    }
    
    // Guardar cambios para poder usar la API
    doc.saveAndClose();
    
    // Obtener el documento con los marcadores
    const apiDoc = Docs.Documents.get(documentId, {
      includeTabsContent: true,
      fields: 'tabs(documentTab(namedRanges))'
    });
    
    // Encontrar los índices de los párrafos
    const tabId = doc.getActiveTab().getId();
    const apiTab = apiDoc.tabs.find(t => 
      String(t.tabProperties?.tabId) === String(tabId)
    );
    
    if (!apiTab) {
      throw new Error('No se encontró el tab activo');
    }
    
    const namedRanges = apiTab.documentTab?.namedRanges || {};
    const ranges = [];
    
    for (let marker of markers) {
      const rangeGroup = namedRanges[marker];
      if (rangeGroup?.namedRanges?.[0]?.ranges?.[0]) {
        const range = rangeGroup.namedRanges[0].ranges[0];
        ranges.push({
          startIndex: Number(range.startIndex),
          endIndex: Number(range.endIndex)
        });
      }
    }
    
    if (ranges.length === 0) {
      throw new Error('No se encontraron los marcadores');
    }
    
    const startIndex = Math.min(...ranges.map(r => r.startIndex));
    const endIndex = Math.max(...ranges.map(r => r.endIndex));
    const tabIdStr = String(tabId);
    
    // Construir las solicitudes a la API
    const requests = [
      // Eliminar viñetas existentes
      {
        deleteParagraphBullets: {
          range: {
            startIndex: startIndex,
            endIndex: endIndex,
            tabId: tabIdStr
          }
        }
      },
      
      // Insertar tabulaciones para establecer el nivel
      ...ranges.map(range => ({
        insertText: {
          location: {
            index: range.startIndex,
            tabId: tabIdStr
          },
          text: '\t'
        }
      })),
      
      // Crear lista con formato a), b), c)
      {
        createParagraphBullets: {
          range: {
            startIndex: startIndex,
            endIndex: endIndex + ranges.length,
            tabId: tabIdStr
          },
          bulletPreset: 'NUMBERED_DECIMAL_ALPHA_ROMAN_PARENS'
        }
      },
      
      // Configurar sangrías
      {
        updateParagraphStyle: {
          range: {
            startIndex: startIndex,
            endIndex: endIndex + ranges.length,
            tabId: tabIdStr
          },
          paragraphStyle: {
            indentStart: { magnitude: 36, unit: 'PT' },
            indentFirstLine: { magnitude: 18, unit: 'PT' },
            indentEnd: { magnitude: 0, unit: 'PT' }
          },
          fields: 'indentStart,indentFirstLine,indentEnd'
        }
      },
      
      // Eliminar marcadores
      ...markers.map(marker => ({
        deleteNamedRange: {
          name: marker,
          tabsCriteria: {
            tabIds: [tabIdStr]
          }
        }
      }))
    ];
    
    // Ejecutar las solicitudes
    Docs.Documents.batchUpdate({
      requests: requests
    }, documentId);
    
    DocumentApp.getUi().alert(`✅ ${ranges.length} párrafos convertidos a incisos.`);
    
  } catch (error) {
    DocumentApp.getUi().alert('❌ Error: ' + error.toString());
  }
}


/**
 * ============================================
 * MENÚ PERSONALIZADO
 * ============================================
 */

/**
 * Crea un menú en la barra de Google Docs
 * Se ejecuta automáticamente al abrir el documento
 */
function onOpen() {
  const ui = DocumentApp.getUi();
  
  ui.createMenu('📝 Incisos')
    .addItem('🔤 Convertir a a), b), c)...', 'aplicarIncisosSimple')
    .addItem('🎨 Convertir con formato (LIST_ITEM)', 'aplicarIncisosConFormato')
    .addSeparator()
    .addItem('⚡ Avanzado (Google Docs API)', 'aplicarIncisosAvanzado')
    .addToUi();
}

/**
 * Función de prueba rápida
 * Útil para pruebas desde el editor de Apps Script
 */
function testIncisos() {
  DocumentApp.getUi().alert(
    'Selecciona el texto en el documento y luego ejecuta:\n' +
    '- aplicarIncisosSimple()\n' +
    '- aplicarIncisosConFormato()\n' +
    '- aplicarIncisosAvanzado()'
  );
}


/**
 * ============================================
 * FUNCIONES ADICIONALES ÚTILES
 * ============================================
 */

/**
 * Elimina todos los incisos del texto seleccionado
 */
function eliminarIncisos() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (!selection) {
    DocumentApp.getUi().alert('Selecciona el texto con incisos a eliminar.');
    return;
  }
  
  const elements = selection.getRangeElements();
  let textoCompleto = '';
  
  for (let element of elements) {
    const el = element.getElement();
    if (el.getType() === DocumentApp.ElementType.TEXT) {
      textoCompleto += el.asText().getText();
    }
  }
  
  // Eliminar patrones como "a) ", "b) ", etc.
  const lineas = textoCompleto.split('\n');
  const lineasLimpias = lineas.map(linea => {
    // Eliminar patrones: a) , b) , c) , etc.
    return linea.replace(/^[a-z]\)\s*/, '').replace(/^[a-z][a-z]\)\s*/, '');
  });
  
  const nuevoTexto = lineasLimpias.join('\n');
  
  // Reemplazar el texto
  const firstElement = elements[0].getElement();
  if (firstElement.getType() === DocumentApp.ElementType.TEXT) {
    const text = firstElement.asText();
    text.setText(nuevoTexto);
    DocumentApp.getUi().alert('✅ Incisos eliminados.');
  }
}


/**
 * Convierte números a letras para incisos
 * Soporta más de 26 incisos (aa, ab, ac...)
 */
function numeroALetra(numero) {
  const letras = 'abcdefghijklmnopqrstuvwxyz';
  if (numero < letras.length) {
    return letras[numero];
  }
  const primera = letras[Math.floor(numero / letras.length) - 1] || '';
  const segunda = letras[numero % letras.length];
  return primera + segunda;
}
