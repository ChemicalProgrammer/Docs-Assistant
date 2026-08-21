/**
 * TEST LAB
 *
 * Prueba completa:
 * 1. Proporciona temporalmente el selector que falta.
 * 2. Ejecuta formatEquationLine().
 * 3. Comprueba la tabla auxiliar recién creada.
 *
 * No utiliza la API avanzada.
 */


/**
 * Devuelve los párrafos afectados por el cursor
 * o por la selección.
 *
 * Esta función falta actualmente en Formatting.gs.
 */
function getStyleTargetParagraphs_() {
  const doc =
    DocumentApp.getActiveDocument();

  const selection =
    doc.getSelection();

  if (selection) {
    return getSelectedParagraphs_();
  }

  const cursor =
    doc.getCursor();

  if (!cursor) {
    return [];
  }

  const paragraph =
    getOwningParagraph_(
      cursor.getElement()
    );

  return paragraph
    ? [paragraph]
    : [];
}


/**
 * Devuelve exclusivamente el párrafo
 * donde está colocado el cursor.
 *
 * También falta actualmente en Formatting.gs
 * y es utilizado por Indentation.
 */
function getCurrentParagraph_() {
  const cursor =
    DocumentApp
      .getActiveDocument()
      .getCursor();

  if (!cursor) {
    return null;
  }

  return getOwningParagraph_(
    cursor.getElement()
  );
}


/**
 * Ejecuta la función real de Equation y comprueba
 * la tabla auxiliar creada.
 *
 * Debes colocar el cursor en una línea real
 * que quieras convertir en ecuación.
 */
/**
 * Comprueba el conteo después de haber creado
 * una tabla auxiliar de ecuación real.
 *
 * No modifica el documento.
 */
function runCurrentTest() {
  const started = Date.now();

  const body =
    getActiveBody_();

  const selectedTable =
    getActiveTable_();

  if (!selectedTable) {
    throw new Error(
      'Place the cursor inside a real table located after the equation.'
    );
  }

  let selectedTableIndex;

  try {
    selectedTableIndex =
      body.getChildIndex(
        selectedTable
      );
  } catch (error) {
    throw new Error(
      'The selected table must be a body-level table.'
    );
  }

  let physicalTables = 0;
  let equationLayoutTablesIgnored = 0;
  let countableTables = 0;

  let rawOrdinal = 0;
  let correctedOrdinal = 0;

  for (
    let childIndex = 0;
    childIndex < body.getNumChildren();
    childIndex++
  ) {
    const child =
      body.getChild(childIndex);

    if (
      child.getType() !==
      DocumentApp.ElementType.TABLE
    ) {
      continue;
    }

    physicalTables++;

    const isEquationLayout =
      testIsEquationLayoutTable_(
        child.asTable()
      );

    if (isEquationLayout) {
      equationLayoutTablesIgnored++;
    } else {
      countableTables++;
    }

    if (
      childIndex <= selectedTableIndex
    ) {
      rawOrdinal++;

      if (!isEquationLayout) {
        correctedOrdinal++;
      }
    }
  }

  const selectedIsEquationLayout =
    testIsEquationLayoutTable_(
      selectedTable
    );

  return {
    ok:
      !selectedIsEquationLayout &&
      equationLayoutTablesIgnored >= 1,

    testId:
      'REAL-GHOST-TABLE-COUNTING',

    physicalTables:
      physicalTables,

    equationLayoutTablesIgnored:
      equationLayoutTablesIgnored,

    countableTables:
      countableTables,

    selectedTableIsEquationLayout:
      selectedIsEquationLayout,

    selectedTableRawOrdinal:
      rawOrdinal,

    selectedTableCorrectedOrdinal:
      correctedOrdinal,

    numbersDifferBy:
      rawOrdinal - correctedOrdinal,

    usesAdvancedDocsApi:
      false,

    modifiesDocument:
      false,

    elapsedMs:
      Date.now() - started
  };
}


/**
 * Detecta la estructura exacta que genera
 * formatEquationLine().
 */
function testIsEquationLayoutTable_(table) {
  try {
    if (
      !table ||
      table.getNumRows() !== 1
    ) {
      return false;
    }

    const row =
      table.getRow(0);

    if (row.getNumCells() !== 3) {
      return false;
    }

    const leftText =
      testNormalizeTableText_(
        row.getCell(0).getText()
      );

    const rightText =
      testNormalizeTableText_(
        row.getCell(2).getText()
      );

    return (
      leftText === '' &&
      /^\.{10,}\s*Equation\s+\d+\s*$/i
        .test(rightText)
    );
  } catch (error) {
    return false;
  }
}


/**
 * Cuenta solamente las tablas físicas de
 * nivel principal, sin inspeccionar su contenido.
 */
function testCountBodyTables_(body) {
  let count = 0;

  for (
    let index = 0;
    index < body.getNumChildren();
    index++
  ) {
    if (
      body.getChild(index).getType() ===
      DocumentApp.ElementType.TABLE
    ) {
      count++;
    }
  }

  return count;
}


/**
 * Normaliza espacios especiales de Google Docs.
 */
function testNormalizeTableText_(value) {
  return String(value || '')
    .replace(/\u00A0/g, ' ')
    .trim();
}
