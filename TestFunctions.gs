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
function runCurrentTest() {
  const started = Date.now();

  const body =
    getActiveBody_();

  const targets =
    getStyleTargetParagraphs_();

  if (targets.length !== 1) {
    throw new Error(
      'Place the cursor in one equation line or select only that line.'
    );
  }

  const source =
    targets[0];

  const top =
    getTopLevelBodyElement_(
      source,
      body
    );

  if (
    !top ||
    top.getType() !==
      DocumentApp.ElementType.PARAGRAPH
  ) {
    throw new Error(
      'The equation line must be a body-level normal paragraph.'
    );
  }

  const sourceIndex =
    body.getChildIndex(top);

  const tablesBefore =
    testCountBodyTables_(body);

  /*
   * Ejecuta la función real de producción.
   */
  const formatterResult =
    formatEquationLine();

  /*
   * formatEquationLine() inserta la tabla
   * en la posición del párrafo original.
   */
  const createdElement =
    body.getChild(sourceIndex);

  const createdIsTable =
    createdElement.getType() ===
      DocumentApp.ElementType.TABLE;

  const createdTable =
    createdIsTable
      ? createdElement.asTable()
      : null;

  const detectedAsEquationLayout =
    createdTable
      ? testIsEquationLayoutTable_(
          createdTable
        )
      : false;

  const tablesAfter =
    testCountBodyTables_(body);

  return {
    ok:
      createdIsTable &&
      detectedAsEquationLayout,

    testId:
      'EQUATION-CREATION-AND-DETECTION',

    equationNumber:
      formatterResult.equationNumber,

    createdAtBodyChildIndex:
      sourceIndex,

    createdIsTable:
      createdIsTable,

    createdRows:
      createdTable
        ? createdTable.getNumRows()
        : null,

    createdFirstRowCells:
      createdTable
        ? createdTable
            .getRow(0)
            .getNumCells()
        : null,

    equationLayoutDetected:
      detectedAsEquationLayout,

    wouldBeCountedAsCaptionTable:
      createdIsTable &&
      !detectedAsEquationLayout,

    physicalTablesBefore:
      tablesBefore,

    physicalTablesAfter:
      tablesAfter,

    physicalTableDelta:
      tablesAfter - tablesBefore,

    usesAdvancedDocsApi:
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
