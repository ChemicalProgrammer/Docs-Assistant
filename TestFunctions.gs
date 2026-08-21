/**
 * TEST LAB
 *
 * Comprueba que las tablas auxiliares utilizadas
 * para presentar ecuaciones no se cuenten como
 * tablas reales.
 *
 * Esta prueba:
 * - No modifica el documento.
 * - No utiliza la API avanzada de Google Docs.
 */
function runCurrentTest() {
  const started = Date.now();

  const body = getActiveBody_();
  const targetTable = getActiveTable_();

  if (!targetTable) {
    throw new Error(
      'Place the cursor inside the table you want to test.'
    );
  }

  /*
   * La numeración de captions trabaja únicamente
   * con tablas de nivel principal en el Body.
   */
  let targetIndex;

  try {
    targetIndex = body.getChildIndex(targetTable);
  } catch (error) {
    throw new Error(
      'The selected table must be a body-level table.'
    );
  }

  let physicalTables = 0;
  let equationLayoutTables = 0;
  let countableTables = 0;

  let rawOrdinalAtTarget = 0;
  let correctedOrdinalAtTarget = 0;

  const tables = [];

  for (
    let childIndex = 0;
    childIndex < body.getNumChildren();
    childIndex++
  ) {
    const child = body.getChild(childIndex);

    if (
      child.getType() !==
      DocumentApp.ElementType.TABLE
    ) {
      continue;
    }

    const table = child.asTable();

    physicalTables++;

    const isEquationLayout =
      testIsEquationLayoutTable_(table);

    if (isEquationLayout) {
      equationLayoutTables++;
    } else {
      countableTables++;
    }

    /*
     * Calcula ambos ordinales hasta la tabla
     * donde está colocado el cursor.
     */
    if (childIndex <= targetIndex) {
      rawOrdinalAtTarget++;

      if (!isEquationLayout) {
        correctedOrdinalAtTarget++;
      }
    }

    tables.push(
      testDescribeTable_(
        table,
        childIndex,
        physicalTables,
        isEquationLayout
      )
    );
  }

  const targetIsEquationLayout =
    testIsEquationLayoutTable_(targetTable);

  return {
    ok: true,
    testId: 'CAPTION-GHOST-TABLE-DETECTION',
    message: targetIsEquationLayout
      ? 'The selected table was identified as an equation layout table.'
      : (
          'Raw table ordinal: ' +
          rawOrdinalAtTarget +
          '. Corrected ordinal: ' +
          correctedOrdinalAtTarget +
          '.'
        ),
    usesAdvancedDocsApi: false,
    modifiesDocument: false,

    totals: {
      physicalTables: physicalTables,
      equationLayoutTablesIgnored:
        equationLayoutTables,
      countableTables: countableTables
    },

    selectedTable: {
      bodyChildIndex: targetIndex,
      isEquationLayout:
        targetIsEquationLayout,
      rawOrdinal:
        rawOrdinalAtTarget,
      correctedOrdinal:
        targetIsEquationLayout
          ? null
          : correctedOrdinalAtTarget
    },

    tables: tables,
    elapsedMs: Date.now() - started
  };
}


/**
 * Detecta exclusivamente la estructura creada por
 * formatEquationLine().
 *
 * Estructura esperada:
 * - Una fila.
 * - Tres celdas.
 * - Primera celda vacía.
 * - Tercera celda con puntos y "Equation N".
 *
 * No depende del ancho del borde porque una tabla
 * auxiliar podría haber sido formateada accidentalmente.
 */
function testIsEquationLayoutTable_(table) {
  try {
    if (!table || table.getNumRows() !== 1) {
      return false;
    }

    const row = table.getRow(0);

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

    /*
     * formatEquationLine() genera:
     *
     * .................... Equation 1
     */
    const hasEquationLabel =
      /^\.{10,}\s*Equation\s+\d+\s*$/i
        .test(rightText);

    return (
      leftText === '' &&
      hasEquationLabel
    );
  } catch (error) {
    return false;
  }
}


/**
 * Indica si una tabla debe participar
 * en la numeración de Table captions.
 */
function testIsCountableCaptionTable_(element) {
  if (
    !element ||
    element.getType() !==
      DocumentApp.ElementType.TABLE
  ) {
    return false;
  }

  return !testIsEquationLayoutTable_(
    element.asTable()
  );
}


/**
 * Genera información diagnóstica de una tabla.
 */
function testDescribeTable_(
  table,
  bodyChildIndex,
  physicalOrdinal,
  isEquationLayout
) {
  let firstRowCells = 0;
  let borderWidth = null;
  let rightCellPreview = '';

  try {
    if (table.getNumRows() > 0) {
      const firstRow = table.getRow(0);

      firstRowCells =
        firstRow.getNumCells();

      if (firstRowCells >= 3) {
        rightCellPreview =
          testNormalizeTableText_(
            firstRow
              .getCell(2)
              .getText()
          ).substring(0, 80);
      }
    }
  } catch (error) {}

  try {
    borderWidth =
      table.getBorderWidth();
  } catch (error) {}

  return {
    physicalOrdinal: physicalOrdinal,
    bodyChildIndex: bodyChildIndex,
    rows: table.getNumRows(),
    firstRowCells: firstRowCells,
    borderWidth: borderWidth,
    equationLayout: isEquationLayout,
    countedAsTable: !isEquationLayout,
    rightCellPreview: rightCellPreview
  };
}


/**
 * Normaliza espacios especiales de Google Docs.
 */
function testNormalizeTableText_(value) {
  return String(value || '')
    .replace(/\u00A0/g, ' ')
    .trim();
}
