/**
 * TEST LAB
 *
 * Prueba actual:
 * mide el formato de tabla existente.
 */
function runCurrentTest() {
  const started = Date.now();
  const table = getActiveTable_();

  if (!table) {
    throw new Error(
      'Place the cursor inside the table you want to test.'
    );
  }

  const rows = table.getNumRows();
  let cells = 0;
  let paragraphs = 0;
  let listItems = 0;

  for (let rowIndex = 0; rowIndex < rows; rowIndex++) {
    const row = table.getRow(rowIndex);
    cells += row.getNumCells();

    for (
      let cellIndex = 0;
      cellIndex < row.getNumCells();
      cellIndex++
    ) {
      const cell = row.getCell(cellIndex);

      for (
        let childIndex = 0;
        childIndex < cell.getNumChildren();
        childIndex++
      ) {
        const type = cell.getChild(childIndex).getType();

        if (
          type === DocumentApp.ElementType.PARAGRAPH
        ) {
          paragraphs++;
        } else if (
          type === DocumentApp.ElementType.LIST_ITEM
        ) {
          listItems++;
        }
      }
    }
  }

  const result = formatSelectedTable();
  const elapsedMs = Date.now() - started;

  return {
    ok: true,
    testId: 'TABLE-BASELINE-DOCUMENTAPP',
    message:
      'Table formatted in ' +
      elapsedMs +
      ' ms',
    usesAdvancedDocsApi: false,
    rows: rows,
    cells: cells,
    paragraphs: paragraphs,
    listItems: listItems,
    elapsedMs: elapsedMs,
    formatterResult: result
  };
}
