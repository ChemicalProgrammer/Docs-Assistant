/**
 * TEST LAB
 *
 * Prueba:
 * formato rápido de tabla usando setAttributes().
 *
 * No utiliza Advanced Google Docs API.
 */
function runCurrentTest() {
  const table = getActiveTable_();

  if (!table) {
    throw new Error(
      'Place the cursor inside the table you want to test.'
    );
  }

  const started = Date.now();

  const result =
    testFastTableFormatting_(table);

  const elapsedMs =
    Date.now() - started;

  return {
    ok: true,
    testId: 'TABLE-FAST-ATTRIBUTES',
    message:
      'Optimized table formatted in ' +
      elapsedMs +
      ' ms',
    usesAdvancedDocsApi: false,
    rows: result.rows,
    cells: result.cells,
    paragraphs: result.paragraphs,
    listItems: result.listItems,
    elapsedMs: elapsedMs
  };
}


/**
 * Candidato optimizado.
 *
 * Agrupa los atributos de celda y párrafo
 * para reducir llamadas a DocumentApp.
 */
function testFastTableFormatting_(table) {
  const PT_PER_IN = 72;
  const attribute =
    DocumentApp.Attribute;

  /*
   * Atributos comunes de las celdas.
   */
  const cellAttributes = {};

  cellAttributes[
    attribute.VERTICAL_ALIGNMENT
  ] = DocumentApp.VerticalAlignment.CENTER;

  cellAttributes[
    attribute.PADDING_TOP
  ] = 0.028 * PT_PER_IN;

  cellAttributes[
    attribute.PADDING_BOTTOM
  ] = 0.028 * PT_PER_IN;

  cellAttributes[
    attribute.PADDING_LEFT
  ] = 0.028 * PT_PER_IN;

  cellAttributes[
    attribute.PADDING_RIGHT
  ] = 0.028 * PT_PER_IN;


  /*
   * Atributos del encabezado.
   */
  const headerAttributes =
    buildTestTableParagraphAttributes_(
      true,
      PT_PER_IN
    );

  /*
   * Atributos del contenido.
   */
  const bodyAttributes =
    buildTestTableParagraphAttributes_(
      false,
      PT_PER_IN
    );


  table.setBorderColor('#000000');
  table.setBorderWidth(1);

  let cells = 0;
  let paragraphs = 0;
  let listItems = 0;

  const rows = table.getNumRows();

  for (
    let rowIndex = 0;
    rowIndex < rows;
    rowIndex++
  ) {
    const row = table.getRow(rowIndex);

    row.setMinimumHeight(
      0.49 * PT_PER_IN
    );

    const isHeader =
      rowIndex === 0;

    const paragraphAttributes =
      isHeader
        ? headerAttributes
        : bodyAttributes;

    const cellCount =
      row.getNumCells();

    cells += cellCount;

    for (
      let cellIndex = 0;
      cellIndex < cellCount;
      cellIndex++
    ) {
      const cell =
        row.getCell(cellIndex);

      /*
       * Padding y alineación vertical
       * en una sola operación.
       */
      cell.setAttributes(
        cellAttributes
      );

      /*
       * Elimina únicamente el color
       * directo de fondo.
       */
      cell.setBackgroundColor(null);

      for (
        let childIndex = 0;
        childIndex < cell.getNumChildren();
        childIndex++
      ) {
        const child =
          cell.getChild(childIndex);

        const type =
          child.getType();

        if (
          type ===
          DocumentApp.ElementType.PARAGRAPH
        ) {
          child
            .asParagraph()
            .setAttributes(
              paragraphAttributes
            );

          paragraphs++;
        } else if (
          type ===
          DocumentApp.ElementType.LIST_ITEM
        ) {
          child
            .asListItem()
            .setAttributes(
              paragraphAttributes
            );

          listItems++;
        }
      }
    }
  }

  return {
    rows: rows,
    cells: cells,
    paragraphs: paragraphs,
    listItems: listItems
  };
}


/**
 * Construye todos los atributos de un párrafo
 * de tabla para aplicarlos en una sola llamada.
 */
function buildTestTableParagraphAttributes_(
  isHeader,
  pointsPerInch
) {
  const attribute =
    DocumentApp.Attribute;

  const attributes = {};

  attributes[
    attribute.HORIZONTAL_ALIGNMENT
  ] = isHeader
    ? DocumentApp.HorizontalAlignment.CENTER
    : DocumentApp.HorizontalAlignment.LEFT;

  attributes[
    attribute.LINE_SPACING
  ] = 1;

  attributes[
    attribute.SPACING_BEFORE
  ] = 0;

  attributes[
    attribute.SPACING_AFTER
  ] = 0;

  attributes[
    attribute.INDENT_START
  ] = 0.05 * pointsPerInch;

  attributes[
    attribute.INDENT_END
  ] = 0;

  attributes[
    attribute.INDENT_FIRST_LINE
  ] = 0.05 * pointsPerInch;

  attributes[
    attribute.FONT_FAMILY
  ] = 'Arial';

  attributes[
    attribute.FONT_SIZE
  ] = 9;

  attributes[
    attribute.BOLD
  ] = Boolean(isHeader);

  return attributes;
}
