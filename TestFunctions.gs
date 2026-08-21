/**
 * TEST LAB
 *
 * Prueba actual:
 * formato rápido de tabla usando DocumentApp.
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
    listIndentLeftInches: 0,
    listHangingInches: 0.05,
    elapsedMs: elapsedMs
  };
}


/**
 * Formatea la tabla agrupando atributos
 * para reducir llamadas a DocumentApp.
 */
function testFastTableFormatting_(table) {
  const PT_PER_IN = 72;

  const attribute =
    DocumentApp.Attribute;

  /*
   * Atributos comunes de todas las celdas.
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
   * Párrafos normales del encabezado.
   */
  const headerParagraphAttributes =
    buildTestTableParagraphAttributes_(
      true,
      false,
      PT_PER_IN
    );


  /*
   * Párrafos normales del contenido.
   */
  const bodyParagraphAttributes =
    buildTestTableParagraphAttributes_(
      false,
      false,
      PT_PER_IN
    );


  /*
   * Listas dentro del encabezado.
   */
  const headerListAttributes =
    buildTestTableParagraphAttributes_(
      true,
      true,
      PT_PER_IN
    );


  /*
   * Listas dentro del contenido.
   */
  const bodyListAttributes =
    buildTestTableParagraphAttributes_(
      false,
      true,
      PT_PER_IN
    );


  /*
   * Formato general de la tabla.
   */
  table.setBorderColor('#000000');
  table.setBorderWidth(1);


  let cells = 0;
  let paragraphs = 0;
  let listItems = 0;

  const rows =
    table.getNumRows();


  for (
    let rowIndex = 0;
    rowIndex < rows;
    rowIndex++
  ) {
    const row =
      table.getRow(rowIndex);

    row.setMinimumHeight(
      0.49 * PT_PER_IN
    );

    const isHeader =
      rowIndex === 0;

    const paragraphAttributes =
      isHeader
        ? headerParagraphAttributes
        : bodyParagraphAttributes;

    const listAttributes =
      isHeader
        ? headerListAttributes
        : bodyListAttributes;

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
       * Elimina el color directo
       * de fondo de la celda.
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
              listAttributes
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
 * Construye los atributos para párrafos
 * y listas dentro de la tabla.
 *
 * Párrafo normal:
 *   Left = 0.05 in
 *   Special indent = None
 *
 * Lista:
 *   Left = 0
 *   Hanging = 0.05 in
 */
function buildTestTableParagraphAttributes_(
  isHeader,
  isListItem,
  pointsPerInch
) {
  const attribute =
    DocumentApp.Attribute;

  const attributes = {};


  /*
   * Alineación horizontal.
   */
  attributes[
    attribute.HORIZONTAL_ALIGNMENT
  ] = isHeader
    ? DocumentApp.HorizontalAlignment.CENTER
    : DocumentApp.HorizontalAlignment.LEFT;


  /*
   * Espaciado.
   */
  attributes[
    attribute.LINE_SPACING
  ] = 1;

  attributes[
    attribute.SPACING_BEFORE
  ] = 0;

  attributes[
    attribute.SPACING_AFTER
  ] = 0;


  /*
   * Indentación.
   */
  if (isListItem) {
    /*
     * Left = 0
     * Hanging = 0.05 in
     *
     * Primera línea/viñeta inicia en 0.
     * El texto y las líneas siguientes
     * inician en 0.05 in.
     */
    attributes[
      attribute.INDENT_FIRST_LINE
    ] = 0;

    attributes[
      attribute.INDENT_START
    ] = 0.05 * pointsPerInch;
  } else {
    /*
     * Párrafo normal:
     * Left = 0.05 in
     * Special indent = None
     */
    const paragraphIndent =
      0.05 * pointsPerInch;

    attributes[
      attribute.INDENT_FIRST_LINE
    ] = paragraphIndent;

    attributes[
      attribute.INDENT_START
    ] = paragraphIndent;
  }

  attributes[
    attribute.INDENT_END
  ] = 0;


  /*
   * Formato de texto.
   */
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
