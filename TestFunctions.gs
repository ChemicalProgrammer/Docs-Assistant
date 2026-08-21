/**
 * TEST LAB
 *
 * Current verification: runs the production table formatter and reports
 * its result. Replace only this function for the next isolated experiment.
 */
function runCurrentTest() {
  const result = formatSelectedTable();

  return {
    ok: true,
    testId: 'PRODUCTION-TABLE-FORMATTER',
    message:
      'Production table formatted in ' +
      Number(result.elapsedMs || 0) +
      ' ms',
    usesAdvancedDocsApi: false,
    rows: result.rows,
    cells: result.cells,
    paragraphs: result.paragraphs,
    listItems: result.listItems,
    listIndentLeftInches: 0,
    listHangingInches: 0.10,
    elapsedMs: result.elapsedMs
  };
}
