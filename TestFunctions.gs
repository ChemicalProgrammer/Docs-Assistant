/**
 * TEST LAB
 *
 * Production verification for Table caption counting. Place the cursor in a
 * real body-level table. The test does not modify the document.
 */
function runCurrentTest() {
  const started = Date.now();
  const body = getActiveBody_();
  const selectedTable = getActiveTable_();

  if (!selectedTable) {
    throw new Error('Place the cursor inside a body-level table.');
  }

  let selectedTableIndex;
  try {
    selectedTableIndex = body.getChildIndex(selectedTable);
  } catch (error) {
    throw new Error('The selected table must be a body-level table.');
  }

  let physicalTables = 0;
  let equationLayoutTablesIgnored = 0;
  let countableTables = 0;
  let rawOrdinal = 0;
  let correctedOrdinal = 0;

  for (let childIndex = 0; childIndex < body.getNumChildren(); childIndex++) {
    const child = body.getChild(childIndex);
    if (child.getType() !== DocumentApp.ElementType.TABLE) continue;

    physicalTables++;
    const isCountable = isCountableCaptionTable_(child);

    if (isCountable) countableTables++;
    else equationLayoutTablesIgnored++;

    if (childIndex <= selectedTableIndex) {
      rawOrdinal++;
      if (isCountable) correctedOrdinal++;
    }
  }

  const selectedIsEquationLayout = isEquationLayoutTable_(selectedTable);

  return {
    ok: !selectedIsEquationLayout,
    testId: 'PRODUCTION-GHOST-TABLE-COUNTING',
    physicalTables: physicalTables,
    equationLayoutTablesIgnored: equationLayoutTablesIgnored,
    countableTables: countableTables,
    selectedTableIsEquationLayout: selectedIsEquationLayout,
    selectedTableRawOrdinal: rawOrdinal,
    selectedTableCorrectedOrdinal: correctedOrdinal,
    numbersDifferBy: rawOrdinal - correctedOrdinal,
    usesAdvancedDocsApi: false,
    modifiesDocument: false,
    elapsedMs: Date.now() - started
  };
}
