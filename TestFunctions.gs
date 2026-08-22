/**
 * TEST LAB
 *
 * Runs the production formatter used by “Format selected styles”.
 * The selected document section is modified.
 */
function runCurrentTest() {
  const result = formatSelectedNamedStyles();
  result.testId = 'PRODUCTION-DETERMINISTIC-SECTION-FORMAT-V4';
  return result;
}
