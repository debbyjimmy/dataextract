function updateMonthlyMetricsFromExternalSheet() {
  // Step 1: Config
  const sourceSheetName = '..';
  const targetSheetName = '..';

  const month = new Date().toLocaleString('default', { month: 'short' });
  const monthMap = {
    Jan: 'Jan', Feb: 'Feb', Mar: 'Mar', Apr: 'April', May: 'May', Jun: 'June',
    Jul: 'July', Aug: 'Aug', Sep: 'Sep', Oct: 'Oct', Nov: 'Nov', Dec: 'Dec'
  };
  const targetMonth = monthMap[month];

  // Step 2: Load source and target sheets from the same spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const targetSheet = ss.getSheetByName(targetSheetName);

  const sourceData = sourceSheet.getDataRange().getValues();
  const targetData = targetSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[4]; // Row 5 (index 4) is where headers are located
  const targetHeaders = targetData[0]; // Row 1 (index 0) is the target headers

  // Step 3: Get column indexes
  const colIndex = (headers, name) => headers.indexOf(name);

  const emailIdxSource = colIndex(sourceHeaders, "Email");
  const effIdx = colIndex(sourceHeaders, "Efficiency Score");
  const accIdx = colIndex(sourceHeaders, "Accuracy Score");

  const emailIdxTarget = colIndex(targetHeaders, "Email");
  const eCol = colIndex(targetHeaders, `e_${targetMonth}`);
  const qaCol = colIndex(targetHeaders, `qa_${targetMonth}`);

  if ([eCol, qaCol].includes(-1)) {
    throw new Error(`Target sheet is missing one or more columns for ${targetMonth}`);
  }

  // Step 4: Map target emails to row indexes
  const emailToTargetRow = {};
  for (let i = 1; i < targetData.length; i++) {
    const email = targetData[i][emailIdxTarget];
    if (email) emailToTargetRow[email.trim().toLowerCase()] = i;
  }

  // Step 5: Transfer data (Efficiency and Accuracy only)
  for (let i = 5; i < sourceData.length; i++) { // Start at row 5 (index 4) for actual data
    const email = sourceData[i][emailIdxSource];
    if (!email) continue;

    const normalizedEmail = email.trim().toLowerCase();
    const targetRow = emailToTargetRow[normalizedEmail];
    if (targetRow === undefined) continue;

    targetSheet.getRange(targetRow + 1, eCol + 1).setValue(sourceData[i][effIdx]);
    targetSheet.getRange(targetRow + 1, qaCol + 1).setValue(sourceData[i][accIdx]);
  }

  Logger.log(`Updated metrics (Efficiency and Accuracy only) for month: ${targetMonth}`);
}
