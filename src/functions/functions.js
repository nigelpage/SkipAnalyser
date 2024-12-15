/**
 * Add two numbers
 * @customfunction
 * @param {string} range range containing skip names in worksheet
 * @param {string} prefix worksheet prefix name (e.g. 'Round ' searches the range in all worksheets that have a name beginning with 'Round ')
 * @returns {string[][]} list of discovered skip names in 'spill down' format.
 */
export function skipnames(range, prefix) {
  const skips = [];
  const validSheets = [];

  function insertIfSheetHasPrefix(strarray, name, prefix) {
    let namePrefix = name.substring(0, prefix.length);
    if (namePrefix == prefix)
      return insertIfNotFound(strarray, name);

    return false;
  }

  function insertIfNotFound(strarray, str) {
    if (!strarray.includes(str)) {
      strarray.push(str);
      return true;
    }

    return false;
  }

  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;

    sheets.load("items/name");
    await context.sync();

    sheets.items.forEach((sheet) => {
      insertIfSheetHasPrefix(validSheets, sheet.name, prefix);
  
      let sheetSkips = sheet.getRange(range);

      sheetSkips.load("values");
      await context.sync();

      sheetSkips.forEach((sheetSkip) => {
        insertIfNotFound(skips, sheetSkip)
        });
    });

  return skips;
  });
};