import * as Excel from "@microsoft/office-js";

/**
 * Discover skips in worksheets
 * @customfunction
 * @param {string} range range containing skip names in worksheet
 * @param {string} prefix worksheet prefix name (e.g. 'Round ' searches the range in all worksheets that have a name beginning with 'Round ')
 * @returns {string[]} list of discovered skip names in 'spill down' format.
 */
export async function skipnames(range, prefix) {
  const skips = [];
  const validSheets = [];

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
      let namePrefix = sheet.name.substring(0, prefix.length);

      if (namePrefix == prefix) {
        insertIfNotFound(validSheets, sheet.name);

        let sheetSkips = sheet.getRange(range);

        sheetSkips.load("values");

        sheetSkips.forEach((sheetSkip) => {
          insertIfNotFound(skips, sheetSkip);
        });
      }
    });
    return skips;
  });
}
