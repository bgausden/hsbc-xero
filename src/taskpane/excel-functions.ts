// eslint-disable-next-line node/no-missing-import
import { TRANSACTION_DATE } from "./taskpane";

/* global Excel */

export function isRange(range: Excel.Range | Error): range is Excel.Range {
  return (range as Excel.Range).address !== undefined;
}

export async function stripTabComma(context: Excel.RequestContext, sheet: Excel.Worksheet) {
  // eslint-disable-next-line no-unused-vars
  let numReplacements = 0;

  try {
    let foundAreas = sheet.findAllOrNullObject(`\t`, { completeMatch: false, matchCase: false }).areas;
    foundAreas.load("items");
    await context.sync();
    let foundRanges = foundAreas.items;
    if (foundRanges) {
      foundRanges.forEach(async (range) => {
        range.load("values");
        await context.sync();
        range.values = range.values.map(row => row.map(value => (value as string).replace(`\t`, ``)));
      });
      numReplacements = foundRanges.length;
    }
  } catch (err) {
    console.error(err);
  }
  return context
    .sync()
    .then(
      () => numReplacements,
      err => {
        new Error(err);
      }
    )
    .catch(() => 0);
}

export async function deleteExtraneousWhitespace(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  usedRange: Excel.Range
): Promise<void | Error> {
  // remove consecutive whitespace in cell values, trim() cell values potentially resulting in cell value = ""
  // eslint-disable-next-line no-unused-vars
  let numReplacements = 0;
  let newValues: any[][] = [];
  try {
    usedRange.load("values");
    await context.sync();
    const values = usedRange.values;
    for (let row = 0; row < usedRange.rowCount; row++) {
      for (let col = 0; col < usedRange.columnCount; col++) {
        let value = values[row][col];
        if (typeof value === "string") {
          value = value.replace(/\s+/, " ").trim();
        }
        newValues[row][col] = value;
      }
    }
    usedRange.values = newValues;

    /* let foundAreas = sheet.findAllOrNullObject(`\t`, { completeMatch: false, matchCase: false }).areas
    foundAreas.load("items")
    await context.sync()
    let foundRanges = foundAreas.items
    if (foundRanges) {
      foundRanges.forEach(async range => {
        range.load("values")
        await context.sync()
        range.values = range.values.map((row => row.map((value) => (value as string).replace(`\t`, ``))))
      })
      numReplacements = foundRanges.length
    } */
  } catch (err) {
    console.error(`deleteExtraneousWhitespace(): ${err}`);
  }
  return context
    .sync<void>()
    .then(
      () => { },
      err => {
        return new Error(err);
      }
    )
    .catch(err => {
      return new Error(err);
    });
}

export async function deleteExtraneousRows(context: Excel.RequestContext, sheet: Excel.Worksheet, usedRange: Excel.Range) {
  let numDeletedRows = 0;
  let numRowsProcessed = 0;
  const headerCol = await findHeaderCol(context, sheet, usedRange, TRANSACTION_DATE);
  if (headerCol instanceof Error) {
    throw Error;
  }
  usedRange.load(["rowCount", "rowIndex"]);
  await context.sync();
  // const headerCol = header.columnIndex
  let rowCount = usedRange.rowCount;
  // usedRange may not start at A1 so start at the first row in usedRange via Excel.Range.rowIndex
  for (let row = usedRange.rowIndex; numRowsProcessed < rowCount; numRowsProcessed++) {
    const cell = sheet.getCell(row, headerCol);
    cell.load(["values"]);
    await context.sync();
    if (cell.values[0][0] === "") {
      let rowToDelete = cell.getEntireRow();
      rowToDelete.delete(Excel.DeleteShiftDirection.up);
      numDeletedRows += 1;
      // dont move the cursor. we are now on a new line and need to process this line
    } else {
      // move the cursor down so we process the next line
      row += 1;
    }
  }
  return context
    .sync()
    .then(
      () => numDeletedRows,
      err => {
        throw new Error(err);
      }
    )
    .catch(err => {
      console.log(`Couldn't delete: ${err}`);
    });
}

export async function concatConsecutiveColValues(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  usedRange: Excel.Range,
  firstColHeaderText: string
): Promise<number | Error> {
  try {
    let numMergedCells = 0;
    const headerRange = await findHeader(context, sheet, usedRange, firstColHeaderText);
    if (headerRange instanceof Excel.Range) {
      headerRange.load(["columnIndex", "address"]);
      usedRange.load(["rowCount", "rowIndex"]);
      await context.sync();
      const headerColIndex = headerRange.columnIndex;
      for (let row = usedRange.rowIndex; row < usedRange.rowCount; row++) {
        let cell = sheet.getCell(row, headerColIndex);
        cell.load(["values", "address"]);
        let adjacentCell = sheet.getCell(row, headerColIndex + 1);
        adjacentCell.load("values");
        // TODO split into two loops if possible avoid context.sync() inside loop
        await context.sync();
        // don't concat headers. strip out all the excess whitespace (file is full of tab characters)
        if (cell.address !== headerRange.address) {
          cell.values = [[`${cell.values}${adjacentCell.values}`.replace(/\s+/g, " ").trim()]];
        }
        numMergedCells += 1;
      }
      return context.sync<number>().then(
        () => numMergedCells,
        err => {
          console.log(`context.sync() failed: ${err}`);
          return new Error(err);
        }
      );
    } else {
      return context.sync<number>().then(
        () => {
          console.error(`Unable to locate ${firstColHeaderText} position in header row`);
          return 0;
        },
        err => {
          console.error(`context.sync() failed: ${err}`);
          return new Error(err);
        }
      );
    }
  } catch (err) {
    console.log(`concatConsecutiveColValues(): ${err}`);
    throw err;
  }
}

export async function findHeader(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  usedRange: Excel.Range,
  headerText: string
): Promise<Excel.Range | Error> {
  try {
    const foundAreas = sheet.findAll(headerText, {
      completeMatch: true,
      matchCase: false // findAll will not match case
    });
    foundAreas.load(["areas", "address"]);
    await context.sync();
    const rangeCollection = foundAreas.areas;
    rangeCollection.load("items");
    await context.sync();
    if (rangeCollection.items) {
      const foundRanges = rangeCollection.items;
      const range = foundRanges[0];
      return context.sync<Excel.Range>().then<Excel.Range, Error>(
        () => {
          return range;
        },
        err => {
          return new Error(err);
        }
      );
    } else {
      return context.sync<Error>().then(
        () => {
          return new Error(`Unable to locate the header ${headerText}`);
        },
        err => {
          console.error(`findHeader: context.sync() failed. Error is ${err}`);
          return new Error(err);
        }
      );
    }
  } catch (err) {
    console.log(`findHeader(): ${err}`);
    return new Error(err);
  }
}

export async function findHeaderCol(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  usedRange: Excel.Range,
  headerText: string
): Promise<number | Error> {
  const headerRange = await findHeader(context, sheet, usedRange, headerText);
  if (headerRange instanceof Excel.Range) {
    headerRange.load("columnIndex");
    return context.sync<number>().then(
      () => headerRange.columnIndex,
      err => {
        return new Error(err);
      }
    );
  } else {
    return context.sync<Error>().finally(() => {
      throw new Error(`Unable to locate header column for ${headerText}`);
    });
  }
}

export async function stripColumnText(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  usedRange: Excel.Range,
  column: string,
  stripText: string
): Promise<number | Error> {
  try {
    let numReplacements = 0;
    const range = await findHeader(context, sheet, usedRange, column);
    if (range instanceof Excel.Range) {
      const col = range.getEntireColumn();
      const res = col.replaceAll(stripText, "", { completeMatch: false, matchCase: true } as Excel.ReplaceCriteria);
      await context.sync();
      numReplacements = res.value;
      return context.sync<number>().then(
        () => numReplacements,
        err => {
          console.error(`stripColumnText: context.sync() failed. Error is ${err}`);
          return new Error(err);
        }
      );
    } else {
      return context.sync<Error>().then(
        () => {
          console.log(`stripColumnText: Unable to locate header "${column}"`);
          return range as Error;
        },
        err => {
          console.error(`stripColumnText: context.sync() failed. Error is ${err}`);
          return new Error(err);
        }
      );
    }
  } catch (err) {
    console.error(`stripColumnText(): ${err}`);
    throw err;
  }
}

export async function fixAmounts(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  usedRange: Excel.Range,
  amountCol: string
): Promise<void | Error> {
  try {
    const colsToSearch = 6;
    let header = await findHeader(context, sheet, usedRange, amountCol);
    if (header instanceof Excel.Range) {
      let amountsRange = header.getOffsetRange(1, 0).getResizedRange(usedRange.rowCount - 1, 0);
      amountsRange.format.fill.color = "pink";
      amountsRange.load(["values", "address", "rowIndex", "columnIndex"]);
      // extend range to include 6 additional columns to the right
      let searchRange = header.getOffsetRange(1, 1).getResizedRange(usedRange.rowCount - 1, colsToSearch);
      searchRange.format.fill.color = "lightBlue";
      //range = range.set({} as Excel.Interfaces.RangeUpdateData)
      searchRange.load(["values", "rowIndex", "columnIndex"]);
      await context.sync();
      const searchValues = searchRange.values;
      const amounts = amountsRange.values;
      const amountsColIndex = amountsRange.columnIndex;
      for (let row = 0; row < searchValues.length; row++) {
        const searchRowValues = searchValues[row];
        for (let col = 0; col < searchRowValues.length; col++) {
          const searchValue = searchRowValues[col];
          if (typeof searchValue === "number") {
            console.log(`row = ${row}`, `amount = "${amounts[row][0]}"`, `value = "${searchValue}"`);
            // TODO match whitespace, unassigned or null string
            if ((amounts[row][0] as string).replace(/\s+/g, "")[0] === "") {
              amounts[row][0] = searchValue;
            } else {
              // amounts has a number in it but so does one of the further out cells.
              // mark both in red
              sheet.getCell(amountsRange.rowIndex + row, amountsColIndex).format.fill.color = "#CC3300";
              sheet.getCell(searchRange.rowIndex + row, searchRange.columnIndex + col).format.fill.color = "#CC3300";
            }
          }
        }
      }
      return context.sync();
    } else {
      return context.sync<Error>().then(
        () => {
          console.log(`fixAmounts(): Unable to locate header "${amountCol}"`);
          return header as Error;
        },
        err => {
          console.error(`fixAmounts(): context.sync() failed. Error is ${err}`);
          return new Error(err);
        }
      );
    }
  } catch (err) {
    console.error(`fixAmounts(): ${err}`);
    throw err;
  }
}
