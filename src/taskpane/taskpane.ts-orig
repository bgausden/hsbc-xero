/* eslint-disable no-unused-vars */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

import {CastingContext} from "csv-parse"
import parse from "csv-parse/lib/sync"
import * as React from 'react';
import {
  MessageBarButton,
  Link,
  Stack,
  StackItem,
  MessageBar,
  MessageBarType,
  ChoiceGroup,
  IStackProps,
} from 'office-ui-fabric-react';
import ReactDOM from "react-dom";

// eslint-disable-next-line no-redeclare
/* global console, document, Excel, Office, FileReader */

const mergeCol = 1 // zero indexed
const POST_DATE = "Post Date"
const TRANSACTION_DATE = "Transaction Date"
const TRANSACTION_DATE_INDEX = 1 // zero indexed
const DESCRIPTION = "Description"
const FOREIGN_CURRENCY_AMOUNT = "Foreign Currency Amount"
const AMOUNT_HKD = "Amount(HKD)"
const SALES = "SALES: "

function is(value: any) {
  return {
    a: function (check: any) {
      if (check.prototype) check = check.prototype.constructor.name
      const type: string = Object.prototype.toString
        .call(value)
        .slice(8, -1)
        .toLowerCase()
      return value != null && type === check.toLowerCase()
    },
  }
}

function isRange(range: Excel.Range | Error): range is Excel.Range {
  return (range as Excel.Range).address !== undefined
}

async function stripTabComma(context: Excel.RequestContext, sheet: Excel.Worksheet, usedRange: Excel.Range) {

  // eslint-disable-next-line no-unused-vars
  let numReplacements = 0

  try {
    let foundAreas = sheet.findAllOrNullObject(`\t`, { completeMatch: false, matchCase: false }).areas
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
    }
  } catch (err) {
    console.error(err)
  }
  return context.sync().then(() => numReplacements, (err) => { new Error(err) })
    .catch(() => 0)
}

async function deleteExtraneousWhitespace(context: Excel.RequestContext, sheet: Excel.Worksheet, usedRange: Excel.Range): Promise<void | Error> {
  // remove consecutive whitespace in cell values, trim() cell values potentially resulting in cell value = ""
  // eslint-disable-next-line no-unused-vars
  let numReplacements = 0
  let newValues: any[][] = []
  try {
    usedRange.load("values")
    await context.sync()
    const values = usedRange.values
    for (let row = 0; row < usedRange.rowCount; row++) {
      for (let col = 0; col < usedRange.columnCount; col++) {
        let value = values[row][col]
        if (typeof value === "string") {
          value = value.replace(/\s+/, " ").trim()
        }
        newValues[row][col] = value

      }
    }
    usedRange.values = newValues

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
    console.error(`deleteExtraneousWhitespace(): ${err}`)
  }
  return context.sync<void>().then(() => { }, (err) => { return new Error(err) })
    .catch((err) => { return new Error(err) })
}

async function deleteExtraneousRows(context: Excel.RequestContext, sheet: Excel.Worksheet, usedRange: Excel.Range) {
  let numDeletedRows = 0
  let numRowsProcessed = 0
  const headerCol = await findHeaderCol(context, sheet, usedRange, TRANSACTION_DATE)
  if (headerCol instanceof Error) { throw Error }
  usedRange.load(["rowCount", "rowIndex"]);
  await context.sync()
  // const headerCol = header.columnIndex
  let rowCount = usedRange.rowCount
  // usedRange may not start at A1 so start at the first row in usedRange via Excel.Range.rowIndex
  for (let row = usedRange.rowIndex; numRowsProcessed < rowCount; numRowsProcessed++) {
    const cell = sheet.getCell(row, headerCol)
    cell.load(["values"])
    await context.sync()
    if (cell.values[0][0] === "") {
      let rowToDelete = cell.getEntireRow()
      rowToDelete.delete(Excel.DeleteShiftDirection.up)
      numDeletedRows += 1
      // dont move the cursor. we are now on a new line and need to process this line
    } else {
      // move the cursor down so we process the next line
      row += 1
    }
  }
  return context.sync()
    .then(() => numDeletedRows, (err) => { throw new Error(err) })
    .catch((err) => { console.log(`Couldn't delete: ${err}`) })
}

async function concatConsecutiveColValues(context: Excel.RequestContext, sheet: Excel.Worksheet, usedRange: Excel.Range, firstColHeaderText: string): Promise<number | Error> {
  try {
    let numMergedCells = 0
    const headerRange = await findHeader(context, sheet, usedRange, firstColHeaderText)
    if (headerRange instanceof Excel.Range) {
      headerRange.load(["columnIndex", "address"])
      usedRange.load(["rowCount", "rowIndex"])
      await context.sync()
      const headerColIndex = headerRange.columnIndex
      for (let row = usedRange.rowIndex; row < usedRange.rowCount; row++) {
        let cell = sheet.getCell(row, headerColIndex)
        cell.load(["values", "address"])
        let adjacentCell = sheet.getCell(row, headerColIndex + 1)
        adjacentCell.load("values")
        // TODO split into two loops if possible avoid context.sync() inside loop
        await context.sync()
        // don't concat headers. strip out all the excess whitespace (file is full of tab characters)
        if (cell.address !== headerRange.address) { cell.values = ([[`${cell.values}${adjacentCell.values}`.replace(/\s+/g, " ").trim()]]) }
        numMergedCells += 1
      }
      return context.sync<number>()
        .then(() => numMergedCells, (err) => {
          console.log(`context.sync() failed: ${err}`)
          return new Error(err)
        })
    } else {
      return context.sync<number>()
        .then(() => {
          console.error(`Unable to locate ${firstColHeaderText} position in header row`)
          return 0
        }, err => {
          console.error(`context.sync() failed: ${err}`)
          return new Error(err)
        })
    }
  } catch (err) {
    console.log(`concatConsecutiveColValues(): ${err}`)
    throw err
  }
}

async function findHeader(context: Excel.RequestContext, sheet: Excel.Worksheet, usedRange: Excel.Range, headerText: string): Promise<Excel.Range | Error> {
  try {
    const foundAreas = sheet.findAll(headerText, {
      completeMatch: true, // findAll will match the whole cell value
      matchCase: false // findAll will not match case
    })
    foundAreas.load(["areas", "address"])
    await context.sync()
    const rangeCollection = foundAreas.areas
    rangeCollection.load("items")
    await context.sync()
    if (rangeCollection.items) {
      const foundRanges = rangeCollection.items
      const range = foundRanges[0]
      return context.sync<Excel.Range>()
        .then<Excel.Range, Error>(() => {
          return range
        }, err => {
          return new Error(err)
        })
    } else {
      return context.sync<Error>()
        .then(() => {
          return new Error(`Unable to locate the header ${headerText}`)
        }, err => {
          console.error(`findHeader: context.sync() failed. Error is ${err}`)
          return new Error(err)
        })
    }
  } catch (err) {
    console.log(`findHeader(): ${err}`)
    return new Error(err)
  }
}

async function findHeaderCol(context: Excel.RequestContext, sheet: Excel.Worksheet, usedRange: Excel.Range, headerText: string): Promise<number | Error> {
  const headerRange = await findHeader(context, sheet, usedRange, headerText)
  if (headerRange instanceof Excel.Range) {
    headerRange.load("columnIndex")
    return context.sync<number>()
      .then(() => headerRange.columnIndex, (err) => {
        return new Error(err)
      })
  } else {
    return context.sync<Error>()
      .finally(() => { throw new Error(`Unable to locate header column for ${headerText}`) })
  }
}

async function stripColumnText(context: Excel.RequestContext, sheet: Excel.Worksheet, usedRange: Excel.Range, column: string, stripText: string): Promise<number | Error> {
  try {
    let numReplacements = 0
    const range = await findHeader(context, sheet, usedRange, column)
    if (range instanceof Excel.Range) {
      const col = range.getEntireColumn()
      const res = col.replaceAll(stripText, "", { completeMatch: false, matchCase: true } as Excel.ReplaceCriteria)
      await context.sync()
      numReplacements = res.value
      return context.sync<number>()
        .then(() => numReplacements, err => {
          console.error(`stripColumnText: context.sync() failed. Error is ${err}`)
          return new Error(err)
        })
    } else {
      return context.sync<Error>()
        .then(() => {
          console.log(`stripColumnText: Unable to locate header "${column}"`)
          return range as Error
        }, (err) => {
          console.error(`stripColumnText: context.sync() failed. Error is ${err}`)
          return new Error(err)
        })
    }
  } catch (err) {
    console.error(`stripColumnText(): ${err}`)
    throw err
  }
}

async function fixAmounts(context: Excel.RequestContext, sheet: Excel.Worksheet, usedRange: Excel.Range, amountCol: string): Promise<void | Error> {
  try {
    const colsToSearch = 6
    let header = await findHeader(context, sheet, usedRange, amountCol)
    if (header instanceof Excel.Range) {
      let amountsRange = header.getOffsetRange(1, 0).getResizedRange(usedRange.rowCount - 1, 0)
      amountsRange.format.fill.color = "pink"
      amountsRange.load(["values", "address", "rowIndex", "columnIndex"])
      // extend range to include 6 additional columns to the right
      let searchRange = header.getOffsetRange(1, 1).getResizedRange(usedRange.rowCount - 1, colsToSearch)
      searchRange.format.fill.color = "lightBlue"
      //range = range.set({} as Excel.Interfaces.RangeUpdateData)
      searchRange.load(["values", "rowIndex", "columnIndex"])
      await context.sync()
      const searchValues = searchRange.values
      const amounts = amountsRange.values
      const amountsColIndex = amountsRange.columnIndex
      for (let row = 0; row < searchValues.length; row++) {
        const searchRowValues = searchValues[row];
        for (let col = 0; col < searchRowValues.length; col++) {
          const searchValue = searchRowValues[col];
          if (typeof searchValue === "number") {
            console.log(`row = ${row}`, `amount = "${amounts[row][0]}"`, `value = "${searchValue}"`)
            // TODO match whitespace, unassigned or null string
            if ((amounts[row][0] as string).replace(/\s+/g, "")[0] === "") {
              amounts[row][0] = searchValue
            } else {
              // amounts has a number in it but so does one of the further out cells.
              // mark both in red
              sheet.getCell(amountsRange.rowIndex + row, amountsColIndex).format.fill.color = "#CC3300"
              sheet.getCell(searchRange.rowIndex + row, searchRange.columnIndex + col).format.fill.color = "#CC3300"
            }
          }
        }
      }
      return context.sync()
    }
    else {
      return context.sync<Error>()
        .then(() => {
          console.log(`fixAmounts(): Unable to locate header "${amountCol}"`)
          return header as Error
        }, (err) => {
          console.error(`fixAmounts(): context.sync() failed. Error is ${err}`)
          return new Error(err)
        })
    }
  } catch (err) {
    console.error(`fixAmounts(): ${err}`)
    throw err
  }
}

const onRecord = ({ raw, record }: { raw: string, record: string[] }, context: CastingContext) => {
  if (context.error && context.error.code === 'CSV_INCONSISTENT_RECORD_LENGTH') {
    let stringRaw = raw as string
    let counter = 3 // zero-based index
    let nThIndex = 0;

    if (counter > 0) {
      while (counter--) {
        // Get the index of the next occurence
        nThIndex = String.prototype.indexOf.call(stringRaw, ",", nThIndex + ",".length,)
      }
    }

    stringRaw = stringRaw.substring(0, nThIndex) + stringRaw.substring(nThIndex + ",".length)
    //result = stringRaw.trim().split(",").map(field => field.replace(/\s+/g, " ").trim())
    //result = stringRaw.split(",")
    let result = parse(stringRaw, {
      raw: true,
      trim: true,
      onRecord: onRecord,
      cast: (value, context) => {
        return value.replace(/\s+/g, " ").trim()
      }
    })
    return result[0]
  }

  // delete rows where there is only data in the 0th column (garbage)
  if (record[TRANSACTION_DATE_INDEX].trim() === "") return null

  return [record[0], record[4], "", `${record[2]} ${record[3]}`]
}

const csvOnload = (reader: FileReader, context: Excel.RequestContext) => {
  return async (e: ProgressEvent) => {
    const raw = reader.result as string
    const rawData = raw.slice(raw.indexOf(`\n`) + 1)
    let data: string[][] = parse(rawData, {
      relax_column_count: true,
      trim: true,
      raw: true,
      cast: (value, context) => {
        return value.replace(/\s+/g, " ").trim()
      },
      onRecord: onRecord
    })
    // replace the header
    data[0] = ["Date", "Amount", "Payee", "Description"]
    console.log(data)
    const sheet = context.workbook.worksheets.getActiveWorksheet()
    const usedRange = sheet.getUsedRangeOrNullObject()
    usedRange.load(["address", "cellCount"])
    await context.sync()
    if (usedRange.address) {
      usedRange.format.fill.color = "lightBlue"
      console.log(usedRange.address)
      
    } else {
      sheet.getRange("A1")
    }
    await context.sync()
  }
}

/* function messageBarTest () {


interface IExampleProps {
  resetChoice?: () => void;
}

const horizontalStackProps: IStackProps = {
  horizontal: true,
  tokens: { childrenGap: 16 },
};
const verticalStackProps: IStackProps = {
  styles: { root: { overflow: 'hidden', width: '100%' } },
  tokens: { childrenGap: 20 },
};


const ErrorExample = (p: IExampleProps) => (
  <MessageBar
    messageBarType={MessageBarType.error}
    isMultiline={false}
    onDismiss={p.resetChoice}
    dismissButtonAriaLabel="Close"
  >
    Error MessageBar with single line, with dismiss button.
    <Link href="www.bing.com" target="_blank">
      Visit our website.
    </Link>
  </MessageBar>
)







const MessageBarBasicExample: React.FunctionComponent = () => {
  const [choice, setChoice] = React.useState<string | undefined>(undefined);
  const showAll = choice === 'all';

  const resetChoice = React.useCallback(() => setChoice(undefined), []);

  return (
    <Stack {...horizontalStackProps}>
      <Stack {...verticalStackProps}>
<ErrorExample resetChoice={resetChoice} />
      </Stack>
    </Stack>
  )
}
ReactDOM.render(<MessageBarBasicExample />, document.getElementById('messagebar'))
} */
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";
    document.getElementById("run")!.onclick = run;
    document.getElementById("load")!.onclick = load;
    //messageBarTest()
    
  }
})

export async function load() {
  try {
    await Excel.run(async context => {
      const file = document.getElementById("file") as HTMLInputElement
      const reader = new FileReader()
      let content: [[string]]

      if (file?.files && file.files[0]) {
        reader.onload = csvOnload(reader, context)
        reader.readAsText(file.files[0])
      }
    })
  } catch (error) {
    console.error(`load(): ${error}`);
  }
}

export async function run() {
  try {
    await Excel.run(async context => {
      // get the used (non-empty) cells in the active worksheet

      const sheet = context.workbook.worksheets.getActiveWorksheet()
      const usedRange = sheet.getUsedRangeOrNullObject()
      context.trackedObjects.add(sheet)
      context.trackedObjects.add(usedRange)

      /* // replace tab+comma with tab
      console.log(`Replaced ${await stripTabComma(context, sheet, usedRange)} tab+comma.`)

      // strip extra whitepace
      let res = deleteExtraneousWhitespace(context, sheet, usedRange)
      if (res instanceof Error) {
        console.error(`Couldn't delete extraneous whitespace.`)
        throw res
      }

      // delete rows where Transaction Date is empty
      console.log(`Deleted ${await deleteExtraneousRows(context, sheet, usedRange)} rows.`)

      // delete Transaction Date column
      let header = await findHeader(context, sheet, usedRange, TRANSACTION_DATE)
      if (header instanceof Excel.Range) {
        header.getEntireColumn().delete(Excel.DeleteShiftDirection.left)
        console.log(`Deleted ${TRANSACTION_DATE} column.`)
      } else {
        console.error(`Failed to locate column for header ${TRANSACTION_DATE}: ${header}`)
        throw header
      }
      await context.sync()

      // concat Description and Foreign Currency Amount columns
      try {
        const concatResult = await concatConsecutiveColValues(context, sheet, usedRange, DESCRIPTION)
        if (typeof concatResult === "number") {
          console.log(`Concatenated ${concatResult} pairs of cells.`)
        } else {
          console.log(`concatConsecutiveColValues failed: ${Error.toString()}`)
          throw Error
        }
      } catch (err) {
        console.error(err)
        // eslint-disable-next-line no-undef
        process.exit(1)
      }

      // delete Foreign Currency column
      header = await findHeader(context, sheet, usedRange, FOREIGN_CURRENCY_AMOUNT)
      if (header instanceof Excel.Range) {
        header.getEntireColumn().delete(Excel.DeleteShiftDirection.left)
        await context.sync()
        console.log(`Deleted ${FOREIGN_CURRENCY_AMOUNT} column.`)
      } else {
        console.log(`Unable to delete column ${FOREIGN_CURRENCY_AMOUNT}`)
        throw header as Error
      }

      // strip "SALES:" from descriptions
      const numStripped = await stripColumnText(context, sheet, usedRange, DESCRIPTION, SALES)
      if (typeof numStripped === "number") {
        console.log(`Stripped "${SALES}" from ${numStripped} cells.`)
      } else {
        console.log(`Stripping text "${SALES}" from ${DESCRIPTION} column failed.`)
        throw numStripped as Error
      }

      // check to see if any amounts have been pushed to the right because of commas in the description column screwed things up
      const numAmtFixed = await fixAmounts(context, sheet, usedRange, AMOUNT_HKD)
      if (typeof numAmtFixed === "number") {
        console.log(`Fixed ${numAmtFixed} ${AMOUNT_HKD} cells.`)
      } else {
        console.log(`Fixing ${AMOUNT_HKD} column failed: ${numAmtFixed}`)
        throw numAmtFixed as Error
      }


      // move Amount (HKD) column after Post Date Column */
      console.log("Done")
    });
  } catch (error) {
    console.error(`run(): ${error}`);
  }
}









