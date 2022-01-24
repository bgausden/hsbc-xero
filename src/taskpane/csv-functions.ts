import { CastingContext } from "csv-parse"
import parse from "csv-parse/lib/sync"
import debug from "debug"
// eslint-disable-next-line node/no-missing-import
import { populateWorksheet } from "./excel-functions"
// eslint-disable-next-line node/no-missing-import
import { SALES, TRANSACTION_DATE_INDEX } from "./taskpane"

/* global Excel */

const PAYMENT = /PAYMENT - THANK YOU.*$/
const RETURN = /RETURN:.*$/
const DESCRIPTION_INDEX = 2 // column index for Description
const AMOUNT_INDEX = 4 // column index for Amount(HKD)

const csvFunctionRaw = debug('csv-functions:raw')
if (process.env.csvFunctionsRawLogging === 'true') {
  debug.enable('csv-functions:raw')
}
const csvFunctionOut = debug('csv-functions:out')
/* if (process.env.csvFunctionsOutLogging === 'true') {
  debug.enable('csv-functions:out')
}  */
const csvFunctionErr = debug('csv-functions:err')
/* if (process.env.csvFunctionsErrLogging === 'true') {
  debug.enable('csv-functions:err')
} */
const collapseWhitespace = (inputString: string): string => { return inputString.replace(/\s+/g, ' ') }

const cast = (value: string) => {
  return value.replace(/\s+/g, " ").replace(SALES, "").trim()
}

const onRecord = ({ raw, record }: { raw: string; record: string[] }, context: CastingContext) => {
  csvFunctionRaw('raw: %s', collapseWhitespace(raw))
  if (context.error && context.error.code === "CSV_INCONSISTENT_RECORD_LENGTH") {
    // find the 3rd comma in the line and excise it as it's probably incorrectly part of the payee's name (and shouldn't be but HSBC are crap so...)
    csvFunctionErr('err: CSV_INCONSISTENT_RECORD_LENGTH. Attempting to fix.')
    let stringRaw = raw as string
    let counter = 3 // zero-based index
    let nThIndex = 0

    if (counter > 0) {
      while (counter--) {
        // Get the index of the next occurence
        nThIndex = stringRaw.indexOf(",", nThIndex + ",".length)
      }
    }

    stringRaw = stringRaw.substring(0, nThIndex) + stringRaw.substring(nThIndex + ",".length)
    csvFunctionRaw('Created new record: %s', stringRaw)
    csvFunctionRaw('Attempting to reparse.')
    // call CSV.parse() again on the newly constructed line. This time should return the correct number of fields.
    let result = parse(stringRaw, {
      raw: true,
      trim: true,
      onRecord: onRecord,
      cast: cast,
    })
    // TODO check for context.error.code - maybe removing the excess comma wasn't enough to parse successfully.
    return result[0]
  } // end handling raw row with excess number of commas e.g. comma in company name

  // Purchase amounts need to be negative for Xero import
  // Payments and rerurns are positive amounts (credits) in Xero
  // HSBC CSV has everything as a positive value
  if (record[DESCRIPTION_INDEX].match(PAYMENT) || record[DESCRIPTION_INDEX].match(RETURN)) {
    // do nothing. The amount is already positive
    //onRecordDebug('Leave value positive: %s', record.toString())
  }
  else {
    // change the value to a negative value
    record[AMOUNT_INDEX] = (Number.parseFloat(record[AMOUNT_INDEX]) * -1).toString()
  }

  // delete rows where there is no data in the transaction date column - not a transaction - not interesting
  if (record[TRANSACTION_DATE_INDEX].trim() === "") {
    csvFunctionErr(record.toString())
    return null
  }

  // return Post Date, Txn Amount, null, Description + Foreign CCY Amt
  csvFunctionOut('out: %s', [record[0], record[4], "", `${record[2]} ${record[3]}`].toString())
  return [record[0], record[4], "", `${record[2]} ${record[3]}`]
}

export const csvOnload = (reader: FileReader, excelContext: Excel.RequestContext) => {
  return async () => {
    const raw = reader.result as string
    const rawData = raw.slice(raw.indexOf(`\n`) + 1)
    let data: string[][] = parse(rawData, {
      relax_column_count: true,
      trim: true,
      raw: true,
      cast: cast,
      onRecord: onRecord,
    })
    // replace the header
    data[0] = ["Date", "Amount", "Payee", "Description"]
    populateWorksheet(data, excelContext)
  }
}
