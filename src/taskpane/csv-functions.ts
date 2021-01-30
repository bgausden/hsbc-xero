import { CastingContext } from "csv-parse";
import parse from "csv-parse/lib/sync";
// eslint-disable-next-line node/no-missing-import
import { populateWorksheet } from "./excel-functions";
// eslint-disable-next-line node/no-missing-import
import { SALES, TRANSACTION_DATE_INDEX } from "./taskpane";

/* global Excel */

const onRecord = ({ raw, record }: { raw: string; record: string[]; }, context: CastingContext) => {
  if (context.error && context.error.code === "CSV_INCONSISTENT_RECORD_LENGTH") {
    let stringRaw = raw as string;
    let counter = 3; // zero-based index
    let nThIndex = 0;

    if (counter > 0) {
      while (counter--) {
        // Get the index of the next occurence
        //nThIndex = String.prototype.indexOf.call(stringRaw, ",", nThIndex + ",".length);
        nThIndex = stringRaw.indexOf(",", nThIndex + ",".length)
      }
    }

    stringRaw = stringRaw.substring(0, nThIndex) + stringRaw.substring(nThIndex + ",".length);
    //result = stringRaw.trim().split(",").map(field => field.replace(/\s+/g, " ").trim())
    //result = stringRaw.split(",")
    let result = parse(stringRaw, {
      raw: true,
      trim: true,
      onRecord: onRecord,
      cast: (value) => {
        return value.replace(/\s+/g, " ").trim();
      }
    });
    return result[0];
  }

  // delete rows where there is only data in the 0th column (garbage)
  if (record[TRANSACTION_DATE_INDEX].trim() === "")
    return null;

  return [record[0], record[4], "", `${record[2]} ${record[3]}`];
};

export const csvOnload = (reader: FileReader, excelContext: Excel.RequestContext) => {
  return async () => {
    const raw = reader.result as string;
    const rawData = raw.slice(raw.indexOf(`\n`) + 1);
    let data: string[][] = parse(rawData, {
      relax_column_count: true,
      trim: true,
      raw: true,
      cast: (value) => {
        return value.replace(/\s+/g, " ").replace(SALES,"").trim()
      },
      onRecord: onRecord
    });
    // replace the header
    data[0] = ["Date", "Amount", "Payee", "Description"];
    console.log(data);
    populateWorksheet(data,excelContext)
  };
};
