/* eslint-disable no-unused-vars */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

import { MessageBarButton, Link, StackItem, ChoiceGroup } from "office-ui-fabric-react";
import { initializeIcons } from "@fluentui/react/lib/Icons";
// eslint-disable-next-line node/no-missing-import
import { displayMessageBar } from "./messagebar";  
// eslint-disable-next-line node/no-missing-import
import { csvOnload } from "./csv-functions";

/* For Webpack so we don't need them in the html */
/* import "@uifabric/react-hooks/dist/react-hooks";
import "office-ui-fabric-core/dist/css/fabric.min.css"; */

/* global document, Excel, Office, FileReader */

const mergeCol = 1; // zero indexed
const POST_DATE = "Post Date";
export const TRANSACTION_DATE = "Transaction Date";
export const TRANSACTION_DATE_INDEX = 1; // zero indexed
const DESCRIPTION = "Description";
const FOREIGN_CURRENCY_AMOUNT = "Foreign Currency Amount";
const AMOUNT_HKD = "Amount(HKD)";
const SALES = "SALES: ";

function is(value: any) {
  return {
    a: function(check: any) {
      if (check.prototype) check = check.prototype.constructor.name;
      const type: string = Object.prototype.toString
        .call(value)
        .slice(8, -1)
        .toLowerCase();
      return value != null && type === check.toLowerCase();
    }
  };
}

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";
    document.getElementById("run")!.onclick = run;
    document.getElementById("load")!.onclick = load;
    initializeIcons();
  }
});

export async function load() {
  try {
    await Excel.run(async context => {
      const file = document.getElementById("file") as HTMLInputElement;
      const reader = new FileReader();
      let content: [[string]];

      if (file?.files && file.files[0]) {
        reader.onload = csvOnload(reader, context);
        reader.readAsText(file.files[0]);
      }
    });
  } catch (error) {
    console.error(`load(): ${error}`);
  }
}

export async function run() {
  try {
    await Excel.run(async context => {
      // get the used (non-empty) cells in the active worksheet

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      context.trackedObjects.add(sheet);
      context.trackedObjects.add(usedRange);

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
      console.log("Done");
    });
  } catch (error) {
    console.error(`run(): ${error}`);
  }
}
