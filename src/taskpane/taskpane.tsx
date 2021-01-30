/* eslint-disable no-unused-vars */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import "../../assets/Glow Corporate Vertical.svg";

import { MessageBarButton, Link, StackItem, ChoiceGroup } from "office-ui-fabric-react";
import { initializeIcons } from "@fluentui/react/lib/Icons";
// eslint-disable-next-line node/no-missing-import
import { displayMessageBar } from "./messagebar";
// eslint-disable-next-line node/no-missing-import
import { csvOnload } from "./csv-functions";

/* global document, Excel, Office, FileReader */

const mergeCol = 1; // zero indexed
const POST_DATE = "Post Date";
export const TRANSACTION_DATE = "Transaction Date";
export const TRANSACTION_DATE_INDEX = 1; // zero indexed
const DESCRIPTION = "Description";
const FOREIGN_CURRENCY_AMOUNT = "Foreign Currency Amount";
const AMOUNT_HKD = "Amount(HKD)";
export const SALES = "SALES: ";

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
    //document.getElementById("run")!.onclick = run;
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
      /*       run();
       */
    });
  } catch (error) {
    console.error(`load(): ${error}`);
  }
}

/* export async function run() {
  try {
    await Excel.run(async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      context.trackedObjects.add(sheet);
      context.trackedObjects.add(usedRange);
    });
  } catch (error) {
    console.error(`run(): ${error}`);
  }
} */
