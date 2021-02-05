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

// eslint-disable-next-line node/no-missing-import
import { csvOnload } from "./csv-functions";
import { initializeIcons } from "@fluentui/react";

/* global document, Excel, Office, FileReader */

const mergeCol = 1; // zero indexed
const POST_DATE = "Post Date";
export const TRANSACTION_DATE = "Transaction Date";
export const TRANSACTION_DATE_INDEX = 1; // zero indexed
const DESCRIPTION = "Description";
const FOREIGN_CURRENCY_AMOUNT = "Foreign Currency Amount";
const AMOUNT_HKD = "Amount(HKD)";
export const SALES = "SALES: ";

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

Office.onReady(info => {
  initializeIcons();
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";
    document.getElementById("load")!.onclick = load;
  }
});
