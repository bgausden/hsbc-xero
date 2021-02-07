/* eslint-disable no-unused-vars */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/Glow Corporate Vert Prot16w.png";
import "../../assets/Glow Corporate Vert Prot32w.png";
import "../../assets/Glow Corporate Vert Prot64w.png";
import "../../assets/Glow Corporate Vert Prot80w.png";

// eslint-disable-next-line node/no-missing-import
import { csvOnload } from "./csv-functions";
import { initializeIcons } from "@fluentui/react/lib/Icons";

/* global document, Excel, Office, FileReader */

export const TRANSACTION_DATE = "Transaction Date";
export const TRANSACTION_DATE_INDEX = 1; // zero indexed
export const SALES = "SALES: ";

export async function load() {
  try {
    await Excel.run(async context => {
      const file = document.getElementById("file") as HTMLInputElement;
      const reader = new FileReader();

      if (file?.files && file.files[0]) {
        reader.onload = csvOnload(reader, context);
        reader.readAsText(file.files[0]);
      }

      context.trackedObjects.add(context.workbook.worksheets.getActiveWorksheet());
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
