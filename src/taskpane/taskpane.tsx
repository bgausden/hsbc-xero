/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/Glow Corporate Vert Prot16w.png"
import "../../assets/Glow Corporate Vert Prot32w.png"
import "../../assets/Glow Corporate Vert Prot64w.png"
import "../../assets/Glow Corporate Vert Prot80w.png"

import { csvOnload } from "./csv-functions"
import { initializeIcons } from "@fluentui/react/lib/Icons"

/* global document, Excel, Office, FileReader */

export const TRANSACTION_DATE_INDEX = 1 // zero indexed
export const SALES = "SALES: "

export async function load(): Promise<void> {
    try {
        await Excel.run(async (context) => {
            const elem = document.getElementById("file")
            if (isHTMLElement(elem)) {
                const file = document.getElementById("file") as HTMLInputElement
                const reader = new FileReader()

                if (file.files) {
                    if (file.files[0]) {
                        reader.onload = csvOnload(reader, context)
                        reader.readAsText(file.files[0])
                    }
                }
            } else {
                throw new Error("file element not found")
            }
            context.trackedObjects.add(
                context.workbook.worksheets.getActiveWorksheet()
            )
        })
    } catch (error) {
        console.error(`load(): ${error}`)
    }
}

function isHTMLElement(elem: HTMLElement | null): elem is HTMLElement {
    return elem !== null && elem.tagName !== undefined
}
Office.onReady((info) => {
    initializeIcons()
    if (info.host === Office.HostType.Excel) {
        let elem = document.getElementById("sideload-msg")
        if (isHTMLElement(elem)) {
            elem.style.display = "none"
        }
        elem = document.getElementById("app-body")
        if (isHTMLElement(elem)) {
            elem.style.display = "flex"
        }
        elem = document.getElementById("load")
        if (isHTMLElement(elem)) {
            elem.onclick = load
        }
    }
})
