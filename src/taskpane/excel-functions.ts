import { displayMessageBar } from "./messagebar"

/* global Excel */

export async function populateWorksheet(
    data: unknown[][],
    context: Excel.RequestContext
): Promise<void> {
    const sheet = context.workbook.worksheets.getActiveWorksheet()
    const usedRange = sheet.getUsedRangeOrNullObject()
    usedRange.load(["address", "cellCount"])
    await context.sync()
    if (usedRange.address) {
        usedRange.format.fill.color = "lightYellow"
        console.log(usedRange.address)
        displayMessageBar(
            `Please load the transaction data into an empty workbook.`
        )
        // TODO auto-hide the messagebar after 10s
    } else {
        sheet
            .getRange("A1")
            .getResizedRange(data.length - 1, data[0].length - 1).values = data
        sheet.getUsedRange().format.autofitColumns()
    }
    await context.sync()
}
