export function main(workbook: ExcelScript.Workbook) {
    shuffle(workbook)
    let shuffled = 0
    while (shuffled < 100 && hasSamePersonDrivingAndNavigating(workbook)) {
        shuffle(workbook)
        shuffled++
    }

    console.log(`Shuffled ${shuffled} times.`)
}

function shuffle(workbook: ExcelScript.Workbook) {
    workbook.getTable("Table1")?.getSort().apply([{ key: 1, ascending: true }]);
}

function hasSamePersonDrivingAndNavigating(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getActiveWorksheet();
    const lastRow = cellValue("F2", sheet)
    const range = sheet.getRange(`B2:C${lastRow}`);
    const values = range.getValues();
    const matchingRows = findMatchingRows(values);
    return matchingRows.length > 0
}

function findMatchingRows(values: (string | number | boolean | null)[][]): number[] {
    const matchingRows: number[] = [];
    for (let row = 0; row < values.length; row++) {
        const [value1, value2] = values[row];
        if (value1 === value2 && value1 !== null && value1 !== "") {
            matchingRows.push(row);
        }
    }
    return matchingRows;
}

function cellValue(address:string, sheet: ExcelScript.Worksheet) {
    return sheet.getRange(address).getValues()[0][0]
}