export function main(workbook: ExcelScript.Workbook) {
    shuffle(workbook)
    let shuffled = 0
    while (shuffled < 100 && hasSamePersonDrivingAndNavigating(workbook)) {
        shuffle(workbook)
        shuffled++
    }

    console.log(`Shuffled an extra ${shuffled} times.`)
}

function shuffle(workbook: ExcelScript.Workbook) {
    workbook.getTable("Table1")?.getSort().apply([{key: 1, ascending: true}]);
}

function hasSamePersonDrivingAndNavigating(workbook: ExcelScript.Workbook) {
    const range = rotationRange(workbook);
    return findMatchingRows(range.getValues()).length > 0
}

function rotationRange(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getActiveWorksheet();
    return sheet.getRange(`B2:C${(cellValue("F2", sheet))}`);
}

function findMatchingRows(rows: (string | number | boolean | null)[][]): number[] {
    return rows.reduce<number[]>((matchingIndexes, [value1, value2], index) => {
        if (isMatchingRow(value1, value2)) {
            matchingIndexes.push(index);
        }
        return matchingIndexes;
    }, []);
}

function isMatchingRow(value1: unknown, value2: unknown): boolean {
    return value1 === value2 && (value1 !== null && value1 !== "");
}

function cellValue(address: string, sheet: ExcelScript.Worksheet) {
    return sheet.getRange(address).getValues()[0][0]
}