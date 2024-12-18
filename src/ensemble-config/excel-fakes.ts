export class FakeWorkbook {
    activeSheet: FakeWorksheet;
    rows: (string | null | number)[][] = [];
    logs: string[] = [];
    table: FakeTable;

    constructor() {
        this.activeSheet = new FakeWorksheet(() => this.rows);
        this.table = new FakeTable(this);
    }

    getTable(_: string) {
        return this.table;
    }

    getActiveWorksheet() {
        return this.activeSheet;
    }

    log(message: string) {
        this.logs.push(message);
    }

    getLogs() {
        return this.logs;
    }
}

class FakeWorksheet {
    private rowsProvider: () => (string | null | number)[][];

    constructor(rowsProvider: () => (string | null | number)[][]) {
        this.rowsProvider = rowsProvider;
    }

    getRange(range: string): FakeRange {
        const match = /C(\d+)/.exec(range);
        const endRow = match ? parseInt(match[1], 10) : 2;
        return new FakeRange(endRow, this.rowsProvider());
    }
}

class FakeTable {
    workbook: FakeWorkbook;
    sortApplyCallCount: number = 0;

    constructor(workbook: FakeWorkbook) {
        this.workbook = workbook;
    }

    getSort() {
        return {
            apply: (_: { key: number; ascending: boolean }[]) => {
                this.sortApplyCallCount++;
                this.workbook.rows.reverse();
            },
        };
    }
}

class FakeRange {
    private endRow: number;
    private rows: (string | null | number)[][];

    constructor(endRow: number, rows: (string | null | number)[][]) {
        this.endRow = endRow;
        this.rows = rows.slice(0, endRow - 1);
    }

    getValues(): (string | null | number)[][] {
        return this.rows;
    }
}