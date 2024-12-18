import { main } from "./shuffle";

describe('The shuffle script', () => {
    let fakeWorkbook: FakeWorkbook;
    let logSpy = vi.spyOn(console, "log").mockImplementation(() => {});

    beforeEach(() => {
        fakeWorkbook = new FakeWorkbook();
        logSpy.mockClear();
    });

    it('should stop shuffling if no matching rows are detected', () => {
        fakeWorkbook.rows = [['A', 'B'], ['C', 'D']]; // Rows are pre-set to prevent matching
        main(fakeWorkbook as any);

        expect(fakeWorkbook.table.sortApplyCallCount).toBe(1); // Called once initially
        expect(logSpy).toHaveBeenCalledWith('Shuffled 0 times.');
    });

    it('should shuffle up to 100 times if matching rows persist', () => {
        fakeWorkbook.rows = [['A', 'A'], ['B', 'B']]; // Matching rows remain
        main(fakeWorkbook as any);

        expect(fakeWorkbook.table.sortApplyCallCount).toBe(101); // Runs shuffle 100 + 1 initial
        expect(logSpy).toHaveBeenCalledWith('Shuffled 100 times.');
    });
});

/**
 * Fake implementations of ExcelScript types with built-in spying
 */

// Fake Workbook class
class FakeWorkbook {
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

// Fake Worksheet class
class FakeWorksheet {
    private rowsProvider: () => (string | null | number)[][];

    constructor(rowsProvider: () => (string | null | number)[][]) {
        this.rowsProvider = rowsProvider;
    }

    getRange(range: string): FakeRange {
        const match = /C(\d+)/.exec(range); // Look for last row in range like `B2:C10`
        const endRow = match ? parseInt(match[1], 10) : 2;
        return new FakeRange(endRow, this.rowsProvider());
    }
}

// Fake Table class with shuffle tracking
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
                // Simulate shuffling by reversing rows
                this.workbook.rows.reverse();
            },
        };
    }
}

// Fake Range class
class FakeRange {
    private endRow: number;
    private rows: (string | null | number)[][];

    constructor(endRow: number, rows: (string | null | number)[][]) {
        this.endRow = endRow;
        this.rows = rows.slice(0, endRow - 1); // Simulate range restriction
    }

    getValues(): (string | null | number)[][] {
        return this.rows;
    }
}