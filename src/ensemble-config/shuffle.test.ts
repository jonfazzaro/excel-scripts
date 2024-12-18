import {main} from "./shuffle";
import {FakeWorkbook} from "./excel-fakes";

describe('The shuffle script', () => {
    let fakeWorkbook: FakeWorkbook;
    let logSpy = vi.spyOn(console, "log").mockImplementation(() => {
    });

    beforeEach(() => {
        fakeWorkbook = new FakeWorkbook();
        logSpy.mockClear();
    });

    describe('given no matching rows', () => {
        beforeEach(() => {
            runWith([['A', 'B'], ['C', 'D']]);
        });
        
        it('does not shuffle extra times', () => {
            expect(fakeWorkbook.table.sortApplyCallCount).toBe(1); 
            expect(logSpy).toHaveBeenCalledWith('Shuffled an extra 0 times.');
        });
    });

    describe('given matching rows that persist', () => {
        beforeEach(() => {
            runWith([['A', 'A'], ['B', 'B']])
        });

        it('shuffles up to 100 times', () => {
            expect(fakeWorkbook.table.sortApplyCallCount).toBe(101); 
            expect(logSpy).toHaveBeenCalledWith('Shuffled an extra 100 times.');
        });
    });

    function runWith(rows: string[][]) {
        fakeWorkbook.rows = rows;
        main(fakeWorkbook as any);
    }
});
