//#region main
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Dashboard')
        .addItem('Build', 'build')
        .addItem('Format', 'formatDashboard')
        .addToUi();
}

function formatDashboard() {
    const db = new DashBuilder(CONFIG);
    db.clearFormat();
    db.format();
}

function build() {
    const db = new DashBuilder(CONFIG);
    db.build();
}
//#endregion main

//#region config
const DATA_CELLS = {
    projectName: 'B2',
    projectId: 'E2',
    totalFee: 'G3',
    architectural: 'H3',
    consultants: 'I3',
    unallocated: 'K3',
    allocatedPct: 'L3',
    assignedPct: 'M3',
    feeSpent: 'N3',
    feeSpentPct: 'O3',
    projectType: 'Q3',
    totalBilled: 'S3',
    remainingToBill: 'T3',
    billedPct: 'U3',
    totalReceived: 'V3',
    remainingToReceive: 'W3',
};
const NUMBER_FORMAT: Record<NumberFormat, string> = {
    acct: '$#,##0.00;($#,##0.00)',
    pct: '0.00%',
};
const COLUMN_DEFS: ColumnDef[] = [
    {
        header: 'Code',
        id: 'projectId',
        dataCell: DATA_CELLS.projectId,
    },
    {
        header: 'Project Name',
        id: 'projectName',
        dataCell: DATA_CELLS.projectName,
    },
    {
        header: 'Project Type',
        id: 'projectType',
        dataCell: DATA_CELLS.projectType,
    },
    {
        header: 'Total Fee',
        id: 'totalFee',
        format: 'acct',
        dataCell: DATA_CELLS.totalFee,
    },
    {
        header: 'Architectural',
        id: 'architectural',
        format: 'acct',
        dataCell: DATA_CELLS.architectural,
    },
    {
        header: 'Consultants',
        id: 'consultants',
        format: 'acct',
        dataCell: DATA_CELLS.consultants,
    },
    {
        header: 'Unallocated ($)',
        id: 'unallocated',
        format: 'acct',
        dataCell: DATA_CELLS.unallocated,
    },
    {
        header: 'Allocated (%)',
        id: 'allocatedPct',
        format: 'pct',
        dataCell: DATA_CELLS.allocatedPct,
    },
    {
        header: 'Assigned (%)',
        id: 'assignedPct',
        format: 'pct',
        dataCell: DATA_CELLS.assignedPct,
    },
    {
        header: 'Fee Spent ($)',
        id: 'feeSpent',
        format: 'acct',
        dataCell: DATA_CELLS.feeSpent,
    },
    {
        header: 'Fee Spent (%)',
        id: 'feeSpentPct',
        format: 'pct',
        dataCell: DATA_CELLS.feeSpentPct,
    },
    {
        header: 'Total Billed',
        id: 'totalBilled',
        format: 'acct',
        dataCell: DATA_CELLS.totalBilled,
    },
    {
        header: 'Remaining to Bill',
        id: 'remainingToBill',
        format: 'acct',
        dataCell: DATA_CELLS.remainingToBill,
    },
    {
        header: '% Billed',
        id: 'billedPct',
        format: 'pct',
        dataCell: DATA_CELLS.billedPct,
    },
    {
        header: 'Total Received',
        id: 'totalReceived',
        format: 'acct',
        dataCell: DATA_CELLS.totalReceived,
    },
    {
        header: 'Remaining to Receive',
        id: 'remainingToReceive',
        format: 'acct',
        dataCell: DATA_CELLS.remainingToReceive,
    },
];

const CONFIG = {
    dashboard: {
        name: 'Dashboard',
        columnDefs: COLUMN_DEFS,
        headerRow: 1,
        shouldProcessSheet: _isSheetNameValidProject,
    },
};
//#endregion config

//#region dash builder
type TConfig = typeof CONFIG;
class DashBuilder {
    headerRow: number;
    columnDefs: TConfig['dashboard']['columnDefs'];
    dashboardName: string;
    shouldProcessSheet: (name: string) => boolean;

    ss: GoogleAppsScript.Spreadsheet.Spreadsheet;
    dashSheet: GoogleAppsScript.Spreadsheet.Sheet;
    sheets: GoogleAppsScript.Spreadsheet.Sheet[];
    sheetCache: Record<string, GoogleAppsScript.Spreadsheet.Sheet | null> = {};
    dashHeaders: string[] = [];

    dataRanges: string[] = []; // [a1:c1, e1]
    rangeToColIds: Record<string, string[]> = {}; // {[a1:c1]: [id1, id2, id3]}

    constructor({
        dashboard: { name, columnDefs, headerRow, shouldProcessSheet },
    }: TConfig) {
        // TODO: validated config:
        // - unique column def ids
        // - valid a1 notation for cells

        // bindings
        this.requireDashSheet = this.requireDashSheet.bind(this);
        this.getSheetByName = this.getSheetByName.bind(this);
        this.getSheetNamesToProcess = this.getSheetNamesToProcess.bind(this);
        this.initColumnCellRefs = this.initColumnCellRefs.bind(this);
        this.processSheet = this.processSheet.bind(this);

        // args
        this.dashboardName = name;
        this.columnDefs = columnDefs;
        this.headerRow = headerRow;
        this.shouldProcessSheet = shouldProcessSheet;

        // initialization
        this.ss = SpreadsheetApp.getActiveSpreadsheet();
        this.dashSheet = this.requireDashSheet();
        this.sheets = this.ss.getSheets();
        this.initColumnCellRefs(columnDefs);
        this.dashHeaders = columnDefs.map((cd) => cd.header);
    }
    //#region private
    private initColumnCellRefs(columnDefs: TConfig['dashboard']['columnDefs']) {
        this.dataRanges = U.toContiguous(columnDefs.map((cd) => cd.dataCell));
        this.rangeToColIds = this.dataRanges.reduce(
            (acc, currentRange) => {
                acc[currentRange] = [];
                columnDefs.forEach((cd) => {
                    if (U.isCellInRange(cd.dataCell, currentRange)) {
                        acc[currentRange].push(cd.id);
                    }
                });
                return acc;
            },
            {} as Record<string, string[]>
        );
    }
    private requireDashSheet() {
        const dashboardSheet = this.getSheetByName(this.dashboardName);
        if (!dashboardSheet) {
            throw new Error('Dashboard Sheet Missing');
        }
        return dashboardSheet;
    }
    private getSheetByName(sheetName: string) {
        const found = this.sheetCache[sheetName];
        if (found) return found;

        this.sheetCache[sheetName] = this.ss.getSheetByName(sheetName);
        return this.sheetCache[sheetName];
    }
    private getSheetNamesToProcess() {
        const allSheetNames = this.sheets.map((s) => s.getName());
        return allSheetNames.filter(this.shouldProcessSheet);
    }
    private processSheet(
        sheetName: string
    ): Record<
        TConfig['dashboard']['columnDefs'][number]['dataCell'],
        string | number | null
    > {
        const sheet = this.getSheetByName(sheetName);
        if (sheet == null)
            throw new Error(`Fatal: Couldn't find sheet ${sheet}`);

        const results: Record<string, any> = {};
        this.columnDefs.forEach((cd) => (results[cd.id] = null));

        this.dataRanges.forEach((range) => {
            const colIds = this.rangeToColIds[range];
            if (!colIds)
                throw new Error(
                    `Fata: Couldn't find column definitions for range: ${range}`
                );
            const v = sheet.getRange(range).getValues()[0]; // returns a 2d array
            colIds.forEach((cid, i) => {
                results[cid] = v[i];
            });
        });
        return results;
    }
    /** apply column formatting */
    private formatDashboardColumns() {
        const lastRow = this.dashSheet.getLastRow();
        const lastCol = this.dashSheet.getLastColumn();
        const dashSheetHeaders = this.dashSheet
            .getRange(this.headerRow, 1, 1, lastCol)
            .getValues()[0];

        // create map of format => column defs
        const formatToColDefs: Record<NumberFormat, ColumnDef[]> = _.groupBy(
            this.columnDefs.filter((c) => c.format != null),
            'format'
        );

        // for each format's associated column defs do the following:
        // 1. find the index for each column def header on the dashboard
        // 2. create cell refs from those indeces
        // 3. merge cell refes into contiguous ranges
        // 4. expand ranges to last row
        const formatToExpandedRanges: Record<NumberFormat, string[]> =
            _.mapValues(formatToColDefs, (colDefs) => {
                const headerIdxs = colDefs
                    .map((cd) => dashSheetHeaders.indexOf(cd.header))
                    .filter((i) => i >= 0);

                const cellRefs = headerIdxs.map((i) => {
                    const res = `${U.indexToA1Col(i)}${this.headerRow}`;
                    return res;
                });

                const ranges = U.toContiguous(cellRefs);

                return ranges.map((range) => U.expandRange(range, lastRow));
            });

        // apply formatting for each range
        for (const entry of Object.entries(formatToExpandedRanges)) {
            const [fmt, ranges] = entry as [NumberFormat, string[]];
            ranges.forEach((range) => {
                const sheetRange = this.dashSheet.getRange(range);
                const fmtString = NUMBER_FORMAT[fmt];
                sheetRange.setNumberFormat(fmtString);
            });
        }
    }
    //#endregion private
    format() {
        const lastColumn = this.dashSheet.getLastColumn();
        const maxRows = this.dashSheet.getMaxRows();
        const maxColumns = this.dashSheet.getMaxColumns();
        this.dashSheet
            .getRange(1, 1, maxRows, maxColumns)
            .setFontFamily('Roboto Mono');

        this.dashSheet.setFrozenColumns(2); // freeze columns
        this.dashSheet.setRowHeight(this.headerRow, 60); // header height
        this.dashSheet // header format
            .getRange(this.headerRow, 1, 1, lastColumn)
            .setFontWeight('bold')
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');

        this.formatDashboardColumns(); // number formatting and such

        // runs last
        this.dashSheet.autoResizeColumns(1, lastColumn);
    }

    clear() {
        const maxRows = this.dashSheet.getMaxRows();
        const maxColumns = this.dashSheet.getMaxColumns();
        this.dashSheet.getRange(1, 1, maxRows, maxColumns).clear();
    }

    clearFormat() {
        const maxRows = this.dashSheet.getMaxRows();
        const maxColumns = this.dashSheet.getMaxColumns();
        this.dashSheet.getRange(1, 1, maxRows, maxColumns).clearFormat();
    }

    build() {
        const sheetsToProcess = this.getSheetNamesToProcess();
        const sheetsData = sheetsToProcess.map(this.processSheet);

        this.clear();
        this.dashSheet
            .getRange(this.headerRow, 1, 1, this.dashHeaders.length)
            .setValues([this.dashHeaders]);

        const START_DATA_ROW = this.headerRow + 1;
        sheetsData.forEach((data, i) => {
            const row = this.dashHeaders.map((header) => {
                const cd = this.columnDefs.find((col) => col.header === header);
                if (!cd) return null;
                return data[cd.id] ?? null;
            });
            this.dashSheet
                .getRange(START_DATA_ROW + i, 1, 1, row.length)
                .setValues([row]);
        });

        this.format();
    }
}
//#endregion dash builder

//#region utils
const U = {
    exrtact4DigitString(name: string): string | null {
        const regex = /\b\d{4}\b/; // Regex to match exactly 4 digits in a string
        const match = name.match(regex);
        return match ? match[0] : null; // Return the matched number or null if not found
    },

    isNextColumn(leftCol: string, rightCol: string): boolean {
        return U.colToNumber(rightCol) === U.colToNumber(leftCol) + 1;
    },

    colToNumber(col: string) {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
        }
        return num;
    },

    cellRefToRowCol(cellRef: string) {
        const [col, row] = cellRef.match(/([A-Z]+)(\d+)/)!.slice(1);
        return { col, row };
    },
    /** transforms an array of cell refs into an array of horizontally contiguous ranges */
    toContiguous(cellRefs: string[]) {
        const sortedRefs = cellRefs.sort((a, b) => {
            const [colA, rowA] = a.match(/([A-Z]+)(\d+)/)!.slice(1);
            const [colB, rowB] = b.match(/([A-Z]+)(\d+)/)!.slice(1);
            return colA.localeCompare(colB) || Number(rowA) - Number(rowB);
        });

        const ranges: string[] = [];
        let start = sortedRefs[0];
        let end = start;

        for (let i = 1; i < sortedRefs.length; i++) {
            const current = sortedRefs[i];
            const [currentCol, currentRow] = current
                .match(/([A-Z]+)(\d+)/)!
                .slice(1);
            const [endCol, endRow] = end.match(/([A-Z]+)(\d+)/)!.slice(1);

            // Check if the current cell is contiguous
            if (currentRow === endRow && U.isNextColumn(endCol, currentCol)) {
                end = current; // Extend the range
            } else {
                // Save the previous range
                ranges.push(start === end ? start : `${start}:${end}`);
                start = current; // Start a new range
                end = start;
            }
        }

        // Push the last range
        ranges.push(start === end ? start : `${start}:${end}`);

        return ranges;
    },

    indexToA1Col(index: number) {
        if (index < 0) {
            throw new Error('Index must be a non-negative integer.');
        }

        let column = '';
        let currentIndex = index;

        while (currentIndex >= 0) {
            const letter = String.fromCharCode((currentIndex % 26) + 65);
            column = letter + column;
            currentIndex = Math.floor(currentIndex / 26) - 1;
        }

        return column;
    },

    a1ToIndex(a1: string): { row: number; col: number } {
        const { col: colStr, row: rowStr } = U.cellRefToRowCol(a1);

        const col = Array.from(colStr).reduce((acc, char) => {
            return acc * 26 + (char.charCodeAt(0) - 'A'.charCodeAt(0) + 1);
        }, 0);
        const row = parseInt(rowStr, 10);

        return { row, col };
    },

    isCellInRange(cell: string, range: string): boolean {
        if (!range.includes(':')) return cell === range;
        const [start, end] = range.split(':');
        const cellIndex = U.a1ToIndex(cell);
        const startIndex = U.a1ToIndex(start);
        const endIndex = U.a1ToIndex(end);

        return (
            cellIndex.row >= startIndex.row &&
            cellIndex.row <= endIndex.row &&
            cellIndex.col >= startIndex.col &&
            cellIndex.col <= endIndex.col
        );
    },
    /** takes a range and a row and expands the range to the provided row */
    expandRange(range: string, expandToRow: number) {
        if (range.includes(':')) {
            const [startRef, endRef] = range.split(':');
            const end = U.cellRefToRowCol(endRef);
            return `${startRef}:${end.col}${expandToRow}`;
        } else {
            // single column
            const ref = range;
            const { col } = U.cellRefToRowCol(ref);
            return `${ref}:${col}${expandToRow}`;
        }
    },
};

type GroupByKey<T> = keyof T | ((item: T) => string | number);
const _ = {
    groupBy<T>(array: T[], key: GroupByKey<T>): Record<string | number, T[]> {
        return array.reduce(
            (result: Record<string | number, T[]>, item: T) => {
                // Determine the group key
                const groupKey =
                    typeof key === 'function'
                        ? key(item)
                        : (item[key] as string | number);

                // Initialize the group if it doesn't exist
                if (!result[groupKey]) {
                    result[groupKey] = [];
                }

                // Add the item to the group
                result[groupKey].push(item);
                return result;
            },
            {} as Record<string | number, T[]>
        );
    },
    mapValues<T, U>(
        obj: Record<string, T>,
        iteratee: (value: T, key: string) => U
    ): Record<string, U> {
        const result: Record<string, U> = {};
        for (const key in obj) {
            if (obj.hasOwnProperty(key)) {
                result[key] = iteratee(obj[key], key);
            }
        }
        return result;
    },
};
//#endregion utils

//#region defaults
function _isSheetNameValidProject(name: string) {
    // by default, a sheet name is a valid project id if it contains a substring that is a number
    const numberSubstring = U.exrtact4DigitString(name);
    return numberSubstring != null && !isNaN(Number(numberSubstring));
}
//#endregion defaults

//#region types
type InvoicingData = {
    totalBilled: number;
    remainingToBill: number;
    billedPct: number;
    totalReceived: number;
    remainingToReceive: number;
};

type SpendData = {
    unallocated: number;
    allocatedPct: number;
    assignedPct: number;
    feeSpent: number;
    feeSpentPct: number;
};

type FeeData = {
    totalFee: number;
    architectural: number;
    consultants: number;
};

type ProjectSheetData = {
    projectId: string;
    projectName: string;
    projectType: string;
} & FeeData &
    SpendData &
    InvoicingData;

type DataKey = keyof ProjectSheetData;
type NumberFormat = 'acct' | 'pct';

type ColumnDef = {
    header: string;
    id: DataKey;
    format?: NumberFormat;
    dataCell: string;
};
//#endregion types
