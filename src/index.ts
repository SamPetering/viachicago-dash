type NullableRecord<T extends Record<string, any>> = {
    [K in keyof T]: T[K] | null;
};

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
type HeaderKey = Exclude<DataKey, 'projectId' | 'projectName'>; // project name and id have no header

type SupplementalProjectData = {
    sheetName: string;
    stage: string;
    pctComplete: number;
    projectedFee: number;
};
type DashboardEntry = NullableRecord<
    SupplementalProjectData & ProjectSheetData
>;

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
const HEADER_CELLS = {
    totalFee: 'G2',
    architectural: 'H2',
    consultants: 'I2',
    unallocated: 'K2',
    allocatedPct: 'L2',
    assignedPct: 'M2',
    feeSpent: 'N2',
    feeSpentPct: 'O2',
    projectType: 'Q2',
    totalBilled: 'S2',
    remainingToBill: 'T2',
    billedPct: 'U2',
    totalReceived: 'V2',
    remainingToReceive: 'W2',
};

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Dashboard')
        .addItem('Build', 'buildDashboard')
        .addItem('Format', 'formatDashboard')
        .addToUi();
}

function formatDashboard() {
    const dashBuilder = new DashBuilder({
        dashboardSheetName: 'Dashboard',
        dataCellMap: DATA_CELLS,
        headersCellMap: HEADER_CELLS,
    });
    dashBuilder.clearFormat();
    dashBuilder.format();
}

function buildDashboard() {
    const dashBuilder = new DashBuilder({
        dashboardSheetName: 'Dashboard',
        dataCellMap: DATA_CELLS,
        headersCellMap: HEADER_CELLS,
    });

    const data = dashBuilder.getData();
    dashBuilder.clear();
    dashBuilder.addDataToSheet(data);
    dashBuilder.format();
}

class ProjectDataSheet {
    sheet: GoogleAppsScript.Spreadsheet.Sheet;
    dataCells: Record<DataKey, string>;
    headerCells: Record<HeaderKey, string>;
    constructor(args: {
        sheet: GoogleAppsScript.Spreadsheet.Sheet;
        dataCellMap: Record<DataKey, string>;
        headersCellMap: Record<HeaderKey, string>;
    }) {
        this.sheet = args.sheet;
        this.dataCells = args.dataCellMap;
        this.headerCells = args.headersCellMap;

        this.getProjectData = this.getProjectData.bind(this);
        this.getRangesWithValues = this.getRangesWithValues.bind(this);
    }
    private getRangesWithValues(
        ranges: string[]
    ): Record<string, (string | number)[]> {
        const values: Record<string, (string | number)[]> = {};

        ranges.forEach((rangeString) => {
            const range = this.sheet.getRange(rangeString);
            const rangeValues = range.getValues(); // Get values as a 2D array
            values[rangeString] = rangeValues.flat(); // Flatten the values and assign to the record
        });

        return values;
    }

    getProjectDataHeaders(): Record<DataKey, string> {
        const headerRanges = U.findContiguousRanges(this.headerCells);
        const rangesHeadersMap = this.getRangesWithValues(headerRanges);
        const headersValuesFlat = Object.values(rangesHeadersMap).flat();
        const gh = (cell: string) =>
            headersValuesFlat[U.getIndexFromRanges(headerRanges, cell)];

        return {
            projectId: 'Project ID',
            projectName: 'Project Name',
            totalFee: U.toString(gh(this.headerCells.totalFee)),
            architectural: U.toString(gh(this.headerCells.architectural)),
            consultants: U.toString(gh(this.headerCells.consultants)),
            unallocated: U.toString(gh(this.headerCells.unallocated)),
            allocatedPct: U.toString(gh(this.headerCells.allocatedPct)),
            assignedPct: U.toString(gh(this.headerCells.assignedPct)),
            feeSpent: U.toString(gh(this.headerCells.feeSpent)),
            feeSpentPct: U.toString(gh(this.headerCells.feeSpentPct)),
            projectType: U.toString(gh(this.headerCells.projectType)),
            totalBilled: U.toString(gh(this.headerCells.totalBilled)),
            remainingToBill: U.toString(gh(this.headerCells.remainingToBill)),
            billedPct: U.toString(gh(this.headerCells.billedPct)),
            totalReceived: U.toString(gh(this.headerCells.totalReceived)),
            remainingToReceive: U.toString(
                gh(this.headerCells.remainingToReceive)
            ),
        };
    }

    getProjectData(): ProjectSheetData {
        const dataRanges = U.findContiguousRanges(this.dataCells);
        const rangesValuesMap = this.getRangesWithValues(dataRanges);
        const dataValuesFlat = Object.values(rangesValuesMap).flat();

        const gd = (cell: string) =>
            dataValuesFlat[U.getIndexFromRanges(dataRanges, cell)];

        return {
            projectName: U.toString(gd(this.dataCells.projectName)),
            projectId: U.toString(gd(this.dataCells.projectId)),
            totalFee: U.toNumber(gd(this.dataCells.totalFee)),
            architectural: U.toNumber(gd(this.dataCells.architectural)),
            consultants: U.toNumber(gd(this.dataCells.consultants)),
            unallocated: U.toNumber(gd(this.dataCells.unallocated)),
            allocatedPct: U.toNumber(gd(this.dataCells.allocatedPct)),
            assignedPct: U.toNumber(gd(this.dataCells.assignedPct)),
            feeSpent: U.toNumber(gd(this.dataCells.feeSpent)),
            feeSpentPct: U.toNumber(gd(this.dataCells.feeSpentPct)),
            projectType: U.toString(gd(this.dataCells.projectType)),
            totalBilled: U.toNumber(gd(this.dataCells.totalBilled)),
            remainingToBill: U.toNumber(gd(this.dataCells.remainingToBill)),
            billedPct: U.toNumber(gd(this.dataCells.billedPct)),
            totalReceived: U.toNumber(gd(this.dataCells.totalReceived)),
            remainingToReceive: U.toNumber(
                gd(this.dataCells.remainingToReceive)
            ),
        };
    }
}
class DashBuilder {
    ss: GoogleAppsScript.Spreadsheet.Spreadsheet;
    dashboardSheetName: string;
    dataCellMap: Record<DataKey, string>;
    headersCellMap: Record<HeaderKey, string>;
    isSheetNameValidProject: (name: string) => boolean;
    extractProjectIdFromSheetName: (name: string) => string;
    HEADER_ROW: number;
    dashSheet: GoogleAppsScript.Spreadsheet.Sheet;

    constructor(args: {
        dashboardSheetName: string;
        dataCellMap: Record<DataKey, string>;
        headersCellMap: Record<HeaderKey, string>;
        isSheetNameValidProject?: (name: string) => boolean;
        extractProjectIdFromSheetName?: (name: string) => string;
        headerRow?: number;
    }) {
        // bindings
        this.getSheetByName = this.getSheetByName.bind(this);
        this.getSheetNameData = this.getSheetNameData.bind(this);
        this.getSheetNameProjectDataMap =
            this.getSheetNameProjectDataMap.bind(this);
        this.getData = this.getData.bind(this);
        this.projectSheetDataToEntry = this.projectSheetDataToEntry.bind(this);
        this.addDataToSheet = this.addDataToSheet.bind(this);

        // args
        this.ss = SpreadsheetApp.getActiveSpreadsheet();
        this.dashboardSheetName = args.dashboardSheetName;
        this.dataCellMap = args.dataCellMap;
        this.headersCellMap = args.headersCellMap;
        this.isSheetNameValidProject =
            args.isSheetNameValidProject ?? DEFAULTS._isSheetNameValidProject;
        this.extractProjectIdFromSheetName =
            args.extractProjectIdFromSheetName ??
            DEFAULTS._extractProjectIdFromSheetName;
        this.HEADER_ROW = args.headerRow ?? 1;

        // load dash
        this.dashSheet = this.requireDashSheet();
    }

    private requireDashSheet() {
        const dashboardSheet = this.getSheetByName(this.dashboardSheetName);
        if (!dashboardSheet) {
            throw new Error('Dashboard Sheet Missing');
        }
        return dashboardSheet;
    }

    private getSheetByName(sheetName: string) {
        return this.ss.getSheetByName(sheetName);
    }

    private getSheetNameData(): Array<{
        sheetNameId: string;
        sheetName: string;
    }> {
        const sheets = this.ss.getSheets();
        const allSheetNames = sheets.map((s) => s.getName());
        return allSheetNames
            .filter(this.isSheetNameValidProject)
            .map((name) => ({
                sheetName: name,
                sheetNameId: this.extractProjectIdFromSheetName(name),
            }));
    }

    private getSheetNameProjectDataMap({
        sheetName,
    }: {
        sheetName: string;
    }): Record<string, ProjectSheetData | null> {
        const sheet = this.getSheetByName(sheetName);
        if (sheet == null) {
            console.warn(`Couldn't find sheet with name: ${sheetName}`);
            return { [sheetName]: null };
        }
        const s = new ProjectDataSheet({
            sheet,
            headersCellMap: this.headersCellMap,
            dataCellMap: this.dataCellMap,
        });
        const pdata = s.getProjectData();
        return { [sheetName]: pdata };
    }

    private getProjectHeaders(sheetName: string): Record<DataKey, string> {
        const sheet = this.getSheetByName(sheetName);
        if (sheet == null) {
            console.warn(`Couldn't find sheet with name: ${sheetName}`);
            return {
                projectId: 'Project ID',
                projectName: 'Project Name',
                totalFee: '',
                architectural: '',
                consultants: '',
                unallocated: '',
                allocatedPct: '',
                assignedPct: '',
                feeSpent: '',
                feeSpentPct: '',
                projectType: '',
                totalBilled: '',
                remainingToBill: '',
                billedPct: '',
                totalReceived: '',
                remainingToReceive: '',
            };
        }
        const s = new ProjectDataSheet({
            sheet,
            headersCellMap: this.headersCellMap,
            dataCellMap: this.dataCellMap,
        });
        return s.getProjectDataHeaders();
    }

    private projectSheetDataToEntry(
        nameDataMap: Record<string, ProjectSheetData | null>
    ): DashboardEntry {
        const pData = Object.values(nameDataMap)[0];
        const sheetName = Object.keys(nameDataMap)[0];
        const projectSheetData = pData ?? {
            projectName: null,
            projectId: null,
            totalFee: null,
            architectural: null,
            consultants: null,
            unallocated: null,
            allocatedPct: null,
            assignedPct: null,
            feeSpent: null,
            feeSpentPct: null,
            projectType: null,
            totalBilled: null,
            remainingToBill: null,
            billedPct: null,
            totalReceived: null,
            remainingToReceive: null,
        };

        const suppData: NullableRecord<SupplementalProjectData> = {
            sheetName,
            stage: null,
            pctComplete: null,
            projectedFee: null,
        };

        return {
            ...projectSheetData,
            ...suppData,
        };
    }

    getData() {
        const sheetNameData = this.getSheetNameData();
        const allSheetNamesProjectData = sheetNameData.map(
            this.getSheetNameProjectDataMap
        );
        const projectHeaders = this.getProjectHeaders(
            sheetNameData[0].sheetName
        );
        return {
            allSheetNamesProjectData,
            headerData: projectHeaders,
        };
    }

    addDataToSheet({
        headerData,
        allSheetNamesProjectData,
    }: {
        headerData: Record<DataKey, string>;
        allSheetNamesProjectData: Record<string, ProjectSheetData | null>[];
    }) {
        const lastRow = this.dashSheet.getLastRow();
        const lastColumn = this.dashSheet.getLastColumn();
        if (lastRow > this.HEADER_ROW) {
            const rowsToClear = lastRow - this.HEADER_ROW - 1;
            console.log({ lastRow, lastColumn, rowsToClear });
            this.dashSheet
                .getRange(this.HEADER_ROW, 1, rowsToClear, lastColumn)
                .clearContent();
        }

        const headers = Object.values(headerData);
        this.dashSheet
            .getRange(this.HEADER_ROW, 1, 1, headers.length)
            .setValues([headers]);

        const START_DATA_ROW = this.HEADER_ROW + 1;
        allSheetNamesProjectData.forEach((sheetNameProjectData, i) => {
            const [sheetName, pData] = Object.entries(sheetNameProjectData)[0];
            if (pData == null) {
                this.dashSheet.appendRow([
                    `null project data for sheet ${sheetName}`,
                ]);
            } else {
                const row = Object.keys(headerData).map(
                    (key) => pData[key as DataKey]
                );
                this.dashSheet
                    .getRange(START_DATA_ROW + i, 1, 1, row.length)
                    .setValues([row]);
            }
        });
    }

    format() {
        const lastColumn = this.dashSheet.getLastColumn();
        const maxRows = this.dashSheet.getMaxRows();
        const maxColumns = this.dashSheet.getMaxColumns();
        this.dashSheet
            .getRange(1, 1, maxRows, maxColumns)
            .setFontFamily('Roboto Mono');
        this.dashSheet.setFrozenColumns(2); // freeze columns
        this.dashSheet.setRowHeight(this.HEADER_ROW, 60); // header height
        this.dashSheet // header format
            .getRange(this.HEADER_ROW, 1, 1, lastColumn)
            .setFontWeight('bold')
            .setVerticalAlignment('middle');
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
}

const U = {
    extractNumberString(name: string): string | null {
        const regex = /\b\d{4}\b/; // Regex to match exactly 4 digits in a string
        const match = name.match(regex);
        return match ? match[0] : null; // Return the matched number or null if not found
    },

    toString(value: any): string {
        if (value === null || value === undefined) {
            return '';
        }
        return String(value);
    },

    toNumber(value: any): number {
        if (value === null || value === undefined) {
            return 0;
        }
        const num = Number(value);
        return isNaN(num) ? 0 : num;
    },

    isNextColumn(colA: string, colB: string): boolean {
        return U.colToNumber(colB) === U.colToNumber(colA) + 1;
    },

    colToNumber(col: string) {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
        }
        return num;
    },

    isValidColumnRange(start: string, end: string) {
        return U.colToNumber(start) <= U.colToNumber(end);
    },
    /** finds contiguous horizontal ranges */
    findContiguousRanges(cellMap: Record<string, string>): string[] {
        const cellRefs = Object.values(cellMap);
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

    getIndexFromRanges(ranges: string[], cell: string) {
        const expandedRanges: string[] = [];
        for (const range of ranges) {
            if (range.includes(':')) {
                const [start, end] = range.split(':');
                const [startCol, startRow] = start
                    .match(/([A-Z]+)(\d+)/)!
                    .slice(1);
                const [endCol] = end.match(/([A-Z]+)(\d+)/)!.slice(1);

                if (!U.isValidColumnRange(startCol, endCol))
                    throw new Error(
                        `Fatal Error: Invalid column range ${startCol}:${endCol}`
                    );

                // fills in columns
                let curCol = startCol;
                const innerColumns: string[] = [];
                while (
                    U.isValidColumnRange(curCol, endCol) &&
                    curCol !== endCol
                ) {
                    curCol = U.getNextColumn(curCol);

                    if (curCol === endCol) break;
                    innerColumns.push(curCol);
                }
                const inner = innerColumns.map((c) => `${c}${startRow}`);
                expandedRanges.push(...[start, ...inner, end]);
            } else {
                expandedRanges.push(range);
            }
        }
        return expandedRanges.indexOf(cell);
    },

    getNextColumn(input: string): string {
        if (input.length === 0) {
            return 'A';
        }

        const chars = input.split('');
        let carry = true;

        // Start from the last character and work backwards
        for (let i = chars.length - 1; i >= 0 && carry; i--) {
            if (chars[i] === 'Z') {
                chars[i] = 'A'; // Reset to 'A' and carry over
            } else {
                chars[i] = String.fromCharCode(chars[i].charCodeAt(0) + 1);
                carry = false;
            }
        }

        if (carry) {
            chars.unshift('A');
        }

        return chars.join('');
    },
};

const DEFAULTS = {
    _isSheetNameValidProject(name: string) {
        // by default, a sheet name is a valid project id if it contains a substring that is a number
        const numberSubstring = U.extractNumberString(name);
        return numberSubstring != null && !isNaN(Number(numberSubstring));
    },

    _extractProjectIdFromSheetName(name: string) {
        // by default, return the first matched number
        const id = U.extractNumberString(name);
        if (id == null)
            throw new Error(
                `Could not extract project id from sheet name ${name}`
            );
        return id;
    },
};
