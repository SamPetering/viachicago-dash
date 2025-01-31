type InvoicingData = {
    totalBilled: number;
    remainingToBill: number;
    pctBilled: number;
    totalReceived: number;
    remainingToReceive: number;
};

type SpendData = {
    unallocated: number;
    pctAllocated: number;
    pctAssigned: number;
    feeSpent: number;
    pctFeeSpent: number;
};

type FeeData = {
    totalFee: number;
    architectural: number;
    consultants: number;
};

type ProjectSheetData = {
    projectId: string;
} & FeeData &
    SpendData &
    InvoicingData;

type DataKey = keyof ProjectSheetData;
type HeaderKey = Exclude<DataKey, 'projectId'>; // project id has no header

type DashboardEntry = {
    sheetName: string;
    projectType: string | null;
    projectName: string;
    stage: string;
    pctComplete: number;
    projectedFee: number;
} & ProjectSheetData;

function main() {
    const dashBuilder = new DashBuilder({
        dashboardSheetName: 'Dashboard',
        dataCellMap: {
            projectId: 'E2',
            totalFee: 'G3',
            architectural: 'H3',
            consultants: 'I3',
            unallocated: 'K3',
            pctAllocated: 'L3',
            pctAssigned: 'M3',
            feeSpent: 'N3',
            pctFeeSpent: 'O3',
            totalBilled: 'S3',
            remainingToBill: 'T3',
            pctBilled: 'U3',
            totalReceived: 'V3',
            remainingToReceive: 'W3',
        },
        headersCellMap: {
            totalFee: 'G2',
            architectural: 'H2',
            consultants: 'I2',
            unallocated: 'K2',
            pctAllocated: 'L2',
            pctAssigned: 'M2',
            feeSpent: 'N2',
            pctFeeSpent: 'O2',
            totalBilled: 'S2',
            remainingToBill: 'T2',
            pctBilled: 'U2',
            totalReceived: 'V2',
            remainingToReceive: 'W2',
        },
    });

    dashBuilder.getProjects();
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

        this.getCellValue = this.getCellValue.bind(this);
        this.getStringCellValue = this.getStringCellValue.bind(this);
        this.getNumberCellValue = this.getNumberCellValue.bind(this);
        this.getInvoicingData = this.getInvoicingData.bind(this);
        this.getSpendData = this.getSpendData.bind(this);
        this.getFeeData = this.getFeeData.bind(this);
        this.getProjectId = this.getProjectId.bind(this);
        this.getProjectData = this.getProjectData.bind(this);
        this.getRangesWithValues = this.getRangesWithValues.bind(this);
    }
    getCellValue(a1Notation: string) {
        return this.sheet.getRange(a1Notation).getValue();
    }
    getStringCellValue(a1Notation: string): string {
        return U.toString(this.getCellValue(a1Notation));
    }
    getNumberCellValue(a1Notation: string): number {
        return U.toNumber(this.getCellValue(a1Notation));
    }
    private getInvoicingData(): InvoicingData {
        return {
            totalBilled: this.getNumberCellValue(this.dataCells.projectId),
            remainingToBill: this.getNumberCellValue(
                this.dataCells.remainingToBill
            ),
            pctBilled: this.getNumberCellValue(this.dataCells.pctBilled),
            totalReceived: this.getNumberCellValue(
                this.dataCells.totalReceived
            ),
            remainingToReceive: this.getNumberCellValue(
                this.dataCells.remainingToReceive
            ),
        };
    }
    private getSpendData(): SpendData {
        return {
            unallocated: this.getNumberCellValue(this.dataCells.unallocated),
            pctAllocated: this.getNumberCellValue(this.dataCells.pctAllocated),
            pctAssigned: this.getNumberCellValue(this.dataCells.pctAssigned),
            feeSpent: this.getNumberCellValue(this.dataCells.feeSpent),
            pctFeeSpent: this.getNumberCellValue(this.dataCells.pctFeeSpent),
        };
    }
    private getFeeData(): FeeData {
        return {
            totalFee: this.getNumberCellValue(this.dataCells.totalFee),
            architectural: this.getNumberCellValue(
                this.dataCells.architectural
            ),
            consultants: this.getNumberCellValue(this.dataCells.consultants),
        };
    }
    private getProjectId(): string {
        return this.getStringCellValue(this.dataCells.projectId);
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

    getProjectHeaders(): Record<HeaderKey, string> {
        const headerRanges = U.findContiguousRanges(this.headerCells);
        const rangesHeadersMap = this.getRangesWithValues(headerRanges);
        const headersValuesFlat = Object.values(rangesHeadersMap).flat();
        const gh = (cell: string) =>
            headersValuesFlat[U.getIndexFromRanges(headerRanges, cell)];

        return {
            totalFee: U.toString(gh(this.headerCells.totalFee)),
            architectural: U.toString(gh(this.headerCells.architectural)),
            consultants: U.toString(gh(this.headerCells.consultants)),
            unallocated: U.toString(gh(this.headerCells.unallocated)),
            pctAllocated: U.toString(gh(this.headerCells.pctAllocated)),
            pctAssigned: U.toString(gh(this.headerCells.pctAssigned)),
            feeSpent: U.toString(gh(this.headerCells.feeSpent)),
            pctFeeSpent: U.toString(gh(this.headerCells.pctFeeSpent)),
            totalBilled: U.toString(gh(this.headerCells.totalBilled)),
            remainingToBill: U.toString(gh(this.headerCells.remainingToBill)),
            pctBilled: U.toString(gh(this.headerCells.pctBilled)),
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
            projectId: U.toString(gd(this.dataCells.projectId)),
            totalFee: U.toNumber(gd(this.dataCells.totalFee)),
            architectural: U.toNumber(gd(this.dataCells.architectural)),
            consultants: U.toNumber(gd(this.dataCells.consultants)),
            unallocated: U.toNumber(gd(this.dataCells.unallocated)),
            pctAllocated: U.toNumber(gd(this.dataCells.pctAllocated)),
            pctAssigned: U.toNumber(gd(this.dataCells.pctAssigned)),
            feeSpent: U.toNumber(gd(this.dataCells.feeSpent)),
            pctFeeSpent: U.toNumber(gd(this.dataCells.pctFeeSpent)),
            totalBilled: U.toNumber(gd(this.dataCells.totalBilled)),
            remainingToBill: U.toNumber(gd(this.dataCells.remainingToBill)),
            pctBilled: U.toNumber(gd(this.dataCells.pctBilled)),
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

    constructor(args: {
        dashboardSheetName: string;
        dataCellMap: Record<DataKey, string>;
        headersCellMap: Record<HeaderKey, string>;
        isSheetNameValidProject?: (name: string) => boolean;
        extractProjectIdFromSheetName?: (name: string) => string;
    }) {
        this.ss = SpreadsheetApp.getActiveSpreadsheet();
        this.dashboardSheetName = args.dashboardSheetName;
        this.dataCellMap = args.dataCellMap;
        this.headersCellMap = args.headersCellMap;
        this.isSheetNameValidProject =
            args.isSheetNameValidProject ?? DEFAULTS._isSheetNameValidProject;
        this.extractProjectIdFromSheetName =
            args.extractProjectIdFromSheetName ??
            DEFAULTS._extractProjectIdFromSheetName;

        this.getSheetNames = this.getSheetNames.bind(this);
        this.getSheetByName = this.getSheetByName.bind(this);
        this.getProjectIdsWithSheetNames =
            this.getProjectIdsWithSheetNames.bind(this);
        this.getProjectSheetData = this.getProjectSheetData.bind(this);
        this.getProjects = this.getProjects.bind(this);
        this.addProjectIdsToDash = this.addProjectIdsToDash.bind(this);
    }

    private getSheetNames() {
        const sheets = this.ss.getSheets();
        const titles = sheets.map((s) => s.getName());
        return titles;
    }
    private getSheetByName(sheetName: string) {
        return this.ss.getSheetByName(sheetName);
    }
    private getProjectIdsWithSheetNames(): Array<{
        sheetNameId: string;
        sheetName: string;
    }> {
        return this.getSheetNames()
            .filter(this.isSheetNameValidProject)
            .map((name) => ({
                sheetName: name,
                sheetNameId: this.extractProjectIdFromSheetName(name),
            }));
    }
    private getProjectSheetData({
        sheetName,
    }: {
        sheetNameId: string;
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
    private getProjectHeaders(
        sheetName: string
    ): Record<HeaderKey, string> | null {
        const sheet = this.getSheetByName(sheetName);
        if (sheet == null) {
            console.warn(`Couldn't find sheet with name: ${sheetName}`);
            return null;
        }
        const s = new ProjectDataSheet({
            sheet,
            headersCellMap: this.headersCellMap,
            dataCellMap: this.dataCellMap,
        });
        return s.getProjectHeaders();
    }
    getProjects() {
        const idsWithNames = this.getProjectIdsWithSheetNames();
        const allProjectSheetData = idsWithNames
            .slice(0, 10)
            .map(this.getProjectSheetData);

        const projectHeaders = this.getProjectHeaders(
            idsWithNames[0].sheetName
        );

        const headerRow = Object.values(projectHeaders);
        const dataRows = allProjectSheetData.map((sheetNameToPsaMap) => {
            const pdata = Object.values(sheetNameToPsaMap)[0];
            return Object.values(pdata);
        });
        console.log(headerRow);
        console.log(dataRows);
        return allProjectSheetData;
    }
    addProjectIdsToDash(ids: string[]) {
        const dashSheet = this.ss.getSheetByName(this.dashboardSheetName);
        if (!dashSheet)
            throw new Error(`Sheet ${this.dashboardSheetName} not found`);
        dashSheet.clear();
        dashSheet.getRange('A1').setValue('Project Code');

        if (ids.length > 0) {
            const range = dashSheet.getRange(2, 1, ids.length, 1);
            range.setValues(ids.map((id) => [id]));
        }
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
    // Function to convert column letters to a column number
    colToNumber(col: string) {
        let num = 0;
        for (let i = 0; i < col.length; i++) {
            num = num * 26 + (col.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
        }
        return num;
    },
    // Function to check if a column range is valid
    isValidColumnRange(start: string, end: string) {
        return U.colToNumber(start) <= U.colToNumber(end);
    },

    /**
     * finds contiguous horizontal ranges
     */
    findContiguousRanges(cellMap: Record<string, string>): string[] {
        const cellRefs = Object.values(cellMap);
        const sortedRefs = cellRefs.sort((a, b) => {
            const [colA, rowA] = a.match(/([A-Z]+)(\d+)/).slice(1);
            const [colB, rowB] = b.match(/([A-Z]+)(\d+)/).slice(1);
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
        const expandedRanges = [];
        for (const range of ranges) {
            if (range.includes(':')) {
                const [start, end] = range.split(':');
                const [startCol, startRow] = start
                    .match(/([A-Z]+)(\d+)/)
                    .slice(1);
                const [endCol] = end.match(/([A-Z]+)(\d+)/).slice(1);

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
        // Check if the input is empty
        if (input.length === 0) {
            return 'A';
        }

        // Convert the input to an array of characters
        const chars = input.split('');
        let carry = true;

        // Start from the last character and work backwards
        for (let i = chars.length - 1; i >= 0 && carry; i--) {
            if (chars[i] === 'Z') {
                chars[i] = 'A'; // Reset to 'A' and carry over
            } else {
                chars[i] = String.fromCharCode(chars[i].charCodeAt(0) + 1);
                carry = false; // No more carry needed
            }
        }

        // If we still have a carry, we need to add a new 'A' at the start
        if (carry) {
            chars.unshift('A');
        }

        return chars.join('');
    },
};

const DEFAULTS = {
    _isSheetNameValidProject(name: string) {
        // by default, a sheet name is a valid project id
        // if it contains a substring that is a number
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
