function main() {
  const dashBuilder = new DashBuilder({ dashboardSheetName: "Dashboard" });
  dashBuilder.getProjects();
}

type ProjectSheetData = {
  sheetName: string;
  projectId: string;
  pctComplete: number;
  projectedFee: number;
  fee: number;
  unallocated: number;
  pctUnallocated: number;
  pctAssigned: number;
  spentFee: number;
  totalBilled: number;
  remainingToBill: number;
  pctBilled: number;
  totalReceived: number;
  remainingToReceive: number;
};

type DashboardEntry = {
  projectType: string | null;
  projectName: string;
  stage: string;
} & Omit<ProjectSheetData, "sheetName">;

class DashBuilder {
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet;
  dashboardSheetName: string;
  isSheetNameValidProject: (name: string) => boolean;
  extractProjectIdFromSheetName: (name: string) => string;

  constructor(args: {
    dashboardSheetName: string;
    isSheetNameValidProject?: (name: string) => boolean;
    extractProjectIdFromSheetName?: (name: string) => string;
  }) {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.dashboardSheetName = args.dashboardSheetName;
    this.isSheetNameValidProject =
      args.isSheetNameValidProject ?? DEFAULTS._isSheetNameValidProject;
    this.extractProjectIdFromSheetName =
      args.extractProjectIdFromSheetName ??
      DEFAULTS._extractProjectIdFromSheetName;
  }

  private getSheetTitles() {
    const sheets = this.ss.getSheets();
    const titles = sheets.map((s) => s.getName());
    return titles;
  }

  private getSheetByName(name: string) {
    const sheets = this.ss.getSheets();
    return sheets.find((s) => s.getName() === name);
  }

  private getSheetByProjectId(id: string) {
    const sheets = this.ss.getSheets();
    return sheets.find((s) => s.getName().includes(id));
  }

  private getProjectSheetData(id: string): ProjectSheetData {
    const sheet = this.getSheetByProjectId(id);
    throw new Error("No implemented");
  }

  private getProjectIdsWithSheetNames(): Array<{
    projectId: string;
    sheetName: string;
  }> {
    return this.getSheetTitles()
      .filter(this.isSheetNameValidProject)
      .map((name) => ({
        sheetName: name,
        projectId: this.extractProjectIdFromSheetName(name),
      }));
  }

  getProjects() {
    const idsWithNames = this.getProjectIdsWithSheetNames();
    // TODO: we have ids and sheet names, for each id with name, get the project sheet data
    console.log(idsWithNames);
  }

  private getProjectIds() {
    return this.getSheetTitles()
      .filter(this.isSheetNameValidProject)
      .map(this.extractProjectIdFromSheetName);
  }

  addProjectIdsToDash(ids: string[]) {
    const dashSheet = this.ss.getSheetByName(this.dashboardSheetName);
    if (!dashSheet)
      throw new Error(`Sheet ${this.dashboardSheetName} not found`);
    dashSheet.clear();
    dashSheet.getRange("A1").setValue("Project Code");

    if (ids.length > 0) {
      const range = dashSheet.getRange(2, 1, ids.length, 1);
      range.setValues(ids.map((id) => [id]));
    }
  }
}

const UTILS = {
  extractNumberString(name: string): string | null {
    const regex = /\b\d{4}\b/; // Regex to match exactly 4 digits in a string
    const match = name.match(regex);
    return match ? match[0] : null; // Return the matched number or null if not found
  },
};

const DEFAULTS = {
  _isSheetNameValidProject(name: string) {
    // by default, a sheet name is a valid project id
    // if it contains a substring that is a number
    const numberSubstring = UTILS.extractNumberString(name);
    return numberSubstring != null && !isNaN(Number(numberSubstring));
  },
  _extractProjectIdFromSheetName(name: string) {
    // by default, return the first matched number
    const id = UTILS.extractNumberString(name);
    if (id == null)
      throw new Error(`Could not extract project id from sheet name ${name}`);
    return id;
  },
};
