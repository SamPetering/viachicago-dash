function main() {
	const dashBuilder = new DashBuilder()
	const ids = dashBuilder.getProjectIds()
	console.log('found ids', ids.join(', '))
	dashBuilder.addProjectIdsToDash(ids)
}

type DashboardEntry = {
	projectId: string;
	projectType: string | null;
	projectName: string;
	stage: string;
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
}

const DASHBOARD_SHEET_NAME = "Dashboard"

function extractNumberString(name: string): string | null {
	const regex = /\b\d{4}\b/; // Regex to match exactly 4 digits in a string
	const match = name.match(regex);
	return match ? match[0] : null; // Return the matched number or null if not found
}

function _isSheetNameValidProject(name: string) {
	// by default, a sheet name is a valid project id
	// if it contains a substring that is a number
	const numberSubstring = extractNumberString(name)
	return numberSubstring != null && !isNaN(Number(numberSubstring))
}

function _extractProjectIdFromSheetName(name: string) {
	// by default, return the first matched number
	const id = extractNumberString(name)
	if (id == null) throw new Error(`Could not extract project id from sheet name ${name}`)
	return id
}

class DashBuilder {
	ss: GoogleAppsScript.Spreadsheet.Spreadsheet
	isSheetNameValidProject: (name: string) => boolean
	extractProjectIdFromSheetName: (name: string) => string

	constructor(args?: {
		isSheetNameValidProject?: (name: string) => boolean;
		extractProjectIdFromSheetName?: (name: string) => string;
	}) {
		this.ss = SpreadsheetApp.getActiveSpreadsheet();
		this.isSheetNameValidProject = args?.isSheetNameValidProject ?? _isSheetNameValidProject;
		this.extractProjectIdFromSheetName = args?.extractProjectIdFromSheetName ?? _extractProjectIdFromSheetName;
	}

	private getSheetTitles() {
		const sheets = this.ss.getSheets();
		const titles = sheets.map(s => s.getName());
		return titles;
	}

	getProjectIds() {
		return this.getSheetTitles()
			.filter(this.isSheetNameValidProject)
			.map(this.extractProjectIdFromSheetName)
	}

	addProjectIdsToDash(ids: string[]) {
		const dashSheet = this.ss.getSheetByName(DASHBOARD_SHEET_NAME)
		if (!dashSheet) throw new Error(`Sheet ${DASHBOARD_SHEET_NAME} not found`)
		dashSheet.clear();
		dashSheet.getRange('A1').setValue('Project Code')

		if (ids.length > 0) {
			const range = dashSheet.getRange(2, 1, ids.length, 1)
			range.setValues(ids.map((id) => [id]))
		}
	}	
}
