/**
 * Clinical Audit System - Google Sheets Automation
 * Fixed version with proper sheet clearing
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const CONFIG = {
	sheets: {
		data: "Data",
		validator: "Validator",
		leadReview: "Lead Review",
		config: "Config",
		auditLog: "Audit Log",
		dashboard: "Dashboard",
	},
	colors: {
		header: "#4285F4",
		success: "#34A853",
		warning: "#FBBC04",
		error: "#EA4335",
		neutral: "#F8F9FA",
	},
};

// ============================================================================
// MAIN MENU AND INITIALIZATION
// ============================================================================

/**
 * Creates custom menu on spreadsheet open
 */
function onOpen() {
	const ui = SpreadsheetApp.getUi();
	ui.createMenu("🔧 Clinical Audit System")
		.addItem("📊 Initialize All Sheets", "initializeAllSheets")
		.addItem("🔄 Refresh Validator", "refreshValidator")
		.addItem("📈 Update Dashboard", "updateDashboard")
		.addSeparator()
		.addItem("📥 Import CSV Data", "showImportDialog")
		.addItem("📤 Export Audit Report", "exportAuditReport")
		.addSeparator()
		.addItem("🔒 Protect Sheets", "protectCriticalRanges")
		.addItem("⚙️ System Settings", "showSettingsDialog")
		.addItem("❓ Help", "showHelp")
		.addToUi();
}

/**
 * Initialize all sheets with proper structure
 */
function initializeAllSheets() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	try {
		// Create sheets if they don't exist
		ensureSheetExists(CONFIG.sheets.data);
		ensureSheetExists(CONFIG.sheets.validator);
		ensureSheetExists(CONFIG.sheets.leadReview);
		ensureSheetExists(CONFIG.sheets.config);
		ensureSheetExists(CONFIG.sheets.auditLog);
		ensureSheetExists(CONFIG.sheets.dashboard);

		// Setup each sheet (with proper clearing)
		setupDataSheet();
		setupValidatorSheet();
		setupLeadReviewSheet();
		setupConfigSheet();
		setupAuditLogSheet();
		setupDashboardSheet();

		// Log initialization
		logAction("System Initialized", "", "All sheets configured");

		SpreadsheetApp.getUi().alert(
			"✅ System Initialized",
			"All sheets have been configured successfully.",
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	} catch (error) {
		SpreadsheetApp.getUi().alert(
			"❌ Initialization Error",
			"Error: " + error.toString() + "\n\nPlease try again or contact support.",
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	}
}

// ============================================================================
// SHEET SETUP FUNCTIONS WITH PROPER CLEARING
// ============================================================================

/**
 * Clear sheet completely including merged cells
 */
function clearSheetCompletely(sheet) {
	// First, unmerge all cells
	const lastRow = sheet.getMaxRows();
	const lastCol = sheet.getMaxColumns();

	if (lastRow > 0 && lastCol > 0) {
		try {
			// Get all merges and break them
			const range = sheet.getRange(1, 1, lastRow, lastCol);
			range.breakApart();
		} catch (e) {
			// Ignore if no merges exist
		}

		// Clear all content and formatting
		sheet.clear();

		// Remove all data validations
		sheet.getRange(1, 1, lastRow, lastCol).clearDataValidations();

		// Remove all conditional formatting
		sheet.clearConditionalFormatRules();
	}
}

/**
 * Ensure a sheet exists
 */
function ensureSheetExists(sheetName) {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName(sheetName);

	if (!sheet) {
		sheet = ss.insertSheet(sheetName);
	}

	return sheet;
}

/**
 * Setup Data sheet structure
 */
function setupDataSheet() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.data
	);

	// Clear sheet completely first
	clearSheetCompletely(sheet);

	// Set headers
	const headers = [
		[
			"claim_id",
			"note_preview",
			"full_note_text",
			"processing_timestamp",
			"processing_status",
			"diagnosis_codes",
			"cpt_codes",
			"medical_necessity",
			"treatment_goals",
			"interventions",
			"progress_noted",
			"plan_documented",
			"therapist_signature",
			"date_of_service",
			"session_duration",
			"compliance_issues",
			"addendum_needed",
			"addendum_reason",
			"validator_decision",
			"validator_notes",
			"validation_timestamp",
			"validator_email",
			"lead_review_status",
			"final_addendum",
			"lead_reviewer",
			"lead_review_timestamp",
			"ai_confidence_score",
		],
	];

	sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

	// Format headers
	sheet
		.getRange(1, 1, 1, headers[0].length)
		.setBackground(CONFIG.colors.header)
		.setFontColor("#FFFFFF")
		.setFontWeight("bold")
		.setBorder(true, true, true, true, true, true);

	// Freeze header row
	sheet.setFrozenRows(1);

	// Set column widths
	sheet.setColumnWidth(1, 120); // claim_id
	sheet.setColumnWidth(2, 200); // note_preview
	sheet.setColumnWidth(3, 100); // full_note_text (will be hidden)

	// Hide full text column
	sheet.hideColumns(3);

	// Auto-resize other columns (but limit the number to avoid timeout)
	for (let i = 4; i <= Math.min(headers[0].length, 10); i++) {
		sheet.autoResizeColumn(i);
	}
}

/**
 * Setup Validator sheet with interactive UI
 */
function setupValidatorSheet() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.validator
	);

	// Clear sheet completely first
	clearSheetCompletely(sheet);

	// Title section
	sheet.getRange("A1:L1").merge();
	sheet
		.getRange("A1")
		.setValue("Clinical Note Validator Dashboard")
		.setFontSize(18)
		.setFontWeight("bold")
		.setHorizontalAlignment("center")
		.setBackground(CONFIG.colors.header)
		.setFontColor("#FFFFFF");

	// Instructions
	sheet.getRange("A2:L2").merge();
	sheet
		.getRange("A2")
		.setValue(
			"Select a claim ID below to begin validation. Review the note and complete the checklist."
		)
		.setFontSize(10)
		.setFontStyle("italic")
		.setHorizontalAlignment("center")
		.setBackground(CONFIG.colors.neutral);

	// Claim selector section
	sheet
		.getRange("A4")
		.setValue("Select Claim ID:")
		.setFontWeight("bold")
		.setFontSize(12);

	sheet
		.getRange("C4")
		.setBackground("#FFF3CD")
		.setBorder(true, true, true, true, false, false)
		.setFontWeight("bold");

	// Add data validation for claim ID (if data exists)
	const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.data
	);
	const lastRow = dataSheet.getLastRow();
	if (lastRow > 1) {
		const validationRange = dataSheet.getRange(2, 1, lastRow - 1, 1);
		const validation = SpreadsheetApp.newDataValidation()
			.setAllowInvalid(false)
			.requireValueInRange(validationRange, true)
			.build();
		sheet.getRange("C4").setDataValidation(validation);
	}

	// Processing info section
	sheet.getRange("E4").setValue("Processing Status:");
	sheet.getRange("G4").setFormula('=IFERROR(VLOOKUP(C4,Data!A:E,5,FALSE),"")');

	sheet.getRange("I4").setValue("AI Confidence:");
	sheet
		.getRange("K4")
		.setFormula('=IFERROR(VLOOKUP(C4,Data!A:AB,28,FALSE)&"%","")');

	// Main content area
	sheet.getRange("A6:L6").merge();
	sheet
		.getRange("A6")
		.setValue("Clinical Note Content")
		.setFontWeight("bold")
		.setFontSize(14)
		.setBackground("#E3F2FD");

	// Note display area
	sheet.getRange("A7:E25").merge();
	sheet
		.getRange("A7")
		.setBorder(true, true, true, true, false, false)
		.setWrap(true)
		.setVerticalAlignment("top")
		.setBackground("#FFFFFF")
		.setFontSize(11)
		.setFormula(
			'=IFERROR(VLOOKUP(C4,Data!A:C,3,FALSE),"Select a claim ID to view note content")'
		);

	// Validation checklist section
	sheet.getRange("G6:L6").merge();
	sheet
		.getRange("G6")
		.setValue("Validation Checklist")
		.setFontWeight("bold")
		.setFontSize(14)
		.setBackground("#E8F5E9");

	// Checklist headers
	const checklistHeaders = [["Item", "AI Value", "Status", "Notes"]];
	sheet
		.getRange("G7:J7")
		.setValues(checklistHeaders)
		.setFontWeight("bold")
		.setBackground("#F5F5F5");

	// Checklist items
	const checklistItems = [
		["Medical Necessity", '=IFERROR(VLOOKUP($C$4,Data!$A:$AB,8,FALSE),"")'],
		["Diagnosis Codes", '=IFERROR(VLOOKUP($C$4,Data!$A:$AB,6,FALSE),"")'],
		["CPT Codes", '=IFERROR(VLOOKUP($C$4,Data!$A:$AB,7,FALSE),"")'],
		[
			"Treatment Goals",
			'=IFERROR(LEFT(VLOOKUP($C$4,Data!$A:$AB,9,FALSE),50),"")',
		],
		[
			"Interventions",
			'=IFERROR(LEFT(VLOOKUP($C$4,Data!$A:$AB,10,FALSE),50),"")',
		],
		["Progress Noted", '=IFERROR(VLOOKUP($C$4,Data!$A:$AB,11,FALSE),"")'],
		["Plan Documented", '=IFERROR(VLOOKUP($C$4,Data!$A:$AB,12,FALSE),"")'],
		["Therapist Signature", '=IFERROR(VLOOKUP($C$4,Data!$A:$AB,13,FALSE),"")'],
		["Date of Service", '=IFERROR(VLOOKUP($C$4,Data!$A:$AB,14,FALSE),"")'],
		["Session Duration", '=IFERROR(VLOOKUP($C$4,Data!$A:$AB,15,FALSE),"")'],
	];

	// Add checklist items
	for (let i = 0; i < checklistItems.length; i++) {
		const row = 8 + i;
		sheet.getRange(row, 7).setValue(checklistItems[i][0]);
		sheet.getRange(row, 8).setFormula(checklistItems[i][1]);

		// Status dropdown
		const statusValidation = SpreadsheetApp.newDataValidation()
			.requireValueInList(["✓ Pass", "✗ Fail", "⚠ Review", "N/A"], true)
			.build();
		sheet
			.getRange(row, 9)
			.setDataValidation(statusValidation)
			.setBackground("#FFF9E6");

		// Notes column
		sheet.getRange(row, 10).setBackground("#F9F9F9");
	}

	// Compliance issues section
	sheet.getRange("G19").setValue("Compliance Issues:").setFontWeight("bold");
	sheet.getRange("G20:L20").merge();
	sheet
		.getRange("G20")
		.setFormula('=IFERROR(VLOOKUP(C4,Data!A:AB,16,FALSE),"")')
		.setBorder(true, true, true, true, false, false)
		.setWrap(true)
		.setBackground("#FFF3CD");

	// AI Recommendation
	sheet.getRange("G22").setValue("AI Recommendation:").setFontWeight("bold");
	sheet.getRange("G23:L23").merge();
	sheet
		.getRange("G23")
		.setFormula(
			'=IF(VLOOKUP(C4,Data!A:AB,17,FALSE)="Yes","⚠ Addendum Recommended: "&VLOOKUP(C4,Data!A:AB,18,FALSE),"✓ No addendum needed")'
		)
		.setBorder(true, true, true, true, false, false)
		.setWrap(true);

	// Validator decision section
	sheet.getRange("A27:L27").merge();
	sheet
		.getRange("A27")
		.setValue("Validator Decision")
		.setFontWeight("bold")
		.setFontSize(14)
		.setBackground("#FFF3CD");

	sheet.getRange("A29").setValue("Decision:").setFontWeight("bold");

	const decisionValidation = SpreadsheetApp.newDataValidation()
		.requireValueInList(
			["Approved", "Needs Addendum", "Rejected", "Pending Review"],
			true
		)
		.build();
	sheet
		.getRange("C29")
		.setDataValidation(decisionValidation)
		.setBackground("#D4EDDA")
		.setFontWeight("bold");

	sheet.getRange("A31").setValue("Validator Notes:").setFontWeight("bold");
	sheet.getRange("A32:L34").merge();
	sheet
		.getRange("A32")
		.setBorder(true, true, true, true, false, false)
		.setWrap(true)
		.setVerticalAlignment("top");

	// Action buttons section
	sheet.getRange("A36").setValue("Actions:").setFontWeight("bold");

	// Submit button (checkbox)
	sheet.getRange("C36").insertCheckboxes();
	sheet.getRange("D36").setValue("Submit Validation").setFontWeight("bold");

	// Clear button
	sheet.getRange("F36").insertCheckboxes();
	sheet.getRange("G36").setValue("Clear Form").setFontWeight("bold");

	// Send to Lead button
	sheet.getRange("I36").insertCheckboxes();
	sheet.getRange("J36").setValue("Send to Lead Review").setFontWeight("bold");

	// Set reasonable column widths
	sheet.setColumnWidths(1, 12, 90);
}

/**
 * Setup Lead Review sheet
 */
function setupLeadReviewSheet() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.leadReview
	);

	// Clear sheet completely first
	clearSheetCompletely(sheet);

	// Title
	sheet.getRange("A1:M1").merge();
	sheet
		.getRange("A1")
		.setValue("Lead Review Dashboard")
		.setFontSize(18)
		.setFontWeight("bold")
		.setHorizontalAlignment("center")
		.setBackground("#DC3545")
		.setFontColor("#FFFFFF");

	// Action buttons row
	sheet.getRange("A2:D2").merge();
	sheet
		.getRange("A2")
		.setValue("Select items and choose action →")
		.setFontStyle("italic");

	sheet
		.getRange("E2")
		.setValue("Approve Selected")
		.setBackground("#28A745")
		.setFontColor("#FFFFFF")
		.setFontWeight("bold")
		.setHorizontalAlignment("center");

	sheet
		.getRange("F2")
		.setValue("Reject Selected")
		.setBackground("#DC3545")
		.setFontColor("#FFFFFF")
		.setFontWeight("bold")
		.setHorizontalAlignment("center");

	sheet
		.getRange("G2")
		.setValue("Generate Report")
		.setBackground("#007BFF")
		.setFontColor("#FFFFFF")
		.setFontWeight("bold")
		.setHorizontalAlignment("center");

	// Headers
	const headers = [
		[
			"Select",
			"Claim ID",
			"Validator",
			"Decision",
			"Validator Notes",
			"Compliance Issues",
			"AI Recommendation",
			"Draft Addendum",
			"Final Addendum",
			"Lead Status",
			"Lead Notes",
			"Reviewed By",
			"Review Date",
		],
	];

	sheet
		.getRange(3, 1, 1, headers[0].length)
		.setValues(headers)
		.setFontWeight("bold")
		.setBackground("#F8D7DA");

	// Set column widths
	sheet.setColumnWidth(1, 50); // Select
	sheet.setColumnWidth(2, 100); // Claim ID
	sheet.setColumnWidth(3, 150); // Validator
	sheet.setColumnWidth(4, 120); // Decision
	sheet.setColumnWidth(5, 200); // Validator Notes
	sheet.setColumnWidth(6, 200); // Compliance Issues
	sheet.setColumnWidth(7, 200); // AI Recommendation
	sheet.setColumnWidth(8, 300); // Draft Addendum
	sheet.setColumnWidth(9, 300); // Final Addendum
	sheet.setColumnWidth(10, 120); // Lead Status
	sheet.setColumnWidth(11, 200); // Lead Notes

	// Freeze header
	sheet.setFrozenRows(3);

	// Add checkboxes for selection (in first data rows)
	if (sheet.getMaxRows() > 3) {
		sheet
			.getRange(4, 1, Math.min(100, sheet.getMaxRows() - 3), 1)
			.insertCheckboxes();
	}

	// Status dropdown validation
	const statusValidation = SpreadsheetApp.newDataValidation()
		.requireValueInList(
			["Pending", "Approved", "Rejected", "Needs Revision"],
			true
		)
		.build();

	if (sheet.getMaxRows() > 3) {
		sheet
			.getRange(4, 10, Math.min(100, sheet.getMaxRows() - 3), 1)
			.setDataValidation(statusValidation);
	}
}

/**
 * Setup Config sheet
 */
function setupConfigSheet() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.config
	);

	// Clear sheet completely first
	clearSheetCompletely(sheet);

	// Title
	sheet.getRange("A1:B1").merge();
	sheet
		.getRange("A1")
		.setValue("System Configuration")
		.setFontSize(16)
		.setFontWeight("bold")
		.setBackground("#6C757D")
		.setFontColor("#FFFFFF");

	// Configuration settings
	const settings = [
		["Setting", "Value"],
		["Lead Reviewer Email", "lead@example.com"],
		["CC Email List", "manager@example.com"],
		["Send Email Notifications", "TRUE"],
		["Auto-Generate Addendum", "TRUE"],
		["Require Lead Approval", "TRUE"],
		["Days Until Archive", "30"],
		["Max Validation Attempts", "3"],
		["Default Session Duration", "60"],
		["Time Zone", "America/New_York"],
		["Date Format", "MM/dd/yyyy"],
		["Export Format", "PDF"],
		["Backup Frequency (days)", "7"],
		["System Version", "1.0.0"],
		["Last Updated", new Date().toLocaleString()],
	];

	sheet.getRange(3, 1, settings.length, 2).setValues(settings);

	// Format headers
	sheet.getRange(3, 1, 1, 2).setFontWeight("bold").setBackground("#E9ECEF");

	// Add data validation for boolean settings
	const booleanValidation = SpreadsheetApp.newDataValidation()
		.requireValueInList(["TRUE", "FALSE"], true)
		.build();

	sheet.getRange("B5:B7").setDataValidation(booleanValidation);

	// Auto-resize columns
	sheet.autoResizeColumn(1);
	sheet.autoResizeColumn(2);
}

/**
 * Setup Audit Log sheet
 */
function setupAuditLogSheet() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.auditLog
	);

	// Clear sheet completely first
	clearSheetCompletely(sheet);

	const headers = [
		["Timestamp", "User", "Action", "Claim ID", "Details", "Result"],
	];

	sheet
		.getRange(1, 1, 1, headers[0].length)
		.setValues(headers)
		.setFontWeight("bold")
		.setBackground("#FFC107");

	sheet.setFrozenRows(1);

	// Set column widths
	sheet.setColumnWidth(1, 150); // Timestamp
	sheet.setColumnWidth(2, 200); // User
	sheet.setColumnWidth(3, 150); // Action
	sheet.setColumnWidth(4, 100); // Claim ID
	sheet.setColumnWidth(5, 300); // Details
	sheet.setColumnWidth(6, 100); // Result
}

/**
 * Setup Dashboard sheet
 */
function setupDashboardSheet() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.dashboard
	);

	// Clear sheet completely first
	clearSheetCompletely(sheet);

	// Title
	sheet.getRange("A1:H1").merge();
	sheet
		.getRange("A1")
		.setValue("Clinical Audit Dashboard")
		.setFontSize(20)
		.setFontWeight("bold")
		.setHorizontalAlignment("center")
		.setBackground("#4285F4")
		.setFontColor("#FFFFFF");

	// Last updated
	sheet.getRange("A2:H2").merge();
	sheet
		.getRange("A2")
		.setValue("Last Updated: " + new Date().toLocaleString())
		.setHorizontalAlignment("center")
		.setFontStyle("italic");

	// Summary statistics section
	sheet.getRange("A4:D4").merge();
	sheet
		.getRange("A4")
		.setValue("Summary Statistics")
		.setFontWeight("bold")
		.setFontSize(14)
		.setBackground("#E3F2FD");

	// Add statistics with formulas
	const stats = [
		["Total Claims:", "=COUNTA(Data!A:A)-1"],
		["Processed:", '=COUNTIF(Data!E:E,"Success")'],
		["Pending Validation:", "=COUNTBLANK(Data!S:S)-1"],
		["Approved:", '=COUNTIF(Data!S:S,"Approved")'],
		["Need Addendum:", '=COUNTIF(Data!S:S,"Needs Addendum")'],
		["Rejected:", '=COUNTIF(Data!S:S,"Rejected")'],
	];

	let row = 5;
	for (const stat of stats) {
		sheet.getRange(row, 1).setValue(stat[0]).setFontWeight("bold");
		sheet.getRange(row, 2).setFormula(stat[1]);
		row++;
	}

	// Processing metrics section
	sheet.getRange("E4:H4").merge();
	sheet
		.getRange("E4")
		.setValue("Processing Metrics")
		.setFontWeight("bold")
		.setFontSize(14)
		.setBackground("#E8F5E9");

	sheet.getRange("E5").setValue("Success Rate:").setFontWeight("bold");
	sheet.getRange("F5").setFormula('=IFERROR(B6/B5*100,0)&"%"');

	sheet.getRange("E6").setValue("Validation Rate:").setFontWeight("bold");
	sheet.getRange("F6").setFormula('=IFERROR((B8+B9+B10)/(B5-1)*100,0)&"%"');

	sheet.getRange("E7").setValue("Addendum Rate:").setFontWeight("bold");
	sheet.getRange("F7").setFormula('=IFERROR(B9/(B8+B9+B10)*100,0)&"%"');
}

// ============================================================================
// CORE FUNCTIONS
// ============================================================================

/**
 * Submit validation from Validator sheet
 */
function submitValidation() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const validatorSheet = ss.getSheetByName(CONFIG.sheets.validator);
	const dataSheet = ss.getSheetByName(CONFIG.sheets.data);

	// Get claim ID
	const claimId = validatorSheet.getRange("C4").getValue();
	if (!claimId) {
		SpreadsheetApp.getUi().alert("Please select a Claim ID first.");
		return;
	}

	// Get validator decision and notes
	const decision = validatorSheet.getRange("C29").getValue();
	const notes = validatorSheet.getRange("A32").getValue();

	if (!decision) {
		SpreadsheetApp.getUi().alert("Please select a validation decision.");
		return;
	}

	// Find claim in Data sheet
	const dataValues = dataSheet.getDataRange().getValues();
	let targetRow = -1;

	for (let i = 1; i < dataValues.length; i++) {
		if (dataValues[i][0] == claimId) {
			targetRow = i + 1;
			break;
		}
	}

	if (targetRow === -1) {
		SpreadsheetApp.getUi().alert("Claim ID not found in data.");
		return;
	}

	// Update Data sheet
	const timestamp = new Date();
	const userEmail = Session.getActiveUser().getEmail();

	dataSheet.getRange(targetRow, 19).setValue(decision);
	dataSheet.getRange(targetRow, 20).setValue(notes);
	dataSheet.getRange(targetRow, 21).setValue(timestamp);
	dataSheet.getRange(targetRow, 22).setValue(userEmail);

	// Log action
	logAction("Validation Submitted", claimId, `Decision: ${decision}`);

	// Show confirmation
	SpreadsheetApp.getActiveSpreadsheet().toast(
		`Validation submitted for Claim ${claimId}`,
		"Success",
		5
	);

	// Clear form
	clearValidatorForm();
}

/**
 * Clear validator form
 */
function clearValidatorForm() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.validator
	);

	sheet.getRange("C4").clearContent();
	sheet.getRange("C29").clearContent();
	sheet.getRange("A32").clearContent();

	// Clear checklist statuses
	for (let i = 8; i <= 17; i++) {
		sheet.getRange(i, 9).clearContent();
		sheet.getRange(i, 10).clearContent();
	}
}

/**
 * Refresh validator
 */
function refreshValidator() {
	clearValidatorForm();
	SpreadsheetApp.getActiveSpreadsheet().toast(
		"Validator refreshed",
		"Success",
		3
	);
}

/**
 * Update dashboard
 */
function updateDashboard() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.dashboard
	);
	sheet.getRange("A2").setValue("Last Updated: " + new Date().toLocaleString());
	SpreadsheetApp.flush();
	SpreadsheetApp.getActiveSpreadsheet().toast(
		"Dashboard updated",
		"Success",
		3
	);
}

/**
 * Log action to Audit Log
 */
function logAction(action, claimId, details) {
	try {
		const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
			CONFIG.sheets.auditLog
		);

		const logEntry = [
			new Date(),
			Session.getActiveUser().getEmail(),
			action,
			claimId || "",
			details || "",
			"Success",
		];

		sheet.appendRow(logEntry);
	} catch (e) {
		// Fail silently if audit log doesn't exist yet
	}
}

/**
 * Protect critical ranges
 */
function protectCriticalRanges() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	try {
		// Protect Data sheet (except validator columns)
		const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
		if (dataSheet) {
			const protection1 = dataSheet
				.getRange("A:R")
				.protect()
				.setDescription("AI Generated Data - Protected");
			protection1.setWarningOnly(true);
		}

		// Protect formulas in Validator sheet
		const validatorSheet = ss.getSheetByName(CONFIG.sheets.validator);
		if (validatorSheet) {
			const protection2 = validatorSheet
				.getRange("A7")
				.protect()
				.setDescription("Formula Protected");
			protection2.setWarningOnly(true);
		}

		logAction("Protection Applied", "", "Critical ranges protected");

		SpreadsheetApp.getUi().alert("Protection applied to critical ranges.");
	} catch (e) {
		SpreadsheetApp.getUi().alert("Error applying protection: " + e.toString());
	}
}

/**
 * Export audit report
 */
function exportAuditReport() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const dataSheet = ss.getSheetByName(CONFIG.sheets.data);

	if (!dataSheet) {
		SpreadsheetApp.getUi().alert(
			"Data sheet not found. Please initialize sheets first."
		);
		return;
	}

	const data = dataSheet.getDataRange().getValues();

	// Calculate statistics
	const stats = {
		total: Math.max(0, data.length - 1),
		processed: 0,
		approved: 0,
		needsAddendum: 0,
		rejected: 0,
		pending: 0,
	};

	for (let i = 1; i < data.length; i++) {
		if (data[i][4] === "Success") stats.processed++;
		if (data[i][18] === "Approved") stats.approved++;
		if (data[i][18] === "Needs Addendum") stats.needsAddendum++;
		if (data[i][18] === "Rejected") stats.rejected++;
		if (!data[i][18]) stats.pending++;
	}

	// Create report
	const report = `
CLINICAL AUDIT REPORT
Generated: ${new Date().toLocaleString()}
================================

SUMMARY STATISTICS
------------------
Total Claims: ${stats.total}
Processed: ${stats.processed} (${
		stats.total > 0 ? ((stats.processed / stats.total) * 100).toFixed(1) : 0
	}%)
Approved: ${stats.approved}
Needs Addendum: ${stats.needsAddendum}
Rejected: ${stats.rejected}
Pending: ${stats.pending}

================================
End of Report
  `;

	// Show report
	const ui = SpreadsheetApp.getUi();
	ui.alert("Audit Report", report, ui.ButtonSet.OK);

	logAction(
		"Report Exported",
		"",
		`Total: ${stats.total}, Processed: ${stats.processed}`
	);
}

/**
 * Show import dialog
 */
function showImportDialog() {
	const html = `
    <div style="padding: 20px;">
      <h3>Import CSV Data</h3>
      <p>To import data:</p>
      <ol>
        <li>Run the Python backend processor</li>
        <li>Data will automatically appear in the Data sheet</li>
        <li>Or manually paste CSV data into the Data sheet</li>
      </ol>
      <p>Make sure your CSV has the required columns.</p>
    </div>
  `;

	SpreadsheetApp.getUi().showModalDialog(
		HtmlService.createHtmlOutput(html).setWidth(400).setHeight(300),
		"Import Data"
	);
}

/**
 * Show settings dialog
 */
function showSettingsDialog() {
	const html = `
    <div style="padding: 20px;">
      <h3>System Settings</h3>
      <p>Configure settings in the Config sheet.</p>
      <p>Key settings include:</p>
      <ul>
        <li>Email notifications</li>
        <li>Lead reviewer email</li>
        <li>Archive settings</li>
        <li>System preferences</li>
      </ul>
    </div>
  `;

	SpreadsheetApp.getUi().showModalDialog(
		HtmlService.createHtmlOutput(html).setWidth(400).setHeight(300),
		"Settings"
	);
}

/**
 * Show help
 */
function showHelp() {
	const html = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>Clinical Audit System - Help</h2>
      
      <h3>Getting Started</h3>
      <ol>
        <li>Initialize all sheets using the menu</li>
        <li>Import your CSV data</li>
        <li>Configure settings in Config sheet</li>
        <li>Begin validation in Validator sheet</li>
      </ol>
      
      <h3>Validator Workflow</h3>
      <ol>
        <li>Select a Claim ID from dropdown</li>
        <li>Review the clinical note</li>
        <li>Check each validation item</li>
        <li>Select your decision</li>
        <li>Add notes if needed</li>
        <li>Click Submit Validation checkbox</li>
      </ol>
      
      <h3>Troubleshooting</h3>
      <p>If you see errors, try:</p>
      <ul>
        <li>Re-initializing sheets</li>
        <li>Checking the Audit Log</li>
        <li>Ensuring Data sheet has content</li>
      </ul>
      
      <p>Version 1.0.0</p>
    </div>
  `;

	SpreadsheetApp.getUi().showModalDialog(
		HtmlService.createHtmlOutput(html).setWidth(400).setHeight(500),
		"System Help"
	);
}

/**
 * Handle edit triggers
 */
function onEdit(e) {
	try {
		const sheet = e.source.getActiveSheet();
		const range = e.range;

		// Check if edit is in Validator sheet
		if (sheet.getName() === CONFIG.sheets.validator) {
			// Submit validation
			if (range.getA1Notation() === "C36" && e.value === "TRUE") {
				submitValidation();
				range.setValue(false);
			}

			// Clear form
			if (range.getA1Notation() === "F36" && e.value === "TRUE") {
				clearValidatorForm();
				range.setValue(false);
			}
		}
	} catch (error) {
		// Fail silently to not interrupt user
		console.error("onEdit error:", error);
	}
}
