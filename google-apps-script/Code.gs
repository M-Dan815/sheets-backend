/**
 * Clinical Audit System - Google Apps Script
 * COMPLETE VERSION with AI Recommendations Display
 * Version 2.2 - Fixed Lead Review Dropdown & Filtered Validator Dropdown
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
		schemaMeta: "Schema Metadata",
	},
	colors: {
		header: "#4285F4",
		success: "#34A853",
		warning: "#FBBC04",
		error: "#EA4335",
		neutral: "#F8F9FA",
		deprecated: "#E0E0E0",
		aiRecommendation: "#E8F5E9",
	},
	coreColumns: [
		"claim_id",
		"note_preview",
		"full_note_text",
		"processing_timestamp",
		"validator_decision",
		"validator_notes",
		"validation_timestamp",
		"validator_email",
		"lead_review_status",
		"final_addendum",
		"lead_reviewer",
		"lead_review_timestamp",
	],
	defaultValidationFields: [
		"medical_necessity",
		"diagnosis_codes",
		"cpt_codes",
		"treatment_goals",
		"interventions",
		"progress_noted",
		"plan_documented",
		"therapist_signature",
		"date_of_service",
		"session_duration",
	],
};

// ============================================================================
// MAIN MENU
// ============================================================================

function onOpen() {
	const ui = SpreadsheetApp.getUi();
	const menu = ui
		.createMenu("🔧 Audit System")
		.addItem("📊 Initialize System", "initializeAllSheets")
		.addItem("🔄 Manual Refresh Dropdowns", "refreshDropdowns")
		.addItem("🔀 Detect Schema Changes", "detectAndMigrateSchema")
		.addSeparator();

	if (isUserAuthorizedLead()) {
		menu
			.addItem("🔍 Filter Lead Review", "filterLeadReview")
			.addItem("🔒 Protect Lead Review", "protectLeadReviewSheet")
			.addSeparator();
	}

	menu
		.addItem("📤 Export Report", "exportAuditReport")
		.addItem("🐛 Debug System", "debugSystem")
		.addItem("❓ Help", "showHelp")
		.addToUi();
}

// ============================================================================
// TRIGGERS
// ============================================================================

function onEdit(e) {
	try {
		if (!e) return;

		const sheet = e.source.getActiveSheet();
		const sheetName = sheet.getName();
		const range = e.range;
		const row = range.getRow();
		const col = range.getColumn();
		const value = range.getValue();

		// Detect new data added to Data sheet
		if (sheetName === CONFIG.sheets.data && row >= 2) {
			const lastRefresh =
				PropertiesService.getScriptProperties().getProperty("lastAutoRefresh");
			const now = new Date().getTime();

			if (!lastRefresh || now - parseInt(lastRefresh) > 5000) {
				Logger.log("New data detected in Data sheet - auto-refreshing");
				PropertiesService.getScriptProperties().setProperty(
					"lastAutoRefresh",
					now.toString()
				);

				detectAndMigrateSchema();
				Utilities.sleep(500);
				refreshDropdowns();
			}
			return;
		}

		if (value !== true) return;

		Logger.log(`Checkbox detected: Sheet=${sheetName}, Row=${row}, Col=${col}`);

		// Validator Sheet checkboxes
		if (sheetName === CONFIG.sheets.validator) {
			if (row === 3 && col === 1) {
				range.setValue(false);
				submitValidation();
				return;
			}
			if (row === 5 && col === 1) {
				range.setValue(false);
				clearValidatorForm();
				return;
			}
		}

		// Lead Review Sheet checkboxes
		if (sheetName === CONFIG.sheets.leadReview) {
			if (row === 3 && col === 1) {
				range.setValue(false);
				generateAddendumTemplate();
				return;
			}
			if (row === 5 && col === 1) {
				range.setValue(false);
				saveDraft();
				return;
			}
			if (row === 7 && col === 1) {
				range.setValue(false);
				approveClaim();
				return;
			}
		}
	} catch (error) {
		Logger.log("Error in onEdit: " + error);
		SpreadsheetApp.getActiveSpreadsheet().toast(
			"Error: " + error.toString(),
			"Processing Failed",
			5
		);
	}
}

// ============================================================================
// DYNAMIC SCHEMA MANAGEMENT
// ============================================================================

function detectAndMigrateSchema() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
	const metaSheet =
		ss.getSheetByName(CONFIG.sheets.schemaMeta) || createSchemaMetaSheet();

	if (!dataSheet || dataSheet.getLastRow() < 1) {
		SpreadsheetApp.getUi().alert(
			"No data found",
			"Data sheet is empty. Import data first.",
			SpreadsheetApp.getUi().ButtonSet.OK
		);
		return;
	}

	try {
		const currentHeaders = dataSheet
			.getRange(1, 1, 1, dataSheet.getLastColumn())
			.getValues()[0];
		const currentSchema = currentHeaders.filter((h) => h);
		const storedSchema = getStoredSchema(metaSheet);

		if (storedSchema.length === 0) {
			saveSchema(metaSheet, currentSchema);
			SpreadsheetApp.getActiveSpreadsheet().toast(
				`Schema initialized with ${currentSchema.length} columns`,
				"Schema Detected",
				3
			);
			updateValidatorWithSchema(currentSchema);
			return;
		}

		const newColumns = currentSchema.filter(
			(col) => !storedSchema.includes(col)
		);
		const removedColumns = storedSchema.filter(
			(col) => !currentSchema.includes(col)
		);

		if (newColumns.length === 0 && removedColumns.length === 0) {
			Logger.log("No schema changes detected");
			return;
		}

		Logger.log(
			`Schema changes: +${newColumns.length} new, ${removedColumns.length} deprecated`
		);

		migrateSchema(
			dataSheet,
			metaSheet,
			currentSchema,
			storedSchema,
			newColumns,
			removedColumns
		);
		updateValidatorWithSchema(currentSchema);
		refreshDropdowns();

		SpreadsheetApp.getUi().alert(
			"Schema Migration Complete",
			`Schema updated!\n\nNew columns: ${newColumns.length}\nDeprecated columns: ${removedColumns.length}\n\nDeprecated columns preserved for compatibility.`,
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	} catch (error) {
		Logger.log("Error in detectAndMigrateSchema: " + error);
		SpreadsheetApp.getUi().alert(
			"Error",
			"Schema detection failed: " + error.toString(),
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	}
}

function migrateSchema(
	dataSheet,
	metaSheet,
	newSchema,
	oldSchema,
	newColumns,
	removedColumns
) {
	const mergedSchema = [...oldSchema];

	newColumns.forEach((col) => {
		if (!mergedSchema.includes(col)) {
			mergedSchema.push(col);
		}
	});

	const lastRow = dataSheet.getLastRow();
	const lastCol = dataSheet.getLastColumn();
	const allData =
		lastRow > 1
			? dataSheet.getRange(2, 1, lastRow - 1, lastCol).getValues()
			: [];

	const columnMapping = {};
	oldSchema.forEach((col, idx) => {
		const newIdx = mergedSchema.indexOf(col);
		if (newIdx !== -1) {
			columnMapping[idx] = newIdx;
		}
	});

	const migratedData = allData.map((row) => {
		const newRow = new Array(mergedSchema.length).fill("");
		row.forEach((value, oldIdx) => {
			const newIdx = columnMapping[oldIdx];
			if (newIdx !== undefined) {
				newRow[newIdx] = value;
			}
		});
		return newRow;
	});

	dataSheet.clear();

	dataSheet
		.getRange(1, 1, 1, mergedSchema.length)
		.setValues([mergedSchema])
		.setBackground(CONFIG.colors.header)
		.setFontColor("#FFFFFF")
		.setFontWeight("bold");

	removedColumns.forEach((col) => {
		const idx = mergedSchema.indexOf(col);
		if (idx !== -1) {
			dataSheet
				.getRange(1, idx + 1)
				.setBackground(CONFIG.colors.deprecated)
				.setNote("DEPRECATED: Preserved for compatibility");
		}
	});

	if (migratedData.length > 0) {
		dataSheet
			.getRange(2, 1, migratedData.length, mergedSchema.length)
			.setValues(migratedData);
	}

	saveSchema(metaSheet, mergedSchema, removedColumns);

	dataSheet.setFrozenRows(1);
	dataSheet.setColumnWidth(1, 120);
	const noteCol = mergedSchema.indexOf("full_note_text") + 1;
	if (noteCol > 0) dataSheet.hideColumns(noteCol);

	Logger.log(`Schema migrated: ${mergedSchema.length} total columns`);
}

function getStoredSchema(metaSheet) {
	const lastRow = metaSheet.getLastRow();
	if (lastRow < 2) return [];

	const schemaData = metaSheet.getRange(2, 1, lastRow - 1, 1).getValues();
	return schemaData
		.map((row) => row[0])
		.filter((col) => col && !col.startsWith("DEPRECATED:"));
}

function saveSchema(metaSheet, schema, deprecatedColumns = []) {
	metaSheet.clear();

	metaSheet
		.getRange(1, 1, 1, 3)
		.setValues([["Column Name", "Status", "Last Updated"]])
		.setBackground(CONFIG.colors.header)
		.setFontColor("#FFFFFF")
		.setFontWeight("bold");

	const timestamp = new Date();
	const schemaData = schema.map((col) => {
		const status = deprecatedColumns.includes(col) ? "DEPRECATED" : "ACTIVE";
		return [col, status, timestamp];
	});

	if (schemaData.length > 0) {
		metaSheet.getRange(2, 1, schemaData.length, 3).setValues(schemaData);
	}

	Logger.log(`Schema saved: ${schema.length} columns`);
}

function createSchemaMetaSheet() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName(CONFIG.sheets.schemaMeta);

	if (!sheet) {
		sheet = ss.insertSheet(CONFIG.sheets.schemaMeta);
		sheet.hideSheet();
	}

	return sheet;
}

function getValidationFields() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const dataSheet = ss.getSheetByName(CONFIG.sheets.data);

	if (!dataSheet || dataSheet.getLastRow() < 1) {
		return CONFIG.defaultValidationFields;
	}

	const headers = dataSheet
		.getRange(1, 1, 1, dataSheet.getLastColumn())
		.getValues()[0];

	const validationFields = headers.filter(
		(col) =>
			col &&
			!CONFIG.coreColumns.includes(col) &&
			!col.startsWith("ai_source_") &&
			col !== "note_preview" &&
			col !== "full_note_text" &&
			col !== "processing_timestamp" &&
			col !== "processing_status" &&
			col !== "compliance_issues" &&
			col !== "addendum_needed" &&
			col !== "addendum_reason" &&
			col !== "ai_confidence_score"
	);

	return validationFields.length > 0
		? validationFields
		: CONFIG.defaultValidationFields;
}

function updateValidatorWithSchema(schema) {
	Logger.log("Updating validator with new schema");
	setupValidatorSheet();
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function getLeadEmail() {
	const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.config
	);
	if (!configSheet) return "";

	const configData = configSheet.getDataRange().getValues();
	for (let i = 0; i < configData.length; i++) {
		if (configData[i][0] === "Lead Reviewer Email") {
			return configData[i][1];
		}
	}
	return "";
}

function isUserAuthorizedLead() {
	const currentUser = Session.getActiveUser().getEmail();
	const leadEmail = getLeadEmail();
	const owner = SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail();
	return currentUser === leadEmail || currentUser === owner;
}

function requireLeadAuthorization() {
	if (!isUserAuthorizedLead()) {
		SpreadsheetApp.getUi().alert(
			"🔒 Access Denied",
			"Only the designated Lead Reviewer can access this function.",
			SpreadsheetApp.getUi().ButtonSet.OK
		);
		return false;
	}
	return true;
}

function logAction(action, claimId, details) {
	try {
		const auditSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
			CONFIG.sheets.auditLog
		);
		if (!auditSheet) return;

		const timestamp = new Date();
		const user = Session.getActiveUser().getEmail();
		const lastRow = auditSheet.getLastRow();
		auditSheet
			.getRange(lastRow + 1, 1, 1, 5)
			.setValues([[timestamp, user, action, claimId || "", details || ""]]);
	} catch (error) {
		Logger.log("Error logging action: " + error);
	}
}

function getColumnIndex(sheet, columnName) {
	const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
	const index = headers.indexOf(columnName);
	return index === -1 ? null : index + 1;
}

function getColumnLetter(columnName) {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
	if (!dataSheet || dataSheet.getLastRow() < 1) return null;

	const headers = dataSheet
		.getRange(1, 1, 1, dataSheet.getLastColumn())
		.getValues()[0];
	const index = headers.indexOf(columnName);
	if (index === -1) return null;

	let letter = "";
	let num = index + 1;
	while (num > 0) {
		const mod = (num - 1) % 26;
		letter = String.fromCharCode(65 + mod) + letter;
		num = Math.floor((num - mod) / 26);
	}
	return letter;
}

// ============================================================================
// INITIALIZATION
// ============================================================================

function initializeAllSheets() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const ui = SpreadsheetApp.getUi();

	const response = ui.alert(
		"Initialize System",
		"This will set up all required sheets.\n\nExisting data will be preserved.\n\nContinue?",
		ui.ButtonSet.YES_NO
	);

	if (response !== ui.Button.YES) return;

	try {
		createSheet(CONFIG.sheets.data);
		createSheet(CONFIG.sheets.validator);
		createSheet(CONFIG.sheets.leadReview);
		createSheet(CONFIG.sheets.config);
		createSheet(CONFIG.sheets.auditLog);
		createSchemaMetaSheet();

		const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
		if (dataSheet.getLastRow() === 0) {
			setupDataSheet();
		}

		Utilities.sleep(500);
		setupValidatorSheet();
		Utilities.sleep(500);
		setupLeadReviewSheet();
		Utilities.sleep(500);
		setupConfigSheet();
		Utilities.sleep(500);
		setupAuditLogSheet();

		if (isUserAuthorizedLead()) {
			Utilities.sleep(500);
			protectLeadReviewSheet();
		}

		ui.alert(
			"✅ Success",
			"System initialized!\n\nNext Steps:\n1. Update Config sheet with lead email\n2. Import data via Python backend\n3. System will auto-detect schema\n4. Dropdowns auto-refresh after validation\n\nLead Review sheet is now protected.",
			ui.ButtonSet.OK
		);
	} catch (error) {
		ui.alert("Error", error.toString(), ui.ButtonSet.OK);
		Logger.log(error);
	}
}

function createSheet(name) {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName(name);
	if (!sheet) {
		sheet = ss.insertSheet(name);
	}
	return sheet;
}

// ============================================================================
// DROPDOWN MANAGEMENT - FIXED TO FILTER BOTH DROPDOWNS
// ============================================================================

function refreshDropdowns() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
	const validatorSheet = ss.getSheetByName(CONFIG.sheets.validator);

	const validatorCell = validatorSheet.getRange("C2");

	const lastRow = dataSheet.getLastRow();

	if (lastRow < 2) {
		Logger.log("No data for dropdowns yet - clearing both");

		// Clear validator dropdown
		validatorCell.clearContent();
		validatorCell.clearDataValidations();
		validatorCell.setValue("No data available");

		// Clear lead review dropdown if authorized
		if (isUserAuthorizedLead()) {
			const leadReviewSheet = ss.getSheetByName(CONFIG.sheets.leadReview);
			if (leadReviewSheet) {
				const leadCell = leadReviewSheet.getRange("C2");
				leadCell.clearContent();
				leadCell.clearDataValidations();
				leadCell.setValue("No data available");
			}
		}
		return;
	}

	try {
		// Get all claim IDs and their validation status
		const claimIdCol = getColumnIndex(dataSheet, "claim_id");
		const validatorDecisionCol = getColumnIndex(
			dataSheet,
			"validator_decision"
		);

		if (!claimIdCol) {
			Logger.log("claim_id column not found");
			validatorCell.clearContent();
			validatorCell.clearDataValidations();
			validatorCell.setValue("Schema not ready");
			return;
		}

		const dataRange = dataSheet.getRange(
			2,
			1,
			lastRow - 1,
			Math.max(claimIdCol, validatorDecisionCol || claimIdCol)
		);
		const allData = dataRange.getValues();

		// Filter UNVALIDATED claims for Validator dropdown
		const unvalidatedClaims = [];

		for (let i = 0; i < allData.length; i++) {
			const claimId = allData[i][claimIdCol - 1];
			const validatorDecision = validatorDecisionCol
				? allData[i][validatorDecisionCol - 1]
				: "";

			// Only include if claim has NO validator decision yet
			if (claimId && !validatorDecision) {
				unvalidatedClaims.push(String(claimId).trim());
			}
		}

		// Clear validator dropdown completely
		validatorCell.clearContent();
		validatorCell.clearDataValidations();

		if (unvalidatedClaims.length === 0) {
			// If no unvalidated claims, show message
			validatorCell.setValue("All claims validated! ✓");
			Logger.log("All claims have been validated");
		} else {
			// Remove duplicates
			const uniqueUnvalidatedClaims = [...new Set(unvalidatedClaims)];

			const validation = SpreadsheetApp.newDataValidation()
				.setAllowInvalid(false)
				.requireValueInList(uniqueUnvalidatedClaims, true)
				.build();

			validatorCell.setDataValidation(validation);

			Logger.log(
				`Validator dropdown: ${uniqueUnvalidatedClaims.length} unvalidated claims`
			);
		}

		// Refresh Lead Review dropdown if authorized
		if (isUserAuthorizedLead()) {
			refreshLeadReviewDropdown();
			filterLeadReview();
		}

		SpreadsheetApp.getActiveSpreadsheet().toast(
			`Validator: ${unvalidatedClaims.length} unvalidated | Lead: Check Lead Review tab`,
			"Dropdowns Refreshed",
			3
		);
	} catch (error) {
		Logger.log("Error in refreshDropdowns: " + error);
	}
}

function refreshLeadReviewDropdown() {
	try {
		const ss = SpreadsheetApp.getActiveSpreadsheet();
		const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
		const leadReviewSheet = ss.getSheetByName(CONFIG.sheets.leadReview);

		if (!leadReviewSheet) {
			Logger.log("Lead Review sheet not found");
			return;
		}

		// CRITICAL: Get the cell first so we can clear it even if no data
		const cell = leadReviewSheet.getRange("C2");

		const lastRow = dataSheet.getLastRow();
		if (lastRow < 2) {
			Logger.log("No data for lead review dropdown - clearing");
			// Clear the dropdown completely when no data
			const protections = cell.getProtections(
				SpreadsheetApp.ProtectionType.RANGE
			);
			protections.forEach((p) => p.remove());
			cell.clear();
			cell
				.setBackground("#FFF3CD")
				.setBorder(true, true, true, true, false, false);
			cell.setValue("No data available");
			return;
		}

		const claimIdCol = getColumnIndex(dataSheet, "claim_id");
		const validatorDecisionCol = getColumnIndex(
			dataSheet,
			"validator_decision"
		);
		const leadReviewStatusCol = getColumnIndex(dataSheet, "lead_review_status");

		if (!claimIdCol || !validatorDecisionCol) {
			Logger.log("Required columns not found - clearing dropdown");
			const protections = cell.getProtections(
				SpreadsheetApp.ProtectionType.RANGE
			);
			protections.forEach((p) => p.remove());
			cell.clear();
			cell
				.setBackground("#FFF3CD")
				.setBorder(true, true, true, true, false, false);
			cell.setValue("Schema not ready");
			return;
		}

		const lastCol = Math.max(
			claimIdCol,
			validatorDecisionCol,
			leadReviewStatusCol || 1
		);
		const allData = dataSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

		const claimsNeedingAddendum = [];

		for (let i = 0; i < allData.length; i++) {
			const claimId = allData[i][claimIdCol - 1];
			const validatorDecision = allData[i][validatorDecisionCol - 1];
			const leadReviewStatus = leadReviewStatusCol
				? allData[i][leadReviewStatusCol - 1]
				: "";

			if (
				validatorDecision === "⚠ Needs Addendum" &&
				leadReviewStatus !== "✅ Approved" &&
				claimId
			) {
				// AGGRESSIVE cleaning: remove all whitespace and special characters that might be hidden
				const cleanClaimId = String(claimId)
					.replace(/[\u200B-\u200D\uFEFF]/g, "") // Remove zero-width characters
					.replace(/\s+/g, " ") // Normalize whitespace
					.trim();

				if (cleanClaimId && !claimsNeedingAddendum.includes(cleanClaimId)) {
					claimsNeedingAddendum.push(cleanClaimId);
				}
			}
		}

		// Remove any range protection on C2 specifically
		const protections = cell.getProtections(
			SpreadsheetApp.ProtectionType.RANGE
		);
		protections.forEach((p) => p.remove());

		// Completely clear the cell
		cell.clear(); // Clear everything including format, validation, content

		// Re-apply the background color
		cell
			.setBackground("#FFF3CD")
			.setBorder(true, true, true, true, false, false);

		// Wait a moment for the clear to take effect
		SpreadsheetApp.flush();

		if (claimsNeedingAddendum.length > 0) {
			// Sort for consistency
			claimsNeedingAddendum.sort();

			Logger.log(
				`Setting dropdown with values: [${claimsNeedingAddendum.join(", ")}]`
			);

			// Create validation with explicit, clean values
			const validation = SpreadsheetApp.newDataValidation()
				.setAllowInvalid(false)
				.requireValueInList(claimsNeedingAddendum, true)
				.build();

			cell.setDataValidation(validation);

			Logger.log(
				`✓ Lead review dropdown updated: ${claimsNeedingAddendum.length} claims`
			);
		} else {
			cell.setValue("No claims need addendum ✓");
			Logger.log("No claims needing addendum for lead review");
		}
	} catch (error) {
		Logger.log("Error in refreshLeadReviewDropdown: " + error);
		throw error; // Re-throw to see full error in UI
	}
}

function filterLeadReview() {
	if (!requireLeadAuthorization()) return;

	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
	const leadReviewSheet = ss.getSheetByName(CONFIG.sheets.leadReview);

	const dataLastRow = dataSheet.getLastRow();
	if (dataLastRow < 2) return;

	const lastRow = leadReviewSheet.getLastRow();
	if (lastRow > 7) {
		leadReviewSheet.getRange(8, 1, lastRow - 7, 4).clearContent();
	}

	const claimIdCol = getColumnIndex(dataSheet, "claim_id");
	const validatorDecisionCol = getColumnIndex(dataSheet, "validator_decision");
	const validatorEmailCol = getColumnIndex(dataSheet, "validator_email");
	const validationTimestampCol = getColumnIndex(
		dataSheet,
		"validation_timestamp"
	);
	const leadReviewStatusCol = getColumnIndex(dataSheet, "lead_review_status");

	if (!claimIdCol || !validatorDecisionCol) return;

	const maxCol = Math.max(
		claimIdCol,
		validatorDecisionCol,
		validatorEmailCol || 1,
		validationTimestampCol || 1,
		leadReviewStatusCol || 1
	);
	const allData = dataSheet.getRange(2, 1, dataLastRow - 1, maxCol).getValues();

	const filteredData = [];
	for (let i = 0; i < allData.length; i++) {
		const validatorDecision = allData[i][validatorDecisionCol - 1];
		const leadReviewStatus = leadReviewStatusCol
			? allData[i][leadReviewStatusCol - 1]
			: "";

		if (
			validatorDecision === "⚠ Needs Addendum" &&
			leadReviewStatus !== "✅ Approved"
		) {
			filteredData.push([
				allData[i][claimIdCol - 1],
				validatorEmailCol ? allData[i][validatorEmailCol - 1] : "",
				validationTimestampCol ? allData[i][validationTimestampCol - 1] : "",
				leadReviewStatus || "Pending",
			]);
		}
	}

	if (filteredData.length > 0) {
		leadReviewSheet
			.getRange(8, 1, filteredData.length, 4)
			.setValues(filteredData);

		for (let i = 0; i < filteredData.length; i++) {
			const status = filteredData[i][3];
			const cell = leadReviewSheet.getRange(8 + i, 4);
			if (status === "Pending" || !status) {
				cell.setBackground(CONFIG.colors.warning);
			}
		}
	}

	// REMOVED: refreshLeadReviewDropdown(); - Already called by refreshDropdowns()
	Logger.log(`Lead review filtered: ${filteredData.length} claims`);
}

// ============================================================================
// VALIDATION PROCESSING
// ============================================================================

function submitValidation() {
	try {
		const ss = SpreadsheetApp.getActiveSpreadsheet();
		const validatorSheet = ss.getSheetByName(CONFIG.sheets.validator);
		const dataSheet = ss.getSheetByName(CONFIG.sheets.data);

		const claimId = validatorSheet.getRange("C2").getValue();
		const overallDecision = validatorSheet.getRange("G2").getValue();

		if (!claimId) {
			SpreadsheetApp.getUi().alert(
				"❌ Missing Claim ID",
				"Please select a claim ID first.",
				SpreadsheetApp.getUi().ButtonSet.OK
			);
			return;
		}

		if (!overallDecision) {
			SpreadsheetApp.getUi().alert(
				"❌ Missing Decision",
				"Please select an overall decision (Approved/Needs Addendum/Rejected).",
				SpreadsheetApp.getUi().ButtonSet.OK
			);
			return;
		}

		const validationFields = getValidationFields();
		const validationData = {};
		let hasFailure = false;
		const addendumLines = [];

		const startRow = 12;
		for (let i = 0; i < validationFields.length; i++) {
			const row = startRow + i;
			const field = validationFields[i];
			const aiValue = validatorSheet.getRange(row, 7).getValue();
			const aiSource = validatorSheet.getRange(row, 8).getValue();
			const passFail = validatorSheet.getRange(row, 9).getValue();
			const addendumLine = validatorSheet.getRange(row, 10).getValue();

			validationData[field] = {
				aiValue: aiValue,
				aiSource: aiSource,
				passFail: passFail,
				addendumLine: addendumLine,
			};

			if (passFail === "✗ Fail") {
				hasFailure = true;
				if (addendumLine) {
					addendumLines.push(`${field}: ${addendumLine}`);
				}
			}
		}

		const decision = overallDecision;
		const validatorNotes =
			addendumLines.length > 0 ? addendumLines.join("\n") : "";

		if (
			(decision === "⚠ Needs Addendum" || decision === "❌ Rejected") &&
			addendumLines.length === 0
		) {
			SpreadsheetApp.getUi().alert(
				"❌ Missing Addendum Lines",
				"Please provide addendum lines for items that need correction.",
				SpreadsheetApp.getUi().ButtonSet.OK
			);
			return;
		}

		const dataLastRow = dataSheet.getLastRow();
		if (dataLastRow < 2) {
			SpreadsheetApp.getUi().alert(
				"❌ No Data",
				"No data found.",
				SpreadsheetApp.getUi().ButtonSet.OK
			);
			return;
		}

		const claimIdCol = getColumnIndex(dataSheet, "claim_id");
		const claimIds = dataSheet
			.getRange(2, claimIdCol, dataLastRow - 1, 1)
			.getValues();

		let rowIndex = -1;
		for (let i = 0; i < claimIds.length; i++) {
			if (String(claimIds[i][0]).trim() === String(claimId).trim()) {
				rowIndex = i + 2;
				break;
			}
		}

		if (rowIndex === -1) {
			SpreadsheetApp.getUi().alert(
				"❌ Claim Not Found",
				`Claim ID "${claimId}" not found.`,
				SpreadsheetApp.getUi().ButtonSet.OK
			);
			return;
		}

		const timestamp = new Date();
		const userEmail = Session.getActiveUser().getEmail();

		const validatorDecisionCol = getColumnIndex(
			dataSheet,
			"validator_decision"
		);
		const validatorNotesCol = getColumnIndex(dataSheet, "validator_notes");
		const validationTimestampCol = getColumnIndex(
			dataSheet,
			"validation_timestamp"
		);
		const validatorEmailCol = getColumnIndex(dataSheet, "validator_email");

		if (validatorDecisionCol)
			dataSheet.getRange(rowIndex, validatorDecisionCol).setValue(decision);
		if (validatorNotesCol)
			dataSheet.getRange(rowIndex, validatorNotesCol).setValue(validatorNotes);
		if (validationTimestampCol)
			dataSheet.getRange(rowIndex, validationTimestampCol).setValue(timestamp);
		if (validatorEmailCol)
			dataSheet.getRange(rowIndex, validatorEmailCol).setValue(userEmail);

		validationFields.forEach((field) => {
			const fieldCol = getColumnIndex(dataSheet, field);
			if (fieldCol && validationData[field]) {
				dataSheet
					.getRange(rowIndex, fieldCol)
					.setValue(validationData[field].aiValue || "");
			}

			const aiSourceCol = getColumnIndex(dataSheet, `ai_source_${field}`);
			if (aiSourceCol && validationData[field].aiSource) {
				dataSheet
					.getRange(rowIndex, aiSourceCol)
					.setValue(validationData[field].aiSource);
			}
		});

		logAction("Validation Submitted", claimId, `Decision: ${decision}`);
		clearValidatorForm();
		refreshDropdowns();

		if (isUserAuthorizedLead()) {
			Utilities.sleep(500);
			refreshLeadReviewDropdown();
			filterLeadReview();
		}

		SpreadsheetApp.getActiveSpreadsheet().toast(
			`✅ Validation submitted for ${claimId}`,
			"Success",
			3
		);
	} catch (error) {
		Logger.log("Error in submitValidation: " + error);
		SpreadsheetApp.getUi().alert(
			"Error",
			"Failed to submit: " + error.toString(),
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	}
}

function clearValidatorForm() {
	try {
		const validatorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
			CONFIG.sheets.validator
		);

		validatorSheet.getRange("C2").clearContent();
		validatorSheet.getRange("G2").clearContent();

		const validationFields = getValidationFields();
		const startRow = 12;

		if (validationFields.length > 0) {
			validatorSheet
				.getRange(startRow, 9, validationFields.length, 2)
				.clearContent();
		}

		SpreadsheetApp.getActiveSpreadsheet().toast("Form cleared", "Success", 2);
	} catch (error) {
		Logger.log("Error in clearValidatorForm: " + error);
	}
}

// ============================================================================
// LEAD REVIEW PROCESSING - FIXED FORMULAS
// ============================================================================

function generateAddendumTemplate() {
	try {
		if (!requireLeadAuthorization()) return;

		const ss = SpreadsheetApp.getActiveSpreadsheet();
		const leadReviewSheet = ss.getSheetByName(CONFIG.sheets.leadReview);
		const dataSheet = ss.getSheetByName(CONFIG.sheets.data);

		const claimId = leadReviewSheet.getRange("C2").getValue();

		if (!claimId || claimId === "No claims need addendum ✓") {
			SpreadsheetApp.getUi().alert(
				"❌ No Claim Selected",
				"Please select a claim ID first.",
				SpreadsheetApp.getUi().ButtonSet.OK
			);
			return;
		}

		const claimIdCol = getColumnIndex(dataSheet, "claim_id");
		const dataLastRow = dataSheet.getLastRow();
		const claimIds = dataSheet
			.getRange(2, claimIdCol, dataLastRow - 1, 1)
			.getValues();

		let rowIndex = -1;
		for (let i = 0; i < claimIds.length; i++) {
			if (String(claimIds[i][0]).trim() === String(claimId).trim()) {
				rowIndex = i + 2;
				break;
			}
		}

		if (rowIndex === -1) {
			SpreadsheetApp.getUi().alert(
				"❌ Claim Not Found",
				"Claim not found.",
				SpreadsheetApp.getUi().ButtonSet.OK
			);
			return;
		}

		const validatorNotesCol = getColumnIndex(dataSheet, "validator_notes");
		const dateOfServiceCol = getColumnIndex(dataSheet, "date_of_service");
		const addendumReasonCol = getColumnIndex(dataSheet, "addendum_reason");

		const validatorNotes = validatorNotesCol
			? dataSheet.getRange(rowIndex, validatorNotesCol).getValue()
			: "";
		const dateOfService = dateOfServiceCol
			? dataSheet.getRange(rowIndex, dateOfServiceCol).getValue()
			: "Not specified";
		const addendumReason = addendumReasonCol
			? dataSheet.getRange(rowIndex, addendumReasonCol).getValue()
			: "";

		const template = `CLINICAL ADDENDUM

Claim ID: ${claimId}
Date of Service: ${dateOfService}
Addendum Date: ${Utilities.formatDate(
			new Date(),
			Session.getScriptTimeZone(),
			"yyyy-MM-dd"
		)}

AI IDENTIFIED ISSUES:
${addendumReason || "No AI issues identified"}

VALIDATOR FINDINGS:
${validatorNotes || "No specific findings provided"}

ADDITIONAL CLINICAL DOCUMENTATION:
[Add required clinical information here]

UPDATED ASSESSMENT:
[Document clinical assessment updates]

TREATMENT MODIFICATIONS:
[Document any changes to treatment plan]

CLINICAL JUSTIFICATION:
[Provide medical necessity justification]

---
Reviewed by: ${Session.getActiveUser().getEmail()}
Date: ${Utilities.formatDate(
			new Date(),
			Session.getScriptTimeZone(),
			"yyyy-MM-dd HH:mm:ss"
		)}
`;

		leadReviewSheet.getRange("G10:H50").getMergedRanges()[0].setValue(template);
		SpreadsheetApp.getActiveSpreadsheet().toast(
			"✅ Template generated",
			"Success",
			2
		);
	} catch (error) {
		Logger.log("Error in generateAddendumTemplate: " + error);
		SpreadsheetApp.getUi().alert(
			"Error",
			error.toString(),
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	}
}

function saveDraft() {
	try {
		if (!requireLeadAuthorization()) return;

		const ss = SpreadsheetApp.getActiveSpreadsheet();
		const leadReviewSheet = ss.getSheetByName(CONFIG.sheets.leadReview);
		const dataSheet = ss.getSheetByName(CONFIG.sheets.data);

		const claimId = leadReviewSheet.getRange("C2").getValue();
		const addendumText = leadReviewSheet
			.getRange("G10:H50")
			.getMergedRanges()[0]
			.getValue();

		if (!claimId || claimId === "No claims need addendum ✓") {
			SpreadsheetApp.getUi().alert(
				"❌ No Claim Selected",
				"Please select a claim ID first.",
				SpreadsheetApp.getUi().ButtonSet.OK
			);
			return;
		}

		if (!addendumText || addendumText.trim() === "") {
			SpreadsheetApp.getUi().alert(
				"❌ No Content",
				"Please enter addendum text before saving.",
				SpreadsheetApp.getUi().ButtonSet.OK
			);
			return;
		}

		const claimIdCol = getColumnIndex(dataSheet, "claim_id");
		const dataLastRow = dataSheet.getLastRow();
		const claimIds = dataSheet
			.getRange(2, claimIdCol, dataLastRow - 1, 1)
			.getValues();

		let rowIndex = -1;
		for (let i = 0; i < claimIds.length; i++) {
			if (String(claimIds[i][0]).trim() === String(claimId).trim()) {
				rowIndex = i + 2;
				break;
			}
		}

		if (rowIndex === -1) {
			SpreadsheetApp.getUi().alert(
				"❌ Claim Not Found",
				"Claim not found.",
				SpreadsheetApp.getUi().ButtonSet.OK
			);
			return;
		}

		const finalAddendumCol = getColumnIndex(dataSheet, "final_addendum");
		if (finalAddendumCol) {
			dataSheet.getRange(rowIndex, finalAddendumCol).setValue(addendumText);
		}

		logAction("Draft Saved", claimId, "Addendum draft saved");
		SpreadsheetApp.getActiveSpreadsheet().toast("✅ Draft saved", "Success", 2);
	} catch (error) {
		Logger.log("Error in saveDraft: " + error);
		SpreadsheetApp.getUi().alert(
			"Error",
			error.toString(),
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	}
}

function approveClaim() {
	try {
		if (!requireLeadAuthorization()) return;

		const ss = SpreadsheetApp.getActiveSpreadsheet();
		const leadReviewSheet = ss.getSheetByName(CONFIG.sheets.leadReview);
		const dataSheet = ss.getSheetByName(CONFIG.sheets.data);

		const claimId = leadReviewSheet.getRange("C2").getValue();
		const addendumText = leadReviewSheet
			.getRange("G10:H50")
			.getMergedRanges()[0]
			.getValue();

		if (!claimId || claimId === "No claims need addendum ✓") {
			SpreadsheetApp.getUi().alert(
				"❌ No Claim Selected",
				"Please select a claim ID first.",
				SpreadsheetApp.getUi().ButtonSet.OK
			);
			return;
		}

		const claimIdCol = getColumnIndex(dataSheet, "claim_id");
		const dataLastRow = dataSheet.getLastRow();
		const claimIds = dataSheet
			.getRange(2, claimIdCol, dataLastRow - 1, 1)
			.getValues();

		let rowIndex = -1;
		for (let i = 0; i < claimIds.length; i++) {
			if (String(claimIds[i][0]).trim() === String(claimId).trim()) {
				rowIndex = i + 2;
				break;
			}
		}

		if (rowIndex === -1) {
			SpreadsheetApp.getUi().alert(
				"❌ Claim Not Found",
				"Claim not found.",
				SpreadsheetApp.getUi().ButtonSet.OK
			);
			return;
		}

		const timestamp = new Date();
		const userEmail = Session.getActiveUser().getEmail();

		const leadReviewStatusCol = getColumnIndex(dataSheet, "lead_review_status");
		const finalAddendumCol = getColumnIndex(dataSheet, "final_addendum");
		const leadReviewerCol = getColumnIndex(dataSheet, "lead_reviewer");
		const leadReviewTimestampCol = getColumnIndex(
			dataSheet,
			"lead_review_timestamp"
		);

		if (leadReviewStatusCol)
			dataSheet.getRange(rowIndex, leadReviewStatusCol).setValue("✅ Approved");
		if (finalAddendumCol)
			dataSheet
				.getRange(rowIndex, finalAddendumCol)
				.setValue(addendumText || "Approved without changes");
		if (leadReviewerCol)
			dataSheet.getRange(rowIndex, leadReviewerCol).setValue(userEmail);
		if (leadReviewTimestampCol)
			dataSheet.getRange(rowIndex, leadReviewTimestampCol).setValue(timestamp);

		logAction("Lead Review Approved", claimId, `Approved by ${userEmail}`);

		leadReviewSheet.getRange("C2").clearContent();
		leadReviewSheet.getRange("G10:H50").getMergedRanges()[0].clearContent();

		filterLeadReview();

		SpreadsheetApp.getActiveSpreadsheet().toast(
			"✅ Claim approved",
			"Success",
			3
		);
	} catch (error) {
		Logger.log("Error in approveClaim: " + error);
		SpreadsheetApp.getUi().alert(
			"Error",
			error.toString(),
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	}
}

function protectLeadReviewSheet() {
	if (!requireLeadAuthorization()) return;

	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const leadReviewSheet = ss.getSheetByName(CONFIG.sheets.leadReview);
	const leadEmail = getLeadEmail();
	const owner = ss.getOwner().getEmail();

	try {
		// Remove all existing protections
		const protections = leadReviewSheet.getProtections(
			SpreadsheetApp.ProtectionType.SHEET
		);
		protections.forEach((p) => p.remove());

		const rangeProtections = leadReviewSheet.getProtections(
			SpreadsheetApp.ProtectionType.RANGE
		);
		rangeProtections.forEach((p) => p.remove());

		// Protect the entire sheet
		const protection = leadReviewSheet
			.protect()
			.setDescription("Lead Review - Restricted Access")
			.setWarningOnly(false);

		// Set unprotected ranges (cells that need to be editable)
		const unprotectedRanges = [
			leadReviewSheet.getRange("A3"), // Generate template checkbox
			leadReviewSheet.getRange("A5"), // Save draft checkbox
			leadReviewSheet.getRange("A7"), // Approve checkbox
			leadReviewSheet.getRange("C2"), // Claim ID dropdown - CRITICAL
			leadReviewSheet.getRange("G10:H50"), // Addendum draft area
		];

		protection.setUnprotectedRanges(unprotectedRanges);

		const editors = [owner];
		if (leadEmail && leadEmail !== owner) {
			editors.push(leadEmail);
		}

		protection.addEditors(editors);

		SpreadsheetApp.getUi().alert(
			"Protection Applied",
			`Lead Review sheet is now protected.\n\nEditable cells: C2 (dropdown) and addendum area\n\nAuthorized users:\n- ${editors.join(
				"\n- "
			)}`,
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	} catch (error) {
		Logger.log("Error in protectLeadReviewSheet: " + error);
		SpreadsheetApp.getUi().alert(
			"Error",
			error.toString(),
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	}
}

// ============================================================================
// SHEET SETUP - WITH AI RECOMMENDATIONS
// ============================================================================

function setupDataSheet() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.data
	);
	sheet.clear();

	const headers = [
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
		"ai_source_diagnosis_codes",
		"ai_source_cpt_codes",
		"ai_source_medical_necessity",
		"ai_source_treatment_goals",
		"ai_source_interventions",
		"ai_source_progress_noted",
		"ai_source_plan_documented",
		"ai_source_therapist_signature",
		"ai_source_date_of_service",
		"ai_source_session_duration",
	];

	sheet
		.getRange(1, 1, 1, headers.length)
		.setValues([headers])
		.setBackground(CONFIG.colors.header)
		.setFontColor("#FFFFFF")
		.setFontWeight("bold");

	sheet.setFrozenRows(1);
	sheet.setColumnWidth(1, 120);
	sheet.setColumnWidth(2, 250);
	sheet.hideColumns(3);
}

function setupValidatorSheet() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.validator
	);
	sheet.clear();

	// Header
	sheet
		.getRange("A1:J1")
		.merge()
		.setValue("📋 CLINICAL NOTE VALIDATOR - WITH AI RECOMMENDATIONS")
		.setFontSize(14)
		.setFontWeight("bold")
		.setHorizontalAlignment("center")
		.setBackground(CONFIG.colors.header)
		.setFontColor("#FFFFFF");

	// Claim selection
	sheet.getRange("A2").setValue("Select Claim:").setFontWeight("bold");
	sheet
		.getRange("C2")
		.setBackground("#FFF3CD")
		.setBorder(true, true, true, true, false, false);

	// Overall Decision
	sheet.getRange("E2").setValue("Overall Decision:").setFontWeight("bold");
	sheet
		.getRange("G2")
		.setBackground("#D4EDDA")
		.setBorder(true, true, true, true, false, false);
	const decisionValidation = SpreadsheetApp.newDataValidation()
		.requireValueInList(
			["✅ Approved", "⚠ Needs Addendum", "❌ Rejected"],
			true
		)
		.build();
	sheet.getRange("G2").setDataValidation(decisionValidation);

	// Action buttons
	sheet.getRange("A3").insertCheckboxes().setValue(false);
	sheet
		.getRange("B3")
		.setValue("✓ Submit Validation")
		.setBackground("#D4EDDA")
		.setFontWeight("bold");

	sheet.getRange("A5").insertCheckboxes().setValue(false);
	sheet
		.getRange("B5")
		.setValue("🔄 Clear Form")
		.setBackground("#FFF3CD")
		.setFontWeight("bold");

	// AI Recommendations Panel
	sheet
		.getRange("A7:J7")
		.merge()
		.setValue("🤖 AI RECOMMENDATIONS")
		.setFontWeight("bold")
		.setBackground(CONFIG.colors.aiRecommendation)
		.setHorizontalAlignment("center");

	sheet
		.getRange("A8:J8")
		.merge()
		.setFormula('=IFERROR(VLOOKUP(TRIM(C2),Data!A:R,18,FALSE),"")')
		.setWrap(true)
		.setFontStyle("italic")
		.setFontSize(9)
		.setBackground("#F1F8E9");

	// LEFT PANEL: Original Clinical Note
	sheet
		.getRange("A10:D10")
		.merge()
		.setValue("ORIGINAL CLINICAL NOTE")
		.setFontWeight("bold")
		.setBackground("#E3F2FD")
		.setHorizontalAlignment("center");

	sheet
		.getRange("A11:D50")
		.merge()
		.setFormula(
			'=IFERROR(VLOOKUP(TRIM(C2),Data!A:C,3,FALSE),"Select a claim ID above")'
		)
		.setWrap(true)
		.setVerticalAlignment("top")
		.setBorder(true, true, true, true, false, false)
		.setFontSize(9);

	// RIGHT PANEL: Validation Table with AI Sources
	sheet
		.getRange("F10:J10")
		.merge()
		.setValue("VALIDATION TABLE (with AI findings)")
		.setFontWeight("bold")
		.setBackground("#E8F5E9")
		.setHorizontalAlignment("center");

	// Validation table headers
	const tableHeaders = [
		["Field", "AI Output", "AI Source", "Pass/Fail", "Addendum Line"],
	];
	sheet
		.getRange(11, 6, 1, 5)
		.setValues(tableHeaders)
		.setFontWeight("bold")
		.setBackground("#F8F9FA")
		.setBorder(true, true, true, true, true, true);

	// Get validation fields and populate
	const validationFields = getValidationFields();
	const startRow = 12;

	const fieldLabels = validationFields.map((field) => [
		field.replace(/_/g, " ").toUpperCase(),
	]);
	if (fieldLabels.length > 0) {
		sheet
			.getRange(startRow, 6, fieldLabels.length, 1)
			.setValues(fieldLabels)
			.setFontWeight("bold");
	}

	// AI Output formulas (Column G)
	for (let i = 0; i < validationFields.length; i++) {
		const row = startRow + i;
		const field = validationFields[i];
		const colLetter = getColumnLetter(field);
		if (colLetter) {
			sheet
				.getRange(row, 7)
				.setFormula(
					`=IFERROR(VLOOKUP(TRIM($C$2),Data!A:${colLetter},MATCH("${field}",Data!$1:$1,0),FALSE),"")`
				);
		}
	}

	// AI Source formulas (Column H)
	for (let i = 0; i < validationFields.length; i++) {
		const row = startRow + i;
		const field = validationFields[i];
		const sourceField = "ai_source_" + field;
		const colLetter = getColumnLetter(sourceField);
		if (colLetter) {
			sheet
				.getRange(row, 8)
				.setFormula(
					`=IFERROR(VLOOKUP(TRIM($C$2),Data!A:${colLetter},MATCH("${sourceField}",Data!$1:$1,0),FALSE),"")`
				);
			sheet
				.getRange(row, 8)
				.setBackground("#FFF9C4")
				.setWrap(true)
				.setFontSize(8);
		} else {
			sheet.getRange(row, 8).setValue("").setBackground("#F8F9FA");
		}
	}

	// Pass/Fail dropdowns (Column I)
	const passFailValidation = SpreadsheetApp.newDataValidation()
		.requireValueInList(["✓ Pass", "✗ Fail", "N/A"], true)
		.build();
	sheet
		.getRange(startRow, 9, validationFields.length, 1)
		.setDataValidation(passFailValidation);

	// Addendum Line (Column J)
	sheet
		.getRange(startRow, 10, validationFields.length, 1)
		.setBackground("#FFF3CD");

	// Set column widths
	sheet.setColumnWidth(1, 80);
	sheet.setColumnWidth(2, 150);
	sheet.setColumnWidth(3, 120);
	sheet.setColumnWidth(4, 350);
	sheet.setColumnWidth(5, 20);
	sheet.setColumnWidth(6, 150);
	sheet.setColumnWidth(7, 120);
	sheet.setColumnWidth(8, 200);
	sheet.setColumnWidth(9, 80);
	sheet.setColumnWidth(10, 200);

	sheet.setFrozenRows(10);
}

function setupLeadReviewSheet() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.leadReview
	);
	sheet.clear();

	// Header
	sheet
		.getRange("A1:D1")
		.merge()
		.setValue("🔍 LEAD REVIEW DASHBOARD")
		.setFontSize(14)
		.setFontWeight("bold")
		.setHorizontalAlignment("center")
		.setBackground("#DC3545")
		.setFontColor("#FFFFFF");

	sheet
		.getRange("E1:H1")
		.merge()
		.setValue("THREE PANEL HORIZONTAL LAYOUT")
		.setFontSize(14)
		.setFontWeight("bold")
		.setHorizontalAlignment("center")
		.setBackground("#DC3545")
		.setFontColor("#FFFFFF");

	// Claim selection and actions
	sheet.getRange("A2").setValue("Select Claim:").setFontWeight("bold");
	sheet
		.getRange("C2")
		.setBackground("#FFF3CD")
		.setBorder(true, true, true, true, false, false);

	sheet.getRange("A3").insertCheckboxes().setValue(false);
	sheet
		.getRange("B3")
		.setValue("🤖 Generate Template")
		.setBackground("#D1ECF1")
		.setFontWeight("bold");

	sheet.getRange("A5").insertCheckboxes().setValue(false);
	sheet
		.getRange("B5")
		.setValue("💾 Save Draft")
		.setBackground("#FFC107")
		.setFontWeight("bold");

	sheet.getRange("A7").insertCheckboxes().setValue(false);
	sheet
		.getRange("B7")
		.setValue("✅ Approve Claim")
		.setBackground("#28A745")
		.setFontColor("#FFF")
		.setFontWeight("bold");

	// LEFT PANEL: Claims Dashboard
	sheet
		.getRange("A9:D9")
		.merge()
		.setValue("CLAIMS NEEDING ADDENDUM")
		.setFontWeight("bold")
		.setBackground("#F8D7DA")
		.setHorizontalAlignment("center");

	const dashboardHeaders = [["Claim ID", "Validator", "Date", "Status"]];
	sheet
		.getRange(10, 1, 1, 4)
		.setValues(dashboardHeaders)
		.setFontWeight("bold")
		.setBackground("#F8F9FA");

	// MIDDLE PANEL: Original Clinical Note - FIXED FORMULA WITH ARRAYFORMULA
	sheet
		.getRange("E9:F9")
		.merge()
		.setValue("ORIGINAL NOTE")
		.setFontWeight("bold")
		.setBackground("#E3F2FD")
		.setHorizontalAlignment("center");

	// Use INDEX/MATCH instead of VLOOKUP for better reliability
	sheet
		.getRange("E10:F50")
		.merge()
		.setFormula(
			'=IFERROR(INDEX(Data!C:C,MATCH(TRIM(C2),Data!A:A,0)),"Select a claim from dropdown above")'
		)
		.setWrap(true)
		.setVerticalAlignment("top")
		.setBorder(true, true, true, true, false, false)
		.setFontSize(9);

	// RIGHT PANEL: Addendum Draft
	sheet
		.getRange("G9:H9")
		.merge()
		.setValue("ADDENDUM DRAFT")
		.setFontWeight("bold")
		.setBackground("#28A745")
		.setFontColor("#FFF")
		.setHorizontalAlignment("center");

	sheet
		.getRange("G10:H50")
		.merge()
		.setWrap(true)
		.setVerticalAlignment("top")
		.setFontFamily("Courier New")
		.setBorder(true, true, true, true, false, false);

	// Set column widths
	sheet.setColumnWidth(1, 120);
	sheet.setColumnWidth(2, 150);
	sheet.setColumnWidth(3, 100);
	sheet.setColumnWidth(4, 80);
	sheet.setColumnWidth(5, 300);
	sheet.setColumnWidth(6, 300);
	sheet.setColumnWidth(7, 300);
	sheet.setColumnWidth(8, 300);

	sheet.setFrozenRows(9);
	sheet.setFrozenColumns(4);
}

function setupConfigSheet() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.config
	);
	sheet.clear();

	sheet
		.getRange("A1:B1")
		.merge()
		.setValue("System Configuration")
		.setFontSize(14)
		.setFontWeight("bold")
		.setBackground("#6C757D")
		.setFontColor("#FFFFFF");

	const settings = [
		["Setting", "Value"],
		["Lead Reviewer Email", "lead@example.com"],
		["CC Email List", "manager@example.com"],
		["Send Notifications", "TRUE"],
		["System Version", "2.2"],
	];

	sheet.getRange(3, 1, settings.length, 2).setValues(settings);
	sheet.getRange(3, 1, 1, 2).setFontWeight("bold").setBackground("#E9ECEF");
	sheet.getRange("B4").setNote("UPDATE THIS with actual lead reviewer email!");
}

function setupAuditLogSheet() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		CONFIG.sheets.auditLog
	);
	sheet.clear();

	const headers = [["Timestamp", "User", "Action", "Claim ID", "Details"]];
	sheet
		.getRange(1, 1, 1, 5)
		.setValues(headers)
		.setFontWeight("bold")
		.setBackground("#FFC107");
	sheet.setFrozenRows(1);
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function debugSystem() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
	const validatorSheet = ss.getSheetByName(CONFIG.sheets.validator);
	const leadReviewSheet = ss.getSheetByName(CONFIG.sheets.leadReview);

	const lastRow = dataSheet ? dataSheet.getLastRow() : 0;
	let firstClaim = "N/A";
	let schemaColumns = 0;

	if (dataSheet && lastRow > 1) {
		firstClaim = dataSheet.getRange(2, 1).getValue();
		schemaColumns = dataSheet.getLastColumn();
	}

	const validationFields = getValidationFields();
	const c2Value = validatorSheet
		? validatorSheet.getRange("C2").getValue()
		: "N/A";
	const leadC2Value = leadReviewSheet
		? leadReviewSheet.getRange("C2").getValue()
		: "N/A";

	const addendumReasonCol = getColumnIndex(dataSheet, "addendum_reason");

	SpreadsheetApp.getUi().alert(
		"Debug System",
		`DATA SHEET:\n` +
			`- Rows: ${lastRow - 1}\n` +
			`- Columns: ${schemaColumns}\n` +
			`- First claim: "${firstClaim}"\n` +
			`- AI recommendations column: ${
				addendumReasonCol ? "Found" : "Missing"
			}\n\n` +
			`SCHEMA:\n` +
			`- Validation fields detected: ${validationFields.length}\n` +
			`- Fields: ${validationFields.slice(0, 3).join(", ")}...\n\n` +
			`VALIDATOR:\n` +
			`- Selected claim: "${c2Value}"\n\n` +
			`LEAD REVIEW:\n` +
			`- Selected claim: "${leadC2Value}"\n` +
			`- User is authorized: ${isUserAuthorizedLead()}\n\n` +
			`${lastRow < 2 ? "❌ No data! Import data first." : "✅ System ready"}`,
		SpreadsheetApp.getUi().ButtonSet.OK
	);
}

function exportAuditReport() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const dataSheet = ss.getSheetByName(CONFIG.sheets.data);
	const auditSheet = ss.getSheetByName(CONFIG.sheets.auditLog);

	const reportName = `Audit Report - ${Utilities.formatDate(
		new Date(),
		Session.getScriptTimeZone(),
		"yyyy-MM-dd HH:mm"
	)}`;
	const newSpreadsheet = SpreadsheetApp.create(reportName);

	// Export Data sheet
	const dataLastRow = dataSheet.getLastRow();
	const dataLastCol = dataSheet.getLastColumn();

	const newDataSheet = newSpreadsheet.getActiveSheet();
	newDataSheet.setName("Audit Data");

	if (dataLastRow > 0 && dataLastCol > 0) {
		const dataValues = dataSheet
			.getRange(1, 1, dataLastRow, dataLastCol)
			.getValues();
		if (dataValues.length > 0) {
			newDataSheet
				.getRange(1, 1, dataValues.length, dataValues[0].length)
				.setValues(dataValues);
		}
	} else {
		newDataSheet.getRange("A1").setValue("No data available");
	}

	// Export Audit Log
	const auditLastRow = auditSheet.getLastRow();
	const auditLastCol = auditSheet.getLastColumn();

	const newAuditSheet = newSpreadsheet.insertSheet("Audit Log");

	if (auditLastRow > 0 && auditLastCol > 0) {
		const auditValues = auditSheet
			.getRange(1, 1, auditLastRow, auditLastCol)
			.getValues();
		if (auditValues.length > 0) {
			newAuditSheet
				.getRange(1, 1, auditValues.length, auditValues[0].length)
				.setValues(auditValues);
		}
	} else {
		newAuditSheet.getRange("A1").setValue("No audit log entries");
	}

	// Create Summary
	const summarySheet = newSpreadsheet.insertSheet("Summary");
	const totalClaims = dataLastRow > 1 ? dataLastRow - 1 : 0;

	let approvedCount = 0;
	let needsAddendumCount = 0;
	let pendingCount = 0;

	if (totalClaims > 0) {
		const decisionCol = getColumnIndex(dataSheet, "validator_decision");
		if (decisionCol) {
			const decisions = dataSheet
				.getRange(2, decisionCol, totalClaims, 1)
				.getValues();
			decisions.forEach((row) => {
				const decision = row[0];
				if (decision === "✅ Approved") approvedCount++;
				else if (decision === "⚠ Needs Addendum") needsAddendumCount++;
				else pendingCount++;
			});
		} else {
			pendingCount = totalClaims;
		}
	}

	// FIX: Ensure all rows have exactly 2 columns
	const summary = [
		["Audit Report Summary", ""], // 2 columns
		["Generated:", new Date()], // 2 columns
		["", ""], // 2 columns
		["Statistics", ""], // 2 columns
		["Total Claims:", totalClaims], // 2 columns
		["Approved:", approvedCount], // 2 columns
		["Needs Addendum:", needsAddendumCount], // 2 columns
		["Pending Review:", pendingCount], // 2 columns
		["", ""], // 2 columns
		[
			"Completion Rate:",
			totalClaims > 0
				? (((totalClaims - pendingCount) / totalClaims) * 100).toFixed(1) + "%"
				: "N/A",
		], // 2 columns
	];

	summarySheet.getRange(1, 1, summary.length, 2).setValues(summary);
	summarySheet
		.getRange("A1:B1")
		.merge()
		.setFontSize(16)
		.setFontWeight("bold")
		.setBackground("#4285F4")
		.setFontColor("#FFFFFF");
	summarySheet.getRange("A2:B2").setBackground("#F8F9FA");
	summarySheet.getRange("A4:B4").setFontWeight("bold").setBackground("#E9ECEF");
	summarySheet
		.getRange("A10:B10")
		.setFontWeight("bold")
		.setBackground("#D4EDDA");

	// Format columns
	summarySheet.setColumnWidth(1, 200);
	summarySheet.setColumnWidth(2, 150);
	summarySheet.getRange("A1:B10").setBorder(true, true, true, true, true, true);

	const url = newSpreadsheet.getUrl();

	if (totalClaims === 0) {
		SpreadsheetApp.getUi().alert(
			"Export Complete",
			`Report created (no data found).\n\nURL: ${url}`,
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	} else {
		SpreadsheetApp.getUi().alert(
			"Export Complete",
			`Report created!\n\nTotal Claims: ${totalClaims}\nApproved: ${approvedCount}\nNeeds Addendum: ${needsAddendumCount}\nPending: ${pendingCount}\n\nURL: ${url}`,
			SpreadsheetApp.getUi().ButtonSet.OK
		);
	}
}
function showHelp() {
	const helpText = `CLINICAL AUDIT SYSTEM v2.2 - HELP

📊 SETUP:
1. Initialize System
2. Update Config with lead email
3. Import data via Python backend
4. Dropdowns auto-populate with unvalidated claims

🔍 VALIDATION:
- Dropdown shows ONLY unvalidated claims
- View AI recommendations at top
- LEFT: View clinical note
- RIGHT: Validation table with AI sources
- Fill Pass/Fail and addendum lines
- Submit to remove from validator queue

👤 LEAD REVIEW:
- Dropdown shows ONLY claims needing addendum
- LEFT: Claims dashboard
- MIDDLE: Original note
- RIGHT: Draft addendum with AI findings
- Approve to remove from queue

🤖 AI FEATURES:
- Shows AI-identified issues
- Displays source references
- Auto-populates addendum template

🔄 AUTO-FILTERING:
- Validator sees only NEW claims
- Lead sees only FLAGGED claims
- Completed claims removed automatically

📤 EXPORT:
Click "Export Report" for full audit data`;

	SpreadsheetApp.getUi().alert(
		"Help",
		helpText,
		SpreadsheetApp.getUi().ButtonSet.OK
	);
}

function removeAllProtections() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const leadReviewSheet = ss.getSheetByName("Lead Review");

	// Remove all protections
	const sheetProtections = leadReviewSheet.getProtections(
		SpreadsheetApp.ProtectionType.SHEET
	);
	sheetProtections.forEach((p) => p.remove());

	const rangeProtections = leadReviewSheet.getProtections(
		SpreadsheetApp.ProtectionType.RANGE
	);
	rangeProtections.forEach((p) => p.remove());

	// Clear C2 completely
	leadReviewSheet.getRange("C2").clear();

	Logger.log("All protections removed and C2 cleared");
}
