/********************************************
* SCOUTHUB CSV IMPORT SCRIPT v3.2
*
* SUMMARY:
* Imports ScoutHub Member Data from a CSV file.
* Uses the CSV Export from ScoutHub and writes the import to a column structure matching your current member data
* so that you can easily compare and update new members and changes.
*
* Features:
* - Config-driven based on "Import Config" sheet
* - Dynamic column mapping
* - Change Tracking and Field Update detection
* - Last import timestamp tracking
* ******************************************
*
* ScoutHub CSV EXPORT INSTRUCTIONS:
* - Login to your ScoutHub portal as an admin.
* - Navigate to Members > Reporting page
* - Click "Generate Report"
* - Select all tick-box fields you want to export
* - Select "Export seasonal data" from "Current" trek
* - Set "Name formatting" as per your preference. Advice: use "First & Last in separate columns".
* **Save as template**. This allows you to use the same report settings next time.
* - Click GENERATE REPORT
* - Save the ScoutHub CSV and upload to a secure Google Drive location.
* - Copy the Google Drive link to the ScoutHub CSV file.
* - This URL will be pasted into the "CSV Drive URL" field of the "Import Config" sheet in the next steps.
********************************************
*
* GOOGLE SHEETS INITIAL SCRIPT SETUP INSTRUCTIONS:
* From this Google Sheet:
* 1. Open the "Extensions" menu, then select "Apps Script".
* 2. In the Apps Script editor, create a new script file ("+ Add a File"):
*    a. Name the script "ScoutHub CSV Importer.gs".
*    b. Delete any existing default code in the Code.gs file so that it's blank.
*    c. Copy the contents of this script file and paste it into the "ScoutHub CSV Importer.gs" file in the Apps Script editor.
*    d. Click the floppy disk icon or select "File > Save" to save your changes.
* 3. Repeat the process to create a second new script file ("+ Add a File"):
*    a. Name the script "ScoutHub CSV Transforms.gs".
*    b. Copy the contents of the Transforms script file and paste it into the "ScoutHub CSV Transforms.gs" file in the Apps Script editor.
*    c. Click the floppy disk icon or select "File > Save" to save your changes.
* 4. Close the Apps Script editor tab and return to your Google Sheet.
* 5. Reload (refresh) your Google Sheet in the browser to activate the new menu.
* 6. You should now see a "CSV Importer" menu in the Google Sheets menu bar.
*
* NOTE: The first time you use the script, you will be prompted to review and authorize permissions for the script to access your Google Sheets and Drive.
********************************************
*
* CSV IMPORTER INITIAL SETUP INSTRUCTIONS
* From this Google Sheet:
* 1. use the "CSV IMPORT" menu to "Initialise/Reset Import Config" to establish your baseline import settings.
* 2. Use the newly created "Import Config" sheet to define your variables.
* 3. Paste your own Google Drive link to the ScoutHub CSV file into the "CSV Drive URL" field.
* 4. Use the "CSV IMPORT" menu to "Refresh Data Columns and Mappings"
* 5. Check and update your column mappings as needed using the mapping table.
* All configuration and mapping is managed via the 'Import Config' sheet.
* ******************************************
*
* CSV IMPORTER USAGE INSTRUCTIONS:
* Use the "CSV IMPORT" menu in this sheet to: 
* 1. "1. Run CSV Import" to import data from the ScoutHub CSV.
* 2. "2. Run All Data Transformations" to normalise and correct all data fields.
* 2b. Optionally, each Transform operation can be run individually from the "ScoutHub Transforms" sub-menu.
* 3. "3. Flag Changes for Review" to highlight and tag differences between imported and existing member records.
********************************************
********************************************/

/*********************************
* 1. CONFIGURATION CONSTANTS
* Defines the name of the configuration sheet.
* All other sheet names are user-configurable via the config sheet.
*********************************/
const CONFIG_SHEET_NAME = 'Import Config';

/*********************************
* 2. MENU AND INITIALIZATION
* Creates hierarchical menu with:
* - Core import functions
* - Data transformation submenu
* - Visual separators
*********************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('⚙️ ScoutHub CSV Importer')
    .addSubMenu(
      ui.createMenu('🛠️ Setup & Config')
        .addItem('1. Initialize Config Sheet', 'createConfigSheet')
    )
    .addSeparator()
    .addItem('📥 1. Import ScoutHub CSV from URL', 'importScoutHubCSV')
    .addItem('✏️ 2. Get/Refresh Column Mappings', 'updateColumnMappings')
    .addItem('🗺️ 3. Map ScoutHub to your Table Structure ', 'mapCSVData')
    .addItem('⚙️ 4. Run All Data Transforms', 'runAllTransforms')
    .addItem('🚩 5. Flag Changes', 'compareAndFlagChanges')
    .addItem('🆔 6. Update Member Numbers', 'updateMemberNumbers')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('🔧 Data Transformations')
        .addItem('🚀 Run All Transforms', 'runAllTransforms')
        .addSeparator()
        .addItem('📞 Normalize Numbers', 'normalizePhoneNumbers')
        .addItem('📞 De-duplicate Mobiles', 'deduplicateMobileNumbers')
        .addItem('📧 De-duplicate Emails', 'deduplicateEmails')
        .addItem('👪 Split Parent Names', 'applyParentNameSplitting')
        .addItem('👪 Populate Preferred Names', 'applyPreferredNamePopulation')
        .addItem('🔢 Sort by Member Number', 'sortByMemberNumber')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu('⚡ Tools')
        .addItem('Clear All Caches', 'clearTransformCaches')
        .addItem('Reset Cache Configurations', 'resetAllConfigsMenu')
        .addItem('Test Phone Normalization', 'testPhoneNormalization')
        .addItem('View Performance Logs', 'showPerformanceLogs')
    )
    .addSeparator()
    .addItem('📘 Documentation', 'showDocumentation')
    .addToUi();
}

// Helper function stubs for new menu items
function showPerformanceLogs() {
  // Implementation would fetch logs from script properties
  SpreadsheetApp.getUi().alert('Performance Logs', 'Feature not yet implemented', SpreadsheetApp.getUi().ButtonSet.OK);
}

function showDocumentation() {
  const url = 'https://support.scouthub.au/data-importer-guide';
  SpreadsheetApp.getUi().showModelessDialog(
    HtmlService.createHtmlOutput(`<a href="${url}" target="_blank">Open Documentation</a>`),
    'User Guide'
  );
}


/*********************************
* 3. SETUP CONFIG SHEET & OTHERS
* Creates or resets the configuration sheet with:
* - User-editable settings for source/target sheet names and import options
* - Dynamic mapping table for column relationships
* - Auto-detects current member sheet columns for mapping
*********************************/
function createConfigSheet() {
  const ss = SpreadsheetApp.getActive();
  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (configSheet) ss.deleteSheet(configSheet);
  configSheet = ss.insertSheet(CONFIG_SHEET_NAME);

  const { memberSource } = getSheetNamesFromConfig();
  const settings = [
    ['Member Source Sheet Name', 'Members'],
    ['Mapped Members Target Sheet Name', 'Scouthub Import'],
    ['ScoutHub CSV Staging Sheet Name', 'ScoutHub CSV'], // Added new setting for staging sheet
    ['CSV Drive URL', 'https://drive.google.com/open?id=1hw3nTnmj9jY9M_D0a_TWI3xxxxQznoc5&usp=drive_fs'],
    ['Source Start Row', '2'],
    ['Target Start Row', '2'],
    ['Last Import Date', '']
  ];

  // Write settings at the top of the config sheet
  configSheet.getRange(1, 1, settings.length, 2)
    .setValues(settings)
    .setBackground('#f0f8ff');

  // Get columns from the member source sheet for mapping
  const targetSheet = ss.getSheetByName(memberSource);
  if (!targetSheet) throw new Error(`'${memberSource}' sheet not found`);
  const targetHeaders = targetSheet.getDataRange().offset(0, 0, 1).getValues()[0];

  // Write mapping table headers and detected member columns
  configSheet.getRange(10, 1, 1, 4)
    .setValues([['Member Sheet Columns', 'Mapped ScoutHub Columns', 'Apply Transform', 'ScoutHub CSV Columns']])
    .setBackground('#e6e6fa');
  configSheet.getRange(11, 1, targetHeaders.length, 1)
    .setValues(targetHeaders.map(h => [h]));
  configSheet.autoResizeColumns(1, 4);
}

function ensureStagingSheetExists() {
  const ss = SpreadsheetApp.getActive();
  const { stagingSheet } = getSheetNamesFromConfig();
  let sheet = ss.getSheetByName(stagingSheet);
  if (!sheet) {
    sheet = ss.insertSheet(stagingSheet);
    sheet.appendRow(['ScoutHub CSV Staging Sheet - Paste Raw CSV Columns Here']);
  }
  return sheet;
}

/*********************************
* 4. CONFIG SHEET VARIABLE NAME RETRIEVALS
* Reads the current source and target sheet names from the config sheet.
* Returns default names if not set.
*********************************/
function getSheetNamesFromConfig() {
  const ss = SpreadsheetApp.getActive();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) return {
    memberSource: 'Members',
    mappedTarget: 'Scouthub Import',
    stagingSheet: 'ScoutHub CSV'
  };
  const values = configSheet.getRange('B1:B3').getValues().flat();
  return {
    memberSource: values[0] || 'Members',
    mappedTarget: values[1] || 'Scouthub Import',
    stagingSheet: values[2] || 'ScoutHub CSV'
  };
}

function getImportConfigParams() {
  const ss = SpreadsheetApp.getActive();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) throw new Error("Config sheet missing");

  // Define each config cell explicitly
  const SHOW_NUMBER_CHANGE_PROMPTS = String(configSheet.getRange("E1").getValue()).toLowerCase() === 'true';
  const AUTO_UPDATE_NUMBER_CHANGES = String(configSheet.getRange("E2").getValue()).toLowerCase() === 'true';
  const DEBUG_MODE = String(configSheet.getRange("E3").getValue()).toLowerCase() === 'true';
  const DATE_INPUT_FORMAT = configSheet.getRange("E4").getValue() || 'International (AU)';

  return {
    SHOW_NUMBER_CHANGE_PROMPTS,
    AUTO_UPDATE_NUMBER_CHANGES,
    DEBUG_MODE,
    DATE_INPUT_FORMAT
  };
}


/*********************************
* 5. FILE ID EXTRACTION
* Extracts the Google Drive file ID from a variety of URL formats,
* including HYPERLINK formulas and direct file links.
*********************************/
function extractDriveFileId(url) {
  const urlString = String(url || '').trim();
  let cleanUrl = urlString;
  if (urlString.startsWith('=HYPERLINK("')) {
    const match = urlString.match(/=HYPERLINK\("([^"]+)"/);
    if (match) cleanUrl = match[1];
  }
  const idMatch = cleanUrl.match(/[\?&]id=([\w-]{25,})/) || cleanUrl.match(/\/d\/([\w-]{25,})/);
  return idMatch ? idMatch[1] : '';
}

/*********************************
* 6. COLUMN MAPPING ENGINE
* - Uses 'Member Sheet Column': 'ScoutHub Column' format for standard mappings
* - Populates mapping table with standard or existing mappings
* - Highlights mapped columns for user clarity
*********************************/
/**
 * Updates the column mappings in the configuration sheet.
 * 
 * This function:
 * - Reads member sheet columns and ScoutHub CSV staging sheet columns.
 * - Loads existing mappings from the config sheet.
 * - Applies standard mappings only if the corresponding ScoutHub CSV column exists (case-insensitive).
 * - Ensures all column header comparisons are case-insensitive.
 * - Writes the updated mapping table back to the config sheet.
 * - Highlights mapped columns for clarity.
 * 
 * Alerts the user on completion or errors.
 */
function updateColumnMappings() {
  const ss = SpreadsheetApp.getActive();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();
  const { memberSource, stagingSheet } = getSheetNamesFromConfig();

  // DEFAULT COLUMN MAPPINGS
  // Format: 'Member Sheet Column': 'ScoutHub Column'
  const STANDARD_MAPPINGS = {
    'Membership Number': 'Membership Number',
    'First Name': 'First name',
    'Preferred Name': 'Preferred Name',
    'Last Name': 'Last name',
    'D.O.B': 'Date of birth',
    'Gender': 'Gender identity',
    'Address': 'Address',
    'Home_Number': 'Home phone',
    'Member Mobile': 'Mobile phone',
    'Member_Email': 'Email address',
    'Application Date': 'Registered on',
    'ScoutHub Payment Class': 'Payment status',
    'ScoutHub Payment Date': 'Payment date',
    'Record Updated': 'Last updated',
    'Record Added': 'Registered on',
    'Parent1_FName': 'Primary Contact Full Name',
    'Parent1_Relationship': 'PEC Relationship to Member',
    'Parent1_Mobile': 'PEC Primary Contact Number',
    'Parent1_Email': 'PEC Email Address',
    'Parent2_FName': 'Alternate Contact Full Name',
    'Parent2_Relationship': 'AEC Relationship to Member',
    'Parent2_Mobile': 'AEC Primary Contact Number',
    'Parent2_Email': 'AEC Email Address',
    'Parent1_Assistance': 'What assistance are you able to provide to our Scout Group?',
    'Medical Info': 'Allergy/Medical/Mobility Details',
    'Photo Permission': 'Do you give permission for UNNAMED photos and media to be shared on our local group social media?'
  };

  try {
    // Get columns from member sheet
    const memberSheet = ss.getSheetByName(memberSource);
    if (!memberSheet) throw new Error(`'${memberSource}' sheet not found`);
    const memberColumns = memberSheet.getRange(1, 1, 1, memberSheet.getLastColumn()).getValues()[0];

    // Get columns from the ScoutHub CSV staging sheet
    const stgSheet = ss.getSheetByName(stagingSheet);
    if (!stgSheet) {
      ui.alert(`The staging sheet "${stagingSheet}" does not exist. Please import or paste CSV data first.`);
      return;
    }
    const stgData = stgSheet.getDataRange().getValues();
    if (stgData.length < 1) {
      ui.alert(`The staging sheet "${stagingSheet}" is empty. Please import or paste CSV data first.`);
      return;
    }
    const scouthubColumns = stgData[0];
    const scouthubLower = scouthubColumns.map(c => c.toLowerCase());

    // Load existing mappings from config sheet
    let existingMappings = {};
    const lastRow = configSheet.getLastRow();
    if (lastRow >= 11) {
      const colAValues = configSheet.getRange('A11:A' + lastRow).getValues().flat();
      const colBValues = configSheet.getRange('B11:B' + lastRow).getValues().flat();
      colAValues.forEach((colA, i) => existingMappings[colA] = colBValues[i]);
    }

    // Clear previous mapping table content
    if (lastRow >= 11) configSheet.getRange(`A11:D${lastRow}`).clearContent();

    // Build new mapping table with validation
    const data = [];
    let mappedCount = 0;
    memberColumns.forEach((memberCol, index) => {
      let mapVal = existingMappings[memberCol] || '';

      // If no existing mapping, try standard mapping only if it exists in ScoutHub CSV columns (case-insensitive)
      if (!mapVal && STANDARD_MAPPINGS[memberCol]) {
        const stdMap = STANDARD_MAPPINGS[memberCol];
        if (scouthubLower.includes(stdMap.toLowerCase())) {
          mapVal = stdMap;
        }
      }

      // Validate that mapping exists in ScoutHub CSV columns (case-insensitive)
      if (mapVal && !scouthubLower.includes(mapVal.toLowerCase())) {
        mapVal = '';
      }

      if (mapVal) mappedCount++;
      data.push([memberCol, mapVal, '', scouthubColumns[index] || '']);
    });

    // Add any extra ScoutHub columns not mapped
    scouthubColumns.slice(memberColumns.length).forEach(scoutCol => {
      data.push(['', '', '', scoutCol]);
    });

    // Write mapping table to config sheet
    if (data.length) configSheet.getRange(11, 1, data.length, 4).setValues(data);

    // Highlight mapped columns in the mapping table for user clarity
    let rules = configSheet.getConditionalFormatRules();
    rules = rules.filter(r => !r.getRanges().some(rg => rg.getColumn() === 4));
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=COUNTIF($B$11:$B$${10 + data.length}, D11) > 0`)
      .setBackground('#c6efce')
      .setRanges([configSheet.getRange('D11:D' + (10 + data.length))])
      .build();
    configSheet.setConditionalFormatRules([...rules, rule]);

    ui.alert(`Mappings Updated\n${memberColumns.length} Member Columns\n${scouthubColumns.length} CSV Columns\n${mappedCount} Mapped`);
  } catch (e) {
    ui.alert(`Error: ${e.message}`);
    Logger.log(e);
  }
}


/*********************************
 * 7. CSV DATA MAPPING TO TARGET

 * Maps raw ScoutHub CSV data from the staging worksheet to the member target sheet.
 *
 * Reads raw CSV data from the configured staging sheet,
 * applies the column mappings defined in the config sheet,
 * and writes the mapped data to the target member sheet.
 * 
 * Matching of column headers is case-insensitive.
 * 
 * Updates the last import timestamp in config on success.
 *********************************/
function mapCSVData() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const { memberSource, mappedTarget, stagingSheet } = getSheetNamesFromConfig();

  try {
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    if (!configSheet) throw new Error('Config sheet missing');

    const stgSheet = ss.getSheetByName(stagingSheet);
    if (!stgSheet) {
      ui.alert(`The staging sheet "${stagingSheet}" does not exist. Please import or paste CSV data first.`);
      return;
    }

    const stgData = stgSheet.getDataRange().getValues();
    if (stgData.length < 1) {
      ui.alert(`The staging sheet "${stagingSheet}" is empty. Please import or paste CSV data first.`);
      return;
    }

    const stgHeaders = stgData[0];
    const headerIndices = stgHeaders.reduce((acc, h, i) => {
      acc[h.toLowerCase()] = i;
      return acc;
    }, {});

    const memberSheet = ss.getSheetByName(memberSource);
    if (!memberSheet) throw new Error(`'${memberSource}' sheet not found`);
    const memberHeaders = memberSheet.getRange(1, 1, 1, memberSheet.getLastColumn()).getValues()[0];

    const lastRow = configSheet.getLastRow();
    const mappingRange = configSheet.getRange(`A11:B${lastRow}`);
    const mappingValues = mappingRange.getValues().filter(([a, b]) => a && b);
    const mappingData = {};
    mappingValues.forEach(([memberCol, scoutCol]) => {
      mappingData[memberCol] = scoutCol;
    });

    const mappedData = stgData.slice(1).map(row => 
      memberHeaders.map(memberCol => {
        const scoutCol = mappingData[memberCol];
        if (!scoutCol) return '';
        const colIndex = headerIndices[scoutCol.toLowerCase()];
        return (colIndex !== undefined) ? row[colIndex] : '';
      })
    );

    let targetSheet = ss.getSheetByName(mappedTarget);
    if (!targetSheet) targetSheet = ss.insertSheet(mappedTarget);
    targetSheet.clear();
    targetSheet.appendRow(memberHeaders);
    if (mappedData.length) {
      targetSheet.getRange(2, 1, mappedData.length, memberHeaders.length).setValues(mappedData);
    }

    configSheet.getRange('B7').setValue(new Date());
    ui.alert(`Mapped ${mappedData.length} records from "${stagingSheet}" to "${mappedTarget}".`);
  } catch (e) {
    ui.alert(`Mapping Failed: ${e.message}`);
    Logger.log(e);
  }
}




/*********************************
* 7. CSV IMPORT ENGINE
* Imports mapped data from the ScoutHub CSV to the target sheet.
* - Uses the mapping table for column alignment
* - Writes headers and data to the target sheet
* - Tracks last import date in the config sheet
*********************************/
function importScoutHubCSV() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const { stagingSheet } = getSheetNamesFromConfig();

  try {
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    if (!configSheet) throw new Error('Config sheet missing');
    
    // Get CSV file URL
    const fileUrl = configSheet.getRange('B4').getValue();
    const fileId = extractDriveFileId(fileUrl);
    
    if (!fileId) {
      ui.alert('No CSV URL', 
              'No CSV Drive URL is defined in the config. Please either:\n' +
              '1. Add a Google Drive URL to the CSV file in the config, or\n' +
              '2. Manually paste your CSV data into the "' + stagingSheet + '" sheet.',
              ui.ButtonSet.OK);
      
      // Ensure staging sheet exists and activate it
      const stgSheet = ensureStagingSheetExists();
      stgSheet.activate();
      return;
    }
    
    // Get the CSV data from Drive
    try {
      const csvContent = DriveApp.getFileById(fileId).getBlob().getDataAsString();
      const csvData = Utilities.parseCsv(csvContent);
      
      // Ensure staging sheet exists
      const stgSheet = ensureStagingSheetExists();
      
      // Clear and populate staging sheet
      stgSheet.clear();
      if (csvData.length > 0) {
        stgSheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
        ui.alert(`CSV imported successfully`, 
                `Data from Drive has been imported to "${stagingSheet}" sheet.\n\n` +
                `Next step: Run "Map CSV Data" to process the imported data.`,
                ui.ButtonSet.OK);
      } else {
        throw new Error('CSV file appears to be empty');
      }
    } catch (csvError) {
      throw new Error(`Failed to import CSV: ${csvError.message}`);
    }
  } catch (e) {
    ui.alert(`Import Failed: ${e.message}`);
    Logger.log(e);
  }
}


/*********************************
* 8. CHANGE TRACKING SYSTEM
* Compares imported records to the member source sheet.
* - Flags membership number changes and tracked field updates
* - Highlights changed cells and tags the Action column
* - Normalizes dates (Excel serials, AU/ISO formats)
* - Groups all debug output into a single log block per member
* Last updated: 2025-05-08
*********************************/
function compareAndFlagChanges() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const { memberSource, mappedTarget } = getSheetNamesFromConfig();

  // Read configuration from Import Config sheet
  const config = getImportConfigParams();
  const SHOW_NUMBER_CHANGE_PROMPTS = config.SHOW_NUMBER_CHANGE_PROMPTS;
  const AUTO_UPDATE_NUMBER_CHANGES = config.AUTO_UPDATE_NUMBER_CHANGES;
  const DEBUG_MODE = config.DEBUG_MODE;
  const DATE_INPUT_FORMAT = config.DATE_INPUT_FORMAT;
  const TARGET_TIMEZONE = ss.getSpreadsheetTimeZone();

  const TRACKED_FIELDS = [
    'Medical Info', 'Member Mobile', 'Member_Email', 'Additional Email Addresses', 
    'Parent1_Mobile', 'Parent1_Email', 'Parent2_Mobile', 'Parent2_Email', 'Address'
  ];
  const HIGHLIGHT_COLOR = '#800080';
  const TEXT_COLOR = '#ffffff';
  const NEW_MEMBER_COLOR = '#e6ffec';
  const UPDATED_NUMBER_COLOR = '#800080';

  // Composite key generator with debug logging
  const createCompositeKey = (row, headers, context, debugLines) => {
    const getValue = (field) => {
      const idx = headers.indexOf(field);
      if (idx === -1) return '';
      const val = row[idx];
      return typeof val === 'string' 
        ? val.trim().normalize('NFKC').replace(/\s+/g, ' ').toLowerCase()
        : String(val).trim().toLowerCase();
    };

    const dob = normalizeDate(
      row[headers.indexOf('D.O.B')],
      context,
      debugLines,
      DATE_INPUT_FORMAT,
      TARGET_TIMEZONE
    );
    const firstName = getValue('First Name');
    const lastName = getValue('Last Name');
    const key = `${firstName}|${lastName}|${dob}`;
    debugLines && debugLines.push(`[${context}] Composite key: ${key}`);
    return {
      membershipId: String(getValue('Membership Number')).trim(),
      nameDob: key
    };
  };

  try {
    const memberSheet = ss.getSheetByName(memberSource);
    const memberData = memberSheet.getDataRange().getValues();
    const memberHeaders = memberData[0];

    // Build multi-key map
    const memberMap = new Map();
    memberData.slice(1).forEach((row, idx) => {
      let debugLines = [];
      const keys = createCompositeKey(row, memberHeaders, 'SOURCE', debugLines);
      const rowNum = idx + 2;
      if (keys.membershipId) {
        memberMap.set(`id:${keys.membershipId}`, { row, rowNum, keys });
      }
      memberMap.set(keys.nameDob, { row, rowNum, keys });
      if (DEBUG_MODE) {
        debugLines.unshift(`--- [SOURCE] Row ${rowNum} ---`);
        console.log(debugLines.join('\n'));
      }
    });

    const targetSheet = ss.getSheetByName(mappedTarget);
    const targetData = targetSheet.getDataRange().getValues();
    let targetHeaders = targetData[0];

    // Ensure Action column exists
    let actionColIdx = targetHeaders.indexOf('Action');
    if (actionColIdx === -1) {
      targetSheet.insertColumnAfter(targetHeaders.length);
      actionColIdx = targetHeaders.length;
      targetSheet.getRange(1, actionColIdx + 1).setValue('Action');
      targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    }

    const actions = [];
    const backgrounds = [];
    const fontColors = [];
    const numberChanges = [];

    for (let rowIndex = 1; rowIndex < targetData.length; rowIndex++) {
      let debugLines = [];
      const row = targetData[rowIndex];
      const targetKeys = createCompositeKey(row, targetHeaders, 'TARGET', debugLines);
      let memberMatch = null;
      let matchType = '';

      // 1. Priority match by exact membership ID
      if (targetKeys.membershipId) {
        memberMatch = memberMap.get(`id:${targetKeys.membershipId}`);
        matchType = 'id';
      }

      // 2. Fallback to name/DOB match
      if (!memberMatch) {
        memberMatch = memberMap.get(targetKeys.nameDob);
        matchType = 'nameDob';
      }

      backgrounds[rowIndex] = new Array(targetHeaders.length).fill(null);
      fontColors[rowIndex] = new Array(targetHeaders.length).fill(null);
      let actionNotes = [];

      if (memberMatch) {
        debugLines.push(`[MATCH] Source key: ${memberMatch.keys.nameDob}`);
        const sourceId = String(memberMatch.row[memberHeaders.indexOf('Membership Number')]).trim();
        const targetId = String(targetKeys.membershipId).trim();

        // ID change detection
        if (sourceId !== targetId && matchType === 'nameDob') {
          debugLines.push(`[ID CHANGE] ${sourceId} -> ${targetId}`);
          debugLines.push(`[ID CHANGE] Source key: ${memberMatch.keys.nameDob}`);
          debugLines.push(`[ID CHANGE] Target key: ${targetKeys.nameDob}`);
          numberChanges.push({
            sourceId,
            targetId,
            name: `${row[targetHeaders.indexOf('First Name')]} ${row[targetHeaders.indexOf('Last Name')]}`,
            rowNum: memberMatch.rowNum
          });
          actionNotes.push('Member ID Changed');
          const idCol = targetHeaders.indexOf('Membership Number');
          backgrounds[rowIndex][idCol] = UPDATED_NUMBER_COLOR;
          fontColors[rowIndex][idCol] = TEXT_COLOR;
        }

        // Field changes
        let changesDetected = false;
        TRACKED_FIELDS.forEach(field => {
          const targetIdx = targetHeaders.indexOf(field);
          const sourceIdx = memberHeaders.indexOf(field);
          if (targetIdx > -1 && sourceIdx > -1) {
            const sourceVal = memberMatch.row[sourceIdx];
            const targetVal = row[targetIdx];
            const normalize = (val) => {
              if (val instanceof Date) return val.getTime();
              return String(val).trim().toLowerCase();
            };
            if (normalize(sourceVal) !== normalize(targetVal)) {
              changesDetected = true;
              debugLines.push(`[FIELD CHANGE] ${field}: "${sourceVal}" -> "${targetVal}"`);
              backgrounds[rowIndex][targetIdx] = HIGHLIGHT_COLOR;
              fontColors[rowIndex][targetIdx] = TEXT_COLOR;
            }
          }
        });
        if (changesDetected) actionNotes.push('Field Updates');
      } else {
        debugLines.push('[NO MATCH] No member found for this composite key.');
        actionNotes.push('New Member');
        backgrounds[rowIndex].fill(NEW_MEMBER_COLOR);
      }

      actions[rowIndex] = [actionNotes.join(', ')];
      if (DEBUG_MODE) {
        debugLines.unshift(`--- [TARGET] Row ${rowIndex + 1} ---`);
        console.log(debugLines.join('\n'));
      }
    }

    // Batch process number changes at the end
    if (numberChanges.length > 0 && SHOW_NUMBER_CHANGE_PROMPTS) {
      const changeList = numberChanges.map((change, index) => 
        `\n${index + 1}. ${change.name}: ${change.sourceId} → ${change.targetId}`
      ).join('');

      const response = ui.alert(
        'Membership Number Changes Detected',
        `Found ${numberChanges.length} number changes:${changeList}\n\nApply all changes?`,
        ui.ButtonSet.YES_NO
      );

      if (response === ui.Button.YES) {
        updateMemberNumbers(memberSource, memberHeaders, numberChanges);
      }
    }

    // Apply formatting
    actions[0] = [targetHeaders[actionColIdx] || 'Action'];
    backgrounds[0] = new Array(targetHeaders.length).fill(null);
    fontColors[0] = new Array(targetHeaders.length).fill(null);

    targetSheet.getRange(1, actionColIdx + 1, actions.length, 1).setValues(actions);
    const lastRow = targetData.length;
    const lastCol = targetHeaders.length;
    targetSheet.getRange(2, 1, lastRow - 1, lastCol)
      .setBackgrounds(backgrounds.slice(1))
      .setFontColors(fontColors.slice(1));

    ui.alert(`${targetData.length - 1} Records Processed.\n${numberChanges.length} ID changes detected.`);
  } catch (e) {
    ui.alert(`Error: ${e.message}`);
    console.error(e);
    Logger.log(e.stack || e);
  }
}

/*********************************
* DATE NORMALISATION UTILITIES
* Handles multiple date formats and regions:
* - Excel serial dates
* - ISO, AU, and US formats
* - Fallback to JavaScript Date parsing
* Configurable via parameters for input format and timezone
*********************************/
function normalizeDate(
  dateValue, 
  context = '', 
  debugLines = null, 
  dateInputFormat = 'International (AU)', 
  targetTimezone = 'Australia/Sydney'
) {
  if (!dateValue) return '';
  try {
    debugLines && debugLines.push(`[${context}] Raw date value: ${dateValue} (${typeof dateValue})`);

    // 1. Handle Excel serial numbers
    if (typeof dateValue === 'number') {
      const excelEpoch = new Date('1899-12-30T00:00:00Z');
      const jsDate = new Date(excelEpoch.getTime() + Math.round(dateValue) * 86400000);
      const isoDate = Utilities.formatDate(jsDate, targetTimezone, 'yyyy-MM-dd');
      debugLines && debugLines.push(`[${context}] Excel serial ${dateValue} → ${isoDate}`);
      return isoDate;
    }

    // 2. Handle string dates based on configured format
    if (typeof dateValue === 'string') {
      let match, year, month, day;
      
      // ISO 8601 Format (YYYY-MM-DD)
      if (dateInputFormat === 'ISO') {
        match = dateValue.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
        if (match) ([, year, month, day] = match);
      }
      // International/AU Format (DD/MM/YYYY or DD/MM/YY)
      else if (dateInputFormat === 'International (AU)') {
        match = dateValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})$/);
        if (match) ([, day, month, year] = match);
      }
      // US Format (MM/DD/YYYY or MM/DD/YY)
      else if (dateInputFormat === 'US') {
        match = dateValue.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4}|\d{2})$/);
        if (match) ([, month, day, year] = match);
      }

      if (match) {
        // Handle 2-digit years
        year = year.length === 2 ? `20${year}` : year;
        // Pad single-digit months/days
        month = month.padStart(2, '0');
        day = day.padStart(2, '0');
        const isoDate = `${year}-${month}-${day}`;
        debugLines && debugLines.push(`[${context}] Formatted ${dateInputFormat} date: ${dateValue} → ${isoDate}`);
        return isoDate;
      }
    }

    // 3. Fallback to JS Date parsing
    const date = dateValue instanceof Date ? dateValue : new Date(dateValue);
    const isoDate = Utilities.formatDate(date, targetTimezone, 'yyyy-MM-dd');
    if (isNaN(date.getTime())) throw new Error('Invalid Date');
    debugLines && debugLines.push(`[${context}] Parsed fallback date: ${dateValue} → ${isoDate}`);
    return isoDate;
  } catch(e) {
    debugLines && debugLines.push(`[${context}] Date error: ${dateValue} → ${e.message}`);
    return 'invalid-date';
  }
}

/*********************************
 * UPDATE MEMBER NUMBERS
 * Sub-function to update member numbers in the source sheet based on the provided changes array.
 * @param {string} memberSource - The name of the member source sheet.
 * @param {Array} memberHeaders - The headers of the member source sheet.
 * @param {Array} numberChanges - Array of objects: {rowNum, targetId, name, sourceId}
*********************************/
function updateMemberNumbers(memberSource, memberHeaders, numberChanges) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const memberSheet = ss.getSheetByName(memberSource);
  let updateCount = 0;
  let changeList = '';
  
  try {
    numberChanges.forEach((change, idx) => {
      memberSheet.getRange(change.rowNum, memberHeaders.indexOf('Membership Number') + 1)
        .setValue(change.targetId);
      updateCount++;
      changeList += `\n${idx + 1}. ${change.name}: ${change.sourceId} → ${change.targetId}`;
    });
    ui.alert(`Successfully updated ${updateCount} membership numbers:${changeList}`);
  } catch (e) {
    ui.alert(`Member number update failed: ${e.message}`);
    Logger.log(e);
  }
}

