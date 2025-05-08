/********************************************
* SCOUTHUB CSV IMPORT SCRIPT v3.1
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
  
  ui.createMenu('⚙️ CSV Importer')
    .addItem('🚀 1. Run Full Import', 'importScoutHubCSV')
    .addItem('🚀 2. Run All Transforms', 'runAllTransforms')
    .addItem('🚀 3. Flag Changes', 'compareAndFlagChanges')
    .addSubMenu(
      ui.createMenu('🛠️ Setup & Config')
        .addItem('1. Initialize Config Sheet', 'createConfigSheet')
        .addItem('2. Update Column Mappings', 'updateColumnMappings')
    )
    .addSubMenu(
      ui.createMenu('🔧 Data Transformations')
	  // ALL TRANSFORMATIONS EXIST IN THE SEPARATE "IMPORT SCOUTHUB TRANSFORMATIONS" SCRIPT
        .addItem('🚀 Run All Transforms', 'runAllTransforms')
        .addSeparator()
        .addItem('📞 Normalize Numbers', 'normalizePhoneNumbers')
        .addItem('📞 De-duplicate Mobiles', 'deduplicateMobileNumbers')
        .addItem('📧 De-duplicate Emails', 'deduplicateEmails')
        .addItem('👪 Split Parent Names', 'applyParentNameSplitting')
        .addItem('👪 Populate Preferred Names', 'applyPreferredNamePopulation')
        .addItem('Sort by Member Number', 'sortByMemberNumber')
    )
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
* 3. CONFIG SHEET CREATION
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
    ['CSV Drive URL', ''],
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

/*********************************
* 4. CONFIG SHEET NAME RETRIEVAL
* Reads the current source and target sheet names from the config sheet.
* Returns default names if not set.
*********************************/
function getSheetNamesFromConfig() {
  const ss = SpreadsheetApp.getActive();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) return {
    memberSource: 'Members',
    mappedTarget: 'Scouthub Import'
  };
  const values = configSheet.getRange('B1:B2').getValues().flat();
  return {
    memberSource: values[0] || 'Members',
    mappedTarget: values[1] || 'Scouthub Import'
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
function updateColumnMappings() {
  const ss = SpreadsheetApp.getActive();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();
  const { memberSource } = getSheetNamesFromConfig();

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

    // Get columns from the ScoutHub CSV
    const fileId = extractDriveFileId(configSheet.getRange('B3').getValue());
    if (!fileId) throw new Error('Invalid Google Drive URL in B3');
    const csvData = Utilities.parseCsv(DriveApp.getFileById(fileId).getBlob().getDataAsString());
    const scouthubColumns = csvData[0];

    // Read any existing user mappings from the config sheet
    let existingMappings = {};
    const lastRow = configSheet.getLastRow();
    if (lastRow >= 11) {
      const colAValues = configSheet.getRange('A11:A' + lastRow).getValues().flat();
      const colBValues = configSheet.getRange('B11:B' + lastRow).getValues().flat();
      colAValues.forEach((colA, i) => existingMappings[colA] = colBValues[i]);
    }

    // Clear previous mapping table
    if (lastRow >= 11) configSheet.getRange(`A11:D${lastRow}`).clearContent();

    // Build new mapping table
    const data = [];
    let mappedCount = 0;
    memberColumns.forEach((memberCol, index) => {
      let mapVal = existingMappings[memberCol] || STANDARD_MAPPINGS[memberCol] || '';
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
* 7. CSV IMPORT ENGINE
* Imports mapped data from the ScoutHub CSV to the target sheet.
* - Uses the mapping table for column alignment
* - Writes headers and data to the target sheet
* - Tracks last import date in the config sheet
*********************************/
function importScoutHubCSV() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const { memberSource, mappedTarget } = getSheetNamesFromConfig();

  try {
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    if (!configSheet) throw new Error('Config sheet missing');
    const settings = configSheet.getRange('B3:B5').getValues().flat();
    const fileId = extractDriveFileId(settings[0]);
    const sourceStartRow = parseInt(settings[1]) || 2;
    const targetStartRow = parseInt(settings[2]) || 2;

    const memberSheet = ss.getSheetByName(memberSource);
    if (!memberSheet) throw new Error(`'${memberSource}' sheet missing`);
    const memberHeaders = memberSheet.getRange(1, 1, 1, memberSheet.getLastColumn()).getValues()[0];
    
    // Build mapping from config sheet (A11:B)
    const mappingData = configSheet.getRange('A11:B' + configSheet.getLastRow())
      .getValues()
      .filter(([a,b]) => a && b)
      .reduce((acc, [a,b]) => (acc[a] = b, acc), {});

    const csvData = Utilities.parseCsv(DriveApp.getFileById(fileId).getBlob().getDataAsString());
    const headerIndices = csvData[0].reduce((acc, h, i) => (acc[h] = i, acc), {});

    // Map each row from CSV to the member sheet columns
    const mappedData = csvData.slice(sourceStartRow - 1).map(row => 
      memberHeaders.map(h => mappingData[h] ? row[headerIndices[mappingData[h]]] : '')
    );

    const targetSheet = ss.getSheetByName(mappedTarget) || ss.insertSheet(mappedTarget);
    targetSheet.clear();
    targetSheet.appendRow(memberHeaders);
    if (mappedData.length) {
      targetSheet.getRange(targetStartRow, 1, mappedData.length, memberHeaders.length)
        .setValues(mappedData);
    }

    configSheet.getRange('B6').setValue(new Date());
    ui.alert(`Imported ${mappedData.length} records`);
  } catch (e) {
    ui.alert(`Import Failed: ${e.message}`);
    Logger.log(e);
  }
}

/*********************************
* 8. CHANGE TRACKING SYSTEM
* Compares imported records to the member source sheet.
* - Flags membership number changes and contact updates
* - Highlights changed cells and tags the Action column
* - Ensures robust date and key normalization
*********************************/
function compareAndFlagChanges() {
  const ss = SpreadsheetApp.getActive();
  const { memberSource, mappedTarget } = getSheetNamesFromConfig();
  const ui = SpreadsheetApp.getUi();

  const TRACKED_FIELDS = [
    'Medical Info', 'Member Mobile', 'Member_Email',
    'Additional Email Addresses', 'Parent1_Mobile',
    'Parent1_Email', 'Parent2_Mobile', 'Parent2_Email',
    'Home_Number', 'Address'
  ];
  const HIGHLIGHT_COLOR = '#800080';
  const TEXT_COLOR = '#ffffff';

  // Robust date normalization for composite key
  const normalizeDOB = (dateValue) => {
    if (!dateValue) return '';
    try {
      const date = new Date(dateValue);
      return Utilities.formatDate(date, 'Australia/Sydney', 'yyyy-MM-dd');
    } catch(e) {
      return '';
    }
  };

  // Composite key generator for matching
  const createCompositeKey = (row, headers) => {
    const firstName = (row[headers.indexOf('First Name')] || '').toString().trim().toLowerCase();
    const lastName = (row[headers.indexOf('Last Name')] || '').toString().trim().toLowerCase();
    const dob = normalizeDOB(row[headers.indexOf('D.O.B')]);
    return `${firstName}|${lastName}|${dob}`;
  };

  try {
    const memberSheet = ss.getSheetByName(memberSource);
    const memberData = memberSheet.getDataRange().getDisplayValues();
    const memberHeaders = memberData[0];
    const memberMap = new Map(memberData.slice(1).map(row => [
      createCompositeKey(row, memberHeaders),
      row
    ]));

    const targetSheet = ss.getSheetByName(mappedTarget);
    const targetData = targetSheet.getDataRange().getDisplayValues();
    let targetHeaders = targetData[0];

    let actionColIdx = targetHeaders.indexOf('Action');
    if (actionColIdx === -1) {
      targetSheet.insertColumnAfter(1);
      actionColIdx = 1;
      targetSheet.getRange(1, 2).setValue('Action');
      targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    }

    const actions = [];
    const backgrounds = [];
    const fontColors = [];

    for (let rowIndex = 1; rowIndex < targetData.length; rowIndex++) {
      const row = targetData[rowIndex];
      const key = createCompositeKey(row, targetHeaders);
      const memberRow = memberMap.get(key);
      let actionNotes = [];
      backgrounds[rowIndex] = new Array(targetHeaders.length).fill(null);
      fontColors[rowIndex] = new Array(targetHeaders.length).fill(null);

      // Membership number change detection
      const memberNumIdx = targetHeaders.indexOf('Membership Number');
      if (memberNumIdx > -1 && memberRow) {
        const oldNum = memberRow[memberHeaders.indexOf('Membership Number')];
        const newNum = row[memberNumIdx];
        if (oldNum !== newNum) {
          actionNotes.push('New Number');
          backgrounds[rowIndex][memberNumIdx] = HIGHLIGHT_COLOR;
          fontColors[rowIndex][memberNumIdx] = TEXT_COLOR;
        }
      }

      // Contact field change detection
      let contactChanged = false;
      TRACKED_FIELDS.forEach(field => {
        const targetIdx = targetHeaders.indexOf(field);
        const memberIdx = memberHeaders.indexOf(field);
        if (targetIdx > -1 && memberIdx > -1 && memberRow) {
          if (row[targetIdx] !== memberRow[memberIdx]) {
            contactChanged = true;
            backgrounds[rowIndex][targetIdx] = HIGHLIGHT_COLOR;
            fontColors[rowIndex][targetIdx] = TEXT_COLOR;
          }
        }
      });
      if (contactChanged) actionNotes.push('Updated Contact');

      actions[rowIndex] = [actionNotes.join(', ')];
    }

    // Fill header row for actions and formatting
    actions[0] = [targetHeaders[actionColIdx] || 'Action'];
    backgrounds[0] = new Array(targetHeaders.length).fill(null);
    fontColors[0] = new Array(targetHeaders.length).fill(null);

    // Update Action column
    targetSheet.getRange(1, actionColIdx + 1, actions.length, 1)
      .setValues(actions);

    // Apply formatting for highlighted changes
    for (let colIndex = 0; colIndex < targetHeaders.length; colIndex++) {
      const colBgs = backgrounds.map(row => row[colIndex] || null);
      const colFonts = fontColors.map(row => row[colIndex] || null);
      targetSheet.getRange(1, colIndex + 1, colBgs.length, 1)
        .setBackgrounds(colBgs.map(bg => [bg]))
        .setFontColors(colFonts.map(fc => [fc]));
    }

    ui.alert('Change Tracking Complete');
  } catch (e) {
    ui.alert(`Change Tracking Failed: ${e.message}`);
    Logger.log(e.stack || e.message);
  }
}
