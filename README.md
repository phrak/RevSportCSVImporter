# ScoutHub CSV Importer & Transforms Module

Welcome to the **ScoutHub CSV Importer & Transforms Module** - a powerful Google Sheets-based toolset designed to streamline importing, transforming, and managing your ScoutHub member data with ease.

---

## Overview

This project consists of two main components:

1. **ScoutHub CSV Import Script v3.1**  
   Imports ScoutHub member data from CSV exports into your Google Sheet, aligning columns to your existing member data for easy comparison and updates.

2. **ScoutHub CSV Import Transforms Module v4.0**  
   Provides data transformation utilities to clean, normalize, and enhance the imported member data with batch operations and error handling.

Together, these components help you maintain accurate, up-to-date membership records with minimal manual effort.

---

## Features

### CSV Import Script

- Config-driven import based on your **Import Config** sheet  
- Dynamic column mapping for flexible data alignment  
- Change tracking and field update detection  
- Last import timestamp tracking  

### Transforms Module

- Configurable target sheet (defined in `Import Config!B2`)  
- Customizable column names and settings in the CONFIG section  
- Single-click batch processing for:  
  - Splitting parent names into first and last names  
  - Automatically populating Preferred Name fields  
  - Sorting by Membership Number  
  - Normalizing phone numbers to Australian formats  
  - Highlighting invalid phone numbers (red cell, white text)  
  - De-duplicating mobile numbers and email addresses  
- Robust error handling with user-friendly alerts and logs  
- Built-in test utility for phone number normalization  
- Performance monitoring for long-running operations  

---

## Getting Started

### Step 1: Export ScoutHub Member Data CSV

1. Log in to your ScoutHub portal as an admin.  
2. Navigate to **Members > Reporting**.  
3. Click **Generate Report**.  
4. Select all tick-box fields you want to export.  
5. Choose **Export seasonal data** from the "Current" trek.  
6. Set **Name formatting** (recommended: "First & Last in separate columns").  
7. **Save as template** for future exports.  
8. Click **GENERATE REPORT**.  
9. Download the CSV file and upload it to a secure Google Drive folder.  
10. Copy the Google Drive shareable link for the CSV file - you will need this in the next steps.

---

### Step 2: Set Up the Google Sheets Scripts

1. Open your Google Sheet where you want to import and manage ScoutHub data.  
2. Go to **Extensions > Apps Script**.  
3. Create a new script file:  
   - Name it `ScoutHub CSV Importer.gs`.  
   - Delete any existing code and paste the **CSV Import Script** code.  
   - Save the file.  
4. Create another new script file:  
   - Name it `ScoutHub CSV Transforms.gs`.  
   - Paste the **Transforms Module** code.  
   - Save the file.  
5. Close the Apps Script editor and refresh your Google Sheet.  
6. You should now see a new menu called **CSV Importer** in the sheet's menu bar.  

> **Note:** On first use, you will be prompted to authorize the script to access your Google Sheets and Drive.

---

### Step 3: Initial Configuration

1. From the **CSV IMPORT** menu, select **Initialise/Reset Import Config** to create your baseline import settings.  
2. Go to the newly created **Import Config** sheet.  
3. Paste your Google Drive CSV file link into the **CSV Drive URL** field.  
4. Use the **CSV IMPORT** menu to **Refresh Data Columns and Mappings**.  
5. Review and adjust column mappings as needed in the **Import Config** sheet.  

---

### Step 4: Importing and Transforming Data

1. Use the **CSV IMPORT** menu to **Run CSV Import** - this pulls your ScoutHub CSV data into the sheet.  
2. Run **All Data Transformations** from the same menu to clean and normalize your data.  
   - Alternatively, run individual transforms from the **ScoutHub Transforms** submenu.  
3. Use **Flag Changes for Review** to highlight differences between imported data and existing member records for easy review.

---

## Tips & Best Practices

- Always save your ScoutHub report as a template to maintain consistent export settings.  
- Keep your Google Drive CSV file in a secure, accessible location.  
- Regularly refresh and review your import configuration to accommodate any changes in your ScoutHub export format.  
- Use the built-in test utilities in the Transforms module to validate phone number normalization and other data transformations.  

---

## Support & Contributions

If you encounter issues or have suggestions for improvements, please open an issue or submit a pull request. Contributions are welcome!

---
