/**
 * Aerial Fiber Attachment Inspection Form — Apps Script backend
 *
 * SETUP:
 * 1. Create (or open) the Google Sheet that will store submissions.
 * 2. Copy its Spreadsheet ID (the long string in the URL between /d/ and /edit)
 *    and paste it into SPREADSHEET_ID below.
 * 3. Create a Google Drive folder for photos and get its Folder ID
 *    (right-click folder > Get link > copy ID from URL)
 *    and paste it into PHOTOS_FOLDER_ID below.
 * 4. In the Apps Script editor: Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone within [your organization]  (G-Suite / Workspace only)
 * 5. Share the resulting URL with inspectors.
 */

const SPREADSHEET_ID = '1eeLkCnbLcWRVjsbbcrHvHZOyBYNb022yrFZE-47SEX4';
const SHEET_NAME = 'Aerial QC Reports';
const PHOTOS_FOLDER_ID = '11CjcCm_-ko5WMBph96MlQHLrK0Ic0JnJ';

// Codes pulled from the "Failure Codes" reference sheet (codes only, no descriptions)
const FAILURE_CODES = [
  'PC1','PC2','PC3','PC4',
  'MC1','MC2','MC3',
  'SL1','SL2',
  'V1','R1',
  'C1','C2','RC1','C3','C4','C5','C6',
  'DM3','M4',
  'G1','G2','G3','G4',
  'GR','L','B1','AC1','TC1',
  'CS1','CS2'
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Aerial Fiber Inspection Form')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Lets Index.html pull in shared HTML/CSS/JS partials if you split them later
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Called from the client to populate the dropdowns
function getFailureCodes() {
  return FAILURE_CODES;
}

/**
 * Called from the client to upload a photo
 * dataUrl: base64 encoded image data
 * fileName: new filename for the photo (e.g., "Pole_123.jpg")
 */
function uploadPhoto(dataUrl, fileName) {
  if (!PHOTOS_FOLDER_ID || PHOTOS_FOLDER_ID === 'PASTE_YOUR_PHOTOS_FOLDER_ID_HERE') {
    throw new Error('Photos folder ID not configured.');
  }

  try {
    const folder = DriveApp.getFolderById(PHOTOS_FOLDER_ID);
    
    // Parse the data URL
    const matches = dataUrl.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
    if (!matches || matches.length !== 3) {
      throw new Error('Invalid image data format.');
    }

    const contentType = matches[1];
    const bytes = Utilities.base64Decode(matches[2]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);

    // Upload to Drive folder
    const file = folder.createFile(blob);
    return { success: true, fileId: file.getId(), url: file.getUrl() };
  } catch (error) {
    throw new Error('Failed to upload photo: ' + error.message);
  }
}

/**
 * Called from the client on Submit.
 * inspectorName: string
 * inspectionDate: string (e.g. '2026-09-04')
 * poles: [{ poleNumber: '123', codes: ['PC1','','','','',''], comment: 'text', hasPhoto: true }, ...]
 */
function submitInspection(inspectorName, inspectionDate, poles) {
  Logger.log('submitInspection called with:');
  Logger.log('  inspectorName: ' + inspectorName);
  Logger.log('  inspectionDate: ' + inspectionDate);
  Logger.log('  poles count: ' + poles.length);
  
  if (!poles || !poles.length) {
    throw new Error('No pole data received.');
  }
  if (!inspectorName) {
    throw new Error('Inspector name is required.');
  }
  if (!inspectionDate) {
    throw new Error('Inspection date is required.');
  }

  try {
    Logger.log('Opening spreadsheet...');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('Spreadsheet opened successfully');
    
    Logger.log('Looking for sheet: ' + SHEET_NAME);
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      Logger.log('Sheet not found, creating new sheet');
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        'Timestamp', 'Inspector Name', 'Inspection Date', 'Pole Number',
        'Code 1', 'Code 2', 'Code 3', 'Code 4', 'Code 5', 'Code 6', 'Comment', 'Photo URL'
      ]);
      sheet.setFrozenRows(1);
      Logger.log('Sheet created with headers');
    } else {
      Logger.log('Sheet found');
    }

    Logger.log('Building row data...');
    const timestamp = new Date();
    const rows = [];
    
    for (let i = 0; i < poles.length; i++) {
      const pole = poles[i];
      const codes = pole.codes || [];
      
      // Pad codes to 6 entries
      const paddedCodes = [];
      for (let j = 0; j < 6; j++) {
        paddedCodes.push(codes[j] || '');
      }
      
      // Check if there was a photo attached
      const photoUrl = pole.hasPhoto ? '(Photo: Pole_' + pole.poleNumber + ')' : '';
      
      const row = [
        timestamp,
        inspectorName,
        inspectionDate,
        pole.poleNumber || ''
      ];
      
      // Add padded codes
      for (let j = 0; j < paddedCodes.length; j++) {
        row.push(paddedCodes[j]);
      }
      
      // Add comment and photo URL
      row.push(pole.comment || '');
      row.push(photoUrl);
      
      rows.push(row);
    }

    Logger.log('Rows prepared: ' + rows.length + ' row(s)');
    
    if (rows.length > 0) {
      const lastRow = sheet.getLastRow();
      Logger.log('Last row in sheet: ' + lastRow);
      const startRow = lastRow + 1;
      Logger.log('Writing to range starting at row: ' + startRow);
      
      sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
      Logger.log('Data written successfully');
    }

    Logger.log('Submit complete');
    return { success: true, count: rows.length };
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    throw new Error('Submission error: ' + error.message);
  }
}