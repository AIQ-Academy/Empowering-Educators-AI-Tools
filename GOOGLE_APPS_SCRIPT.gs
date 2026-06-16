// ════════════════════════════════════════════════════════════════════════════════
// AIQ ACADEMY - REGISTRATION FORM HANDLER
// Google Apps Script for saving student registrations to Google Sheets
// ════════════════════════════════════════════════════════════════════════════════

// ── CONFIGURATION ──
// Update these values with your own Google Sheet ID and sheet name
const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID'; // Replace with your Google Sheet ID
const SHEET_NAME = 'Registrations'; // Name of the sheet where data will be saved

// ── MAIN HANDLER ──
function doPost(e) {
  try {
    // Parse the incoming data
    const data = JSON.parse(e.postData.contents);
    
    // Get the spreadsheet and sheet
    const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    
    // If sheet doesn't exist, create it with headers
    if (!sheet) {
      const newSheet = spreadsheet.insertSheet(SHEET_NAME);
      addHeaders(newSheet);
      sheet = newSheet;
    }
    
    // Append the registration data
    const row = [
      data.timestamp || new Date().toISOString(),
      data.fullName || '',
      data.age || '',
      data.class || '',
      data.location || '',
      data.phone || '',
      data.email || '',
      data.courseName || ''
    ];
    
    sheet.appendRow(row);
    
    // Return success response
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Registration saved successfully'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    // Return error response
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── ADD HEADERS TO NEW SHEET ──
function addHeaders(sheet) {
  const headers = [
    'Timestamp',
    'Full Name',
    'Age',
    'Class / Grade',
    'Location / School',
    'Phone Number',
    'Email Address',
    'Course Name'
  ];
  
  sheet.appendRow(headers);
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a3fa6');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  
  // Set column widths
  sheet.setColumnWidth(1, 180); // Timestamp
  sheet.setColumnWidth(2, 150); // Full Name
  sheet.setColumnWidth(3, 80);  // Age
  sheet.setColumnWidth(4, 120); // Class
  sheet.setColumnWidth(5, 150); // Location
  sheet.setColumnWidth(6, 130); // Phone
  sheet.setColumnWidth(7, 180); // Email
  sheet.setColumnWidth(8, 200); // Course Name
}

// ── OPTIONAL: GET ALL REGISTRATIONS ──
function getAllRegistrations() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const registrations = [];
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    registrations.push({
      timestamp: data[i][0],
      fullName: data[i][1],
      age: data[i][2],
      class: data[i][3],
      location: data[i][4],
      phone: data[i][5],
      email: data[i][6],
      courseName: data[i][7]
    });
  }
  
  return registrations;
}

// ── OPTIONAL: GET REGISTRATION COUNT ──
function getRegistrationCount() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) return 0;
  
  return sheet.getLastRow() - 1; // Subtract 1 for header row
}
