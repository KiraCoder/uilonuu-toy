/**
 * Google Apps Script for Wedding RSVP Form
 *
 * This script receives RSVP submissions from the wedding invitation website
 * and saves them to a Google Spreadsheet.
 *
 * Setup Instructions:
 * 1. Create a new Google Spreadsheet
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code and paste this entire file
 * 4. Click the disk icon to save
 * 5. Click "Deploy" > "New deployment"
 * 6. Select "Web app" as deployment type
 * 7. Set "Execute as" to "Me"
 * 8. Set "Who has access" to "Anyone"
 * 9. Click "Deploy"
 * 10. Copy the Web App URL
 * 11. Paste the URL into index.html where it says 'YOUR_GOOGLE_SCRIPT_URL_HERE'
 */

// The name of the sheet where RSVPs will be stored
const SHEET_NAME = 'RSVP Responses';

/**
 * Handles POST requests from the wedding invitation form
 */
function doPost(e) {
  try {
    // Parse the incoming JSON data
    const data = JSON.parse(e.postData.contents);

    // Get or create the spreadsheet sheet
    const sheet = getOrCreateSheet();

    // Prepare the row data
    const timestamp = new Date(data.timestamp);
    const name = data.name || '';
    const rsvpStatus = getRsvpLabel(data.rsvp);
    const userAgent = data.userAgent || '';

    // Append the data to the sheet
    sheet.appendRow([
      timestamp,
      name,
      rsvpStatus,
      data.rsvp,
      userAgent
    ]);

    // Return success response
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: 'RSVP received successfully'
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // Log error for debugging
    console.error('Error processing RSVP:', error);

    // Return error response
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'Error processing RSVP: ' + error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles GET requests (optional - for testing)
 */
function doGet(e) {
  return ContentService
    .createTextOutput('Wedding RSVP Form Backend is running!')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Get or create the RSVP sheet
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  // If sheet doesn't exist, create it
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);

    // Set up headers
    const headers = [
      'Timestamp',
      'Name / Аты-жөнү',
      'Response / Жооп',
      'Response Code',
      'User Agent'
    ];

    sheet.appendRow(headers);

    // Format the header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4C463F');
    headerRange.setFontColor('#FAF9F7');

    // Set column widths
    sheet.setColumnWidth(1, 180); // Timestamp
    sheet.setColumnWidth(2, 250); // Name
    sheet.setColumnWidth(3, 200); // Response
    sheet.setColumnWidth(4, 120); // Response Code
    sheet.setColumnWidth(5, 200); // User Agent

    // Freeze the header row
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Convert RSVP code to readable label in Kyrgyz
 */
function getRsvpLabel(code) {
  const labels = {
    'yes': 'Келемин (Coming)',
    'both': 'Жубайым менен келемин (Coming with spouse)',
    'no': 'Келе албаймын (Cannot attend)'
  };

  return labels[code] || code;
}

/**
 * Optional: Function to get statistics
 * You can run this manually to get RSVP statistics
 */
function getRSVPStatistics() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log('No RSVP data found');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const stats = {
    total: data.length - 1, // Exclude header
    coming: 0,
    comingWithSpouse: 0,
    notComing: 0
  };

  // Count responses (skip header row)
  for (let i = 1; i < data.length; i++) {
    const rsvpCode = data[i][3]; // Response Code column

    if (rsvpCode === 'yes') {
      stats.coming++;
    } else if (rsvpCode === 'both') {
      stats.comingWithSpouse++;
    } else if (rsvpCode === 'no') {
      stats.notComing++;
    }
  }

  // Calculate estimated guest count
  stats.estimatedGuests = stats.coming + (stats.comingWithSpouse * 2);

  Logger.log('RSVP Statistics:');
  Logger.log('Total Responses: ' + stats.total);
  Logger.log('Coming Alone: ' + stats.coming);
  Logger.log('Coming with Spouse: ' + stats.comingWithSpouse);
  Logger.log('Cannot Attend: ' + stats.notComing);
  Logger.log('Estimated Total Guests: ' + stats.estimatedGuests);

  return stats;
}

/**
 * Optional: Send email notification when new RSVP is received
 * Uncomment and configure to use
 */
/*
function sendEmailNotification(name, rsvpStatus) {
  const recipient = 'your-email@example.com'; // Change to your email
  const subject = 'New Wedding RSVP from ' + name;
  const body = 'You received a new RSVP:\n\n' +
               'Name: ' + name + '\n' +
               'Response: ' + rsvpStatus + '\n\n' +
               'View all RSVPs: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();

  MailApp.sendEmail(recipient, subject, body);
}
*/
