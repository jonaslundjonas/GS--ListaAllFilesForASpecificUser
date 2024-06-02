/**
 * Google Apps Script for listing Google Drive files owned by a specific user.
 * 
 * This script generates a report in a Google Sheet with the following details:
 * - Document Title
 * - Username
 * - Document ID
 * - Last modified Date
 * - Document Owner
 * - Document Type
 * - Creation Date
 * - Modified in 3 month Period
 * - Created more than 3 months ago and Modified Recently
 * - Shared with others
 * - Number of shares
 *
 * This script is great for identifying how many files are not in use to help clean out unused files 
 * and stay GDPR compliant.
 *
 * If the script encounters a timeout or exceeds running time limits, you can use the 'Continue Listing Files'
 * function to resume the process without losing progress.
 *
 * Requirements:
 * - A Google Sheet to run the script.
 * - Google Apps Script authorization to access Google Drive and Google Sheets.
 * 
 * Instructions:
 * 1. Open a Google Sheet.
 * 2. Click on 'Extensions' -> 'Apps Script'.
 * 3. Copy and paste this script into the Apps Script editor.
 * 4. Save the script.
 * 5. Reload the Google Sheet.
 * 6. Use the 'Drive File List' menu to start, continue, or reset the listing process.
 * 
 * Written by Jonas Lund 2024
 */

var MONTH_LIMIT = 3; // Number of months to go back from today's date with the created date
var SPECIFIC_USER = 'jonas.lund@academedia.se'; // Email of the specific user

// Function to create a custom menu in the Google Sheets UI
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Drive File List')
    .addItem('Start Listing Files', 'startListingFiles')
    .addItem('Continue Listing Files', 'continueListingFiles')
    .addItem('Reset and Start Over', 'resetAndStartOver')
    .addToUi();
}

// Function to start the listing process and initialize the sheet
function startListingFiles() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear(); // Clear the sheet
  var headers = ['Document Title', 'Username', 'Document ID', 'Last modified Date', 'Document Owner', 'Document Type', 'Creation Date', 'Modified in 3 month Period', 'Created more than 3 months ago and Modified Recently', 'Shared with others', 'Number of shares'];
  sheet.appendRow(headers);
  
  // Make the header row bold
  var range = sheet.getRange(1, 1, 1, headers.length);
  range.setFontWeight('bold');
  
  // Freeze the first row
  sheet.setFrozenRows(1);

  listFiles(); // Call the function to list files
}

// Function to continue listing files if the process was interrupted
function continueListingFiles() {
  listFiles();
}

// Function to reset the sheet and start the listing process over
function resetAndStartOver() {
  startListingFiles(); // Call the startListingFiles function to reset and start over
}

// Function to list files owned by the specified user
function listFiles() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var query = "mimeType != 'application/vnd.google-apps.folder' and trashed = false";
  if (SPECIFIC_USER) {
    query += " and '" + SPECIFIC_USER + "' in owners";
  }
  
  var files = [];
  var pageToken = null;
  
  // Retrieve files in a paginated manner
  do {
    var response = Drive.Files.list({
      q: query,
      pageToken: pageToken,
      fields: "nextPageToken, files(id, name, mimeType, owners, createdTime, modifiedTime)"
    });
    
    files = files.concat(response.files);
    pageToken = response.nextPageToken;
  } while (pageToken);

  var currentDate = new Date();
  var pastDate = new Date();
  pastDate.setMonth(currentDate.getMonth() - MONTH_LIMIT);

  // Process each file and append the data to the sheet
  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    var createdTime = new Date(file.createdTime);
    var modifiedTime = new Date(file.modifiedTime);
    var modifiedInPeriod = modifiedTime >= pastDate;
    var createdMoreThan3MonthsAgo = createdTime < pastDate;
    var createdMoreThan3MonthsAgoAndModifiedRecently = createdMoreThan3MonthsAgo && modifiedInPeriod;
    var shareInfo = getShareInfo(file.id);
    
    var row = [
      file.name,
      SPECIFIC_USER,
      file.id,
      file.modifiedTime,
      file.owners[0].emailAddress,
      file.mimeType,
      file.createdTime,
      modifiedInPeriod ? "Yes" : "No",
      createdMoreThan3MonthsAgoAndModifiedRecently ? "Yes" : "No",
      shareInfo.sharedWithOthers ? "Yes" : "No",
      shareInfo.numberOfShares
    ];
    sheet.appendRow(row);
  }
}

// Function to get sharing information for a file
function getShareInfo(fileId) {
  try {
    var permissions = Drive.Permissions.list(fileId, {fields: 'permissions(emailAddress, role)'}).permissions;
    var numberOfShares = 0;
    var sharedWithOthers = false;
    
    // Count the number of shares excluding the owner
    if (permissions) {
      for (var i = 0; i < permissions.length; i++) {
        if (permissions[i].role !== 'owner') {
          numberOfShares++;
          sharedWithOthers = true;
        }
      }
    }
    
    return { sharedWithOthers: sharedWithOthers, numberOfShares: numberOfShares };
  } catch (e) {
    return { sharedWithOthers: false, numberOfShares: 0 };
  }
}
