/**
 * DoveRefer — LinkedIn Referral Message Generator
 * Google Apps Script Web App Controller
 *
 * This file serves the web application and handles optional server-side actions,
 * such as logging referral requests to a linked Google Sheet (if history tracking is desired).
 */

/**
 * Serves the HTML frontend when the user visits the published web app URL.
 * 
 * @param {Object} e - HTTP request event object.
 * @return {HtmlOutput} The rendered HTML page.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('DoveRefer — LinkedIn Referral Message Generator')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Optional Google Sheets integration:
 * Logs generated referral metadata to a linked Google Sheet for tracking.
 * This runs ONLY if a spreadsheet is bound or specified, and requires zero paid APIs.
 * 
 * @param {Object} data - Payload containing employeeName, companyName, jobPosition, etc.
 * @return {Object} Status response indicating success or failure.
 */
function logReferralToSheet(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return { success: false, message: 'No active Google Sheet linked. Run inside a sheet-bound Apps Script project or provide Spreadsheet ID.' };
    }
    
    var sheetName = 'Referral Logs';
    var sheet = ss.getSheetByName(sheetName);
    
    // Create sheet and header row if not present
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        'Timestamp',
        'Employee Name',
        'Company Name',
        'Job Position',
        'Requisition ID',
        'My Experience',
        'Skills',
        'Applicant Name'
      ]);
      // Style header row with light purple background
      sheet.getRange(1, 1, 1, 8)
        .setFontWeight('bold')
        .setBackground('#f3e8ff')
        .setFontColor('#4c1d95');
      sheet.setFrozenRows(1);
    }
    
    // Append the referral request record
    sheet.appendRow([
      new Date().toLocaleString(),
      data.employeeName || 'N/A',
      data.companyName || 'N/A',
      data.jobPosition || 'N/A',
      data.requisitionId || 'None',
      data.experience || '2 years',
      data.skills || 'None',
      data.applicantName || 'Anonymous'
    ]);
    
    return { success: true, message: 'Referral logged successfully.' };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}
