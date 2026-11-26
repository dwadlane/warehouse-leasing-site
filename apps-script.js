// ============================================================
// PASIG WAREHOUSE LEAD TRACKER, Google Apps Script
// ============================================================
// Paste this into: Extensions > Apps Script in your Google Sheet
// Then: Deploy > New deployment > Web app > Anyone can access
// Copy the web app URL and paste it into your listing page HTML
// ============================================================

const SHEET_NAME = 'Leads';
const ANALYTICS_SHEET = 'Analytics';
const NOTIFY_EMAIL = 'dwad@gaps2gains.com';

// ─── Handle form submissions (POST) and page views (GET) ───
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.type === 'pageview') {
      return logPageView(data);
    }

    if (data.type === 'chatbot') {
      return logChatInteraction(data);
    }

    // Default: lead submission
    return logLead(data);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // Return analytics summary
  try {
    const action = e.parameter.action;
    if (action === 'stats') {
      return getStats();
    }
    // Default: log a page view via GET (pixel tracking)
    logPageView({
      page: e.parameter.page || '/',
      referrer: e.parameter.ref || '',
      ua: e.parameter.ua || ''
    });
    // Return 1x1 transparent pixel
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── Log a new lead ───
function logLead(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Create Leads sheet with headers if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Timestamp',
      'First Name',
      'Last Name',
      'Company',
      'Email',
      'Phone',
      'Type',
      'Space Required',
      'Message',
      'Source',
      'Status',
      'Notes',
      'Follow-Up Date'
    ]);
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, 13);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#0f1b2d');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    // Set column widths
    sheet.setColumnWidth(1, 160); // Timestamp
    sheet.setColumnWidth(2, 120); // First Name
    sheet.setColumnWidth(3, 120); // Last Name
    sheet.setColumnWidth(4, 180); // Company
    sheet.setColumnWidth(5, 220); // Email
    sheet.setColumnWidth(6, 140); // Phone
    sheet.setColumnWidth(7, 140); // Type
    sheet.setColumnWidth(8, 160); // Space Required
    sheet.setColumnWidth(9, 300); // Message
    sheet.setColumnWidth(10, 100); // Source
    sheet.setColumnWidth(11, 100); // Status
    sheet.setColumnWidth(12, 200); // Notes
    sheet.setColumnWidth(13, 120); // Follow-Up
  }

  const timestamp = new Date().toLocaleString('en-PH', { timeZone: 'Asia/Manila' });

  sheet.appendRow([
    timestamp,
    data.firstName || '',
    data.lastName || '',
    data.company || '',
    data.email || '',
    data.phone || '',
    data.inquirerType || '',
    data.spaceRequired || '',
    data.message || '',
    data.source || 'Website Form',
    'New',
    '',
    ''
  ]);

  // Send email notification
  sendLeadNotification(data, timestamp);

  // Log as conversion in analytics
  logAnalyticsEvent('form_submission', data.source || 'website');

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Lead captured successfully'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Log page view ───
function logPageView(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ANALYTICS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(ANALYTICS_SHEET);
    sheet.appendRow([
      'Timestamp',
      'Event',
      'Page',
      'Referrer',
      'User Agent',
      'Session ID'
    ]);
    const headerRange = sheet.getRange(1, 1, 1, 6);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1a6bff');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  const timestamp = new Date().toLocaleString('en-PH', { timeZone: 'Asia/Manila' });
  sheet.appendRow([
    timestamp,
    'pageview',
    data.page || '/',
    data.referrer || '',
    data.ua || '',
    data.sessionId || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Log chatbot interaction ───
function logChatInteraction(data) {
  logAnalyticsEvent('chatbot_interaction', data.question || '');

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Analytics helper ───
function logAnalyticsEvent(event, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ANALYTICS_SHEET);
  if (!sheet) return;

  const timestamp = new Date().toLocaleString('en-PH', { timeZone: 'Asia/Manila' });
  sheet.appendRow([timestamp, event, detail, '', '', '']);
}

// ─── Email notification ───
function sendLeadNotification(data, timestamp) {
  const subject = '🔥 New Warehouse Lead: ' + (data.company || data.firstName || 'Unknown');

  const body = [
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    '  NEW LEAD, Pasig Warehouse Listing',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    '',
    'Time:           ' + timestamp,
    'Name:           ' + (data.firstName || '') + ' ' + (data.lastName || ''),
    'Company:        ' + (data.company || 'Not specified'),
    'Email:          ' + (data.email || 'Not provided'),
    'Phone:          ' + (data.phone || 'Not provided'),
    'Type:           ' + (data.inquirerType || 'Not specified'),
    'Space Required: ' + (data.spaceRequired || 'Not specified'),
    '',
    'Message:',
    data.message || '(No message)',
    '',
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
    'Source: ' + (data.source || 'Website Form'),
    '',
    '→ View all leads: https://docs.google.com/spreadsheets/d/' +
      SpreadsheetApp.getActiveSpreadsheet().getId(),
    '',
    'Reply to this email or call them back ASAP!',
  ].join('\n');

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: subject,
    body: body,
    replyTo: data.email || NOTIFY_EMAIL
  });
}

// ─── Stats endpoint ───
function getStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadsSheet = ss.getSheetByName(SHEET_NAME);
  const analyticsSheet = ss.getSheetByName(ANALYTICS_SHEET);

  let totalLeads = 0, totalViews = 0, totalChats = 0;

  if (leadsSheet) {
    totalLeads = Math.max(0, leadsSheet.getLastRow() - 1);
  }

  if (analyticsSheet) {
    const data = analyticsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === 'pageview') totalViews++;
      if (data[i][1] === 'chatbot_interaction') totalChats++;
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      totalLeads: totalLeads,
      totalViews: totalViews,
      totalChats: totalChats
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
