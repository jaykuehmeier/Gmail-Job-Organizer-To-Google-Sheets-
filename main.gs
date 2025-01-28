// Configuration
const EMAIL_SEARCH_TERMS = [
  'thank you for applying',
  'application received',
  'application confirmation',
  'thank you for your application'
];

function setupTrackingSheet() {
  // Get the active spreadsheet that this script is bound to
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if Applications sheet exists
  let sheet = ss.getSheetByName('Applications');
  if (!sheet) {
    sheet = ss.insertSheet('Applications');
    
    // Set up headers
    const headers = [
      'Date Applied',
      'Company',
      'Position',
      'Status',
      'Last Updated',
      'Email Subject',
      'Email Link',
      'Notes'
    ];
    
    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers])
      .setBackground('#4285f4')
      .setFontColor('white')
      .setFontWeight('bold');
    
    // Create status dropdown
    const statusRange = sheet.getRange('D2:D1000');
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'Applied',
        'Application Received',
        'Under Review',
        'Interview Scheduled',
        'Interview Completed',
        'Rejected',
        'Offer Received'
      ], true)
      .build();
    statusRange.setDataValidation(statusRule);
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Set column widths
    sheet.setColumnWidth(1, 100);  // Date
    sheet.setColumnWidth(2, 150);  // Company
    sheet.setColumnWidth(3, 200);  // Position
    sheet.setColumnWidth(4, 130);  // Status
    sheet.setColumnWidth(5, 100);  // Last Updated
    sheet.setColumnWidth(6, 300);  // Email Subject
    sheet.setColumnWidth(7, 250);  // Email Link
    sheet.setColumnWidth(8, 300);  // Notes
  }
  
  return sheet;
}

function findNewApplications() {
  const sheet = setupTrackingSheet();
  const existingData = sheet.getDataRange().getValues();
  const existingEmails = existingData.map(row => row[6]); // Email Link column
  
  // Create search query for all relevant terms
  const searchQueries = EMAIL_SEARCH_TERMS.map(term => `subject:"${term}"`);
  const searchQuery = searchQueries.join(' OR ');
  
  // Search for threads matching any of the terms
  const threads = GmailApp.search(searchQuery);
  let newApplications = 0;
  
  threads.forEach(thread => {
    const messages = thread.getMessages();
    const firstMessage = messages[0];
    const emailLink = thread.getPermalink();
    
    // Skip if already in spreadsheet
    if (!existingEmails.includes(emailLink)) {
      const subject = firstMessage.getSubject();
      const date = firstMessage.getDate();
      const from = firstMessage.getFrom();
      
      // Extract company and position
      const companyName = extractCompanyName(subject, from);
      const position = extractPosition(subject);
      
      // Add new row
      sheet.appendRow([
        date,
        companyName,
        position,
        'Application Received',
        date,
        subject,
        emailLink,
        ''  // Notes
      ]);
      
      newApplications++;
    }
  });
  
  // Format any new rows
  if (newApplications > 0) {
    formatSheet(sheet);
  }
  
  // Show results
  const ui = SpreadsheetApp.getUi();
  if (newApplications > 0) {
    ui.alert(`Found ${newApplications} new job application${newApplications === 1 ? '' : 's'}!`);
  } else {
    ui.alert('No new job applications found.');
  }
}

function formatSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;  // Skip if only header row exists
  
  // Format dates
  const dateRanges = sheet.getRange(`A2:A${lastRow}`);
  dateRanges.setNumberFormat('MM/dd/yyyy');
  
  const lastUpdateRanges = sheet.getRange(`E2:E${lastRow}`);
  lastUpdateRanges.setNumberFormat('MM/dd/yyyy');
  
  // Make links clickable
  const linkRange = sheet.getRange(`G2:G${lastRow}`);
  linkRange.setWrap(true)
    .setFontColor('#1155cc')
    .setFontLine('underline');
}

function extractCompanyName(subject, from) {
  // Try to extract company name from the "thank you for applying" email
  const fromParts = from.split('@');
  if (fromParts.length > 1) {
    // Remove common recruitment system domains
    const domain = fromParts[1].split('.')[0]
      .replace('greenhouse', '')
      .replace('lever', '')
      .replace('workday', '')
      .replace('recruiting', '');
    
    if (domain) {
      return domain.charAt(0).toUpperCase() + domain.slice(1);
    }
  }
  
  // Try to find company name in subject
  const subjectParts = subject.toLowerCase().split('thank you for applying');
  if (subjectParts.length > 1 && subjectParts[0]) {
    return subjectParts[0].trim();
  }
  
  return 'Company Name TBD';
}

function extractPosition(subject) {
  // Try to extract position from various email formats
  const positionPatterns = [
    /applied for(.*?)at/i,
    /applying for(.*?)at/i,
    /application for(.*?)position/i,
    /application for(.*?)role/i
  ];
  
  for (const pattern of positionPatterns) {
    const match = subject.match(pattern);
    if (match && match[1]) {
      return match[1].trim();
    }
  }
  
  return 'Position TBD';
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Job Applications')
    .addItem('Check for New Applications', 'findNewApplications')
    .addSeparator()
    .addItem('Set Up Daily Check', 'createDailyTrigger')
    .addToUi();
}

function createDailyTrigger() {
  // Remove any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create new daily trigger
  ScriptApp.newTrigger('findNewApplications')
    .timeBased()
    .everyDays(1)
    .atHour(9)  // 9 AM
    .create();
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('Daily check has been scheduled for 9 AM');
}
