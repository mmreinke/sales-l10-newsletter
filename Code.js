/**
 * Creates custom menu when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Newsletter')
    .addItem('Send email now', 'sendNewsletter')
    .addToUi();
}

/**
 * Gets headlines from the Headlines tab
 * @returns {Array} Array of headline objects with content and owner
 */
function getHeadlines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headlinesSheet = ss.getSheetByName('Headlines');

  if (!headlinesSheet) {
    Logger.log('Headlines sheet not found');
    return [];
  }

  // Get data starting from A2 (column A = highlights, column B = owner)
  const lastRow = headlinesSheet.getLastRow();
  if (lastRow < 2) {
    return []; // No data
  }

  const range = headlinesSheet.getRange(2, 1, lastRow - 1, 2);
  const values = range.getValues();

  // Filter out empty rows and format data
  const headlines = values
    .filter(row => row[0] && row[0].toString().trim() !== '')
    .map(row => ({
      highlight: row[0],
      owner: row[1] || 'Unknown'
    }));

  return headlines;
}

/**
 * Gets rock progress from the Rock Progress tab
 * @returns {Array} Array of rock objects with rock/owner, goal, and progress
 */
function getRockProgress() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rockSheet = ss.getSheetByName('Rock Progress');

  if (!rockSheet) {
    Logger.log('Rock Progress sheet not found');
    return [];
  }

  // Get data starting from A2 (column A = rock/owner, column B = goal, column C = progress)
  const lastRow = rockSheet.getLastRow();
  if (lastRow < 2) {
    return []; // No data
  }

  const range = rockSheet.getRange(2, 1, lastRow - 1, 3);
  const values = range.getValues();

  // Filter out empty rows and format data
  const rocks = values
    .filter(row => row[0] && row[0].toString().trim() !== '')
    .map(row => ({
      rockOwner: row[0],
      goal: row[1] || '',
      progress: row[2] || ''
    }));

  return rocks;
}

/**
 * Formats and sends the weekly newsletter email
 */
function sendNewsletter() {
  const recipient = 'madison@globaloring.com';
  const sender = 'madison@globaloring.com';

  // Get today's date for subject line
  const today = new Date();
  const dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy');
  const subject = `L10 - Sales: Newsletter ${dateString}`;

  // Get data from sheets
  const headlines = getHeadlines();
  const rocks = getRockProgress();

  // Build email body
  let body = 'Weekly Newsletter\n\n';

  // Add Headlines section
  body += '=== HEADLINES ===\n\n';
  if (headlines.length > 0) {
    headlines.forEach((item, index) => {
      body += `${index + 1}. ${item.highlight}\n`;
      body += `   Owner: ${item.owner}\n\n`;
    });
  } else {
    body += 'No headlines this week.\n\n';
  }

  // Add Rock Progress section
  body += '\n=== ROCK PROGRESS ===\n\n';
  if (rocks.length > 0) {
    rocks.forEach((item, index) => {
      body += `${index + 1}. ${item.rockOwner}\n`;
      body += `   Goal: ${item.goal}\n`;
      body += `   Progress: ${item.progress}\n\n`;
    });
  } else {
    body += 'No rock progress this week.\n\n';
  }

  // Send email
  try {
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: body
    });

    SpreadsheetApp.getUi().alert('Newsletter sent successfully to ' + recipient);
    Logger.log('Newsletter sent successfully');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error sending email: ' + error.message);
    Logger.log('Error sending email: ' + error);
  }
}
