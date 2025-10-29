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
 * Generates HTML email body with brand styling
 * @param {Array} headlines - Array of headline objects
 * @param {Array} rocks - Array of rock progress objects
 * @param {string} dateString - Formatted date string
 * @returns {string} HTML formatted email body
 */
function generateHtmlBody(headlines, rocks, dateString) {
  const headlineBlue = '#0093d0';
  const primaryGreen = '#00993d';
  const secondaryGreen = '#00b259';

  let html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link href="https://fonts.googleapis.com/css2?family=Oswald:wght@400;500;600;700&family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">
  <style>
    body {
      margin: 0;
      padding: 0;
      font-family: 'Open Sans', Arial, sans-serif;
      background-color: #f5f5f5;
      color: #333333;
    }
    .container {
      max-width: 600px;
      margin: 0 auto;
      background-color: #ffffff;
    }
    .header {
      background-color: #ffffff;
      padding: 30px 20px;
      text-align: left;
    }
    .header h1 {
      font-family: 'Oswald', Arial, sans-serif;
      font-size: 32px;
      font-weight: 600;
      color: ${primaryGreen};
      margin: 0 0 10px 0;
      text-transform: uppercase;
      letter-spacing: 1px;
    }
    .header .date {
      font-family: 'Open Sans', Arial, sans-serif;
      font-size: 14px;
      color: #666666;
    }
    .content {
      padding: 30px 20px;
    }
    .section-title {
      font-family: 'Oswald', Arial, sans-serif;
      font-size: 24px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      margin: 0 0 20px 0;
      padding-bottom: 10px;
    }
    .section-title.headlines {
      color: ${headlineBlue};
      border-bottom: 3px solid ${headlineBlue};
    }
    .section-title.rocks {
      color: ${primaryGreen};
      border-bottom: 3px solid ${secondaryGreen};
    }
    .headlines-table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 30px;
      border: 1px solid #000000;
    }
    .headlines-table th {
      font-family: 'Oswald', Arial, sans-serif;
      font-size: 16px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      background-color: ${headlineBlue};
      color: #ffffff;
      padding: 12px 15px;
      text-align: left;
      border: 1px solid #000000;
    }
    .headlines-table td {
      font-family: 'Open Sans', Arial, sans-serif;
      font-size: 14px;
      padding: 12px 15px;
      border: 1px solid #000000;
      vertical-align: top;
    }
    .headlines-table tbody tr:nth-child(odd) {
      background-color: #ffffff;
    }
    .headlines-table tbody tr:nth-child(even) {
      background-color: #d9d9d9;
    }
    .headlines-table td:first-child {
      font-weight: 700;
      color: #000000;
      width: 30%;
    }
    .headlines-table td:last-child {
      width: 70%;
      line-height: 1.6;
      color: #000000;
    }
    .headlines-table th {
      white-space: nowrap;
    }
    .item {
      margin-bottom: 25px;
      padding: 15px;
      background-color: #f9f9f9;
      border-left: 4px solid ${secondaryGreen};
      border-radius: 4px;
    }
    .item-number {
      font-family: 'Oswald', Arial, sans-serif;
      font-size: 18px;
      font-weight: 600;
      color: ${primaryGreen};
      margin-bottom: 8px;
    }
    .item-content {
      font-family: 'Open Sans', Arial, sans-serif;
      font-size: 15px;
      line-height: 1.6;
      color: #333333;
      margin-bottom: 8px;
    }
    .item-meta {
      font-family: 'Open Sans', Arial, sans-serif;
      font-size: 13px;
      color: #666666;
      font-weight: 600;
    }
    .item-label {
      color: ${primaryGreen};
      font-weight: 700;
    }
    .no-items {
      font-family: 'Open Sans', Arial, sans-serif;
      font-size: 14px;
      color: #999999;
      font-style: italic;
      padding: 15px;
    }
    .footer {
      background-color: #f0f0f0;
      padding: 20px;
      text-align: center;
      font-family: 'Open Sans', Arial, sans-serif;
      font-size: 12px;
      color: #666666;
    }
    .section-divider {
      margin: 40px 0 30px 0;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Sales Team Newsletter</h1>
      <div class="date">Week of ${dateString}</div>
    </div>

    <div class="content">
      <!-- Headlines Section -->
      <h2 class="section-title headlines">Headlines</h2>`;

  if (headlines.length > 0) {
    html += `
      <table class="headlines-table">
        <thead>
          <tr>
            <th style="white-space: nowrap; width: 30%; text-align: center;">Team Member</th>
            <th style="width: 70%;">Headline</th>
          </tr>
        </thead>
        <tbody>`;

    let rowIndex = 0;
    headlines.forEach((item) => {
      // Split headline by line breaks
      const headlineParts = item.highlight.toString().split('\n').filter(part => part.trim() !== '');

      // Create a row for each part of the headline
      headlineParts.forEach((headlinePart) => {
        const bgColor = (rowIndex % 2 === 0) ? '#f0f0f0' : '#ffffff';
        html += `
          <tr style="background-color: ${bgColor};">
            <td style="font-family: 'Open Sans', Arial, sans-serif; font-weight: 700; color: #000000; width: 30%; text-align: center;"><strong>${item.owner}</strong></td>
            <td style="font-family: 'Open Sans', Arial, sans-serif; color: #000000; width: 70%;">${headlinePart.trim()}</td>
          </tr>`;
        rowIndex++;
      });
    });

    html += `
        </tbody>
      </table>`;
  } else {
    html += `<div class="no-items">No headlines this week.</div>`;
  }

  html += `

      <!-- Rock Progress Section -->
      <div class="section-divider"></div>
      <h2 class="section-title rocks">Rock Progress</h2>`;

  if (rocks.length > 0) {
    rocks.forEach((item, index) => {
      html += `
      <div class="item">
        <div class="item-number">${index + 1}.</div>
        <div class="item-content">${item.rockOwner}</div>
        <div class="item-meta"><span class="item-label">Goal:</span> ${item.goal}</div>
        <div class="item-meta"><span class="item-label">Progress:</span> ${item.progress}</div>
      </div>`;
    });
  } else {
    html += `<div class="no-items">No rock progress this week.</div>`;
  }

  html += `
    </div>

    <div class="footer">
      Global O-Ring and Seal | Sales Team L10
    </div>
  </div>
</body>
</html>`;

  return html;
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

  // Generate HTML body
  const htmlBody = generateHtmlBody(headlines, rocks, dateString);

  // Send email
  try {
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: htmlBody
    });

    SpreadsheetApp.getUi().alert('Newsletter sent successfully to ' + recipient);
    Logger.log('Newsletter sent successfully');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error sending email: ' + error.message);
    Logger.log('Error sending email: ' + error);
  }
}
