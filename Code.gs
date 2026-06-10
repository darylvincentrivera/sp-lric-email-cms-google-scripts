function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('LRIC Content Mailer')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getAssetsFolderId(customFolderId) {
  if (customFolderId && customFolderId.trim() !== "") {
    return customFolderId.trim();
  }
  return "1Pj4BR-MiuOBaVZFUZPA_Q2VgJNflaY9_";
}

function getOAuthToken() {
  return ScriptApp.getOAuthToken();
}

function getTargetSpreadsheet() {
  return SpreadsheetApp.openById("1KAssPljYfiBWQpd2tpIGGTtEQCADCYj9lyVBaipEX7k");
}

// Normalize all keys to lowercase so loadSystemSettings() can match them
function getSystemSettings() {
  const ss = getTargetSpreadsheet();
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  let settings = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      settings[data[i][0].toString().toLowerCase()] = data[i][1];
    }
  }
  return settings;
}

// Accept flat string values instead of {url, enabled} objects
function saveSystemSettings(settingsObj) {
  const ss = getTargetSpreadsheet();
  let sheet = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  sheet.clearContents();
  sheet.appendRow(['Key', 'Value']);
  for (let key in settingsObj) {
    sheet.appendRow([key, settingsObj[key]]);
  }
}

function getDashboardMetrics() {
  const ss = getTargetSpreadsheet();
  const recipientSheet = ss.getSheetByName('Recipients') || ss.insertSheet('Recipients');
  if (recipientSheet.getLastRow() === 0) {
    recipientSheet.appendRow(['Email', 'Name', 'Status', 'Date Added']);
  }
  const recipientData = recipientSheet.getDataRange().getValues();
  let activeRecipients = 0;
  for (let i = 1; i < recipientData.length; i++) {
    if (recipientData[i][2] && recipientData[i][2].toString().trim().toLowerCase() === "active") {
      activeRecipients++;
    }
  }

  const campaignSheet = ss.getSheetByName('Campaigns') || ss.insertSheet('Campaigns');
  if (campaignSheet.getLastRow() === 0) {
    campaignSheet.appendRow(['Timestamp', 'Subject', 'HTML Body']);
  }
  const totalPublished = campaignSheet.getLastRow() > 1 ? campaignSheet.getLastRow() - 1 : 0;
  return {
    totalPublished: totalPublished,
    activeRecipients: activeRecipients,
    systemStatus: "Operational",
    statusColor: "text-emerald-400"
  };
}

function getRecipientList() {
  const ss = getTargetSpreadsheet();
  const sheet = ss.getSheetByName('Recipients');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return data.map((row, index) => ({
    rowId: index + 2,
    email: row[0],
    name: row[1],
    status: row[2] || 'Active',
    date: row[3] ? Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "yyyy-MM-dd") : 'N/A'
  }));
}

function getCampaignList() {
  const ss = getTargetSpreadsheet();
  const sheet = ss.getSheetByName('Campaigns');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data.map((row) => ({
    timestamp: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : 'N/A',
    subject: row[1],
    htmlBody: row[2] || ''
  })).reverse();
}

function addNewRecipient(name, email) {
  const ss = getTargetSpreadsheet();
  let sheet = ss.getSheetByName('Recipients');
  if (!sheet) {
    sheet = ss.insertSheet('Recipients');
    sheet.appendRow(['Email', 'Name', 'Status', 'Date Added']);
  }
  sheet.appendRow([email, name, 'Active', new Date()]);
  return { success: true };
}

function toggleUserStatus(rowId, currentStatus) {
  const ss = getTargetSpreadsheet();
  const sheet = ss.getSheetByName('Recipients');
  const newStatus = (currentStatus === 'Active') ? 'Unsubscribed' : 'Active';
  sheet.getRange(rowId, 3).setValue(newStatus);
  return newStatus;
}

function updateRecipientDetails(rowId, name, email) {
  const ss = getTargetSpreadsheet();
  const sheet = ss.getSheetByName('Recipients');
  sheet.getRange(rowId, 1).setValue(email);
  sheet.getRange(rowId, 2).setValue(name);
  return { success: true };
}

function executeBroadcast(subject, htmlBody) {
  const ss = getTargetSpreadsheet();
  const sheet = ss.getSheetByName('Recipients');
  if (!sheet || sheet.getLastRow() <= 1) {
    throw new Error("Aborted: No recipient data rows detected inside the sheet tab.");
  }

  const startRow = 2;
  const numRows = sheet.getLastRow() - startRow + 1;
  const data = sheet.getRange(startRow, 1, numRows, 3).getValues();
  let sentCount = 0;
  const errors = [];

  const safeHtmlBody = htmlBody ? String(htmlBody).trim() : "";
  if (!safeHtmlBody || safeHtmlBody === "") {
    throw new Error("Aborted: Compiled HTML data payload was empty.");
  }

  data.forEach(function(row, index) {
    const email = row[0] ? row[0].toString().trim() : "";
    const name  = row[1] ? row[1].toString().trim() : "Subscriber";
    let status  = row[2] ? row[2].toString().trim().toLowerCase() : "active";

    if (status === "") status = "active";
    if (!email || status !== "active") return;

    const personalizedHtml = safeHtmlBody.replace(/{{Name}}/g, name);

    try {
      GmailApp.sendEmail(email, subject, "Please view this email in an HTML compatible client.", {
        htmlBody: personalizedHtml
      });
      sentCount++;
    } catch (err) {
      const msg = `Row ${index + 2} (${email}): ${err.message}`;
      Logger.log(`Delivery Error — ${msg}`);
      errors.push(msg);
    }
  });

  try {
    let campaignSheet = ss.getSheetByName('Campaigns') || ss.insertSheet('Campaigns');
    if (campaignSheet.getLastRow() === 0) {
      campaignSheet.appendRow(['Timestamp', 'Subject', 'HTML Body']);
    }
    campaignSheet.appendRow([new Date(), subject, safeHtmlBody]);
  } catch (logError) {
    Logger.log(`Registry Error: ${logError.message}`);
  }

  if (errors.length > 0 && sentCount === 0) {
    throw new Error("All deliveries failed. First error: " + errors[0]);
  }

  return sentCount;
}
