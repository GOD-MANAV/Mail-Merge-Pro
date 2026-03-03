/**
 * ==============================================================================
 * MAIL MERGE PRO - COMPLETE GOOGLE APPS SCRIPT
 * ==============================================================================
 * A professional mail merge solution for Google Sheets + Gmail
 * 
 * Instructions:
 * 1. Open Google Sheets → Extensions → Apps Script
 * 2. Delete the default Code.gs content
 * 3. Paste this entire file
 * 4. Create ONE new HTML file called "Sidebar.html" and paste the HTML code
 * 5. Update appsscript.json with the required scopes
 * 6. Save and refresh your Google Sheet
 * ==============================================================================
 */

// ==================== MENU & UI ====================

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📧 Mail Merge Pro')
    .addItem('Open Mail Merge', 'showSidebar')
    .addSeparator()
    .addItem('Templates', 'showTemplateManager')
    .addItem('Scheduled Emails', 'showScheduledEmailsDialog')
    .addItem('Activity Monitor', 'showActivityMonitor')
    .addItem('Scheduled Campaigns', 'showScheduledCampaignsManager')
    .addSeparator()
    .addItem('About', 'showAbout')
    .addToUi();
}

/**
 * Shows the main sidebar UI
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Mail Merge Pro')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows templates dialog (uses Options.html)
 */
function showTemplateManager() {
  const template = HtmlService.createTemplateFromFile('Options');
  template.initialTab = 'templates';
  const html = template.evaluate().setWidth(500).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Templates');
}

/**
 * Shows about dialog
 */
function showAbout() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: 'Google Sans', Arial, sans-serif; padding: 20px; text-align: center;">
      <h2 style="color: #1a73e8;">📧 Mail Merge Pro</h2>
      <p>Version 1.0.0</p>
      <p>A professional mail merge solution for job hunting and cold emailing.</p>
      <hr>
      <p style="color: #666; font-size: 12px;">Built with ❤️ for recruiters, HRs, and job seekers</p>
    </div>
  `).setWidth(300).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'About');
}

/**
 * Shows scheduled emails dialog (uses Options.html)
 */
function showScheduledEmailsDialog() {
  const template = HtmlService.createTemplateFromFile('Options');
  template.initialTab = 'scheduled';
  const html = template.evaluate().setWidth(500).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Scheduled Emails');
}

/**
 * Shows activity monitor dialog (uses Options.html)
 */
function showActivityMonitor() {
  const template = HtmlService.createTemplateFromFile('Options');
  template.initialTab = 'activity';
  const html = template.evaluate().setWidth(550).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Activity Monitor');
}

function showScheduledCampaignsManager() {
  const template = HtmlService.createTemplateFromFile('Options');
  template.initialTab = 'campaigns';
  const html = template.evaluate().setWidth(600).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Scheduled Campaigns');
}

// ==================== SPREADSHEET FUNCTIONS ====================

/**
 * Get current spreadsheet info
 */
function getCurrentSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    id: ss.getId(),
    name: ss.getName(),
    url: ss.getUrl()
  };
}

/**
 * Get all sheets in the current spreadsheet
 */
function getSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  return sheets.map(sheet => ({
    name: sheet.getName(),
    index: sheet.getIndex(),
    rowCount: sheet.getLastRow(),
    colCount: sheet.getLastColumn()
  }));
}

/**
 * Get column headers from a specific sheet
 */
function getColumns(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet || sheet.getLastRow() === 0) {
    return [];
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.map((header, index) => ({
    name: header || `Column ${index + 1}`,
    index: index,
    letter: columnToLetter(index + 1)
  }));
}

/**
 * Get row count (excluding header)
 */
function getRowCount(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return 0;
  
  const lastRow = sheet.getLastRow();
  return Math.max(0, lastRow - 1);
}

/**
 * Get all data from a sheet
 */
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet || sheet.getLastRow() === 0) {
    return { headers: [], data: [] };
  }
  
  const range = sheet.getDataRange();
  const values = range.getValues();
  
  return {
    headers: values[0],
    data: values.slice(1)
  };
}

/**
 * Get preview data (first 3 rows)
 */
function getPreviewData(sheetName, emailCol, nameCol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }
  
  const lastRow = Math.min(4, sheet.getLastRow());
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  return data.map((row, index) => {
    const rowObj = {};
    headers.forEach((header, i) => {
      rowObj[header] = row[i];
    });
    return {
      rowNumber: index + 2,
      email: row[emailCol] || '',
      name: row[nameCol] || '',
      data: rowObj
    };
  });
}

/**
 * Convert column index to letter
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// ==================== USER PROPERTIES ====================

/**
 * Save user preference
 */
function saveUserPreference(key, value) {
  PropertiesService.getUserProperties().setProperty(key, JSON.stringify(value));
}

/**
 * Get user preference
 */
function getUserPreference(key) {
  const value = PropertiesService.getUserProperties().getProperty(key);
  return value ? JSON.parse(value) : null;
}

/**
 * Get user's email address
 */
function getUserEmail() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (email && email.includes('@')) {
      return email;
    }
    // Fallback: try effective user
    const effectiveEmail = Session.getEffectiveUser().getEmail();
    if (effectiveEmail && effectiveEmail.includes('@')) {
      return effectiveEmail;
    }
    return '';
  } catch (e) {
    return '';
  }
}

/**
 * Get sender aliases (Send-as addresses) from Gmail API
 * Uses Gmail Advanced Service to get actual verified aliases
 */
function getSenderAliases() {
  try {
    // Use Gmail API to get send-as aliases
    const sendAsSettings = Gmail.Users.Settings.SendAs.list('me');
    const aliases = [];
    
    if (sendAsSettings && sendAsSettings.sendAs) {
      sendAsSettings.sendAs.forEach(sendAs => {
        // Only include verified aliases (verificationStatus === 'accepted' or isPrimary)
        if (sendAs.isPrimary || sendAs.verificationStatus === 'accepted') {
          aliases.push({
            email: sendAs.sendAsEmail,
            name: sendAs.displayName || '',
            isPrimary: sendAs.isPrimary || false,
            replyTo: sendAs.replyToAddress || '',
            isDefault: sendAs.isDefault || false
          });
        }
      });
    }
    
    // Sort: primary first, then alphabetically
    aliases.sort((a, b) => {
      if (a.isPrimary) return -1;
      if (b.isPrimary) return 1;
      return a.email.localeCompare(b.email);
    });
    
    return aliases;
  } catch (e) {
    // Fallback to Session API if Gmail API fails
    Logger.log('Gmail API error, falling back: ' + e.message);
    try {
      const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
      if (email && email.includes('@')) {
        return [{
          email: email,
          name: getUserDisplayName(),
          isPrimary: true
        }];
      }
    } catch (e2) {
      Logger.log('Fallback also failed: ' + e2.message);
    }
    return [];
  }
}

/**
 * Add a custom sender alias
 */
function addSenderAlias(email, name) {
  try {
    const aliases = getUserPreference('customSenderAliases') || [];
    // Check if already exists
    if (!aliases.find(a => a.email === email)) {
      aliases.push({ email, name });
      saveUserPreference('customSenderAliases', aliases);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Remove a custom sender alias  
 */
function removeSenderAlias(email) {
  try {
    let aliases = getUserPreference('customSenderAliases') || [];
    aliases = aliases.filter(a => a.email !== email);
    saveUserPreference('customSenderAliases', aliases);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== DRIVE FUNCTIONS ====================

/**
 * Get files and folders from a Drive folder
 */
function getDriveFiles(folderId) {
  try {
    let folder;
    let parentId = null;
    
    if (folderId) {
      folder = DriveApp.getFolderById(folderId);
      // Get parent folder for "back" navigation
      const parents = folder.getParents();
      if (parents.hasNext()) {
        parentId = parents.next().getId();
      }
    } else {
      folder = DriveApp.getRootFolder();
    }
    
    const result = {
      currentFolderId: folder.getId(),
      currentFolderName: folder.getName(),
      parentFolderId: parentId,
      items: []
    };
    
    // Get folders first
    const folders = folder.getFolders();
    while (folders.hasNext()) {
      const subFolder = folders.next();
      result.items.push({
        id: subFolder.getId(),
        name: subFolder.getName(),
        mimeType: 'application/vnd.google-apps.folder',
        isFolder: true
      });
    }
    
    // Then get files
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      result.items.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: file.getMimeType(),
        size: file.getSize(),
        url: file.getUrl(),
        isFolder: false
      });
    }
    
    // Sort: folders first, then files (both alphabetically)
    result.items.sort((a, b) => {
      if (a.isFolder && !b.isFolder) return -1;
      if (!a.isFolder && b.isFolder) return 1;
      return a.name.localeCompare(b.name);
    });
    
    return result;
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Get file as blob for attachment
 */
function getFileBlob(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    return {
      name: file.getName(),
      mimeType: file.getMimeType(),
      bytes: Utilities.base64Encode(file.getBlob().getBytes())
    };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Save quick attach resume ID
 */
function saveQuickAttachResume(fileId, fileName) {
  saveUserPreference('quickAttachResume', { fileId, fileName });
}

/**
 * Get quick attach resume
 */
function getQuickAttachResume() {
  return getUserPreference('quickAttachResume');
}

/**
 * Clear quick attach resume
 */
function clearQuickAttachResume() {
  PropertiesService.getUserProperties().deleteProperty('quickAttachResume');
}

// ==============================================================================
// EMAIL SERVICE
// ==============================================================================

/**
 * Request to cancel email sending
 * Sets a flag that the sendEmails loop will check
 */
function requestCancelSending() {
  try {
    PropertiesService.getUserProperties().setProperty('cancelSending', 'true');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Clear the cancel sending flag
 */
function clearCancelFlag() {
  PropertiesService.getUserProperties().deleteProperty('cancelSending');
}

/**
 * Check if cancellation was requested
 */
function isCancelRequested() {
  return PropertiesService.getUserProperties().getProperty('cancelSending') === 'true';
}

/**
 * Main function to send emails
 */
function sendEmails(config) {
  // Clear any previous cancel flag at the start
  clearCancelFlag();
  const {
    sheetName,
    emailColumn,
    nameColumn,
    titleColumn,
    ccColumn,
    bccColumn,
    subject,
    body,
    attachments,
    statusColumn,
    scheduledTime,
    senderEmail,
    startRow,
    endRow,
    delayOption,
    delayMin,
    delayMax,
    fixedDelay
  } = config;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return { success: false, error: 'Sheet not found' };
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();
  const actualEndRow = endRow ? Math.min(endRow, lastRow) : lastRow;
  const actualStartRow = startRow || 2;
  
  const results = {
    total: 0,
    sent: 0,
    failed: 0,
    cancelled: false,
    errors: []
  };
  
  // Get attachment blobs
  const attachmentBlobs = [];
  if (attachments && attachments.length > 0) {
    for (const att of attachments) {
      try {
        if (att.fileId) {
          const file = DriveApp.getFileById(att.fileId);
          attachmentBlobs.push(file.getBlob());
        } else if (att.bytes) {
          const blob = Utilities.newBlob(
            Utilities.base64Decode(att.bytes),
            att.mimeType,
            att.name
          );
          attachmentBlobs.push(blob);
        }
      } catch (e) {
        results.errors.push({ row: 0, error: `Attachment error: ${e.message}` });
      }
    }
  }
  
  // Process each row
  for (let row = actualStartRow; row <= actualEndRow; row++) {
    // Check if cancellation was requested
    if (isCancelRequested()) {
      results.cancelled = true;
      updateStatus(sheet, row, statusColumn, 'Cancelled by user', headers.length);
      // Mark remaining rows as cancelled
      for (let remainingRow = row + 1; remainingRow <= actualEndRow; remainingRow++) {
        updateStatus(sheet, remainingRow, statusColumn, 'Cancelled - Not sent', headers.length);
      }
      clearCancelFlag();
      break;
    }
    
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowObj = {};
    headers.forEach((header, i) => {
      rowObj[header] = rowData[i];
    });
    
    // Get email(s)
    const emailStr = rowData[emailColumn] || '';
    const emails = parseMultipleEmails(emailStr);
    
    if (emails.length === 0) {
      // More detailed skip message for debugging
      const debugInfo = emailStr ? `(found: "${String(emailStr).substring(0, 30)}...")` : '(cell empty)';
      updateStatus(sheet, row, statusColumn, `Skipped - No valid email ${debugInfo}`, headers.length);
      continue;
    }
    
    // Get name, title and apply smart greeting
    // Use typeof check to properly handle column index 0
    const hasNameCol = typeof nameColumn === 'number';
    const hasTitleCol = typeof titleColumn === 'number';
    const name = hasNameCol ? String(rowData[nameColumn] || '').trim() : '';
    const title = hasTitleCol ? String(rowData[titleColumn] || '').trim() : '';
    const greeting = getSmartGreeting(name, title);
    
    // Get CC and BCC
    const cc = ccColumn !== undefined ? (rowData[ccColumn] || '') : '';
    const bcc = bccColumn !== undefined ? (rowData[bccColumn] || '') : '';
    
    // Replace variables in subject and body
    const processedSubject = replaceVariables(subject, rowObj, greeting);
    const processedBody = replaceVariables(body, rowObj, greeting);
    
    results.total++;
    
    try {
      // Build email options
      const emailOptions = {
        to: emails.join(','),
        subject: processedSubject,
        htmlBody: processedBody,
        name: getUserDisplayName()
      };
      
      // Use sender alias if specified
      if (senderEmail && senderEmail.includes('@')) {
        emailOptions.from = senderEmail;
      }
      
      if (cc) emailOptions.cc = cc;
      if (bcc) emailOptions.bcc = bcc;
      if (attachmentBlobs.length > 0) emailOptions.attachments = attachmentBlobs;
      
      // Handle scheduled send
      if (scheduledTime) {
        const draft = GmailApp.createDraft(
          emailOptions.to,
          emailOptions.subject,
          '',
          emailOptions
        );
        storeDraftForSchedule(draft.getId(), scheduledTime, row);
        updateStatus(sheet, row, statusColumn, `Scheduled: ${formatDateTime(new Date(scheduledTime))}`, headers.length);
      } else {
        // Send immediately
        GmailApp.sendEmail(emailOptions.to, emailOptions.subject, '', emailOptions);
        updateStatus(sheet, row, statusColumn, `Sent: ${formatDateTime(new Date())}`, headers.length);
      }
      
      results.sent++;
      
    } catch (e) {
      results.failed++;
      results.errors.push({ row: row, email: emails[0], error: e.message });
      updateStatus(sheet, row, statusColumn, `Failed: ${e.message}`, headers.length);
    }
    
    // Add delay between emails to avoid rate limiting and spam filters
    if (row < actualEndRow) {
      let delayMs = 100; // Default minimal delay
      if (delayOption === 'random') {
        // Random delay between delayMin and delayMax seconds
        const min = delayMin || 2;
        const max = delayMax || 6;
        delayMs = (Math.floor(Math.random() * (max - min + 1)) + min) * 1000;
      } else if (delayOption === 'fixed' && fixedDelay) {
        delayMs = fixedDelay * 1000;
      } else if (delayOption === 'none') {
        delayMs = 100; // Minimal delay for processing
      }
      Utilities.sleep(delayMs);
    }
  }
  
  return results;
}

/**
 * Send a single test email
 */
function sendTestEmail(config) {
  const { subject, body, attachments } = config;
  
  // Get recipient email directly from Session
  let toEmail = '';
  try {
    toEmail = Session.getActiveUser().getEmail();
    if (!toEmail || !toEmail.includes('@')) {
      toEmail = Session.getEffectiveUser().getEmail();
    }
  } catch (e) {
    return { success: false, error: 'Could not get your email address. Please check script permissions.' };
  }
  
  if (!toEmail || !toEmail.includes('@')) {
    return { success: false, error: 'Could not get your email address. Please authorize the script properly.' };
  }
  
  try {
    const emailOptions = {
      htmlBody: body,
      name: getUserDisplayName()
    };
    
    if (attachments && attachments.length > 0) {
      const attachmentBlobs = [];
      for (const att of attachments) {
        if (att.fileId) {
          const file = DriveApp.getFileById(att.fileId);
          attachmentBlobs.push(file.getBlob());
        } else if (att.bytes) {
          const blob = Utilities.newBlob(
            Utilities.base64Decode(att.bytes),
            att.mimeType,
            att.name
          );
          attachmentBlobs.push(blob);
        }
      }
      if (attachmentBlobs.length > 0) {
        emailOptions.attachments = attachmentBlobs;
      }
    }
    
    GmailApp.sendEmail(toEmail, subject, '', emailOptions);
    return { success: true, sentTo: toEmail };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Send a test email to a specific address (for when Session.getActiveUser doesn't work)
 */
function sendTestEmailWithAddress(config) {
  const { to, subject, body, attachments } = config;
  
  if (!to || !to.includes('@')) {
    return { success: false, error: 'Please provide a valid email address.' };
  }
  
  try {
    const emailOptions = {
      htmlBody: body,
      name: getUserDisplayName()
    };
    
    if (attachments && attachments.length > 0) {
      const attachmentBlobs = [];
      for (const att of attachments) {
        if (att.fileId) {
          const file = DriveApp.getFileById(att.fileId);
          attachmentBlobs.push(file.getBlob());
        } else if (att.bytes) {
          const blob = Utilities.newBlob(
            Utilities.base64Decode(att.bytes),
            att.mimeType,
            att.name
          );
          attachmentBlobs.push(blob);
        }
      }
      if (attachmentBlobs.length > 0) {
        emailOptions.attachments = attachmentBlobs;
      }
    }
    
    GmailApp.sendEmail(to, subject, '', emailOptions);
    return { success: true, sentTo: to };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Parse multiple emails from a string
 */
function parseMultipleEmails(emailStr) {
  if (!emailStr) return [];
  
  // Convert to string if not already
  let str = String(emailStr).trim();
  if (!str) return [];
  
  // Normalize: replace ALL newlines/carriage returns with comma first
  str = str.replace(/\r\n/g, ',').replace(/\r/g, ',').replace(/\n/g, ',');
  
  // Also handle spaces that might separate emails (but only if it looks like email boundary)
  str = str.replace(/\s+(?=[a-zA-Z0-9._%+-]+@)/g, ',');
  
  return str
    .split(/[,;]+/)
    .map(e => e.trim())
    .filter(e => e && isValidEmail(e));
}

/**
 * Validate email format
 */
function isValidEmail(email) {
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}

/**
 * Get smart greeting based on name and title (Mr./Mrs.)
 * Logic:
 * - Title + Name: "Dear [Name] Sir/Ma'am"
 * - No Title + Name: "Dear [Name]"
 * - Title + No Name: "Dear Sir/Ma'am" (based on Mr./Mrs.)
 * - No Title + No Name: "Dear Sir/Ma'am"
 */
function getSmartGreeting(name, title) {
  const trimmedName = name ? name.toString().trim() : '';
  const trimmedTitle = title ? title.toString().trim().toLowerCase() : '';
  
  // Determine if title indicates Sir or Ma'am
  const isMr = trimmedTitle.includes('mr') && !trimmedTitle.includes('mrs') && !trimmedTitle.includes('ms');
  const isMrs = trimmedTitle.includes('mrs') || trimmedTitle.includes('ms') || trimmedTitle.includes('miss');
  
  if (trimmedName) {
    // Has name
    if (isMr) {
      return `Dear ${trimmedName} Sir`;
    } else if (isMrs) {
      return `Dear ${trimmedName} Ma'am`;
    } else {
      // No title, just use name
      return `Dear ${trimmedName}`;
    }
  } else {
    // No name
    if (isMr) {
      return 'Dear Sir';
    } else if (isMrs) {
      return "Dear Ma'am";
    } else {
      // No name, no title
      return "Dear Sir/Ma'am";
    }
  }
}

/**
 * Replace variables in template
 */
function replaceVariables(template, rowData, greeting) {
  if (!template) return '';
  
  let result = template;
  
  // Replace greeting variable
  result = result.replace(/\{\{greeting\}\}/gi, greeting);
  
  // Replace column variables
  for (const [key, value] of Object.entries(rowData)) {
    const regex = new RegExp(`\\{\\{${escapeRegex(key)}\\}\\}`, 'gi');
    result = result.replace(regex, value || '');
  }
  
  return result;
}

/**
 * Escape special regex characters
 */
function escapeRegex(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Update status cell
 */
function updateStatus(sheet, row, statusColumn, status, totalColumns) {
  if (statusColumn === undefined || statusColumn === null || statusColumn === -1) {
    const statusColIndex = totalColumns + 1;
    if (sheet.getRange(1, statusColIndex).getValue() !== 'Email Status') {
      sheet.getRange(1, statusColIndex).setValue('Email Status');
    }
    sheet.getRange(row, statusColIndex).setValue(status);
  } else {
    sheet.getRange(row, statusColumn + 1).setValue(status);
  }
}

/**
 * Get user's display name
 */
function getUserDisplayName() {
  try {
    const email = Session.getActiveUser().getEmail();
    const name = email.split('@')[0].replace(/[._]/g, ' ');
    return name.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
  } catch (e) {
    return '';
  }
}

/**
 * Format date time for display
 */
function formatDateTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM dd, yyyy HH:mm');
}

/**
 * Store draft for scheduled sending
 */
function storeDraftForSchedule(draftId, scheduledTime, row) {
  const schedules = getUserPreference('scheduledEmails') || [];
  schedules.push({
    draftId: draftId,
    scheduledTime: scheduledTime,
    row: row,
    created: new Date().toISOString()
  });
  saveUserPreference('scheduledEmails', schedules);
  createScheduleTrigger();
}

/**
 * Create trigger to check scheduled emails
 */
function createScheduleTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'processScheduledEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  ScriptApp.newTrigger('processScheduledEmails')
    .timeBased()
    .everyMinutes(1)
    .create();
}

/**
 * Process scheduled emails (called by trigger)
 */
function processScheduledEmails() {
  const schedules = getUserPreference('scheduledEmails') || [];
  const now = new Date().getTime();
  const remaining = [];
  
  for (const schedule of schedules) {
    if (new Date(schedule.scheduledTime).getTime() <= now) {
      try {
        const drafts = GmailApp.getDrafts();
        for (const draft of drafts) {
          if (draft.getId() === schedule.draftId) {
            draft.send();
            break;
          }
        }
      } catch (e) {
        Logger.log('Error sending scheduled email: ' + e.message);
      }
    } else {
      remaining.push(schedule);
    }
  }
  
  saveUserPreference('scheduledEmails', remaining);
  
  if (remaining.length === 0) {
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'processScheduledEmails') {
        ScriptApp.deleteTrigger(trigger);
      }
    }
  }
}

/**
 * Get all scheduled emails for display in sidebar
 */
function getScheduledEmails() {
  try {
    const schedules = getUserPreference('scheduledEmails') || [];
    const result = [];
    
    for (const schedule of schedules) {
      const scheduledDate = new Date(schedule.scheduledTime);
      const now = new Date();
      const timeRemaining = scheduledDate.getTime() - now.getTime();
      
      // Check if draft still exists
      let draftSubject = 'Loading...';
      try {
        const draft = GmailApp.getDraft(schedule.draftId);
        if (draft) {
          const message = draft.getMessage();
          draftSubject = message.getSubject() || '(No subject)';
        }
      } catch (e) {
        draftSubject = '(Draft deleted)';
      }
      
      result.push({
        id: schedule.draftId,
        subject: draftSubject,
        scheduledTime: schedule.scheduledTime,
        scheduledTimeFormatted: formatDateTime(scheduledDate),
        timeRemaining: timeRemaining > 0 ? timeRemaining : 0,
        row: schedule.row,
        created: schedule.created
      });
    }
    
    // Sort by scheduled time (soonest first)
    result.sort((a, b) => new Date(a.scheduledTime) - new Date(b.scheduledTime));
    
    return { success: true, emails: result };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Cancel a scheduled email by draft ID
 */
function cancelScheduledEmail(draftId) {
  try {
    const schedules = getUserPreference('scheduledEmails') || [];
    const remaining = schedules.filter(s => s.draftId !== draftId);
    
    // Delete the draft from Gmail
    try {
      const draft = GmailApp.getDraft(draftId);
      if (draft) {
        draft.deleteDraft();
      }
    } catch (e) {
      // Draft may already be deleted, continue
    }
    
    saveUserPreference('scheduledEmails', remaining);
    
    // Remove trigger if no more scheduled emails
    if (remaining.length === 0) {
      const triggers = ScriptApp.getProjectTriggers();
      for (const trigger of triggers) {
        if (trigger.getHandlerFunction() === 'processScheduledEmails') {
          ScriptApp.deleteTrigger(trigger);
        }
      }
    }
    
    return { success: true, message: 'Scheduled email cancelled' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Cancel all scheduled emails
 */
function cancelAllScheduledEmails() {
  try {
    const schedules = getUserPreference('scheduledEmails') || [];
    
    // Delete all drafts
    for (const schedule of schedules) {
      try {
        const draft = GmailApp.getDraft(schedule.draftId);
        if (draft) {
          draft.deleteDraft();
        }
      } catch (e) {
        // Continue even if draft deletion fails
      }
    }
    
    // Clear scheduled emails
    saveUserPreference('scheduledEmails', []);
    
    // Remove trigger
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'processScheduledEmails') {
        ScriptApp.deleteTrigger(trigger);
      }
    }
    
    return { success: true, message: 'All scheduled emails cancelled' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Get remaining email quota for the day
 */
function getRemainingDailyQuota() {
  try {
    return MailApp.getRemainingDailyQuota();
  } catch (e) {
    return 500;
  }
}

/**
 * Check if we can send X emails
 */
function canSendEmails(count) {
  const remaining = getRemainingDailyQuota();
  return {
    canSend: remaining >= count,
    remaining: remaining,
    requested: count
  };
}

// ==============================================================================
// TEMPLATE SERVICE
// ==============================================================================

const TEMPLATE_FOLDER_NAME = 'Mail Merge Pro Templates';

/**
 * Get or create template folder in Google Drive
 */
function getTemplateFolder() {
  const folders = DriveApp.getFoldersByName(TEMPLATE_FOLDER_NAME);
  
  if (folders.hasNext()) {
    return folders.next();
  }
  
  return DriveApp.createFolder(TEMPLATE_FOLDER_NAME);
}

/**
 * Save a template to Drive
 */
function saveTemplate(templateData) {
  try {
    const folder = getTemplateFolder();
    const fileName = `${templateData.name}.json`;
    
    const existingFiles = folder.getFilesByName(fileName);
    if (existingFiles.hasNext()) {
      const file = existingFiles.next();
      file.setContent(JSON.stringify(templateData, null, 2));
      return { success: true, id: file.getId(), message: 'Template updated' };
    }
    
    const file = folder.createFile(fileName, JSON.stringify(templateData, null, 2), 'application/json');
    return { success: true, id: file.getId(), message: 'Template saved' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Get all saved templates
 */
function getTemplates() {
  try {
    const folder = getTemplateFolder();
    const files = folder.getFiles();
    const templates = [];
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      // Check for JSON files
      if (fileName.endsWith('.json')) {
        try {
          const content = JSON.parse(file.getBlob().getDataAsString());
          templates.push({
            id: file.getId(),
            name: content.name || fileName.replace('.json', ''),
            subject: content.subject || '',
            body: content.body || '',
            created: file.getDateCreated().toISOString(),
            modified: file.getLastUpdated().toISOString()
          });
        } catch (parseError) {
          // Skip files that can't be parsed
          Logger.log('Could not parse template file: ' + fileName);
        }
      }
    }
    
    // Sort by modified date (newest first)
    templates.sort((a, b) => new Date(b.modified) - new Date(a.modified));
    
    return templates;
  } catch (e) {
    Logger.log('getTemplates error: ' + e.message);
    return [];
  }
}

/**
 * Load a specific template
 */
function loadTemplate(templateId) {
  try {
    const file = DriveApp.getFileById(templateId);
    const content = JSON.parse(file.getBlob().getDataAsString());
    return { success: true, template: content };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Delete a template
 */
function deleteTemplate(templateId) {
  try {
    const file = DriveApp.getFileById(templateId);
    file.setTrashed(true);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Rename a template
 */
function renameTemplate(templateId, newName) {
  try {
    const file = DriveApp.getFileById(templateId);
    const content = JSON.parse(file.getBlob().getDataAsString());
    content.name = newName;
    file.setContent(JSON.stringify(content, null, 2));
    file.setName(`${newName}.json`);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Get default job hunting templates
 */
function getDefaultTemplates() {
  return [
    {
      name: 'Job Application - General',
      subject: 'Application for {{Position}} Role - {{Name}}',
      body: `<p>{{greeting}},</p>
<p>I hope this email finds you well. I am writing to express my strong interest in the <b>{{Position}}</b> position at <b>{{Company}}</b>.</p>
<p>With my background in [Your Experience], I believe I would be a great fit for your team. I am particularly excited about [What excites you about the company].</p>
<p>I have attached my resume for your review. I would welcome the opportunity to discuss how my skills and experiences align with your team's needs.</p>
<p>Thank you for considering my application. I look forward to hearing from you.</p>
<p>Best regards,<br>{{SenderName}}</p>`
    },
    {
      name: 'Networking - HR Connection',
      subject: 'Quick Question about Opportunities at {{Company}}',
      body: `<p>{{greeting}},</p>
<p>I came across your profile while researching opportunities at <b>{{Company}}</b>, and I was impressed by the company's work in [Industry/Field].</p>
<p>I'm currently exploring new opportunities and would love to learn more about the culture and potential openings at {{Company}}. Would you have 15 minutes for a brief chat?</p>
<p>I've attached my resume for your reference. Please let me know if there's a convenient time for you.</p>
<p>Thank you for your time!</p>
<p>Best regards,<br>{{SenderName}}</p>`
    },
    {
      name: 'Cold Email - Recruiter',
      subject: 'Experienced {{YourRole}} Interested in {{Company}}',
      body: `<p>{{greeting}},</p>
<p>I'm reaching out because I'm actively looking for new opportunities, and {{Company}} caught my attention for its [Reason].</p>
<p>A bit about me:</p>
<ul>
<li>[X] years of experience in [Your Field]</li>
<li>Expertise in [Key Skills]</li>
<li>Previously worked at [Notable Companies]</li>
</ul>
<p>I'd love to explore if there might be a fit on your team. Would you be open to a brief conversation?</p>
<p>Resume attached for your reference.</p>
<p>Thanks,<br>{{SenderName}}</p>`
    }
  ];
}

/**
 * Install default templates
 */
function installDefaultTemplates() {
  const defaults = getDefaultTemplates();
  const results = [];
  
  for (const template of defaults) {
    const result = saveTemplate(template);
    results.push(result);
  }
  
  return results;
}

// ==============================================================================
// SCRIPT EXECUTION MONITORING
// ==============================================================================

/**
 * Save execution status for tracking
 */
function saveExecutionStatus(status) {
  try {
    const props = PropertiesService.getUserProperties();
    props.setProperty('currentExecution', JSON.stringify(status));
    
    // Also save to history
    let history = JSON.parse(props.getProperty('executionHistory') || '[]');
    
    // Update existing or add new
    const existingIndex = history.findIndex(h => h.id === status.id);
    if (existingIndex >= 0) {
      history[existingIndex] = status;
    } else {
      history.unshift(status);
    }
    
    // Keep only last 20 executions
    history = history.slice(0, 20);
    props.setProperty('executionHistory', JSON.stringify(history));
  } catch (e) {
    Logger.log('Error saving execution status: ' + e.message);
  }
}

/**
 * Save execution to history
 */
function saveExecutionVal(exec) {
  try {
    const props = PropertiesService.getUserProperties();
    const historyStr = props.getProperty('executionHistory');
    let history = historyStr ? JSON.parse(historyStr) : [];
    
    // Add new execution to start
    history.unshift(exec);
    
    // Keep last 20
    history = history.slice(0, 20);
    props.setProperty('executionHistory', JSON.stringify(history));
  } catch (e) {
    Logger.log('Error saving execution history: ' + e.message);
  }
}

/**
 * Clear current execution (when complete)
 */
function clearCurrentExecution() {
  PropertiesService.getUserProperties().deleteProperty('currentExecution');
}

/**
 * Get script execution status for the monitor
 */
function getScriptExecutionStatus() {
  try {
    const props = PropertiesService.getUserProperties();
    
    // Get current running execution  
    const currentStr = props.getProperty('currentExecution');
    const current = currentStr ? JSON.parse(currentStr) : null;
    
    // Get history
    const historyStr = props.getProperty('executionHistory');
    let history = historyStr ? JSON.parse(historyStr) : [];
    
    // Format for display
    const running = [];
    if (current && current.status === 'Running') {
      running.push(formatExecutionForDisplay(current));
    }
    
    // Format history (exclude currently running)
    const formattedHistory = history
      .filter(h => h.status !== 'Running' || !current || h.id !== current.id)
      .slice(0, 10)
      .map(formatExecutionForDisplay);
    
    return {
      success: true,
      running: running,
      history: formattedHistory
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Format execution data for display
 */
function formatExecutionForDisplay(exec) {
  const now = new Date();
  const startTime = new Date(exec.startTime);
  const elapsed = Math.floor((now - startTime) / 1000);
  
  let timeRemaining = null;
  if (exec.status === 'Running' && exec.sent > 0 && exec.total > 0) {
    const avgTimePerEmail = elapsed / exec.sent;
    const remaining = (exec.total - exec.sent) * avgTimePerEmail;
    timeRemaining = formatDuration(remaining);
  }
  
  return {
    id: exec.id,
    subject: exec.subject || 'Email Campaign',
    status: exec.status,
    total: exec.total || 0,
    sent: exec.sent || 0,
    failed: exec.failed || 0,
    startTime: formatDateTime(startTime),
    elapsed: formatDuration(elapsed),
    timeRemaining: timeRemaining,
    endTime: exec.endTime ? formatDateTime(new Date(exec.endTime)) : null
  };
}

/**
 * Format date to readable string
 */
function formatDateTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM dd, yyyy 'at' hh:mm a");
}

/**
 * Format duration in seconds to readable string
 */
function formatDuration(seconds) {
  if (seconds < 60) return Math.round(seconds) + 's';
  const mins = Math.floor(seconds / 60);
  const secs = Math.round(seconds % 60);
  if (mins < 60) return mins + 'm ' + secs + 's';
  const hours = Math.floor(mins / 60);
  return hours + 'h ' + (mins % 60) + 'm';
}

// ==============================================================================
// SCHEDULED CAMPAIGN FEATURE
// ==============================================================================

/**
 * Schedule an email campaign to run at a specific time
 * Creates a time-based trigger that will execute sendEmails with stored config
 */
/**
 * Helper to get/create campaign folder
 */
function getCampaignFolder() {
  const folders = DriveApp.getFoldersByName("Mail Merge Campaigns");
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder("Mail Merge Campaigns");
}

/**
 * Cleanup function to resolve Property Quota issues immediately
 */
function cleanupStorage() {
  try {
    const props = PropertiesService.getUserProperties();
    // Clear old execution history if it exists
    props.deleteProperty('executionHistory');
    // Clear old scheduled campaigns if corrupted
    // props.deleteProperty('scheduledCampaigns'); // Optional: only if really stuck
    return { success: true, message: 'Storage cleaned' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Schedule an email campaign to run at a specific time
 * Creates a time-based trigger that will execute sendEmails with stored config
 */
function scheduleEmailCampaign(config, scheduledTimeISO) {
  try {
    // 1. Check for immediate cleanup if quota is tight
    cleanupStorage();

    const campaignId = 'campaign_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    const scheduledDate = new Date(scheduledTimeISO);
    
    if (scheduledDate <= new Date()) {
      return { success: false, error: 'Scheduled time must be in the future' };
    }
    
    // 2. Save heavy config to Drive
    // Remove complex objects like blobs if needed, or serialize them properly
    // Note: Config usually contains base64 strings for attachments which is fine for JSON
    const folder = getCampaignFolder();
    const fileName = "config_" + campaignId + ".json";
    const file = folder.createFile(fileName, JSON.stringify(config), MimeType.PLAIN_TEXT);
    
    // 3. Store LIGHTWEIGHT campaign metadata
    const campaigns = JSON.parse(PropertiesService.getUserProperties().getProperty('scheduledCampaigns') || '[]');
    campaigns.push({
      id: campaignId,
      // config: config, // <--- DO NOT STORE FULL CONFIG HERE
      driveFileId: file.getId(), // Store reference instead
      scheduledTime: scheduledTimeISO,
      subject: config.subject,
      recipientCount: (config.endRow || 0) - (config.startRow || 2) + 1,
      createdAt: new Date().toISOString(),
      status: 'scheduled'
    });
    PropertiesService.getUserProperties().setProperty('scheduledCampaigns', JSON.stringify(campaigns));
    
    // 4. Create time-based trigger
    const trigger = ScriptApp.newTrigger('runScheduledCampaign')
      .timeBased()
      .at(scheduledDate)
      .create();
    
    // 5. Store trigger ID with campaign
    const updatedCampaigns = JSON.parse(PropertiesService.getUserProperties().getProperty('scheduledCampaigns') || '[]');
    const campaign = updatedCampaigns.find(c => c.id === campaignId);
    if (campaign) {
      campaign.triggerId = trigger.getUniqueId();
    }
    PropertiesService.getUserProperties().setProperty('scheduledCampaigns', JSON.stringify(updatedCampaigns));
    
    return { 
      success: true, 
      campaignId: campaignId,
      scheduledTime: formatDateTime(scheduledDate),
      message: 'Campaign scheduled for ' + formatDateTime(scheduledDate)
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Callback function for scheduled campaign trigger
 * This runs at the scheduled time and sends the emails
 */
function runScheduledCampaign(e) {
  try {
    const triggerId = e.triggerUid;
    const campaigns = JSON.parse(PropertiesService.getUserProperties().getProperty('scheduledCampaigns') || '[]');
    
    // Find the campaign for this trigger
    const campaignIndex = campaigns.findIndex(c => c.triggerId === triggerId);
    if (campaignIndex === -1) {
      Logger.log('No campaign found for trigger: ' + triggerId);
      return;
    }
    
    const campaign = campaigns[campaignIndex];
    let config = campaign.config;

    // Retrieve config from Drive if present (New Method)
    if (campaign.driveFileId) {
      try {
        const file = DriveApp.getFileById(campaign.driveFileId);
        const content = file.getBlob().getDataAsString();
        config = JSON.parse(content);
        
        // Clean up file immediately? Or keep for history?
        // Let's keep it until execution is confirmed done, or delete now if we have memory.
        // Doing delete at end.
      } catch (driveError) {
        Logger.log("Error reading config from Drive: " + driveError.message);
        // If fail, we can't run.
        campaign.status = 'failed';
        campaign.results = { error: 'Could not load campaign configuration.' };
        PropertiesService.getUserProperties().setProperty('scheduledCampaigns', JSON.stringify(campaigns));
        return;
      }
    }
    
    if (!config) {
         Logger.log('No config available for campaign ' + campaign.id);
         return;
    }

    // Update status to running
    campaign.status = 'running';
    campaign.startedAt = new Date().toISOString();
    PropertiesService.getUserProperties().setProperty('scheduledCampaigns', JSON.stringify(campaigns));
    
    // Initialize execution tracking
    const executionId = 'exec_' + Date.now();
    const execution = {
      id: executionId,
      status: 'Running',
      startTime: new Date().toISOString(),
      total: campaign.recipientCount,
      sent: 0,
      failed: 0,
      subject: campaign.subject
    };
    PropertiesService.getUserProperties().setProperty('currentExecution', JSON.stringify(execution));

    // Run the email sending
    // We need to inject the executionId into config so sendEmails can update it
    config.executionId = executionId; // Use the loaded config
    const result = sendEmails(config);
    
    // Finalize execution
    execution.status = 'Completed';
    execution.endTime = new Date().toISOString();
    execution.sent = result.sent;
    execution.failed = result.failed;
    
    // Save to history
    saveExecutionVal(execution); // Using helper if available, or manual save
    PropertiesService.getUserProperties().deleteProperty('currentExecution');
    
    // Update campaign with results
    campaign.status = 'completed';
    campaign.completedAt = new Date().toISOString();
    campaign.results = {
      sent: result.sent,
      failed: result.failed,
      total: result.total
    };
    PropertiesService.getUserProperties().setProperty('scheduledCampaigns', JSON.stringify(campaigns));
    
    // Cleanup Drive File
    if (campaign.driveFileId) {
        try {
            DriveApp.getFileById(campaign.driveFileId).setTrashed(true);
        } catch(e) { /* ignore cleanup error */ }
    }

    // Delete the trigger (it already fired)
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(trigger);
        break;
      }
    }
    
    Logger.log('Scheduled campaign completed: ' + campaign.id + ', sent: ' + result.sent);
    
  } catch (e) {
    Logger.log('Error running scheduled campaign: ' + e.message);
  }
}

/**
 * Get list of scheduled campaigns
 */
function getScheduledCampaigns() {
  try {
    const campaigns = JSON.parse(PropertiesService.getUserProperties().getProperty('scheduledCampaigns') || '[]');
    const now = new Date();
    
    // Format campaigns for display
    const result = campaigns.map(c => ({
      id: c.id,
      subject: c.subject || 'Email Campaign',
      scheduledTime: c.scheduledTime,
      scheduledTimeFormatted: c.scheduledTime ? Utilities.formatDate(new Date(c.scheduledTime), Session.getScriptTimeZone(), "MMM dd, yyyy 'at' hh:mm a") : 'Unknown',
      timeRemaining: c.scheduledTime ? Math.max(0, new Date(c.scheduledTime).getTime() - now.getTime()) : 0,
      recipientCount: c.recipientCount || 0,
      status: c.status || 'scheduled',
      createdAt: c.createdAt,
      results: c.results || null
    }));
    
    // Sort: Running > Scheduled (Asc) > Completed/Cancelled (Desc)
    result.sort((a, b) => {
      const getScore = (s) => {
        if (s === 'running') return 0;
        if (s === 'scheduled') return 1;
        return 2;
      };
      
      const scoreA = getScore(a.status);
      const scoreB = getScore(b.status);
      
      if (scoreA !== scoreB) return scoreA - scoreB;
      
      // If both are scheduled, sort ascending (soonest first)
      if (a.status === 'scheduled') {
        return new Date(a.scheduledTime) - new Date(b.scheduledTime);
      }
      // Otherwise (completed/cancelled), sort descending (newest first)
      return new Date(b.scheduledTime) - new Date(a.scheduledTime);
    });
    
    return { success: true, campaigns: result };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Cancel a scheduled campaign
 */
function cancelScheduledCampaign(campaignId) {
  try {
    const campaigns = JSON.parse(PropertiesService.getUserProperties().getProperty('scheduledCampaigns') || '[]');
    const campaignIndex = campaigns.findIndex(c => c.id === campaignId);
    
    if (campaignIndex === -1) {
      return { success: false, error: 'Campaign not found' };
    }
    
    const campaign = campaigns[campaignIndex];
    
    // Delete the trigger if it exists
    if (campaign.triggerId) {
      const triggers = ScriptApp.getProjectTriggers();
      for (const trigger of triggers) {
        if (trigger.getUniqueId() === campaign.triggerId) {
          ScriptApp.deleteTrigger(trigger);
          break;
        }
      }
    }
    
    // Remove campaign from list
    campaigns.splice(campaignIndex, 1);
    PropertiesService.getUserProperties().setProperty('scheduledCampaigns', JSON.stringify(campaigns));
    
    return { success: true, message: 'Campaign cancelled' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Reschedule a campaign
 */
function rescheduleCampaign(campaignId, newTimeISO) {
  try {
    const campaigns = JSON.parse(PropertiesService.getUserProperties().getProperty('scheduledCampaigns') || '[]');
    const campaignIndex = campaigns.findIndex(c => c.id === campaignId);
    
    if (campaignIndex === -1) {
      return { success: false, error: 'Campaign not found' };
    }
    
    const campaign = campaigns[campaignIndex];
    if (campaign.status !== 'scheduled') {
      return { success: false, error: 'Can only reschedule pending campaigns' };
    }

    const scheduledDate = new Date(newTimeISO);
    if (scheduledDate <= new Date()) {
      return { success: false, error: 'Scheduled time must be in the future' };
    }
    
    // Delete old trigger
    if (campaign.triggerId) {
      const triggers = ScriptApp.getProjectTriggers();
      for (const trigger of triggers) {
        if (trigger.getUniqueId() === campaign.triggerId) {
          ScriptApp.deleteTrigger(trigger);
          break;
        }
      }
    }
    
    // Create new trigger
    const newTrigger = ScriptApp.newTrigger('runScheduledCampaign')
      .timeBased()
      .at(scheduledDate)
      .create();
      
    // Update campaign
    campaign.scheduledTime = newTimeISO;
    campaign.triggerId = newTrigger.getUniqueId();
    
    PropertiesService.getUserProperties().setProperty('scheduledCampaigns', JSON.stringify(campaigns));
    
    return { success: true, message: 'Campaign rescheduled' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Request cancellation of current sending process
 * This sets a flag that sendEmails checks
 */
function requestCancelSending() {
  PropertiesService.getUserProperties().setProperty('cancelSending', 'true');
  return { success: true };
}

/**
 * Clear completed campaigns from history (older than 7 days)
 */
function clearOldCampaigns() {
  try {
    const campaigns = JSON.parse(PropertiesService.getUserProperties().getProperty('scheduledCampaigns') || '[]');
    const oneWeekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
    
    const remaining = campaigns.filter(c => {
      if (c.status === 'scheduled') return true; // Keep pending
      if (c.completedAt && new Date(c.completedAt) > oneWeekAgo) return true; // Keep recent completed
      return false;
    });
    
    PropertiesService.getUserProperties().setProperty('scheduledCampaigns', JSON.stringify(remaining));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==============================================================================
// LEGACY GMAIL SCHEDULING SUPPORT (Placeholder for stability)
// ==============================================================================

function getScheduledEmails() {
  return { success: true, emails: [] };
}

function getScheduledEmailsCount() {
  return 0; 
}

function cancelScheduledEmail(id) {
  return { success: true };
}

function cancelAllScheduledEmails() {
  return { success: true };
}