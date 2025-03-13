/**
 * Creates a backup copy of the entire March Madness spreadsheet in a dedicated backup folder
 * The backup will include a date stamp in the filename
 */
function createSpreadsheetBackup() {
  try {
    // Get the active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Format date for the filename
    const now = new Date();
    const dateString = Utilities.formatDate(now, Session.getScriptTimeZone(), 
      'yyyy-MM-dd_HH-mm');
    
    // Create a new filename with date
    const originalName = ss.getName();
    const newName = `${originalName}_Backup_${dateString}`;
    
    // Create a copy of the spreadsheet
    const originalFile = DriveApp.getFileById(ss.getId());
    const backup = originalFile.makeCopy(newName);
    
    // Get or create the backup folder
    const backupFolder = getOrCreateBackupFolder();
    
    // Move the backup to the backup folder
    const file = DriveApp.getFileById(backup.getId());
    backupFolder.addFile(file);
    
    // If the file is also in another folder (like My Drive), remove it from there
    const parents = file.getParents();
    while (parents.hasNext()) {
      const parent = parents.next();
      if (parent.getId() !== backupFolder.getId()) {
        parent.removeFile(file);
      }
    }
    
    // Log success message with the new file's URL
    const backupUrl = backup.getUrl();
    logBackupInfo('Backup Created', 'SUCCESS', `Backup created: ${newName}`, backupUrl);
    
    console.log(`Backup created successfully: ${newName}`);
    console.log(`URL: ${backupUrl}`);
    console.log(`Stored in folder: ${backupFolder.getName()}`);
    
    return {
      success: true,
      name: newName,
      url: backupUrl,
      folderId: backupFolder.getId(),
      folderName: backupFolder.getName()
    };
  } catch (error) {
    console.error("Error creating backup:", error.message, error.stack);
    
    // Log error to the sheet
    logBackupInfo('Backup Failed', 'ERROR', `Error: ${error.message}`);
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Gets or creates a dedicated backup folder in Drive
 */
function getOrCreateBackupFolder() {
  // Check if we've already stored the folder ID
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('BACKUP_FOLDER_ID');
  
  if (folderId) {
    try {
      // Try to get the folder with the stored ID
      const folder = DriveApp.getFolderById(folderId);
      return folder;
    } catch (e) {
      // Folder no longer exists, we'll create a new one
      console.log("Stored backup folder not found, creating new one.");
    }
  }
  
  // Get the spreadsheet name to use in the backup folder name
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssName = ss.getName();
  
  // Create the backup folder name
  const folderName = `${ssName} - Backups`;
  
  // Check if the folder already exists
  const folderIterator = DriveApp.getFoldersByName(folderName);
  if (folderIterator.hasNext()) {
    const folder = folderIterator.next();
    // Store the folder ID for future use
    props.setProperty('BACKUP_FOLDER_ID', folder.getId());
    return folder;
  }
  
  // Create a new folder
  const folder = DriveApp.createFolder(folderName);
  
  // Store the folder ID for future use
  props.setProperty('BACKUP_FOLDER_ID', folder.getId());
  
  return folder;
}

/**
 * Set a specific folder for backups
 */
function setBackupFolder() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set Backup Folder',
    'Enter the ID of the Google Drive folder where backups should be stored:\n' +
    '(You can find this in the folder URL: https://drive.google.com/drive/folders/FOLDER_ID_HERE)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const folderId = response.getResponseText().trim();
    
    try {
      // Verify the folder exists
      const folder = DriveApp.getFolderById(folderId);
      
      // Save the folder ID to script properties
      PropertiesService.getScriptProperties().setProperty('BACKUP_FOLDER_ID', folderId);
      
      ui.alert('Success', `Backup folder set to: ${folder.getName()}`, ui.ButtonSet.OK);
      logBackupInfo('Backup Settings', 'SUCCESS', `Backup folder set to: ${folder.getName()}`);
    } catch (e) {
      ui.alert('Error', 'Could not find a folder with that ID. Please check the ID and try again.', ui.ButtonSet.OK);
    }
  }
}

/**
 * Clear the backup folder setting and revert to default
 */
function resetBackupFolder() {
  PropertiesService.getScriptProperties().deleteProperty('BACKUP_FOLDER_ID');
  
  // Create a new default folder
  const folder = getOrCreateBackupFolder();
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('Backup Folder Reset', `Backup folder has been reset to: ${folder.getName()}`, ui.ButtonSet.OK);
  
  logBackupInfo('Backup Settings', 'INFO', `Backup folder reset to: ${folder.getName()}`);
}

/**
 * Log backup information to a sheet
 */
function logBackupInfo(action, status, details, url = '') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get or create backup log sheet
    let backupLogSheet = ss.getSheetByName('Backup Log');
    if (!backupLogSheet) {
      backupLogSheet = ss.insertSheet('Backup Log');
      backupLogSheet.appendRow(['Timestamp', 'Action', 'Status', 'Details', 'Backup URL']);
      backupLogSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f3f3f3');
    }
    
    // Add log entry
    backupLogSheet.appendRow([
      new Date(),
      action,
      status,
      details,
      url
    ]);
    
    // Format the URL as a hyperlink
    if (url) {
      const lastRow = backupLogSheet.getLastRow();
      backupLogSheet.getRange(lastRow, 5).setFormula(`=HYPERLINK("${url}","Open Backup")`);
    }
  } catch (error) {
    console.error('Failed to log backup info:', error.message);
  }
}

/**
 * Creates a scheduled trigger to run backups automatically
 */
function createDailyBackupTrigger() {
  // Remove any existing backup triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'createSpreadsheetBackup') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new trigger to run at midnight
  ScriptApp.newTrigger('createSpreadsheetBackup')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();
    
  console.log("Daily backup trigger created");
  logBackupInfo('Backup Schedule', 'SUCCESS', 'Daily backup trigger created to run at midnight');
  
  // Show confirmation to user
  SpreadsheetApp.getUi().alert('Daily Backup Scheduled', 'A backup will be created automatically each day at midnight.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Creates a weekly backup trigger
 */
function createWeeklyBackupTrigger() {
  // Remove any existing backup triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'createSpreadsheetBackup') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new trigger to run weekly (Sundays at midnight)
  ScriptApp.newTrigger('createSpreadsheetBackup')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(0)
    .create();
    
  console.log("Weekly backup trigger created");
  logBackupInfo('Backup Schedule', 'SUCCESS', 'Weekly backup trigger created to run Sundays at midnight');
  
  // Show confirmation to user
  SpreadsheetApp.getUi().alert('Weekly Backup Scheduled', 'A backup will be created automatically each Sunday at midnight.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Removes all backup triggers
 */
function removeBackupTriggers() {
  // Remove any existing backup triggers
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'createSpreadsheetBackup') {
      ScriptApp.deleteTrigger(triggers[i]);
      count++;
    }
  }
  
  console.log(`${count} backup triggers removed`);
  logBackupInfo('Backup Schedule', 'INFO', `${count} backup triggers removed`);
  
  // Show confirmation to user
  if (count > 0) {
    SpreadsheetApp.getUi().alert('Backup Schedule Removed', 'All scheduled backups have been canceled.', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('No Schedules Found', 'There were no backup schedules to remove.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * List all current backup files
 */
function listBackups() {
  try {
    // Get the backup folder
    const backupFolder = getOrCreateBackupFolder();
    
    // Get all files in the backup folder
    const fileIterator = backupFolder.getFiles();
    const backups = [];
    
    while (fileIterator.hasNext()) {
      const file = fileIterator.next();
      backups.push({
        name: file.getName(),
        date: file.getDateCreated(),
        url: file.getUrl(),
        id: file.getId()
      });
    }
    
    // Sort backups by date (newest first)
    backups.sort((a, b) => b.date - a.date);
    
    // Get the active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create or clear the backup list sheet
    let backupListSheet = ss.getSheetByName('Backup List');
    if (!backupListSheet) {
      backupListSheet = ss.insertSheet('Backup List');
      backupListSheet.appendRow(['Backup Name', 'Creation Date', 'URL', 'File ID']);
      backupListSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f3f3f3');
    } else {
      // Clear existing content except header
      if (backupListSheet.getLastRow() > 1) {
        backupListSheet.getRange(2, 1, backupListSheet.getLastRow() - 1, 4).clear();
      }
    }
    
    // Add backup list to sheet
    for (const backup of backups) {
      backupListSheet.appendRow([
        backup.name,
        backup.date,
        backup.url,
        backup.id
      ]);
      
      // Format the URL as a hyperlink
      const lastRow = backupListSheet.getLastRow();
      backupListSheet.getRange(lastRow, 3).setFormula(`=HYPERLINK("${backup.url}","Open Backup")`);
    }
    
    // Format the sheet
    backupListSheet.autoResizeColumns(1, 4);
    
    // Add folder information at the top
    backupListSheet.insertRowBefore(1);
    backupListSheet.getRange('A1:D1').merge();
    backupListSheet.getRange('A1').setValue(`Backup Folder: ${backupFolder.getName()}`);
    
    backupListSheet.insertRowBefore(2);
    backupListSheet.getRange('A2:D2').merge();
    backupListSheet.getRange('A2').setFormula(`=HYPERLINK("https://drive.google.com/drive/folders/${backupFolder.getId()}","Open Backup Folder")`);
    
    // Highlight the header
    backupListSheet.getRange('A3:D3').setFontWeight('bold').setBackground('#f3f3f3');
    
    SpreadsheetApp.getUi().alert(`Found ${backups.length} backups in folder "${backupFolder.getName()}"`);
    
    return backups;
  } catch (error) {
    console.error("Error listing backups:", error.message, error.stack);
    SpreadsheetApp.getUi().alert('Error', `Failed to list backups: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return [];
  }
}

/**
 * Add backup options to the menu
 */
/*function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Backup');
  
  menu.addItem('Create Backup Now', 'createSpreadsheetBackup')
    .addItem('List All Backups', 'listBackups')
    .addSeparator()
    .addItem('Set Custom Backup Folder', 'setBackupFolder')
    .addItem('Reset to Default Backup Folder', 'resetBackupFolder')
    .addSeparator()
    .addItem('Schedule Daily Backups', 'createDailyBackupTrigger')
    .addItem('Schedule Weekly Backups', 'createWeeklyBackupTrigger')
    .addItem('Remove Backup Schedule', 'removeBackupTriggers')
    .addToUi();
}
*/