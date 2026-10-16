/**
 * Serves the HTML frontend when visiting the web app URL.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('EA Task Command Center')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper to get the target sheet tab.
 */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('Tasks') || ss.getSheets()[0];
}

/**
 * Formats Date objects safely into YYYY-MM-DD strings.
 */
function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val);
}

/**
 * Reads all tasks from the Sheet and returns them.
 */
function getTasks() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const tasks = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; 

    tasks.push({
      rowIndex: i + 1,
      id: String(row[0]),
      name: String(row[1] || ''),
      description: String(row[2] || ''),
      owner: String(row[3] || ''),
      priority: String(row[4] || 'Medium'),
      dueDate: formatDate(row[5]),
      status: String(row[6] || 'To Do'),
      createdDate: formatDate(row[7]),
      completedDate: formatDate(row[8])
    });
  }
  return tasks;
}

/**
 * Appends a new task row into the Google Sheet.
 */
function addTask(task) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  
  const nextNum = data.length < 2 ? 1 : data.length;
  const taskId = 'TSK-' + String(nextNum).padStart(3, '0');
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  const newRow = [
    taskId, task.name, task.description, task.owner, 
    task.priority, task.dueDate, task.status || 'To Do', today, ''
  ];
  
  sheet.appendRow(newRow);
  return { success: true, taskId: taskId };
}

/**
 * Updates the Status and Completed_Date of an existing task.
 */
function updateTaskStatus(taskId, newStatus) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(taskId).trim()) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, 7).setValue(newStatus); 
      
      if (newStatus === 'Done') {
        const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        sheet.getRange(rowNum, 9).setValue(today); 
      } else {
        sheet.getRange(rowNum, 9).clearContent();
      }
      return { success: true };
    }
  }
  throw new Error('Task ID ' + taskId + ' not found.');
}
