function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カレンダー同期')
    .addItem('今すぐ同期', 'backupCalendarEvents')
    .addToUi();
}

function backupCalendarEvents() {
  // Get the target calendar address from script properties
  const scriptProperties = PropertiesService.getScriptProperties();
  const calendarId = scriptProperties.getProperty('CALENDAR_ID');
  
  if (!calendarId) {
    throw new Error('Calendar ID not found in Script Properties. Please set CALENDAR_ID property.');
  }

  // Get the target calendar
  const calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    throw new Error('Calendar not found with the provided ID.');
  }

  // Get the target spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // Clear existing content
  sheet.clear();

  // Set headers
  const headers = ['タイトル', '開始日時', '終了日時', '場所', '説明', '参加者', 'イベントID'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Get today's date at start of day
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Get events from today onwards
  const events = calendar.getEvents(today, new Date(2099, 11, 31));

  // Prepare data for spreadsheet
  const eventData = events.map(event => [
    event.getTitle(),
    event.getStartTime(),
    event.getEndTime(),
    event.getLocation() || '',
    event.getDescription() || '',
    event.getGuestList().map(guest => guest.getEmail()).join(', '),
    event.getId()
  ]);

  // Write data to spreadsheet if there are events
  if (eventData.length > 0) {
    sheet.getRange(2, 1, eventData.length, headers.length).setValues(eventData);
  }

  // Format the sheet
  sheet.autoResizeColumns(1, headers.length);
  
  // Format date columns
  const dateRange = sheet.getRange(2, 2, eventData.length, 2);
  dateRange.setNumberFormat('yyyy/MM/dd HH:mm');
}

/**
 * Sets up the script properties with the calendar ID
 * @param {string} calendarId - The calendar ID to backup
 */
function setupCalendarId(calendarId) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('CALENDAR_ID', calendarId);
}

/**
 * Sets up daily trigger at 6:00 AM
 */
function setupDailyTrigger() {
  // Delete existing triggers first to avoid duplicates
  deleteTriggers();
  
  // Create a new trigger to run at 6:00 AM every day
  ScriptApp.newTrigger('backupCalendarEvents')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
}

/**
 * Deletes all existing triggers
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
}
