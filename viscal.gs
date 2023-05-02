// Google App Script for Google Sheets - Bi-directional Sync with Google Calendar
//Copyright (c) 2023 - Robert McLellan

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.

// Global variables for Options sheet
const OPTIONS_SHEET_NAME = 'Options';
const VARIABLES = [
  {name: 'Calendar ID', defaultValue: 'primary'},
  {name: 'Start Date', defaultValue: '2023-01-01'},
  {name: 'End Date', defaultValue: '2023-12-31'},
  {name: 'Search Query', defaultValue: ''},
  {name: 'Filter Variable', defaultValue: ''},
  {name: 'Sync Fields', defaultValue: 'calendarId,eventId,title,start,end,location,description,busyStatus,ColorId'},
];

// Global variables for Google Calendar sheet
const GOOGLE_CALENDAR_SHEET_NAME = 'Google Calendar';
const ACTION_COLUMN = 1;
const LAST_ACTION_COLUMN = 2;
const CALENDAR_ID_COLUMN = 3;
const EVENT_ID_COLUMN = 4;

// Global variables for Deleted Events sheet
const DELETED_EVENTS_SHEET_NAME = 'Deleted Events';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Google Calendar Sync')
    .addItem('Refresh', 'refreshSheet')
    .addItem('Synchronize', 'synchronizeSheet')
    .addToUi();

  setActionColumnValidation();
}

function initializeOptionsSheet() {
  let sheet = getOrCreateSheet(OPTIONS_SHEET_NAME);
  let options = sheet.getRange(1, 1, VARIABLES.length, 2);

  for (let i = 0; i < VARIABLES.length; i++) {
    options.getCell(i + 1, 1).setValue(VARIABLES[i].name);
    if (!options.getCell(i + 1, 2).getValue()) {
      options.getCell(i + 1, 2).setValue(VARIABLES[i].defaultValue);
    }
  }
}

function getOrCreateSheet(sheetName) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  }
  return sheet;
}

function refreshSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Warning', 'All data on the ' + GOOGLE_CALENDAR_SHEET_NAME + ' sheet will be wiped out. Do you want to continue?', ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    const sheet = getOrCreateSheet(GOOGLE_CALENDAR_SHEET_NAME);
    sheet.clear();
    initializeOptionsSheet();

    const options = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OPTIONS_SHEET_NAME);
    let startDate = options.getRange(2, 2).getValue();
    let endDate = options.getRange(3, 2).getValue();
    const searchQuery = options.getRange(4, 2).getValue();
    const filterVariable = options.getRange(5, 2).getValue();
    const syncFields = options.getRange(6, 2).getValue().split(',');
    const calendarId = options.getRange(1, 2).getValue(); // Read the calendar ID from the Options sheet

    startDate = new Date(startDate);
    endDate = new Date(endDate);

    if (isNaN(startDate) || isNaN(endDate)) {
      ui.alert('Error', 'Invalid date values. Please check the start and end dates in the Options sheet.', ui.ButtonSet.OK);
      return;
    }

    const events = Calendar.Events.list(calendarId, {
      timeMin: startDate.toISOString(),
      timeMax: endDate.toISOString(),
      q: searchQuery,
      singleEvents: true,
      orderBy: 'startTime',
    }).items;

    sheet.getRange(1, ACTION_COLUMN).setValue('ACTION');
    sheet.getRange(1, LAST_ACTION_COLUMN).setValue('LAST ACTION');
    sheet.getRange(1, CALENDAR_ID_COLUMN).setValue('CALENDAR');
    for (let i = 0; i < syncFields.length; i++) {
      sheet.getRange(1, CALENDAR_ID_COLUMN + i).setValue(syncFields[i]);
    }

    let row = 2;
    events.forEach(event => {
      if (!filterVariable || eventMatchesFilter(event, filterVariable)) {
        sheet.getRange(row, ACTION_COLUMN).setValue('SYNCED');
        sheet.getRange(row, LAST_ACTION_COLUMN).setValue('');
        for (let i = 0; i < syncFields.length; i++) {
          const fieldValue = getEventFieldValue(event, syncFields[i]);
          sheet.getRange(row, CALENDAR_ID_COLUMN + i).setValue(fieldValue);
        }
        row++;
      }
    });

    sheet.hideColumns(EVENT_ID_COLUMN);
    sheet.hideColumns(CALENDAR_ID_COLUMN);
    sheet.protect().setUnprotectedRanges([sheet.getRange(2, ACTION_COLUMN, sheet.getLastRow() - 1, 1), sheet.getRange(2, LAST_ACTION_COLUMN, sheet.getLastRow() - 1, 1)]);
  }
}



function eventMatchesFilter(event, filterVariable) {
  // Add logic to check if the event matches the filter criteria. Return true if it matches, false otherwise.
  // For example, if filtering by the event location, check if the event location matches the filter value.
  return true;
}

function getEventFieldValue(event, field, calendarId) {
  switch (field) {
    case 'eventId':
      return event.getId();
    case 'title':
      return event.summary;
    case 'location':
      return event.location;
    case 'description':
      return event.description;
    case 'start':
      return event.start.dateTime || event.start.date;
    case 'end':
      return event.end.dateTime || event.end.date;
    case 'busyStatus':
      return event.transparency === 'transparent' ? 'FREE' : 'BUSY';
    case 'calendarId':
      return calendarId; // Return the calendarId passed as a parameter
    case 'colorId':
      return event.colorId;
    default:
      return '';
  }
}


function synchronizeSheet() {
  //const ui = SpreadsheetApp.getUi();
  const calendarSheet = getOrCreateSheet(GOOGLE_CALENDAR_SHEET_NAME);
  const optionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OPTIONS_SHEET_NAME);
  const calendarId = optionsSheet.getRange(1, 2).getValue(); // Read the calendar ID from the Options sheet
  const syncFields = optionsSheet.getRange(6, 2).getValue().split(',');
  const numRows = calendarSheet.getLastRow() - 1; // Excluding the header row
  const dataRange = calendarSheet.getRange(2, 1, numRows, calendarSheet.getLastColumn());
  const data = dataRange.getValues();
  // Unlock the "Last Action" column
  //const lastActionRange = calendarSheet.getRange(2, LAST_ACTION_COLUMN, numRows);
  //const protection = lastActionRange.protect();
  //protection.remove();
  
  // Unlock LAST ACTION field
  calendarSheet.getRange(2, LAST_ACTION_COLUMN, numRows).setNumberFormat('@');

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const action = row[ACTION_COLUMN - 1];
    const eventId = row[EVENT_ID_COLUMN - 1];
    
    if (!isValidActionValue(action)) {
      calendarSheet.getRange(2 + i, LAST_ACTION_COLUMN).setValue('INVALID ACTION VALUE');
      continue;
    }

    const eventFields = {};
    for (let j = 0; j < syncFields.length; j++) {
      const field = syncFields[j];
      if (field !== 'eventId') {
        eventFields[field] = row[CALENDAR_ID_COLUMN + j - 1];
      }
    }

    if (!eventFields.title || !eventFields.start || !eventFields.end) {
      calendarSheet.getRange(2 + i, LAST_ACTION_COLUMN).setValue('Title, start, and end required');
      continue;
    }

    let eventExists = false;
    if (eventId) {
      try {
        Calendar.Events.get(calendarId, eventId);
        eventExists = true;
      } catch (e) {
        eventExists = false;
      }
    }

    if (action === 'DELETE') {
      if (!eventId) {
        calendarSheet.getRange(2 + i, LAST_ACTION_COLUMN).setValue('NO EVENT ID');
      } else if (!eventExists) {
        calendarSheet.getRange(2 + i, LAST_ACTION_COLUMN).setValue('EVENT DOES NOT EXIST');
      } else {
        storeDeletedEvent(eventId, calendarId, syncFields);
        Calendar.Events.remove(calendarId, eventId);
        calendarSheet.getRange(2 + i, ACTION_COLUMN).setValue('SYNCED');
        calendarSheet.getRange(i + 2, LAST_ACTION_COLUMN).setValue('EVENT DELETED');
        setRowColorByAction(calendarSheet, i + 1, action);
      }
    } else if (action === 'ADD' || (action === 'UPDATE' && !eventExists)) {
      if (!eventId) {
        const newEvent = createEventFromFields(eventFields);
        const createdEvent = Calendar.Events.insert(newEvent, calendarId);
        calendarSheet.getRange(2 + i, EVENT_ID_COLUMN).setValue(createdEvent.getId());
        calendarSheet.getRange(2 + i, ACTION_COLUMN).setValue('SYNCED');
        calendarSheet.getRange(2 + i, LAST_ACTION_COLUMN).setValue('CREATED');
        syncRowWithEvent(calendarSheet, 1 + i, createdEvent);
        setRowColorByAction(calendarSheet, i + 1, action);
      } else if (eventExists) {
        calendarSheet.getRange(2 + i, ACTION_COLUMN).setValue('UPDATE');
        calendarSheet.getRange(2 + i, LAST_ACTION_COLUMN).setValue('EVENT ID EXISTS');
        setRowColorByAction(calendarSheet, i + 1, action);
      } else {
        const newEvent = createEventFromFields(eventFields);
        const createdEvent = Calendar.Events.insert(newEvent, calendarId);
        calendarSheet.getRange(2 + i, EVENT_ID_COLUMN).setValue(createdEvent.getId());
        calendarSheet.getRange(2 + i, ACTION_COLUMN).setValue('SYNCED');
        calendarSheet.getRange(2 + i, LAST_ACTION_COLUMN).setValue('CREATED');
        syncRowWithEvent(calendarSheet, 1 + i, createdEvent);
        setRowColorByAction(calendarSheet, i + 1, action);
      }
    } else if (action === 'UPDATE' && eventExists) {
      const updatedEvent = updateEventFromFields(eventId, calendarId, eventFields);
      calendarSheet.getRange(2 + i, ACTION_COLUMN).setValue('SYNCED');
      calendarSheet.getRange(2 + i, LAST_ACTION_COLUMN).setValue('UPDATED');
      syncRowWithEvent(calendarSheet, 1 + i, updatedEvent);
      setRowColorByAction(calendarSheet, i + 1, action);
    }
  }
  function syncRowWithEvent(sheet, rowIndex, event) {
    for (let j = 0; j < syncFields.length; j++) {
      const field = syncFields[j];
      if (field !== 'eventId') {
        const fieldValue = getEventFieldValue(event, field);
        sheet.getRange(rowIndex + 1, CALENDAR_ID_COLUMN + j).setValue(fieldValue);
      }
    }
  }
  // Lock the "Last Action" column
  // Not currently Working
  //const lastActionRange = calendarSheet.getRange(2, LAST_ACTION_COLUMN, numRows);
  //const protection = lastActionRange.protect();
  //protection.setDescription('Last Action column locked');
  //protection.setWarningOnly(false);
}

function createEventFromFields(fields) {
  const event = {
    summary: fields.title,
    location: fields.location,
    description: fields.description,
    start: {
      dateTime: new Date(fields.start).toISOString(),
      timeZone: CalendarApp.getDefaultCalendar().getTimeZone(),
    },
    end: {
      dateTime: new Date(fields.end).toISOString(),
      timeZone: CalendarApp.getDefaultCalendar().getTimeZone(),
    },
    transparency: fields.busyStatus === 'FREE' ? 'transparent' : 'opaque',
    colorId: fields.colorId,
  };
  return event;
}

function updateEventFromFields(eventId, calendarId, fields) {
  const event = Calendar.Events.get(calendarId, eventId);

  event.summary = fields.title;
  event.location = fields.location;
  event.description = fields.description;
  event.start.dateTime = new Date(fields.start).toISOString();
  event.end.dateTime = new Date(fields.end).toISOString();
  event.transparency = fields.busyStatus === 'FREE' ? 'transparent' : 'opaque';
  event.colorId = fields.colorId;

  const updatedEvent = Calendar.Events.update(event, calendarId, eventId);
  return updatedEvent;
}


function storeDeletedEvent(eventId, calendarId, syncFields) {
  const event = Calendar.Events.get(calendarId, eventId);
  const deletedEventsSheet = getOrCreateSheet(DELETED_EVENTS_SHEET_NAME);

  if (deletedEventsSheet.getLastRow() === 0) {
    for (let i = 0; i < syncFields.length; i++) {
      deletedEventsSheet.getRange(1, i + 1).setValue(syncFields[i]);
    }
  }

  const newRow = deletedEventsSheet.getLastRow() + 1;
  for (let i = 0; i < syncFields.length; i++) {
    const fieldValue = getEventFieldValue(event, syncFields[i], calendarId); // Pass calendarId
    deletedEventsSheet.getRange(newRow, i + 1).setValue(fieldValue);
  }
}

function isValidActionValue(action) {
  const validActions = ['UPDATE', 'DELETE', 'ADD', 'SYNCED'];
  return validActions.includes(action);
}

function setActionColumnValidation() {
  const calendarSheet = getOrCreateSheet(GOOGLE_CALENDAR_SHEET_NAME);
  const numRows = calendarSheet.getMaxRows() - 1; // Excluding the header row

  // Create a data validation rule
  const allowedValues = ['UPDATE', 'DELETE', 'ADD', 'SYNCED'];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(allowedValues, true)
    .setAllowInvalid(false)
    .setHelpText('Only UPDATE, DELETE, ADD, and SYNCED values are allowed.')
    .build();

  // Apply the rule to the ACTION column
  const actionValidationRange = calendarSheet.getRange(2, ACTION_COLUMN, numRows);
  actionValidationRange.setDataValidation(rule);
}

function setRowColorByAction(sheet, rowIndex, action) {
  const rowRange = sheet.getRange(rowIndex + 1, 1, 1, sheet.getLastColumn());
  let color;

  switch (action) {
    case 'SYNCED':
      color = '#90ee90'; // Light green
      break;
    case 'ADD':
      color = '#add8e6'; // Light blue
      break;
    case 'UPDATE':
      color = '#ffff99'; // Light yellow
      break;
    case 'DELETE':
      color = '#ffcccc'; // Light red
      break;
    default:
      color = null;
  }

  rowRange.setBackground(color);
}
