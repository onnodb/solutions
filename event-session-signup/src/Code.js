// Copyright 2019 Google LLC
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     https://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

/**
 * A special function that inserts a custom menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Conference')
      .addItem('Set up conference', 'setUpConference')
      .addItem('Open registration form', 'openForm_')
      .addItem('Open calendar', 'openCalendar_')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Advanced')
        .addItem('Reset conference', 'resetConference_')
      )
      .addToUi();
}

/**
 * A set-up function that uses the conference data in the spreadsheet to create
 * Google Calendar events, a Google Form, and a trigger that allows the script
 * to react to form responses.
 * This function can also be run to update an existing calendar/form with 
 * modified sessions.
 */
function setUpConference() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Conference Setup');
  var range = sheet.getDataRange();
  var values = range.getValues();
  setUpCalendar_(values, range);
  setUpForm_(ss, values);
}

/** 
 * Reset our "memory", so that next time we'll create a new calendar and registration
 * form.
 */
function resetConference_() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Are you sure you want to reset the conference registration? ' +
    'This will unlink any existing calendar and form, and cause a new ' +
    'calendar and form to be created next time you use "Set up conference".',
    ui.ButtonSet.OK_CANCEL);
  if (response == ui.Button.CANCEL) {
    return;
  }

  PropertiesService.getScriptProperties().deleteProperty('calendarId');
  PropertiesService.getScriptProperties().deleteProperty('formId');
  
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    Logger.log("Deleting project trigger...")
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

/**
 * Creates a Google Calendar with events for each conference session in the
 * spreadsheet, then writes the event IDs to the spreadsheet for future use.
 * @param {Array<string[]>} values Cell values for the spreadsheet range.
 * @param {Range} range A spreadsheet range that contains conference data.
 */
function setUpCalendar_(values, range) {
  var cal = null;
  var calId = PropertiesService.getScriptProperties().getProperty('calendarId');
  if ((calId === null) || (CalendarApp.getCalendarById(calId) === null)) {
    // Need to create new calendar.
    cal = CalendarApp.createCalendar('LUMICKS Academy Calendar');
    calId = cal.getId();
    // Store the ID for the Calendar, which is needed to retrieve events by ID.
    PropertiesService.getScriptProperties().setProperty('calendarId', calId);
  }
  else {
    // Update existing calendar.
    cal = CalendarApp.getCalendarById(calId);
  }

  // NOTE: We do NOT delete any events from the calendar. If a session gets deleted,
  // the calendar has to be updated manually. (This is intended.)
  for (var i = 1; i < values.length; i++) {  // start at 1 to skip the sheet's header row
    var session = values[i];
    var title = session[0];
    var start = joinDateAndTime_(session[1], session[2]);
    var end = joinDateAndTime_(session[1], session[3]);
    var location = session[4];
    var event = cal.getEventById(session[5]);
    if (event == null) {
      // Create new event on calendar. We use this circumvent approach in order to
      // be able to generate a Hangouts link with it.
      // See https://stackoverflow.com/a/65216722/1037
      var eventPayload = {
        "summary": title, 
        "start": {"dateTime": start.toISOString()},
        "end": {"dateTime": end.toISOString()},
        "location": location,
        "conferenceData": {
          "createRequest": {
            "conferenceSolutionKey": {
              "type": "hangoutsMeet",
            }, 
            "requestId": Utilities.getUuid(),
          }
        }
      };
      var eventReturn = Calendar.Events.insert(
        eventPayload, 
        calId, 
        {
          "conferenceDataVersion": 1,
          "sendUpdates": "all",
        });
      event = cal.getEventById(eventReturn.id);
    } else {
      // Update event settings to be in sync with sheet data.
      event.setTitle(title);
      event.setTime(start, end);
      event.setLocation(location);
    }

    event.setGuestsCanSeeGuests(true);

    // Store event ID for future updates.
    session[5] = event.getId();
  }
  range.setValues(values);
}

/**
 * Creates a single Date object from separate date and time cells.
 *
 * @param {Date} date A Date object from which to extract the date.
 * @param {Date} time A Date object from which to extract the time.
 * @return {Date} A Date object representing the combined date and time.
 */
function joinDateAndTime_(date, time) {
  date = new Date(date);
  date.setHours(time.getHours());
  date.setMinutes(time.getMinutes());
  return date;
}

/**
 * Creates a Google Form that allows respondents to select which conference
 * sessions they would like to attend, grouped by date and start time in the
 * caller's time zone.
 *
 * @param {Spreadsheet} ss The spreadsheet that contains the conference data.
 * @param {Array<String[]>} values Cell values for the spreadsheet range.
 */
function setUpForm_(ss, values) {
  // Group the sessions by date and time so that they can be passed to the form.
  var schedule = {};
  // Start at 1 to skip the header row.
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var day = session[1].toLocaleDateString('en-GB');
    var time = session[2].toLocaleTimeString('en-GB');
    if (!schedule[day]) {
      schedule[day] = {};
    }
    if (!schedule[day][time]) {
      schedule[day][time] = [];
    }
    schedule[day][time].push(session[0]);
  }

  // Create the form and add a multiple-choice question for each timeslot.
  var form = null;
  var formId = PropertiesService.getScriptProperties().getProperty('formId');
  try {
    if (formId !== null) {
      form = FormApp.openById(formId);
    }
  } catch(e) {
    form = null;
  }
  if (form === null) {
    // Create new form.
    form = FormApp.create("LUMICKS Academy Registration");
    formId = form.getId();
    PropertiesService.getScriptProperties().setProperty('formId', formId);

    form.addTextItem().setTitle('Name').setRequired(true);
    form.addTextItem().setTitle('Email').setRequired(true);

    // Run our `onFormSubmit` code whenever the form gets submitted.
    ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit()
        .create();
  }

  // Create a new empty sheet for receiving registrations. (So each time we run, we create
  // a fresh sheet, keeping old ones around for reference.)
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // (Re-)create questions for all sessions.
  form.getItems(FormApp.ItemType.MULTIPLE_CHOICE).forEach(function(item) {
    form.deleteItem(item);
  });
  form.getItems(FormApp.ItemType.SECTION_HEADER).forEach(function(item) {
    form.deleteItem(item);
  });
  Object.keys(schedule).forEach(function(day) {
    var header = form.addSectionHeaderItem().setTitle('Sessions for ' + day);
    Object.keys(schedule[day]).forEach(function(time) {
      var item = form.addMultipleChoiceItem().setTitle(time + ' ' + day)
          .setChoiceValues(schedule[day][time]);
    });
  });
}

/**
 * A trigger-driven function that sends out calendar invitations and a
 * personalized Google Docs itinerary after a user responds to the form.
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {
  var user = {name: e.namedValues['Name'][0], email: e.namedValues['Email'][0]};
  Logger.log(`OnFormSubmit: ${user.name} -- ${user.email}`);
  Logger.log(e.namedValues);

  // Grab the session data again so that we can match it to the user's choices.
  var response = [];
  var values = SpreadsheetApp.getActive().getSheetByName('Conference Setup')
      .getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var title = session[0];
    var day = session[1].toLocaleDateString('en-GB');
    var time = session[2].toLocaleTimeString('en-GB');
    var timeslot = time + ' ' + day;

    // For every selection in the response, find the matching timeslot and title
    // in the spreadsheet and add the session data to the response array.
    if (e.namedValues[timeslot] && (e.namedValues[timeslot].indexOf(title) !== -1)) {
      Logger.log(`${user.name} registered for ${session}`)
      response.push(session);
    }
  }
  sendInvites_(user, response);
}

/**
 * Add the user as a guest for every session he or she selected.
 * @param {object} user An object that contains the user's name and email.
 * @param {Array<String[]>} response An array of data for the user's session choices.
 */
function sendInvites_(user, response) {
  var calId = ScriptProperties.getProperty('calendarId');
  for (var i = 0; i < response.length; i++) {
    addGuestAndSendEmail_(calId, response[i][5], user.email);
  }
}

function addGuestAndSendEmail_(calendarId, eventId, newGuest) {
  Logger.log(`addGuestAndSendEmail(${calendarId}, ${eventId}, ${newGuest})`);

  // In the Calendar Advanced API, we need to strip the "@google.com" suffix from the
  // end of the Event ID, or it won't work [0].
  // [0]: https://stackoverflow.com/a/55509409
  var calEventId = eventId.replace('@google.com', '');
  var event = Calendar.Events.get(calendarId, calEventId);
  Logger.log(event);
  var attendees = 'attendees' in event ? event.attendees : [];
  attendees.push({email: newGuest});

  var resource = { attendees: attendees };
  var args = { sendUpdates: "all" };

  Calendar.Events.patch(resource, calendarId, calEventId, args);
}

function openUrl_(url) {
  var js = `<script>window.open('${url}', '_blank'); google.script.host.close();</script>`;
  var html = HtmlService.createHtmlOutput(js)
    .setHeight(10)
    .setWidth(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Now loading.')
}

function openForm_() {
  var formId = PropertiesService.getScriptProperties().getProperty('formId');
  if (formId === null) {
    SpreadsheetApp.getUi().alert('No registration form has been set up yet.');
    return;
  }

  var form = FormApp.openById(formId);
  const formUrl = form.getEditUrl();
  openUrl_(formUrl);
}

function openCalendar_() {
  var calId = PropertiesService.getScriptProperties().getProperty('calendarId');
  if (calId === null) {
    SpreadsheetApp.getUi().alert('No calendar has been set up yet.');
    return;
  }

  const calUrl = `https://calendar.google.com/calendar/?cid=${calId}`;
  openUrl_(calUrl);
}
