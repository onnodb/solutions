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
 
// GLOBAL VARS
var HRS_START = 4;
var HRS_END = 8;
var HRLY_PAY_COLUMN = 9;
var APPROVAL_COLUMN = 11;
var EMAIL_COLUMN = 3;

/**
* Creates the menu item "Run Functions" so user can run scripts on drop-down.
*/
function onOpen(e) {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Timesheets')
        .addItem('Calculate Pay', 'calculatePay')
        .addItem('Add Approval Column', 'addApprovalColumn')
        .addItem('Send Emails', 'sendEmails')
        .addToUi();
}

/**
* Runs the script "Send Emails". Sends e mails to every employee notifying 
* them whether or not their timesheet has been approved.
* Checks + updates "NOTIFIED" status so employers do not send e mails more than once.
*/
function sendEmails() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    // lastCol here is the NOTIFIED column
  
    // goes through the whole email column
    for (i =2; i <= lastRow; i++) { 
      // don't notify twice
      var notifiedStatus = sheet.getRange(i, lastCol).getValue();
      if (notifiedStatus == 'NOTIFIED') {
        continue;
      }
           
      var range = sheet.getRange(i, 3);
      var email = range.getValue();
      var dropdownValue = sheet.getRange(i,APPROVAL_COLUMN);    
      
      // approval email
      if (dropdownValue.getValue() == 'APPROVED') {
        var message = 'Your timesheet has been approved';
        var subject = 'TimeSheet Approval';
      } else if (dropdownValue.getValue() == 'NOT APPROVED') {
        var message = 'NOT APPROVED';
        var subject = 'TimeSheet not Approved';
      } else if (dropdownValue.getValue() == 'IN PROGRESS') {
        continue;
      } 
      
      MailApp.sendEmail(email, subject, message);
      var notifiedUpdate = sheet.getRange(i, lastCol).setValue('NOTIFIED');
    }
}


/**
* Runs the "Add Approval Column" script, adding an approval column at the end of the sheet, 
* made entirely of checkboxes.
* Adds "NOTIFIED STATUS" column indicating whether or not an employee has yet been e mailed.
*/
function addApprovalColumn() { 
  // adds new column
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  sheet.insertColumnAfter(lastCol);
  
  // sets new column title
  var cell = sheet.getRange(1, lastCol+1);
  cell.setBackground('ffe3ee').setValue('APPROVAL');
  
  // make sure approval column is all drop-down menu
  for (i = 2; i <= lastRow; i++) {
    var dropdownBox = sheet.getRange(i,lastCol+1);
    var dropdownValues = ['APPROVED', 'NOT APPROVED', 'IN PROGRESS'];
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(dropdownValues).build();
    dropdownBox.setDataValidation(rule);
    dropdownBox.setValue('IN PROGRESS');
  }
  
  // add the notified column
  sheet.insertColumnAfter(APPROVAL_COLUMN); //global
  var cell = sheet.getRange(1, APPROVAL_COLUMN + 1);
  cell.setBackground('ffe3ee').setValue('NOTIFIED STATUS'); 
  
  // sets notified status to NOT NOTIFIED
  for (var i = 2; i <= lastRow; i++) {
    var notifiedCell = sheet.getRange(i, APPROVAL_COLUMN + 1);
    var notifiedValues = ['NOTIFIED', 'NOT NOTIFIED'];
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(notifiedValues).build();
    notifiedCell.setDataValidation(rule);
    notifiedCell.setValue('NOT NOTIFIED');
  }
}

/**
* Runs the "Calculate Pay" script, creating a new column with each employee's 
* pay based on their reported hours and hourly wage.
*/
function calculatePay() {
  // adds new column
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  sheet.insertColumnAfter(lastCol);
  
  // sets new column title
  var cell = sheet.getRange(1, lastCol+1);
  cell.setValue('WEEKLY PAY');
  
  // goes through every employee
  for (var i = 2; i <= lastRow; i++) {
    var totalHours = 0;
    var numHrs = 0;
    
    // move across the hour columns
    for (var j = HRS_START; j <= HRS_END; j++) {
      var numHrsCell = sheet.getRange(i, j);
      numHrs = numHrs + numHrsCell.getValue();
    }
    
    // calulates pay
    var hrlyRateCell = sheet.getRange(i, HRLY_PAY_COLUMN);
    var hrlyRate = hrlyRateCell.getValue();
    var weeklyPay = numHrs * hrlyRate;
    var setPay = sheet.getRange(i, lastCol+1);
    setPay.setValue(weeklyPay);
  }
}
