//TODO - If there is a time condition user change and closes the app then it fails - we need to have a priodic "Syncing between backup and helper somehow"
//TODO - ranges protection can be multiple cells and not one by one for each row there is 7 can be 1 or 2. A:C E:K for example

/* Deploy Instructions:
// 1.Duplicate Master-> Helper
// 2.Protect Helper
// 3.Protect Master first row
// 4.Protect Master (With except)
*/

var masterSheetName = "ראשי" // sheet where the cells are protected from updates
var helperSheetName = "גיבוי" // sheet where the values are copied for later checking
var taskCell = 'D';
var creatorCol = 'E';
var creatorEmailCol = 'F';
var fillerCol = 'G';
var fillerEmailCol = 'H';
var descriptionCol = 'I';
var creatorCommentsCol = 'J';
var fillerPhoneCol = 'M';
var taskStatusCol = 'N';

var missionStringSize = 15; //Extract only part of the mission for the description

var creatorField = 6; //"F"
var fillerField = 8; //"H"
var ui = SpreadsheetApp.getUi();

/**
 * Test function for onEdit. Passes an event object to simulate an edit to
 * a cell in a spreadsheet.
 * Check for updates: https://stackoverflow.com/a/16089067/1677912
 */
function test_onEdit() {
  onEdit({
    user : Session.getActiveUser().getEmail(),
    source : SpreadsheetApp.getActiveSpreadsheet(),
    range : SpreadsheetApp.getActiveSpreadsheet().getActiveCell(),
    value : SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue(),
    authMode : "LIMITED"
  });
}

function onEdit(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getActiveSheet();
  if (masterSheet.getName() != masterSheetName) return;
  
  //Check if its just open and we should restrict it
  if(!checkIfHasAccess(e)) {
    return;
  }
  
  var columnsToGiveAccess = [];
  if(e.range.getColumn() === fillerField) {
    columnsToGiveAccess = [fillerCol, fillerEmailCol, descriptionCol, fillerPhoneCol, taskStatusCol];
  }else if(e.range.getColumn() === creatorField) {
    columnsToGiveAccess = [fillerCol, fillerEmailCol, descriptionCol, fillerPhoneCol, taskStatusCol, creatorCol, creatorEmailCol, creatorCommentsCol];
  }

  // Only grant permissions if set this array
  if(columnsToGiveAccess.length > 0) {
    if(grantPermission(e, columnsToGiveAccess)) {
      showAlert("Access granted","User can now edit those columns " + columnsToGiveAccess.toString());
    }
  }
}

  // This function prevents cells from being updated. When a user edits a cell on the master sheet,
  // it is checked against the same cell on a helper sheet. If the value on the helper sheet is
  // empty, the new value is stored on both sheets.
  // If the value on the helper sheet is not empty, it is copied to the cell on the master sheet,
  // effectively undoing the change.
  // The exception is that the first few rows and the first few columns can be left free to edit by
  // changing the firstDataRow and firstDataColumn variables below to greater than 1.
  // To create the helper sheet, go to the master sheet and click the arrow in the sheet's tab at
  // the tab bar at the bottom of the browser window and choose Duplicate, then rename the new sheet
  // to Helper.
  // To change a value that was entered previously, empty the corresponding cell on the helper sheet,
  // then edit the cell on the master sheet.
  // You can hide the helper sheet by clicking the arrow in the sheet's tab at the tab bar at the
  // bottom of the browser window and choosing Hide Sheet from the pop-up menu, and when necessary,
  // unhide it by choosing View > Hidden sheets > Helper.
function checkIfHasAccess(e) {
  var row = e.range.getRow();

  var creatorEmailContent = getCellContent(creatorEmailCol + row);
  var fillerEmailContent = getCellContent(fillerEmailCol + row);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getActiveSheet();
  var masterRange = masterSheet.getActiveRange();
  
  Logger.log(masterRange.getHeight());
  if(masterRange.getHeight() > 1) {
    undoActiveRange();
    showAlert("Can\'t edit","Please edit one line at a time");
    return false;
  }
  
  var helperSheet = ss.getSheetByName(helperSheetName);
  var helperRange = helperSheet.getRange(masterRange.getA1Notation());
  var newValue = masterRange.getValues();
  var oldValue = helperRange.getValues();
  Logger.log("newValue " + newValue);
  Logger.log("oldValue " + oldValue);
  
  if(creatorEmailContent == "" && fillerEmailContent == "") {
    undoActiveRange();
    showAlert("Can\'t edit","This task doesn\'t have creator email and filler yet");
    return false;
  }
  
  helperRange.setValues(newValue);
  
  return true;
}

function undoActiveRange() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getActiveSheet();
  var masterRange = masterSheet.getActiveRange();
  var helperSheet = ss.getSheetByName(helperSheetName);
  var helperRange = helperSheet.getRange(masterRange.getA1Notation());
  var newValue = masterRange.getValues();
  var oldValue = helperRange.getValues();
  masterRange.setValues(oldValue);
}

function grantPermission(e, columnsToGiveAccess) {
  var newValue = e.range.getValue(); //Better this then e.value which on copy return other results
  Logger.log("Value:" + e.value);
  Logger.log("Actual Value:" + newValue);
  Logger.log("Old Value:" + e.oldValue);
  
/*  if(typeof(e.value) !== "string") {
    Logger.log("Not string trying:");
    Logger.log(e.value);
   return; //TODO - what happenes when multiple cells, deleting a cell, Pasting a cell
  }*/
  
  if(!validateEmail(newValue)) {
    e.range.setValue(e.oldValue || '');
    showAlert("","Invalid email " + newValue + " must be burner@somecompany.something");
    return false;
  }
  
  if(showYesNoCancel("Please Check if the email is correct?",newValue) != ui.Button.YES) {
    e.range.setValue(e.oldValue || '');
    return false;
  }
  
  var fillerEmailCol = newValue;
  
  for(var i=0;i<columnsToGiveAccess.length;i++) {
    var range = getRange(columnsToGiveAccess[i] + e.range.getRow());
    protectRange(range, fillerEmailCol);
  }
  
  // Value is inserted for the first time
  if(newValue && !e.oldValue) {
      Logger.log("Value is inserted for the first time");
  }
  return true;
}

function getCellContent(cellName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var range = sheet.getRange(cellName);
  return range.getCell(1,1).getValue();
}

// Can accept "A2:B3" for example OR "B5"
function getRange(rangeName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  return sheet.getRange(rangeName); 
}

function protectRange(range, newEditorEmail) {
  // Get range task
  var taskContent = getCellContent(taskCell + range.getRow());
  
  // TODO - Should not protect if already protected
  var protection = range.protect().setDescription(taskContent.substr(0, missionStringSize) + ' ' + newEditorEmail);

 // Ensure the current user is an editor before removing others. Otherwise, if the user's edit
 // permission comes from a group, the script will throw an exception upon removing the group.
 //var me = Session.getEffectiveUser();
 //protection.addEditor(me);
 protection.addEditor(newEditorEmail);
   
 //protection.removeEditors(protection.getEditors());
 //if (protection.canDomainEdit()) {
   //protection.setDomainEdit(false);
 //}
 
}
