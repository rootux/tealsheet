//TODO - If there is a time condition user change and closes the app then it fails - we need to have a priodic "Syncing between backup and helper somehow"
//TODO - ranges protection can be multiple cells and not one by one for each row there is 7 can be 1 or 2. A:C E:K for example
//TODO - Anonymouse user can edit what they want

/* Deploy Instructions:
// 1.Duplicate Master-> Helper
// 2.Protect Helper
// 3.Hide Helper
// 4.Protect Master first row
// 5.Protect Master (With except)
*/

var masterSheetName = "ראשי" // sheet where the cells are protected from updates
var helperSheetName = "גיבוי" // sheet where the values are copied for later checking

var mainProtectionName = "Main";
var labelsCol = 'B';
var taskCell = 'C';
var creatorCol = 'D';
var creatorEmailCol = 'E';
var fillerCol = 'F';
var fillerEmailCol = 'G';
var descriptionCol = 'H';
var creatorCommentsCol = 'I';
var fillerPhoneCol = 'L';
var taskStatusCol = 'M';
var extraCol = 'N';
var extraCol2 = 'O';
var communityCommentsCol = 'J';
var contactsCol = 'K';


// Translations
var isEmailCorrect = "אנא בדקי שנית האם זה האימייל שלך?";
var invalidEmail = "אימייל לא תקין\r\n";
var emailMustBe = "\r\n האימייל צריך להיות\r\nburner@somecompany.com";
var accessGranted = "יש לך כעת גישת עריכה";
var editableColumns = "את יכולה כעת לערוך את העמודות הבאות בשורה:\r\n";

var missionStringSize = 15; //Extract only part of the mission for the description

var creatorField = 5; //"E"
var fillerEmailField = 7; //"G"
var communityCommentsField = 10; //"J"

var ui = SpreadsheetApp.getUi();

function removeAllProtection() {
 var ss = SpreadsheetApp.getActive();
 var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
 for (var i = 0; i < protections.length; i++) {
   var protection = protections[i];
   protection.remove();
 } 
}

function getMainProtection(protections) {
  for(var i=0;i<protections.length; i++) {
    if(protections[i].getDescription() === mainProtectionName) {
      return protections[i];
    }
  }
}


/**
 * Test function for onEdit. Passes an event object to simulate an edit to
 * a cell in a spreadsheet.
 * Check for updates: https://stackoverflow.com/a/16089067/1677912
 */
function test_onEdit() {
  Logger.log("OnEdit");
  onEdit({
    user : Session.getActiveUser().getEmail(),
    source : SpreadsheetApp.getActiveSpreadsheet(),
    range : SpreadsheetApp.getActiveSpreadsheet().getActiveCell(),
    value : SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue(),
    authMode : "LIMITED"
  });
    Logger.log("OnEdit Finished");
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
  if(e.range.getColumn() === fillerEmailField) {
    columnsToGiveAccess = [labelsCol, fillerCol, fillerEmailCol, descriptionCol, fillerPhoneCol, taskStatusCol, extraCol, extraCol2];
  }

  // Only grant permissions if set this array
  if(columnsToGiveAccess.length > 0) {
    if(grantPermission(e, columnsToGiveAccess)) {
      unprotectRange(columnsToGiveAccess, e); //Set as unprotected in main sheet
      // Popping the last two columns (Use as extra - should not be displayed)
      columnsToGiveAccess.pop();
      columnsToGiveAccess.pop();
      showAlert(accessGranted, editableColumns + columnsToGiveAccess.toString());
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
  
  if(row.getColumn == fillerEmailField || row.getColumn == communityCommentsField) {
   return true; //Anyone can edit the filler email field 
  }

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
 
  var errorText = "";
  if (fillerEmailContent == "") {
    errorText = "This task doesn\'t have filler email yet";
  }
  
  if(errorText) { 
    undoActiveRange();
    showAlert("Can\'t edit",errorText);
    return false;
  }
  
  // Apply changes to backup sheet
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
  Logger.log("Granting permission");
  var newValue = e.range.getValue(); //Better this then e.value which on copy return other results
  Logger.log("Value:" + e.value);
  Logger.log("Actual Value:" + newValue);
  Logger.log("Old Value:" + e.oldValue);
  
  if(!validateEmail(newValue)) {
    e.range.setValue(e.oldValue || '');
    showAlert("",invalidEmail + newValue + emailMustBe);
    return false;
  }
  
  if(showYesNoCancel(isEmailCorrect,newValue) != ui.Button.YES) {
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
  Logger.log("Granted permission successfully");
  return true;
}

// Set all the needed unprotected columns
function testUnprotectEmptycolumns() {
  var columnsToUnprotect = [communityCommentsCol, contactsCol, fillerEmailCol];
  unprotectRange(columnsToUnprotect);
}

function unprotectRange(columnsToGiveAccess, e) {  
  Logger.log("Unprotecting Range...");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var protections = ss.getProtections(SpreadsheetApp.ProtectionType.SHEET);  
  var mainProtection = getMainProtection(protections);
  assertOrThrow(mainProtection, "Cant find main protection");
  Logger.log("Found main protection");
  
  var unprotectedArray = mainProtection.getUnprotectedRanges();
  for(var i=0;i<columnsToGiveAccess.length;i++) {
    var range;
    if(e == null) { //Unprotect the whole row
      range = getRange(columnsToGiveAccess[i] +":"+ columnsToGiveAccess[i]);
    }else {
      range = getRange(columnsToGiveAccess[i] + e.range.getRow());
    }
    unprotectedArray.push(range);
  }
  Logger.log(unprotectedArray);
  mainProtection.setUnprotectedRanges(unprotectedArray);
    Logger.log("Unprotected range successfully");
}


function protectRange(range, newEditorEmail) {
  // Get range task
  var taskContent = getCellContent(taskCell + range.getRow());
  
  // TODO - Should not protect if already protected
  var protection = range.protect().setDescription(taskContent.substr(0, missionStringSize) + ' ' + newEditorEmail);
  protection.addEditor(newEditorEmail);
}
