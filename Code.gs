/* Deploy Instructions:
// 1.Protect Master - Give it description 'Main'
// 2.Run once testUnprotectEmptycolumns()
// Then you can test with simulateOnEdit - Make sure you have an active sheet open
*/

var masterSheetName = "ראשי" // sheet where the cells are protected from updates

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

var adminsEmails = ['galbra@gmail.com','orkin2@gmail.com', 'amenly@gmail.com', 'bakmic@gmail.com'];


// Translations
var isEmailCorrect = "אנא בדקי שנית האם זה האימייל שלך?";
var invalidEmail = "אימייל לא תקין\r\n";
var emailMustBe = "\r\n האימייל צריך להיות\r\nburner@somecompany.com";
var accessGranted = "יש לך כעת גישת עריכה";
var editableColumns = "את יכולה כעת לערוך את העמודות הבאות בשורה:\r\n";

var missionStringSize = 15; //Extract only part of the mission for the description

var creatorEmailField = 5; //"E"
var fillerEmailField = 7; //"G"
var communityCommentsField = 10; //"J"

var ui = SpreadsheetApp.getUi();

function testRemoveAllProtection() {
 var ss = SpreadsheetApp.getActive();
 var protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
 for (var i = 0; i < protections.length; i++) {
   var protection = protections[i];
   protection.remove();
 } 
}

// Set all the needed unprotected columns
function testUnprotectEmptycolumns() {
  var columnsToUnprotect = [communityCommentsCol, contactsCol];
  unprotectRange(columnsToUnprotect);
}

function getMainProtection(protections) {
  for(var i=0;i<protections.length; i++) {
    if(protections[i].getDescription() === mainProtectionName) {
      return protections[i];
    }
  }
}

// Add a custom menu to the spreadsheet.
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('מידברנרות')
      .addItem('הפעלי הרשאות', 'simulateOnEdit')
      .addToUi();
}

// This needs to be bind to the "Current project trigger" To actually perform actions as the program admin
// TODO - For anonymous user - cant use any user interface boxes - Try to identify and undo the operation
function onEditAsAdmin(e) {
  Logger.log("On edit as admin");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getActiveSheet();
  if (masterSheet.getName() != masterSheetName) return;
  
  //Check if its just open and we should restrict it
  if(!checkIfHasAccess(e)) {
    return;
  }
  
  Logger.log("Have access");
  var columnsToGiveAccess = [];
  if(e.range.getColumn() === fillerEmailField) {
    columnsToGiveAccess = [labelsCol, fillerCol, fillerEmailCol, descriptionCol, fillerPhoneCol, taskStatusCol, extraCol, extraCol2];
  }else if(e.range.getColumn() === creatorEmailField) {
    columnsToGiveAccess = [labelsCol, fillerCol, fillerEmailCol, descriptionCol, fillerPhoneCol, taskStatusCol, extraCol, extraCol2, creatorCol, creatorCommentsCol];
  }

  // Only grant permissions if set this array
  if(columnsToGiveAccess.length > 0) {
    if(grantPermission(columnsToGiveAccess, e)) {
      unprotectRange(columnsToGiveAccess,e); //Set as unprotected in main sheet
      // Popping the last two columns (Use as extra - should not be displayed)
      columnsToGiveAccess.pop();
      columnsToGiveAccess.pop();
      showAlert(accessGranted, editableColumns + columnsToGiveAccess.toString());
    }
  }
  Logger.log("Finished");
}


/**
 * Test function for onEdit. Passes an event object to simulate an edit to
 * a cell in a spreadsheet.
 * Check for updates: https://stackoverflow.com/a/16089067/1677912
 */
function simulateOnEdit() {
  if(!isAdmin()) {
    return
  }
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

function isAdmin() {
  var userEmail = Session.getActiveUser().getEmail();
  for(var i=0;i<adminsEmails.length;i++) {
    if(adminsEmails[i] === userEmail) {
      return true;
    }
  }
  return false;
}

function onEdit(e){
  Logger.log("On edit");
  if(!isAdmin()) {
      Logger.log("Not admin");
    return;
  }
  return onEditAsAdmin(e);
}

function checkIfHasAccess(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getActiveSheet();
  var masterRange = masterSheet.getActiveRange();
  
  // TODO Should only be triggered if from menu
  if(masterRange.getHeight() > 1) {
//    showAlert("Can\'t edit","Please edit one line at a time");
    return false;
  }
  
  return true;
}

function grantPermission(columnsToGiveAccess, e) {
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
