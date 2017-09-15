var openTillHaveValue = 8; //"H"
  var ui = SpreadsheetApp.getUi();;

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
  Logger.log(e.range.getColumn());
  //TODO - Notice that old value doesn't work if you edit more then a single cell
  if(e.range.getColumn() === openTillHaveValue) {
    checkForPermission(e);
  }
}

function checkForPermission(e) {
  Logger.log("Value:" + e.value);
  Logger.log("Old Value:" + e.oldValue);
  
  if(typeof(e.value) !== "string") {
   return; //TODO - what happenes when multiple cells, deleting a cell, Pasting a cell
  }
  
  if(!validateEmail(e.value)) {
    showAlert("","Invalid email " + e.value + " must be burner@somecompany.something");
    e.range.setValue(e.oldValue || '');
    return;
  }
  
  if(showYesNoCancel("Please Check if the email is correct?",e.value) != ui.Button.YES) {
    e.range.setValue(e.oldValue || '');
    return;
  }
  
  // TODO: should protect not this range but others
  protectRange(e.range, e.value);
  
    
  // Value is inserted for the first time
  if(e.value && !e.oldValue) {
      Logger.log("Value is inserted for the first time");
  }
}

function protectRange(range, newEditorEmail) {
  // TODO - Should not protect if already protected
  var protection = range.protect(); //.setDescription('Sample protected range');

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
