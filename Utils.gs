function validateEmail(email) {
    var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(email);
}

function showAlert(title, text) {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  return ui.alert(
     title,
     text,
      ui.ButtonSet.OK);
}

function showYesNoCancel(title, text) {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  return ui.alert(
     title,
     text,
      ui.ButtonSet.YES_NO_CANCEL);
}

function isEmptyArrays(arrays) {
  if(arrays.constructor === Array && arrays.length > 0) {
    for(var i=0;i<arrays.length;i++) {
      if(arrays[i].length > 0 && (arrays[i][0] != "")) {
          return false; 
      }
    }
  }
  return true;
}