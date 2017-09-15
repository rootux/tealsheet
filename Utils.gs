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