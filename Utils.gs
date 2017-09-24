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

function assertOrThrow(statement, errorArgs) {
    if (!statement) {
        showAlert("Error", errorArgs);
        throw new Error(errorArgs)
    }
}

function logToCell(range, text) {
  getRange(range).setValue(text);
}

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


function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}