function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Achita')
    .addItem('ÎNCEPE', 'startScript')
    .addToUi();
}

function startScript() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Introduceți numărul liniei pentru a verifica:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.CANCEL) {
    ui.alert('Operația a fost anulată.');
    return;
  }
  
  var rowNum = parseInt(response.getResponseText());
  if (isNaN(rowNum) || rowNum < 1) {
    ui.alert('Număr de linie nevalid.');
    return;
  }
  var id = sheet.getRange(rowNum, 2).getValue();
  var valueG = sheet.getRange(rowNum, 7).getValue();

  if (!id || !valueG) {
    ui.alert('Date nu au fost găsite.');
    return;
  }
  var targetSpreadsheet = SpreadsheetApp.openById('ID_Таблицы');
  var targetSheet = targetSpreadsheet.getSheetByName('Лист');
  var targetRange = targetSheet.getRange('A:A');
  var values = targetRange.getValues();
  
  var foundRow = -1;
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == id) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow == -1) {
    ui.alert('ID-ul nu a fost găsit în foaia țintă.');
    return;
  }
  var columns = [4, 7, 10, 13, 16, 19, 22, 25, 28, 31];
  var data = {};
  for (var i = 0; i < columns.length; i++) {
    data[columns[i]] = targetSheet.getRange(foundRow, columns[i]).getValue();
  }

  var dataText = '';
  for (var key in data) {
    dataText += 'Coloană ' + targetSheet.getRange(1, key).getA1Notation() + ': ' + data[key] + '\n';
  }
  var dataResponse = ui.alert('Date:\n' + dataText, ui.ButtonSet.OK_CANCEL);
  
  if (dataResponse == ui.Button.CANCEL) {
    ui.alert('Operația a fost anulată.');
    return;
  }
  var columnResponse = ui.prompt('Introdu litera coloanei din care să scadă valoarea:', ui.ButtonSet.OK_CANCEL);

  if (columnResponse.getSelectedButton() == ui.Button.CANCEL) {
    ui.alert('Operația a fost anulată.');
    return;
  }
  
  var columnLetter = columnResponse.getResponseText().toUpperCase();
  var columnIndex = getColumnIndex(columnLetter);

  if (!columnIndex || columns.indexOf(columnIndex) === -1) {
    ui.alert('Coloană nevalidă.');
    return;
  }
  var currentValue = targetSheet.getRange(foundRow, columnIndex).getValue();
  var newValue = currentValue - valueG;
  targetSheet.getRange(foundRow, columnIndex).setValue(newValue);

  ui.alert('Operația este finalizată. Noua valoare: ' + newValue);
}

function getColumnIndex(columnLetter) {
  return columnLetter.charCodeAt(0) - 64;

}
