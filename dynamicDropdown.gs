function onEdit(){
  
  // Диапазон для "Доходы" в листе "Справочники"
  var incomeRange = 'C2:C'
  
  // Диапазон для "Расходы" в листе "Справочники"
  var expenseRange = 'G2:G'
  
  // Диапазон для "Контрагенты дохода"
  var incomeContragentsRange = 'E2:E'
  
  // Диапазон для "Контрагенты расхода"
  var expenseContragentsRange = 'I2:I'
  
  // Позиция колонки "Тип операции"
  var operationTypeCol = 6
  
  // Позиция колонки "Статья"
  var validationCol = 7
  
  // Позиция колонки "Контрагент"
  var contragentCol = 10
  
  // Название страницы "Справочники"
  var dictionaryListName = 'Справочники'
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var datass = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dictionaryListName);
  
  var activeRange = ss.getActiveRange();
  
  if (ss.getSheetName() == 'Cashflow' && activeRange.getColumn() >= 2 && activeRange.getColumn() <= 5 && activeRange.getRow() > 1){
    
    var numRows = activeRange.getNumRows();
    
    for (var i = 1; i <= numRows; i++) {
      var activeRow = activeRange.getCell(i, 1).getRow();
      var operationType = ss.getRange(activeRow, operationTypeCol)
      var operationTypeValue = operationType.getCell(1, 1).getValue();
      if (operationTypeValue == 'Доходы') {
        var validationField = ss.getRange(activeRow, validationCol);
        validationField.getCell(1, 1).clearDataValidations();
        var validationRange = datass.getRange(incomeRange);
        var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
        validationField.getCell(1, 1).setDataValidation(validationRule);
        
        var validationField = ss.getRange(activeRow, contragentCol);
        validationField.getCell(1, 1).clearDataValidations();
        var validationRange = datass.getRange(incomeContragentsRange);
        var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
        validationField.getCell(1, 1).setDataValidation(validationRule);
      }
      else if (operationTypeValue == 'Расходы') {
        var validationField = ss.getRange(activeRow, validationCol);
        validationField.getCell(1, 1).clearDataValidations();
        var validationRange = datass.getRange(expenseRange);
        var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
        validationField.getCell(1, 1).setDataValidation(validationRule);
        
        var validationField = ss.getRange(activeRow, contragentCol);
        validationField.getCell(1, 1).clearDataValidations();
        var validationRange = datass.getRange(expenseContragentsRange);
        var validationRule = SpreadsheetApp.newDataValidation().requireValueInRange(validationRange).build();
        validationField.getCell(1, 1).setDataValidation(validationRule);
      }
      else {
        var validationField = ss.getRange(activeRow, validationCol);
        validationField.getCell(1, 1).clearDataValidations();
        
        var validationField = ss.getRange(activeRow, contragentCol);
        validationField.getCell(1, 1).clearDataValidations();
      }
    }
 
  }
  
}
