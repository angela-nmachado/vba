function main(workbook: ExcelScript.Workbook) {
    // Obtenha a célula e a planilha ativas.
    let selectedCell = workbook.getActiveCell();
    let selectedSheet = workbook.getActiveWorksheet();
  
  
  
    // TODO: escreva o código ou use o botão inserir ação abaixo.
    // Insert at range 11:11 on selectedSheet, move existing cells down
  
    // Paste to range C10 on selectedSheet from range C9 on selectedSheet
    selectedSheet.getRange("12:22").copyFrom(selectedSheet.getRange("1:11"), ExcelScript.RangeCopyType.all, false, false);
    // Paste to range C10 on selectedSheet from range C9 on selectedSheet
    selectedSheet.getRange("2:5").insert(ExcelScript.InsertShiftDirection.down);
  
  
  }
  