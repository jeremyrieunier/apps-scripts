function onOpen() {
 var ui = SpreadsheetApp.getUi();
 ui.createMenu('Refresh Data')
  .addItem('Refresh Now', 'dataRefresh')
  .addToUi();
}


function dataRefresh() {
  var source_1 = SpreadsheetApp.openById("SOURCESHEET_ID");  //add sourcesheet ID here
  var sourceSheet_1 = source_1.getSheetByName("SOURCESHEET_NAME");  //add sourcesheet sheet name(tab) here
  var row_1 = sourceSheet_1.getMaxRows()
  var col_1 = sourceSheet_1.getMaxColumns();

  var rowStart = 1; //starting point for mid row position
  
  var sourceRange_1 = sourceSheet_1.getRange(rowStart,1,row_1-rowStart,col_1);
  var values_1 = sourceRange_1.getValues();

  var target_1 = SpreadsheetApp.getActiveSpreadsheet();    //targetsheet is the current active sheet
  var targetSheet_1 = target_1.getSheetByName("TARGETSHEET_NAME");     //add name of targetsheet sheet name(tab) here
  var targetRange_1 = targetSheet_1.getRange(rowStart,1,row_1-rowStart,col_1);

  targetRange_1.setValues(values_1);


}
