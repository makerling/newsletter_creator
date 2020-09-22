function sort() {

  //Variable for column to sort first
  
  var sortFirst = 4; //index of column to be sorted by; 1 = column A, 2 = column B, etc.
  var sortFirstAsc = true; //Set to false to sort descending
  
  //Variables for column to sort second
 
  var sortSecond = 3;
  var sortSecondAsc = true;
  
  //Number of header rows
  
  var headerRows = 1; 

  /** End Variables for customization**/
  
  /** Begin sorting function **/

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(active_sheet);
  var range = sheet.getRange(headerRows+1, 1, sheet.getMaxRows()-headerRows, sheet.getLastColumn());
  range.sort([{column: sortFirst, ascending: sortFirstAsc}, {column: sortSecond, ascending: sortSecondAsc}]);
}
