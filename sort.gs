function sort() {
  
  /**  Variables for customization:
  
  Each column to sort takes two variables: 
      1) the column index (i.e. column A has a colum index of 1
      2) Sort Ascending -- default is to sort ascending. Set to false to sort descending
  
  **/

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
