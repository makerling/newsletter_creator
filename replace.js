function editColumn(templateColumnNum,arrayItem) {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('mail merge');
  var lrow = ss.getLastRow();
  
  // data is stored in 'newsletter_data' sheet of spreadsheet
  var ss2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('newsletter_data');
  var lrow2 = ss2.getLastRow();  
  
  // Set the scope of the array from the Spreadsheet
  var rngFirstNameCouple = ss.getRange(2, templateColumnNum, lrow - 1, 1);
  var dataFirstNameCouple = rngFirstNameCouple.getValues();
  
  // row, column, numRows, numColumns
  var data2 = ss2.getRange(2,1, lrow2 - 1, 5)
  var data2values = data2.getValues();    
  Logger.log('data2values is: ' + data2values)
  
  // Code to convert email address to desired value for 2nd, 3rd and 4th column of template
  for (var d=0; d < dataFirstNameCouple.length; d++) {
    for ( i=0; i<data2values.length; i++ ) {
      for ( j=0; j<data2values[i][0].length; j++ ) {
        if (data2values[i][0] === dataFirstNameCouple[d][0]) {
          dataFirstNameCouple[d][0] = data2values[i][arrayItem];
        }
      }
    }
  }
  rngFirstNameCouple.setValues(dataFirstNameCouple);
}

// runs 3 times with the 3 values in this function, once for each column that needs changing in the template
function replace(){
  editColumn('2','4'); //FIRST NAMES: template column number (starts with "1), array item (starts with "0");
  Logger.log('**********************************');
  editColumn('3','3'); //LAST NAME: template column number (starts with "1), array item (starts with "0");
  Logger.log('**********************************');
  editColumn('4','1'); //TEAM: template column number (starts with "1), array item (starts with "0");
}