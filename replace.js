function editColumn(templateColumnNum,arrayItem) {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(active_sheet);
  var lrow = ss.getLastRow();

  var ss2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(data);
  var lrow2 = ss2.getLastRow();
  
  // Set the scope of the array from the Spreadsheet
  var rngFirstNameCouple = ss.getRange(2, templateColumnNum, lrow - 1, 1);
  var dataFirstNameCouple = rngFirstNameCouple.getValues();

  // row, column, numRows, numColumns
  var data2 = ss2.getRange(1,templateColumnNum, lrow2 -1, 1)
  var data2values = data2.getValues();
  
  // Code to convert email address to desired value for 2nd, 3rd and 4th column of template
  // iterate through the list of names in template, and match email address in array and populate the columns with corresponding data from array
  for (var d=0; d < dataFirstNameCouple.length; d++) {
    //for ( i=0; i<name.length; i++ ) {
    for ( i=0; i<data2values.length; i++ ) {
      // for ( j=0; j<name[i][0].length; j++ ) {
      for ( j=0; j<data2values[i][0].length; j++ ) {
        //if the email address (first item) in template matches the email address of the array (0 = first item of matched array), replace the email address in column to desired result (last name, team, etc)
        if (name[i][0] === dataFirstNameCouple[d][0]) {
        dataFirstNameCouple[d][0] = name[i][arrayItem];
        //Logger.log("search term is: \"" + dataFirstNameCouple[d][0] + "\" ** First name is: " + name[i][2] + " ** Last name is: " + name[i][3]) 
        //Logger.log("******")
        }
      }
    }
  }
  rngFirstNameCouple.setValues(dataFirstNameCouple);
}

// 3 times with the 3 values in this function, once for each column that needs changing in the template
function replace(){
  editColumn('2','4'); //FIRST NAMES: template column number (starting with "1), array item (starting with "0");
  Logger.log('**********************************');
  editColumn('3','3'); //LAST NAME: template column number (starting with "1), array item (starting with "0");
  Logger.log('**********************************');
  editColumn('4','1'); //TEAM: template column number (starting with "1), array item (starting with "0");
}