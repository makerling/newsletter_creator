//this array variable stores all the members of the entity with the following information: email address, Team, first name, last name, first name(s) of household unit
//first names of household unit are used because members didn't want their last name put on the newsletter, but last name is needed internally for sorting the entries
var name = new Array ( );
  //name[0] = new Array ('name@email.com','team name','first name','last name','first & last name of couple');
 name[0] = new Array ('optimizer@vanderling.com','T1','Joe','Bradson','Joe & Donna Bradson');
 name[1] = new Array ('voip.ms@vanderling.com','T2','Jack','Greg','Jack & Sarah Greg');
 name[2] = new Array ('zipcar@vanderling.com','T3','Brad','Jackson','Brad & Darlene Jackson');
  
//This function populates the "editColumn" function, it will run 3 times with the 3 values in this function, once for each column that needs changing in the template
function replace(){
  editColumn('2','4'); //FIRST NAMES: template column number (starting with "1), array item (starting with "0");
  Logger.log('**********************************');
  editColumn('3','3'); //LAST NAME: template column number (starting with "1), array item (starting with "0");
  Logger.log('**********************************');
  editColumn('4','1'); //TEAM: template column number (starting with "1), array item (starting with "0");
}

function editColumn(templateColumnNum,arrayItem) {
  var names = new Array (name);

  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(active_sheet);
  var lrow = ss.getLastRow();
  
  // Set the scope of the array from the Spreadsheet
  var rngFirstNameCouple = ss.getRange(2, templateColumnNum, lrow - 1, 1);
  var dataFirstNameCouple = rngFirstNameCouple.getValues();
  
  // Code to convert email address to desired value for 2nd, 3rd and 4th column of template
  // iterate through the list of names in template, and match email address in array and populate the columns with corresponding data from array
  for (var d=0; d < dataFirstNameCouple.length; d++) {
    for ( i=0; i<name.length; i++ ) {
      for ( j=0; j<name[i][0].length; j++ ) {
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
