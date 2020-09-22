// FOR NEW Account need to do following:
// enable > Resources/Advanced Google Services > Gmail API and Drive API 
// (also check to make sure they are activated in Google Cloud Platform API Dashboard)
//******************************************************************************************************

// adds menu items in spreadsheet when opened
function onOpen(e) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  // When the user clicks on "Newsletter" then "1 Insert Email text", the function pullingEmails is executed.
  menuEntries.push({name: "1 Insert Email text", functionName: "pullingEmails"});
  menuEntries.push(null); // line separator
  menuEntries.push({name: "2 Send as Word doc my Email", functionName: "sendingEmails"});

  ss.addMenu("Newsletter", menuEntries);
}

//***********************************************************************************

function pullingEmails(){
  emailToSpreadsheet();
  replace();
  sort();
}

function sendingEmails(){
  mailMergeToDocs();
  addImages();
  emailAsDocx();
}

function isOdd(num) { return num % 2;}
