//FOR NEW Account need to do following:
//enable > Resources/Advanced Google Services > Gmail API and Drive API (also check on link to make sure they are activated in Google Cloud Platform API Dashboard)
/** Build a menu item
From https://developers.google.com/apps-script/guides/menus#menus_for_add-ons_in_google_docs_or_sheets 
ToDo:
- check names with communicator list 
- catch inline pics
- change green color to real green, remove alternative colors in rows
**/

//******************************************************************************************************

function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createMenu('communicator'); 

var menu1 = "'1 Insert Email text', 'pullingEmails'"
var menu2 = "'2 Send as Word doc my Email', 'sendingEmails'"

if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    // Add a normal menu item (works in all authorization modes).
        menu.addItem(menu1);  
        menu.addItem(menu2);                
  } else {
    // Add a menu item based on properties (doesn't work in AuthMode.NONE).
    var properties = PropertiesService.getDocumentProperties();
    var workflowStarted = properties.getProperty('workflowStarted');
    if (workflowStarted) {
        menu.addItem(menu1);  
        menu.addItem(menu2);    
    } else {
        menu.addItem(menu1);  
        menu.addItem(menu2);   
  }
   menu.addToUi();
  }
}

//******************************************************************************************************

//need to change manually to picsFolderId value
//only the owner of the file may trash or delete the file
//function DeleteOldFiles() {
//  var Folders = new Array(
//    ''
//  );
//  var Files;

//  for each (var FolderID in Folders) {
//    Folder = DriveApp.getFolderById(FolderID)
//    Files = Folder.getFiles();

//    while (Files.hasNext()) {
//      var File = Files.next();

//      File.setTrashed(true); // Places the file int the Trash folder
      //Drive.Files.remove(File.getId()); // Permanently deletes the file
//      Logger.log('File ' + File.getName() + ' was deleted.');
//    }
//  }
//}

//******************************************************************************************************

function pullingEmails(){
  //DeleteOldFiles();
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
