function mailMergeToDocs() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = spreadsheet.getSheetByName(active_sheet);
  //getting count of rows
  var lastRow = ss.getLastRow();
  var dataRange = ss.getRange(1, 1, lastRow, 3).getValues();
  var rowsData = [];
  dataRange.forEach(function(row){
      var rowNumberTotal = Object.keys(rowsData).length;      
      rowsData.push([row[0],row[1],row[2]]);
    });  
  for (var row in dataRange)
  {
    var length = Object.keys(rowsData).length;
  }

  //define header cell style which we will use while adding cells in header row
  //Backgroud color, text bold, white
  var headerStyle = {};
  headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#59b300';
  headerStyle[DocumentApp.Attribute.BOLD] = true;
  headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  
  //Style for the cells other than header row
  var cellStyleDark = {};
  cellStyleDark[DocumentApp.Attribute.BACKGROUND_COLOR] = '#ddefcc';
  cellStyleDark[DocumentApp.Attribute.BOLD] = false;
  cellStyleDark[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  
  var cellStyleLight = {};
  cellStyleLight[DocumentApp.Attribute.BACKGROUND_COLOR] = '#eef7e5';
  cellStyleLight[DocumentApp.Attribute.BOLD] = false;
  cellStyleLight[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';

  //By default, each paragraph had space after, so we will change the paragraph style to add zero space
  var paraStyle = {};
  paraStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
  paraStyle[DocumentApp.Attribute.LINE_SPACING] = 1;
  
  //get the document
  var doc = DocumentApp.openById(fchCommunicatorTemplateId); //or
  //deletes the existing text of the file so that it can be a blank slate for this document merge
  doc.setText('');   
  
  //get the body section of document
  var body = doc.getBody();
  body.setMarginLeft(50)
  body.setMarginRight(50)
  body.setMarginBottom(50)
  body.setMarginTop(50)
  
  //Add a table in document
  var table = body.appendTable();
  
  //Create 5 rows and 4 columns
  for(var i=0; i<lastRow; i++){
    var tr = table.appendTableRow();
    
    //add 4 cells in each row
    for(var j=0; j<3; j++){
      //var td = tr.appendTableCell('Cell '+i+j);
      //getRange needs to have numbers beginning with 1 not 0
      var Team = ss.getRange(i+1,4).getValue();
      var Name = ss.getRange(i+1,2).getValue();

      //DocumentApp.getActiveDocument().getBody().insertImage(0, img); 
      var Message = ss.getRange(i+1,5).getValue();  
      if(j == 0) var td = tr.appendTableCell(Team);
      if(j == 1) var td = tr.appendTableCell(Name);
      if(j == 2) var td = tr.appendTableCell(Message);  
      
      if(i == 0 || j == 0) td.setAttributes(headerStyle);
      else if (isOdd(i) == 1) td.setAttributes(cellStyleDark);
      else td.setAttributes(cellStyleLight);
      
      //Apply the para style to each paragraph in cell
      var paraInCell = td.getChild(0).asParagraph();
      paraInCell.setAttributes(paraStyle);
      //Setting alignment
      paraInCell.setAlignment(DocumentApp.HorizontalAlignment.CENTER);  
      td.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      if( i > 0 && j == 2) 
        paraInCell.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    }
  }
 
 //get the current date written out for title
 var date = new Date(),
     locale = "en-us",
     month = date.toLocaleString(locale, { month: "short" });
 var monthReal = month.match(/^\w+ \d+, \d+/g);
 //removing stray double spaces //removing double commas (breaks importing pictures)
 body.replaceText(" +", " ").replaceText("FCH Communicator dated:", "\t\tNewsletter for Acme Company");
 //setting the width of columns
 table.setColumnWidth(0,55)
 table.setColumnWidth(1,80)
 table.setColumnWidth(2,400)
 table.setBorderColor('#ffffff');
  //Save and close the document
  doc.saveAndClose();
}
