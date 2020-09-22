//need to do for loop on each imageIDs to catch >1 picture attachment
function addImages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var LastRow = sheet.getLastRow();
  // When the "numRows" argument is used, only a single column of data is returned.
  var range = sheet.getRange(2, 6, LastRow);
  var values = range.getValues();
  var imageIDs = []
  
  // Prints 3 values from the first column, starting from row 1.
  for (var row in values) {
    for (var col in values[row]) {
      var rowRemovedComma = values[row][col].replace(/,\s*$/, "");
      Logger.log('rowRemovedComma is:' + rowRemovedComma)
      if (rowRemovedComma != '')
      {
        imageIDs.push(rowRemovedComma);
      }
   }
 }
  //Logger.log('imageIDszzzz is: ' + imageIDs)
  //Logger.log('rowRemovedCommazzzz is:' + rowRemovedComma)
  //Logger.log(imageIDs)
  var imageIDsLength = imageIDs.length;
  for(i = 0; i < imageIDsLength; i++ )
    imageIDs[i] && imageIDs.push(imageIDs[i]);  // copy non-empty values to the end of the array
    imageIDs.splice(0 , imageIDsLength);  // cut the array and leave only the non-empty values
    Logger.log('comma removed imagedIDs array is: ' + imageIDs)   
    Logger.log('imageIDsLength is: ' + imageIDsLength)
  var doc = DocumentApp.openById(fchCommunicatorTemplateId);
  var tables = doc.getTables();
  for (var k in tables)
  {
    var table = tables[k];
    var tablerows=table.getNumRows();
    for ( var row = 0; row < tablerows; ++row ) {
      var tablerow = table.getRow(row);
      var cell=2
      var celltext = tablerow.getChild(cell).getText();
      for (var image = 0; image < imageIDsLength; image++) {
        Logger.log('imageID is: ' + imageIDs[image])
        Logger.log('celltext is: ' + celltext)
        if(celltext.match(imageIDs[image]) !=null) {
          
          var imageIDarray = imageIDs[image].split(',');
          Logger.log('imageIDarray is: ' + imageIDarray)
          Logger.log('#1: ' + imageIDarray[0] + '#2: ' + imageIDarray[1] + '#3: ' + imageIDarray[2])          
          
          for (var n = 0; n < imageIDarray.length; n++) {
            Logger.log('imageIDarray is: ' + imageIDarray[n])
            var img = DriveApp.getFileById(imageIDarray[n]).getBlob();
            var cellImage = table.getCell(row, cell).insertImage(0, img);           
            table.replaceText(imageIDarray[n], "");
            
            //get the dimensions of the image AFTER having inserted it to fix
            //its dimensions afterwards
            var originW =  cellImage.getWidth();
            var originH = cellImage.getHeight();
            var newW = originW;
            var newH = originH;
            var ratio = originW/originH
            maxWidth = 300
            if(originW>maxWidth){
              newW = maxWidth;
              newH = parseInt(newW/ratio);
            }
            cellImage.setWidth(newW).setHeight(newH); 
          }
        }
      }
    }
  }
  //removing some stray commas after replacing IDs with pictures (leaves the commas for someone submitting multiple pics)
  var body = doc.getBody();
  body.replaceText(",,+", "")
  body.replaceText("^(, )", "")
  doc.saveAndClose();
}
