function emailToSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = spreadsheet.getSheetByName(active_sheet);
  
  // clearing spreadsheet of data before pulling text
  var LastRow = ss.getLastRow();
  if (LastRow > 1) ss.getRange('A2:G' + LastRow).clearContent();  
  
  // filters messages newer than 9 days, eleiminates emails sent from staff member 
  var labelFilter = 'label:newsletter newer_than:9d -from:' + currentUser; 
  var msgs = Gmail.Users.Messages.list('me', {q:labelFilter}).messages;
  var out = [], row = [];
  msgs.forEach(function (e)
  {
    var dat = GmailApp.getMessageById(e.id).getDate();
    //need to strip out the names and <> from the from address, lookbehind doesn't work with flavor of regex
    var raw_from = GmailApp.getMessageById(e.id).getFrom();
    var from = raw_from.match(/[^@<\s\"]+@[^@\s>\"]+/).toString();
    var msgplain = GmailApp.getMessageById(e.id).getPlainBody();
    var attachments = GmailApp.getMessageById(e.id).getAttachments();    
    //var links = '', linksfinal = ''
    var links = [], linksfinal = [];
    if(attachments[0] == 'GmailAttachment')
    {      
      for(var k in attachments)
      {
        var contentType = attachments[k].getContentType()
        Logger.log('content type is: ' + contentType)
        if (contentType == ("image/jpeg" || "image/jpg" || "image/bmp" || "image/gif" || "image/png" || "image/svg")) { 
          Logger.log('What type of attachment is this attachment: ' + contentType)
          var attachment = attachments[k];
          var attch = attachment.copyBlob();
          var folder = DriveApp.getFolderById(picsFolderId)
          var link = folder.createFile(attch).getId(); 
          Logger.log('link for pics is: *********** ' + link)  
          links.push(link)              
          linksfinal.push(links);
          var linksfinal2 = '' + linksfinal
          var linksfinal3 = linksfinal2.replace(/,*$/, "");
          var msgfinal = linksfinal3 + " " + msgplain
          var row = [from,from,from,from,msgfinal,linksfinal3,dat];
          Logger.log(row)
        } else
        {
          Logger.log('unsupported image file, need to figure out manually what it is: ' + contentType)
          var linksfinal2 = ""
          var row = [from,from,from,from,msgplain,'',dat];            
          }
      }
    } else
    {
        var linksfinal2 = ""
        var row = [from,from,from,from,msgplain,'',dat];          
    }  
    out.push(row);  
    GmailApp.getMessageById(e.id).markRead();
  })
  ss.getRange(ss.getLastRow()+1,1,out.length,out[0].length).setValues(out);
}
