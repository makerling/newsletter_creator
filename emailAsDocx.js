function emailAsDocx() {

  var url  = 'https://docs.google.com/document/d/'+fchCommunicatorTemplateId+'/export?format=docx';
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });

  var fileName = 'FCH Communicator dated ' + currentDate + '.docx';
  var blobs   = [response.getBlob().setName(fileName)];

  GmailApp.sendEmail(currentUser, 
                     "FCH Communicator for " + currentDate + " to be formatted", 
                     "Attached is the Communicator dated: " + currentDate + " and is ready to be formatted and sent out.",
    {
      attachments: blobs
    }
  );
}
