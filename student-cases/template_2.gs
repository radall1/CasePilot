function onOpen() {
  const ui = DocumentApp.getUi();
  const menu = ui.createMenu('Disseminate');
  
  menu.addItem(`Send Email`, 'sendReport');
  menu.addToUi();
}

function sendReport() {

  var placeholderKEY = "{{insertMaterialsRequest}}"                         // insert the placeholder key 

  var currentDocumentId = DocumentApp.getActiveDocument().getId();             // Finding ID of current document 
  var currentDocumentName = DocumentApp.getActiveDocument().getName();         // Finding Name of current document
  var groupEmail = "honorcouncil@stolaf.edu";                                  // Update with your Google Group email address

  var doc = DocumentApp.openById(currentDocumentId);                           // Opening the current document  
  var content = doc.getBody().getText();                                       // Obtaining the text of the document 

 
  var emailIndex = content.indexOf("<email:");                                 // Obtaining the email address to send the email to 
  if (emailIndex !== -1) {
    var address = content.substring(emailIndex + 7);                          // Start after "<email:"
    var endIndex = address.indexOf(">");                                      // End with ">"

    if (endIndex !== -1) {
      address = address.substring(0, endIndex);                                // email stored in `address`
    }
  }

  
  var InvSIndex = content.indexOf("Case Investigator");                         // Obtaining the name of the case investigator 
  var InvEIndex = content.indexOf("St. Olaf Honor Council");
  
  if (InvSIndex !== -1 && InvEIndex !== -1 && InvSIndex < InvEIndex) {
    var InvName = content.substring(InvSIndex + 17, InvEIndex).trim();          // Investigator name stored in `InvName`
  } 


  var ToSIndex = content.indexOf("Dear") + 5;                                   // Obtaining the name of the receipent
  var ToEIndex = content.indexOf(":");
  
  if (ToSIndex !== -1 && ToEIndex !== -1 && ToSIndex < ToEIndex) {
    var ToName = content.substring(ToSIndex, ToEIndex).trim();                  // Receipent name stored in `ToName`
  } 

  // this forms the PDF of the current attachment to be used as an attachment 
  var pdf = DriveApp.getFileById(currentDocumentId).getAs(MimeType.PDF).setName(currentDocumentName + ".pdf");

  
  var message = {                                                               // email content 
    to: address,
    subject: "CONFIDENTIAL: Honor Council Case " + currentDocumentName[0]+currentDocumentName[1]+currentDocumentName[2]+currentDocumentName[3]+currentDocumentName[4]+currentDocumentName[5]+currentDocumentName[6]+currentDocumentName[7],
    body: "Dear " + ToName + ",\n\nPlease read the attached memo in its entirety. Let me know if you have any questions or concerns.\n\nSincerely,\n" + InvName + "\nHonor Council Case Investigator",
    name:"Honor Council",
    attachments: [pdf]
  };

  
  GmailApp.sendEmail(message.to, message.subject, message.body, {                // this actually sends the email
    from: groupEmail,
    name: message.name,
    attachments: message.attachments
  });

  // this forms the PDF of the current attachment to be stored for Internal purposes
  var document = DriveApp.getFileById(currentDocumentId);
  var folder = document.getParents().next();                                      // Get the parent folder of the document
  var pdfBlob = document.getAs(MimeType.PDF);                                     // conversion to pdf
  var pdfFile = folder.createFile(pdfBlob);                                       // creates the file inside the parent folder

  var pdfFileId = pdfFile.getId();                                                // retains the ID of the formed PDF

  // this forms the name of the Case Report
  var documentName = currentDocumentName[0]+currentDocumentName[1]+currentDocumentName[2]+currentDocumentName[3]+currentDocumentName[4]+currentDocumentName[5]+currentDocumentName[6]+currentDocumentName[7] + " - Case Report";


  var files = DriveApp.getFilesByName(documentName);                               // using the name above, this finds the ID
  var documentId = "";

  while (files.hasNext()) {                                                       // checks all the files to find the correct one
    var file = files.next();
    if (file.getName() === documentName) {
      documentId = file.getId();
      break;
    }
  }

  if (documentId !== "") {
    
    var pdfDocumentUrl = "https://drive.google.com/file/d/" + pdfFileId + "/edit";  // creates link using ID

    var targetDocument = DocumentApp.openById(documentId);                          // opens case report
    var content = targetDocument.getBody();                                         // access text of case report
    content.replaceText(placeholderKEY, pdfDocumentUrl);                            // places the link in the case report
    targetDocument.saveAndClose();                                                  // saves case report
  }                                                                                 

  DriveApp.getFileById(currentDocumentId).setTrashed(true);                         // deletes current file

}

