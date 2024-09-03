// This Google App Script is used to create a Google Sheet from variables specified in a Google Doc.
// That Google Sheet can be populated with data that substitutes for the variables in the Google Doc.
// The Doc with the subsittuted variables will then be turned into an email and sent to users specified in the Google Sheet.

// This function is called when the Google Sheet is opened.
function onOpen() {
    var ui = DocumentApp.getUi();
    ui.createMenu('Email')
        .addItem('Create Spreadsheet', 'createSpreadsheet')
        .addItem('Create Single Draft', 'testEmail')
        .addItem('Create Drafts', 'createDrafts')
        // Commenting to prevent accidentally sending!
        //.addItem('Send Email', 'sendEmail')
        .addToUi();
    }

function doc_to_html(document_Id)
    {
     var id = document_Id;
     var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+id+"&exportFormat=html";
     var param = 
            {
              method      : "get",
              headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
              muteHttpExceptions:true,
            };
     var html = UrlFetchApp.fetch(url,param).getContentText();
     return html
    }

function createSpreadsheet() {
    sheet=getSpreadSheet();
    updateSpreadsheet();
    // Display a dialog box with a link to the spreadsheet
    sheetURL = sheet.getUrl();
    var html = HtmlService.createHtmlOutput('<a href="' + sheetURL + '" target="_blank">Click here to open the spreadsheet</a>');
    DocumentApp.getUi().showModalDialog(html, 'Spreadsheet Created');
}

// This function is called when the "Create Spreadsheet" menu item is selected.
function getSpreadSheet() {
    // Look in the same directory as the doc for a file called DocName + "Email Spreadsheet".
    // If it exists, open it. If not, create it.
    var doc = DocumentApp.getActiveDocument();
    var docName = doc.getName();
    var id=doc.getId();
    var folder = DriveApp.getFileById(id).getParents().next();
    var files = folder.getFilesByName(docName + " (Email Spreadsheet)");
    var file;
    if (files.hasNext()) {
        file = files.next();
    }
    else {
        file = SpreadsheetApp.create(docName + " (Email Spreadsheet)");
        updateSpreadsheet();
    }
    // Write the vars to the header and freeze that row
    
    var sheet = SpreadsheetApp.openById(file.getId());
    return sheet;
}

function updateSpreadsheet(){
    var vars = getVars();
    if (vars==null){
        vars=["To","CC","BCC","Reply-To","Subject"];
    }
    else{
    vars = ["To","CC","BCC","Reply-To","Subject",...vars];
    }
    // Clear spreadsheet
    var sheet = getSpreadSheet().getSheets()[0];
    sheet.clear();
    // Update the spreadsheet with the latest variables from the doc
    sheet.getRange(1, 1, 1, vars.length).setValues([vars]);
    sheet.setFrozenRows(1);
}


function getVars(){
    // Read the word doc for variables (ex: {{var1}}) and return them as an array
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var text = body.getText();
    var vars = text.match(/{{\w+}}/g);
    return vars;
}

// This function is called when the "Test Email" menu item is selected.
// Creates an email for the first row of the spreadsheet and puts it in the user's drafts folder.
function prepareEmail(headers,vars,html) {
    var fieldTo=vars[headers.indexOf("To")];
    var fieldCC=vars[headers.indexOf("CC")];
    var fieldBCC=vars[headers.indexOf("BCC")];
    var fieldReplyTo=vars[headers.indexOf("Reply-To")];
    var fieldSubject=vars[headers.indexOf("Subject")];
    var replacementHeaders=headers.slice(headers.indexOf("Subject")+1);
    var replacementVars=vars.slice(headers.indexOf("Subject")+1);
    var replacements=makeDict(replacementHeaders,replacementVars);
    var body=DocumentApp.getActiveDocument().getBody().getText();    
    console.log(html);
    var emailBody = body.replace(/{{\w+}}/g, function(all) {
        return replacements[all];
    });
    var html = html.replace(/{{\w+}}/g, function(all) {
        return replacements[all];
    });
    return [fieldTo,fieldCC,fieldBCC,fieldReplyTo,fieldSubject,emailBody,html];
}

function testEmail() {
    var sheet = getSpreadSheet().getSheets()[0];;
    var data = sheet.getRange(1, 1, 2, sheet.getLastColumn()).getValues();
    var headers=data[0];
    var vars=data[1];
    //Get HTML Body now to avoid rate limiting
    var html=doc_to_html(DocumentApp.getActiveDocument().getId());
    [fieldTo,fieldCC,fieldBCC,fieldReplyTo,fieldSubject,emailBody,html]=prepareEmail(headers,vars,html);
    GmailApp.createDraft(fieldTo, fieldSubject, emailBody, {cc: fieldCC, bcc: fieldBCC, replyTo: fieldReplyTo, htmlBody: html});
}

function sendEmail() {
    var sheet = getSpreadSheet().getSheets()[0];;
    var data = sheet.getDataRange().getValues();
    var headers=data[0];
    //Get HTML Body now to avoid rate limiting
    var html=doc_to_html(DocumentApp.getActiveDocument().getId());
    for (var i = 1; i < data.length; i++) {
        var vars=data[i];
        [fieldTo,fieldCC,fieldBCC,fieldReplyTo,fieldSubject,emailBody,html]=prepareEmail(headers,vars,html);
        GmailApp.sendEmail(fieldTo, fieldSubject, emailBody, {cc: fieldCC, bcc: fieldBCC, replyTo: fieldReplyTo, htmlBody: html});
    }
}

function createDrafts(){
    var sheet = getSpreadSheet().getSheets()[0];;
    var data = sheet.getDataRange().getValues();
    var headers=data[0];
    //Get HTML Body now to avoid rate limiting
    var html=doc_to_html(DocumentApp.getActiveDocument().getId());
    for (var i = 1; i < data.length; i++) {
        var vars=data[i];
        [fieldTo,fieldCC,fieldBCC,fieldReplyTo,fieldSubject,emailBody,html]=prepareEmail(headers,vars,html);
        GmailApp.createDraft(fieldTo, fieldSubject, emailBody, {cc: fieldCC, bcc: fieldBCC, replyTo: fieldReplyTo, htmlBody: html});
    }
}

function makeDict(headers,vars){
    // Create dictionary of variables and their replacements
    var replacements = {};
    for (var i = 0; i < headers.length; i++) {
        replacements[headers[i]]=vars[i];
    }
    return replacements;
}