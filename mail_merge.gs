/***************************
This script provides a Mail Marge feature by merging values in google sheet into google doc
 
You can add this script to any spreadshit, please make sure 1st row has header names you will be referncing in your google doc.
eg. header1 and header2 can be referenced in doc as
ex: "Hi, [header1], you now have [header2] left"

*/  


// We will store docID to be able to use it multiple times
var scriptProperties = PropertiesService.getScriptProperties();

/**********************************************************************************
This function is responsible for actuall merge action.
 */
function doMerge() {
  var docLink = scriptProperties.getProperty("docLink");
  var regex = /[-\w]{25,}/;   // Find ID in URL
  var idFromLink = docLink.match(regex);
  var selectedTemplateId = idFromLink;

  timezone = "GMT+" + new Date().getTimezoneOffset()/60
  var date = Utilities.formatDate(new Date(), timezone, "dd-MM-yyyy HH:mm"); // "yyyy-MM-dd'T'HH:mm:ss'Z'"

  var templateFile = DriveApp.getFileById(selectedTemplateId);
  var mergedFile = templateFile.makeCopy();  // We will be making a new file to preserve template.

  mergedFile.setName(date+" "+templateFile.getName()+"_done");// new file's name
  var mergedDoc = DocumentApp.openById(mergedFile.getId());
  var bodyElement = mergedDoc.getBody();// find text we work with
  var bodyCopy = bodyElement.copy();// make a cope

  bodyElement.clear();

  var sheet = SpreadsheetApp.getActiveSheet();//current sheet

  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var fieldNames = values[0];//First row of the sheet must be the the field names

  for (var i = 1; i < numRows; i++) {//data values start from the second row of the sheet
    var row = values[i];
    var body = bodyCopy.copy();
    
    for (var f = 0; f < fieldNames.length; f++) {
      body.replaceText("\\[" + fieldNames[f] + "\\]", row[f]);//replace [fieldName] with the respective data value
    }
    
    var numChildren = body.getNumChildren();//number of the contents in the template doc
    
    for (var c = 0; c < numChildren; c++) {//Go over all the content of the template doc, and replicate it for each row of the data.
      var child = body.getChild(c);
      child = child.copy();
      if (child.getType() == DocumentApp.ElementType.HORIZONTALRULE) {
        mergedDoc.appendHorizontalRule(child);
      } else if (child.getType() == DocumentApp.ElementType.INLINEIMAGE) {
        mergedDoc.appendImage(child.getBlob());
      } else if (child.getType() == DocumentApp.ElementType.PARAGRAPH) {
        mergedDoc.appendParagraph(child);
      } else if (child.getType() == DocumentApp.ElementType.LISTITEM) {
        mergedDoc.appendListItem(child);
      } else if (child.getType() == DocumentApp.ElementType.TABLE) {
        mergedDoc.appendTable(child);
      } else {
        Logger.log("Unknown element type: " + child);
      }
    }
    
    mergedDoc.appendPageBreak();//Appending page break. Each row will be merged into a new page.

  }
}

/**********************************************************************************
This function will ingest doc url and grab a file ID from it.
 */
function getTemplate() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Google Doc template needed", "Please paste your file link", ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.YES) {
    var docLink = response.getResponseText();
    scriptProperties.setProperty("docLink",docLink); // Save for future.

    } else if (response.getSelectedButton() == ui.Button.NO) {
      Logger.log('No URL provided');
      } else {
        Logger.log('Window was closed.');
      }
}

/**********************************************************************************
This code adds user facing menu that should be used by them to operate this script.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
  {
    name : "Link template",
    functionName : "getTemplate"
  },
  {
    name : "Fill template",
    functionName : "doMerge"
  }];
  spreadsheet.addMenu("Merge", entries);
}
