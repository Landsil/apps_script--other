/* This is a set of functions that will perform acions for you within G Suite
 * They require structurised sheet so running "create_template" command will take care of that
 *
 *
 */

//*************************************************************
// Create Menu in sheet when it's open
var ui = SpreadsheetApp.getUi();
function onOpen() {
  ui.createMenu("G Suite Actions")
  .addItem("Create Drive Template", "create_drive_template")
  .addItem("Check files", "file_check_all")
  .addItem("Transfer all files", "transfer_all")
  .addItem("Give edit", "edit_to_all")
  .addToUi();
};


//***********************************************************************************************************
// Create template sheet that will be used later on.
function create_drive_template() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.insertSheet().setName("AUTO_drive");
  var AUTO_drive = spreadsheet.getSheetByName("AUTO_drive");
  
  // Formating
  AUTO_drive.setFrozenRows(1) // header
  AUTO_drive.getRange("1:1").activate();
  AUTO_drive.getActiveRangeList().setHorizontalAlignment("center").setFontWeight("bold"); // center and bold
  AUTO_drive.getRange(1, 1, AUTO_drive.getMaxRows(), AUTO_drive.getMaxColumns()).activate();
  AUTO_drive.getActiveRangeList().setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);  // Clip when text to long
  AUTO_drive.setColumnWidth(4, 140);
  AUTO_drive.setColumnWidth(7, 120);  // Column size
  AUTO_drive.setColumnWidth(8, 140);
  
  // Content
  AUTO_drive.getRange("A1").activate();
  AUTO_drive.getCurrentCell().setValue("FileID");
  AUTO_drive.getRange("B1").activate();
  AUTO_drive.getCurrentCell().setValue("Transfer to");
  AUTO_drive.getRange("C1").activate();
  AUTO_drive.getCurrentCell().setValue("File Name");
  AUTO_drive.getRange("D1").activate();
  AUTO_drive.getCurrentCell().setValue("Current Owner");
  AUTO_drive.getRange("E1").activate();
  AUTO_drive.getCurrentCell().setValue("Editors");
  AUTO_drive.getRange("F1").activate();
  AUTO_drive.getCurrentCell().setValue("Viewers");
  AUTO_drive.getRange("G1").activate();
  AUTO_drive.getCurrentCell().setValue("Sharing Access");
  AUTO_drive.getRange("H1").activate();
  AUTO_drive.getCurrentCell().setValue("Sharing Permission");
  AUTO_drive.getRange("J1").activate();
  AUTO_drive.getCurrentCell().setValue("Status");
  
  // Test data
  AUTO_drive.getRange("A2").activate();
  AUTO_drive.getCurrentCell().setValue("drive_ID_1_test");
  AUTO_drive.getRange("A3").activate();
  AUTO_drive.getCurrentCell().setValue("drive_ID_2_test");
  AUTO_drive.getRange("A4").activate();
  AUTO_drive.getCurrentCell().setValue("drive_ID_3_test");
  
  
} 


//***********************************************************************************************************
// This function will try to pull data about file
// TODO: Make it work with URL and return file ID while at it.
function file_check_all() {
  // Load sheet and data fileds.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var AUTO_drive = SpreadsheetApp.setActiveSheet(ss.getSheetByName("AUTO_drive"));
  AUTO_drive.getRange('B2:J').clear();
  var sheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("AUTO_drive"));
    var startRow = 2; 
    var numRows = sheet.getLastRow() - 1; 
    var dataRange = sheet.getRange(startRow, 1, numRows); 
    var drive_ID_values = dataRange.getValues();
  
    for ( var i in drive_ID_values) {
      var drive_ID = drive_ID_values[i];
      var spreadsheetRow = startRow + Number(i);
      sheet.getRange(spreadsheetRow, 2, spreadsheetRow, 6).setValue("");
      if (drive_ID != "") {
        try {
          data = pull_file_info(drive_ID);
          if (data["status"] == 404) { 
            sheet.getRange(spreadsheetRow, 10).setValue("404 returned ??");
                } else {
                    sheet.getRange(spreadsheetRow, 3).setValue(data["file_name"]);
                    sheet.getRange(spreadsheetRow, 4).setValue(data["owner"]);
                    sheet.getRange(spreadsheetRow, 5).setValue(data["editors"]);
                    sheet.getRange(spreadsheetRow, 6).setValue(data["viewers"]);
                    sheet.getRange(spreadsheetRow, 7).setValue(data["SharingAccess"]);
                    sheet.getRange(spreadsheetRow, 8).setValue(data["SharingPermission"]);
                  }
            } catch(err) {
                Logger.log(err);
                sheet.getRange(spreadsheetRow, 10).setValue("error / not owned by a user?");
                }
       }
    }
}

//*************************************************************
// This is the function that does the actuall work of calling google to ask
// https://developers.google.com/apps-script/reference/drive/drive-app
// https://developers.google.com/apps-script/reference/drive/file
// TODO: make it work with files in Shared Drives, provide more info?
function pull_file_info(drive_ID) {
  //Find file
  var file = DriveApp.getFileById(drive_ID);
  // Check info about ...
  var owner = file.getOwner().getEmail();
  var editors = file.getEditors();
  var viewers = file.getViewers();
  var SharingAccess = file.getSharingAccess();
  var SharingPermission = file.getSharingPermission();
  var file_name = file.getName();
  
  var editors_list = [];
  for (var i = 0; i < editors.length; i++) {
    editors_list.push(editors[i].getEmail());
  }
  var editors_list = editors_list.toString();
  
  var viewers_list = [];
  for (var i = 0; i < viewers.length; i++) {
    viewers_list.push(viewers[i].getEmail());
  }
  var viewers_list = viewers_list.toString();
  
  // Return dictonary
  var data = {
    "owner": owner,
    "editors": editors_list,
    "viewers": viewers_list,
    "SharingAccess": SharingAccess,
    "SharingPermission": SharingPermission,
    "file_name": file_name
             };
  Logger.log(data);
  
  return data
}

//***********************************************************************************************************
// Change file owner of all the files
function transfer_all() {
  // Load sheet and data fileds.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("AUTO_drive"));
  var startRow = 2; 
  var numRows = sheet.getLastRow() - 1; 
  var dataRange = sheet.getRange(startRow, 1, numRows, 2) 
  var data = dataRange.getValues();

  for (i in data) {
    var row = data[i];
    try {
      response_data = change_owner(row[0], row[1]);
      status = "done";
    } catch(err) {
      Logger.log(err);
      status = err;
    }
    sheet.getRange(startRow + Number(i), 10).setValue(status);
  }
}


//*************************************************************
// This is the function that does the actuall work of calling google to change file owner
// https://developers.google.com/apps-script/reference/drive/file
function change_owner(drive_ID,new_owner){
  //Find file
  var file = DriveApp.getFileById(drive_ID);
  // Change owner
  file.setOwner(new_owner);
}


//***********************************************************************************************************
// Give edit access to all of the files
function edit_to_all() {
  // Load sheet and data fileds.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName("AUTO_drive"));
  var startRow = 2; 
  var numRows = sheet.getLastRow() - 1; 
  var dataRange = sheet.getRange(startRow, 1, numRows, 2) 
  var data = dataRange.getValues();

  for (i in data) {
    var row = data[i];
    try {
      response_data = add_editor(row[0], row[1]);
      status = "done";
    } catch(err) {
      Logger.log(err);
      status = err;
    }
    sheet.getRange(startRow + Number(i), 10).setValue(status);
  }
}


//*************************************************************
// This is the function that does the actuall work of calling google to add person as owner
// https://developers.google.com/apps-script/reference/drive/drive-app
// https://developers.google.com/apps-script/reference/drive/file
function add_editor(drive_ID,new_owner){
  //Find file
  var file = DriveApp.getFileById(drive_ID);
  // Change owner
  file.addEditor(new_owner);
}
