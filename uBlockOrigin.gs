// Related to
// https://github.com/gorhill/uBlock/wiki/Deploying-uBlock-Origin
// http://raymondhill.net/ublock/adminSetting.html


/* 1. Save the script
 * 2. Refresh Sheet
 * 3. Use new manu called "uBlock" --> Create Template
 * 4. You will be asked to give access to script
 * 5. In the template sheet add your website to column "A" under "Websites" and use uBlock --> Generate JSON
 */ 


// Create Menu in sheet when it's open
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
    name : "Create Template",
    functionName : "create_template"
    },
    {
    name : "Generate JSON",
    functionName : "make_JSON"
    }

  ];
  sheet.addMenu('uBlock', entries);
}


// Create template sheet that will be used later on for everything.
function create_template() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//  var yourNewSheet = activeSpreadsheet.getSheetByName("AUTO_data");

  var yourNewSheet = activeSpreadsheet.insertSheet();
  yourNewSheet.setName("AUTO_data");
  var AUTO_data = activeSpreadsheet.getSheetByName("AUTO_data");
  
  // Create header
  AUTO_data.setFrozenRows(1)
  // Bold and center header
  AUTO_data.getRange("1:1").activate();
  AUTO_data.getActiveRangeList().setHorizontalAlignment("center").setFontWeight("bold");
  // Content
  AUTO_data.getRange("A1").activate();
  AUTO_data.getCurrentCell().setValue("Websites");
  
  AUTO_data.getRange("A2").activate();
  AUTO_data.getCurrentCell().setValue("page1");
  AUTO_data.getRange("A3").activate();
  AUTO_data.getCurrentCell().setValue("page2");
  AUTO_data.getRange("A4").activate();
  AUTO_data.getCurrentCell().setValue("page3");
  
  AUTO_data.getRange("C2").activate();
  AUTO_data.getCurrentCell().setValue("File content -->");
  AUTO_data.getRange("C3").activate();
  AUTO_data.getCurrentCell().setValue("Create a txt file with this string and use in G Suite");
  AUTO_data.getRange("C4").activate();
  AUTO_data.getCurrentCell().setValue("https://admin.google.com/ac/chrome/apps/user?org=01czz19p1dzbqrd&f=ID.cjpalhdlnbpafiamejdnhcphjbkeiagm");
}


//************************************************************************************************************************************
// Take list of websites and asseble them into correct format
function make_JSON() {
  // Read data in column
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var AUTO_data = activeSpreadsheet.getSheetByName("AUTO_data");
  
  // Read data from "A2:A" and flatten from 2D array to 1D array
  var values =  AUTO_data.getRange("A2:A"+AUTO_data.getLastRow()).getValues().flat();
  Logger.log("values")
  Logger.log(values)
  
  var whitelist_values = values
  Logger.log("")
  Logger.log("whitelist_values")
  Logger.log(whitelist_values)
  
  var netWhitelist_values = values.join("\n");
  Logger.log("")
  Logger.log("netWhitelist_values")
  Logger.log(netWhitelist_values)
  
  
  var input = {"adminSettings":
                { "Value":
                  {"whitelist": whitelist_values,
               "netWhitelist": netWhitelist_values
                  }
                }
              }
  Logger.log("")
  Logger.log("input")
  Logger.log(input)

  var output = JSON.stringify(JSON.stringify(input));
  
  AUTO_data.getRange("D2").activate();
  AUTO_data.getCurrentCell().setValue(output);
  
}
