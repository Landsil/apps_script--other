/** This is a script to use with your google sheet, this means you need either gmail or G Suite account.
 * This was tested on G Suite account ( you may need assistance from your system admin )
 * To work this code requires you to deploy it as public app with full access to anonymous accounts
 * Publish > Deploy as web app >
 * Remember to re-publish your app after making changes
 * "Current web app URL:" === Link to use as webhook
 * eg. link: https://script.google.com/macros/s/i34etgh8iydufg8iheysgfhdieyr/exec
 */
 
//this is a Function that starts when the webapp receives a POST request from external app
// Read data here
function onPOST(e) {
  var params = JSON.stringify(e.postData.contents);
  params = JSON.parse(params);
  var myData = JSON.parse(e.postData.contents);

// Assign JSON data to variables
  var id = myData.id;
  var actor_type = myData.actor_type;
  var action = myData.action;
  var object_type = myData.object_type;
  var success = myData.success;
  var message = myData.message;
  var created_at = myData.created_at; 

// Load active sheet ( TODO: change to load named sheet ) , get last empty (none 1st) row and post data there.
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = Math.max(sheet.getLastRow(),1);
  sheet.insertRowAfter(lastRow);
  var timestamp = new Date();
  
// Here you specify to what columnt will specific content go
  sheet.getRange(lastRow + 1, 1).setValue(timestamp);   // << This is not taken from JSON
  sheet.getRange(lastRow + 1, 2).setValue(created_at);
  sheet.getRange(lastRow + 1, 3).setValue(id);
  sheet.getRange(lastRow + 1, 4).setValue(actor_type);
  sheet.getRange(lastRow + 1, 5).setValue(action);
  sheet.getRange(lastRow + 1, 6).setValue(success);
  sheet.getRange(lastRow + 1, 7).setValue(object_type);
  sheet.getRange(lastRow + 1, 8).setValue(message);  
  
// This is your debug/help. It will post whole JSON into this cell
  sheet.getRange(lastRow + 1, 9).setValue(params);
  
  SpreadsheetApp.flush();   // This will aply all changes in consistent maner.
  return HtmlService.createHtmlOutput("post request received");
}


// Obwious down-side of this script is that at some point sheet will be too big to load in a resonable manner.
