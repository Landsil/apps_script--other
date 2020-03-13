/**
 *This will send SMS using your Twilio account
 * Made based on https://www.twilio.com/blog/2016/02/send-sms-from-a-google-spreadsheet.html
 * 
 */


// Create a menu to be activate funtions
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : 'SMS',
    functionName : 'sendAll'
  },
  {
    name : 'Lookup',
    functionName : 'lookupAll'
  }
                ];
  sheet.addMenu('Twilio_API', entries);
}


// Specify all "Properties" that will hold your tokens.
var scriptProperties = PropertiesService.getScriptProperties();
var messages_url_prop = scriptProperties.getProperty('messages_url_prop');
var from_number_prop = scriptProperties.getProperty('from_number_prop');
var Authorization_prop = scriptProperties.getProperty('Authorization_prop');

// Testing bit for https://www.twilio.com/docs/sms/services#alphanumeric-sender-id
var service_SID_prop = scriptProperties.getProperty('service_SID_prop');  // From https://www.twilio.com/console/sms/services 

// ******************************************************************************************

// This function sends sms. This is the actuall work horse that makes the API call and sends it.
function sendSms(to, body) {
  var messages_url = messages_url_prop;

  var payload = {
    "To": to,
    "Body" : body,
  //  "From" : from_number_prop,   unless you are actually using it
    "MessagingServiceSid" : service_SID_prop   //  Coment out if you aren't using it.
  };

  var options = {
    "method" : "post",
    "payload" : payload
  };

  options.headers = { 
    "Authorization" : "Basic " + Utilities.base64Encode(Authorization_prop)
  };

  UrlFetchApp.fetch(messages_url, options);
}

// This will send 1 SMS per row in "SMS-send" ( that should be a name of your sheet)
// It's just a loop using previous funcion and updating status
function sendAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName('SMS-send'));
  var startRow = 2; 
  var numRows = sheet.getLastRow() - 1; 
  var dataRange = sheet.getRange(startRow, 1, numRows, 2) 
  var data = dataRange.getValues();

  for (i in data) {
    var row = data[i];
    try {
      response_data = sendSms(row[0], row[1]);
      status = "sent";
    } catch(err) {
      Logger.log(err);
      status = err;
    }
    sheet.getRange(startRow + Number(i), 3).setValue(status);
  }
}



// ******************************************************************************************
// Made Based on https://www.twilio.com/blog/2016/03/how-to-look-up-and-verify-phone-numbers-in-google-spreadsheets-with-javascript.html

// Again this is actuall function doing the API call
function lookup(phoneNumber) {
    var lookupUrl = "https://lookups.twilio.com/v1/PhoneNumbers/" + phoneNumber + "?Type=carrier"; 

    var options = {
        "method" : "get"
    };

    options.headers = {    
        "Authorization" : "Basic " + Utilities.base64Encode(Authorization_prop)
    };

    var response = UrlFetchApp.fetch(lookupUrl, options);
    var data = JSON.parse(response); 
    Logger.log(data); 
    return data; 
}

// This fucntion does the processing for all rows of data.
function lookupAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.setActiveSheet(ss.getSheetByName('number_check'));
    var startRow = 2; 
    var numRows = sheet.getLastRow() - 1; 
    var dataRange = sheet.getRange(startRow, 1, numRows); 
    var phoneNumbers = dataRange.getValues();

    for (var i in phoneNumbers) {
        var phoneNumber = phoneNumbers[i]; 
        var spreadsheetRow = startRow + Number(i); 
        sheet.getRange(spreadsheetRow, 2, spreadsheetRow, 6).setValue("");
        if (phoneNumber != "") { 
            try { 
                data = lookup(phoneNumber);
                if (data['status'] == 404) { 
                    sheet.getRange(spreadsheetRow, 2).setValue("not found");
                } else {
                    sheet.getRange(spreadsheetRow, 2).setValue("found");
                    sheet.getRange(spreadsheetRow, 3).setValue(data['carrier']['type']);
                    sheet.getRange(spreadsheetRow, 4).setValue(data['carrier']['name']);
                    sheet.getRange(spreadsheetRow, 5).setValue(data['country_code']);
                    sheet.getRange(spreadsheetRow, 6).setValue(data['national_format']);
                }  
            } catch(err) {
                Logger.log(err);
                sheet.getRange(spreadsheetRow, 2).setValue('lookup error');
            }
        }
    }
}
