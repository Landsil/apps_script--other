function getSignature() {
  //pretty basic function for testing
  if ( startupChecks()) { return; }
  var email = SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue().toString();
  if ( email === "" ) {
    Browser.msgBox("No email selected", "Please select a cell containing a user's email" , Browser.Buttons.OK);
    return;
  }
  var result = authorisedUrlFetch(email, {});
  Browser.msgBox(result.getContentText());
}

function setIndividualSignature() {
  Logger.log('[%s]\t Starting setIndividualSignature run', Date());
  if ( startupChecks()) { return; }
  var userData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Summary');
  var template = getTemplate();
  
  var row = SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getRow();
  if (userData.getRange(row, 1).isBlank() === true) {
    var msg = "Please select a cell on a row containing the user who's signature you wish to update";
    Browser.msgBox('No valid user selected', msg, Browser.Buttons.OK);
  } else {
    setSignature(template, userData, row);
  } 
  Logger.log('[%s]\t Completed setIndividualSignature run', Date());
}

function setAllSignatures() {
  Logger.log('[%s]\t Starting setAllSignatures run', Date());
  if ( startupChecks()) { return; }
  var userData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Summary');

  var template = getTemplate();
  
  //Go through each user listing
  for ( row = 2; row <= userData.getLastRow() ; row++) { 
    setSignature(template, userData, row);
  }
  Logger.log('[%s]\t Completed setAllSignatures run', Date());
}

function getTemplate(){
  Logger.log('[%s]\t Getting Template', Date());
  var settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Signature Settings');
  var template = settings.getRange(2, 1).getValue().toString();
  
  //Substitute the company wide variables into the template 
  template = substituteVariablesFromRow(template, settings, 2);

  return template;
}

function setSignature(template, userData, row){
    var groupData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Signature Group Settings'); 
    
    //Google Apps Scripts always deals in ranges even if you just want one cell
    //getValue returns an object, so convert it to a string
    var email = userData.getRange(row, 1).getValue().toString();
    
    //quick exit if the user isn't in the domain
    if (!checkUserIsValid(email)){ 
      Logger.log('[%s]\t Skipping user %s',Date(),email);
      return;
    }
  
    //Substitute in group variables, e.g those for groups of users
    //this must be done before filling out user specific data as it was added after initial design
    Logger.log('[%s]\t Substituting Group Variables for user %s',Date(),email);
    var signature = substituteGroupVariables(template, userData, groupData, row);
    
    //Fill out the template with the data from the user's row to form the signatures
    Logger.log('[%s]\t Substituting Individual Variables for user %s',Date(),email);
    signature = substituteVariablesFromRow(signature, userData, row);
    
    //The API docs say there is a 10,000 character limit 
    //https://developers.google.com/google-apps/email-settings/#updating_a_signature
    if (signature.length > 10000) { Browser.msgBox('signature over 10000 characters for:' + email); }
    
    Logger.log('[%s]\t Sending signature for user %s',Date(),email);
    sendSignature(email, signature);
    Logger.log('[%s]\t Processing complete for user %s',Date(),email);
}

function substituteVariablesFromRow(text, sheet, row) {
  //Generating two lists avoids the need to do lots of individual calls to the sheet
  var tags = sheet.getSheetValues(1, 1, 1, sheet.getLastColumn())[0];
  var values = sheet.getSheetValues(row, 1, 1, sheet.getLastColumn())[0];
  for ( v = 0; v < values.length; v++){
    text = tagReplace(tags[v],values[v],text);
  }  
  return text;
}

function substituteGroupVariables(text, dataSheet, lookupSheet, row) {
  //this function is still not great but at least it makes use of getSheet
  var tags = dataSheet.getSheetValues(1, 1, 1, dataSheet.getLastColumn())[0];
  var values = dataSheet.getSheetValues(row, 1, 1, dataSheet.getLastColumn())[0];
  var GroupVariables = lookupSheet.getSheetValues(1, 1, lookupSheet.getLastRow(),1);
  
  //for each GroupVariable
  for (j = 0; j < GroupVariables.length ; j += 3){
    
    //find the column for later changing the value
    for (i = 0; i < tags.length; i++){
      if (tags[i] === GroupVariables[j][0]){
        
        //and build a lookup table to switch it out
        var lookupTable = lookupSheet.getSheetValues(j+1,2,2,lookupSheet.getLastColumn()-1);
        for ( k=0;k<lookupTable[0].length;k++) {
          if (values[i] === lookupTable[0][k]){
            text = tagReplace(tags[i], lookupTable[1][k], text);
          }
        }
        
      }
    }
  }
  
  return text;
}

function sanitize(text){
  var invalid = ["[","^","$",".","|","?","*","+","(",")"];
  for(m=0;m<invalid.length;m++){
      text = text.replace(invalid[m],"\\"+invalid[m]);
  }
  return text;
}


function tagReplace(tag, value, text){
  var regOpen = sanitize(UserProperties.getProperty('regOpen'));
  var tagOpen = sanitize(UserProperties.getProperty('tagOpen'));
  var regClose = sanitize(UserProperties.getProperty('regClose'));
  var tagClose = sanitize(UserProperties.getProperty('tagClose'));
  
  
  var regex = new RegExp("(.*)"+regOpen+'(.*?)'+tagOpen+tag+tagClose+'(.*?)'+regClose+"(.*)","g");
  value = value.toString().replace("$","\\$");
  if ((value !== "")) { value = "$2"+value+"$3"; }
  value = "$1"+value+"$4";
  
  //I'm sure this can be avoided by making the regex more complicated, but this will do for now
  for(q=0; ((text.match(regex)) && q<128); q++ ){
    text = text.replace(regex,value);
  }
  
  return text;
}

function sendSignature(email, signature) {
  // https://developers.google.com/google-apps/email-settings/#updating_a_signature
  var requestData = {
    'method': 'PUT',
    'contentType': 'application/atom+xml',
    'payload': getPayload(signature)
  };
  var result = authorisedUrlFetch(email, requestData);
  if (result.getResponseCode() != 200) {
    var msg = 'There was an error sending ' + email + "'s signature to Google";
    Browser.msgBox('Error settings signature', msg, Browser.Buttons.OK);
  }
}

function checkUserIsValid(user){
  var userList = UserManager.getAllUsers();
  for ( u=0 ; u < userList.length ; u++ ) {
    if (userList[u].getEmail() === user){ return true; }
  }
  return false;
}

function getPayload(signature) {
  //First line is needed for XML, second isn't but we might as well do it for consistency
  signature = signature.replace(/&/g, '&amp;').replace(/</g, '&lt;');
  signature = signature.replace(/>/g, '&gt;').replace(/'/g, '&apos;').replace(/"/g, '&quot;');
  
  //Unfortunately when inside app script document.createElement doesn't work so lets just hardcode the XML for now
  var xml = '<?xml version="1.0" encoding="utf-8"?>' +
    '<atom:entry xmlns:atom="http://www.w3.org/2005/Atom" xmlns:apps="http://schemas.google.com/apps/2006" >' +
    '<apps:property name="signature" value="'+signature+'" /></atom:entry>';
  return xml;
}

function authorisedUrlFetch(email, requestData) {
  //takes request data and wraps oauth authentication around it before sending out the request
  // https://developers.google.com/apps-script/class_oauthconfig
  // http://support.google.com/a/bin/answer.py?hl=en&hlrm=en&answer=162105
  // https://developers.google.com/apps-script/articles/picasa_google_apis#section2a
  // The scope from https://developers.google.com/google-apps/email-settings/ has to be URIcomponent encoded
  
  var oAuthConfig = UrlFetchApp.addOAuthService('google');
  oAuthConfig.setConsumerSecret(UserProperties.getProperty('oAuthConsumerSecret')); 
  oAuthConfig.setConsumerKey(UserProperties.getProperty('oAuthClientID'));
  oAuthConfig.setRequestTokenUrl('https://www.google.com/accounts/OAuthGetRequestToken?scope=https%3A%2F%2Fapps-apis.google.com%2Fa%2Ffeeds%2Femailsettings%2F');
  oAuthConfig.setAuthorizationUrl('https://www.google.com/accounts/OAuthAuthorizeToken');
  oAuthConfig.setAccessTokenUrl('https://www.google.com/accounts/OAuthGetAccessToken');
  UrlFetchApp.addOAuthService(oAuthConfig);
  
  requestData['oAuthServiceName'] = 'google';
  requestData['oAuthUseToken'] = 'always';
  
  var emailParts = email.split('@');
  var url = 'https://apps-apis.google.com/a/feeds/emailsettings/2.0/' + emailParts[1] + '/' + emailParts[0] + '/signature';
  var result = UrlFetchApp.fetch(url, requestData);
  if ( result.getResponseCode() != 200 ) {
    //Do some logging if something goes wrong
    //Too deep to give the user a meaningful error though so pass the result back up anyway
    Logger.log('Error on fetch on' + url);
    Logger.log(requestData);
    Logger.log(result.getResponseCode());
    Logger.log(result.getHeaders());
    Logger.log(result.getContentText());
  }
  return result;
}

function onOpen() {
  //add a toolbar and list the functions you want to call externally
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({name: 'Set All Signatures', functionName: 'setAllSignatures'});
  menuEntries.push({name: 'Set Individual Signature', functionName: 'setIndividualSignature'});
  menuEntries.push({name: 'Get Signature', functionName: 'getSignature'});
  ss.addMenu('Signatures', menuEntries);
}

function startupChecks() {
  //Check that everything that is needed to run is there
  //I don't check that any of it makes sense, just that it exists.
  var requiredProperties = [];
  
  //the help text looks pretty terrible but it is better than nothing
  var oAuthHelp = 'Goto https://code.google.com/apis/console#:access and register as an "Installed application" \n'+
                  'Then add the ClientID to authorised 3rd party clients \n'+
                  'With scope https://apps-apis.google.com/a/feeds/emailsettings/ \n'+
                  'The script may then need authorising, this can be done by running one of the scripts from the script editor';
  requiredProperties.push({name: 'oAuthClientID', help: oAuthHelp});
  requiredProperties.push({name: 'oAuthConsumerSecret', help: oAuthHelp});
  requiredProperties.push({name: 'regOpen', help: 'A character or sequence to go before sections to be substituded, e.g ${'});
  requiredProperties.push({name: 'regClose', help: 'A character or sequence to go after sections that will be substituted, e.g } or }$'});
  requiredProperties.push({name: 'tagOpen', help: 'A character or sequence to go before tags to be substituded, e.g {'});
  requiredProperties.push({name: 'tagClose', help: 'A character or sequence to go after tags that will be substituted, e.g } or }$'});
  
  var requiredSheets = [];
  requiredSheets.push({name: 'Summary', help: 'A "Summary" sheet must exist that contains a 1 header row and 1 row per user, with no gaps in either the 1st column or row, the 1st row must be the users usernames'});
  requiredSheets.push({name: 'Signature Settings', help: 'A "Signature Settings" sheet must exist that contains a the template in cell 2A and then has 1 header row and 1 row per company wide variable, with no empty header cells'});
  requiredSheets.push({name: 'Signature Group Settings', help: 'A "Signature Group Settings" sheet must exist that contains 3 Rows (setting values, what to substitute, comments) with every third row containing a column header'});
  
  var fail = false; 
  for ( s = 0; s < requiredProperties.length; s++) {
    var property = UserProperties.getProperty(requiredProperties[s].name);
    if (property == null) {
      var title = 'Script Property ' + requiredProperties[s].name + ' is required';
      var prompt = requiredProperties[s].help;
      var newValue = Browser.inputBox(title, prompt, Browser.Buttons.OK_CANCEL);
      if ((newValue === '') || (newValue === 'cancel')) {
        fail = true;
      } else {
        UserProperties.setProperty(requiredProperties[s].name, newValue);
      } 
    }
  } 

  for ( s = 0; s < requiredSheets.length; s++) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(requiredSheets[s].name);
    if (sheet == null) {
      fail = true;
      var title = 'Sheet ' + requiredSheets[s].name + ' is required';
      var prompt = requiredSheets[s].help;
      Browser.msgBox(title, prompt, Browser.Buttons.OK);
    }
  }
  
  return fail;    
}
