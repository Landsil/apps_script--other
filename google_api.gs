// Menu options
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : 'Users',
    functionName : 'downloadUsers'
  }];
  sheet.addMenu('Download', entries);
}


/**
 * Lists users in a G Suite domain. ( this version is valid up to 500 )
 * Create a spreedsheet, name one sheet "AUTO_users" enable API's as needed.
 * You will have to go to: Resources >> Advanced Google Services >> Admin Directory API - enable 
 */
 
// Pulls data from G Suite
function downloadUsers() {
    var optionalArgs = {
    customer: 'my_customer',
    maxResults: 500, //API limit of 500 TODO: update script to pull data thru paging (100 per page)
    orderBy: 'email'
  };

// Assemble data
// Ths will read responce and organise it to redable form.
  var response = AdminDirectory.Users.list(optionalArgs);
  console.log(response);
  var params = JSON.stringify(response.users);
  var data = JSON.parse(params);
  
    
  // Per every user put relevant data in correct column
  // Start by fiding correct sheet ( in this case it's named "AUTO_users" )
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var AUTO_users = SpreadsheetApp.setActiveSheet(ss.getSheetByName('AUTO_users'));
  
  // Clear content except header all the way to "K" column. TODO make it find cells with content and cleare those.
  AUTO_users.getRange('A2:K').clear();
  
  // This marks where to post data. Starts after header.
  var lastRow = Math.max(AUTO_users.getRange(2, 1).getLastRow(),1);
    
  // Populate sheet
  for(var i = 0; i < data.length; i++ )
  {
    // Sheet var name, get last lost + previus content, columnt. Set value based on position in JSON
    AUTO_users.getRange(lastRow + i, 1).setValue(data[i].orgUnitPath);
    AUTO_users.getRange(lastRow + i, 2).setValue(data[i].name.fullName);
    AUTO_users.getRange(lastRow + i, 3).setValue(data[i].primaryEmail);
    
    // This data sits in an array in JSON, you have to specify all steps to get there. Put it in >> (things||"" << to post empty cell if there is no data.
    var title = (data[i] && data[i].organizations && data[i].organizations[0] && data[i].organizations[0].title)||""; AUTO_users.getRange(lastRow + i, 4).setValue(title);
    var department = (data[i] && data[i].organizations && data[i].organizations[0] && data[i].organizations[0].department)||""; AUTO_users.getRange(lastRow + i, 5).setValue(department);
    var phone = (data[i] && data[i].phones && data[i].phones[0] && data[i].phones[0].value)||""; AUTO_users.getRange(lastRow + i, 6).setValue(phone);
    var manager = (data[i] && data[i].relations && data[i].relations[0] && data[i].relations[0].value)||""; AUTO_users.getRange(lastRow + i, 7).setValue(manager);
    
    AUTO_users.getRange(lastRow + i, 8).setValue(data[i].lastLoginTime);
    
    //debug >> Full answer
   //  AUTO_users.getRange(lastRow + i, 10).setValue(params);

  }
  
// This actually posts data when it's ready.
  AUTO_users.sort(1);
SpreadsheetApp.flush();
}
