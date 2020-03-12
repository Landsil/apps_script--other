// Menu options
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : 'Users',
    functionName : 'downloadUsers'
  },
  {
    name : 'Groups',
    functionName : 'downloadGroups'
  },
  {
    name : 'ChromeOS',
    functionName : 'downloadChromeOS'
  }
                ];
  sheet.addMenu('Download', entries);
}


/*******************************************************************************************************************************************
 * Lists all ChromeOS in a G Suite domain.
 * Create a spreedsheet, name one sheer "AUTO_ChromeOS" enable API's as needed.
 * You will need to enable at least Direcory API and admin SDK
 */
// Pulls Device data from G Suite
function downloadChromeOS() {
  var pageToken;
  var page;
  
  // Position in sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var AUTO_ChromeOS = SpreadsheetApp.setActiveSheet(ss.getSheetByName('AUTO_ChromeOS'));
  
  // Clear content except header all the way to "K" column. TODO make it find cells with content and cleare those.
  AUTO_ChromeOS.getRange('A2:Z').clear();
  
  // This decided where to post. Starts after header.
  var lastRow = Math.max(AUTO_ChromeOS.getRange(2, 1).getLastRow(),1);
  var index = 0;
  
  // Run the reqeust
  do {
    page = AdminDirectory.Chromeosdevices.list('my_customer',{
    maxResults: 50,
    projection: 'FULL',
    pageToken: pageToken
  });

//************************
// Assemble Device data
  var params = JSON.stringify(page.chromeosdevices);
  var data = JSON.parse(params);
  
  // Populate sheet
    if (data) {
      for(var i = 0; i < data.length; i++ ){
        
        // Sheet var name, get last lost + previus content, columnt. Set value based on position in JSON
        // Some of the data sits in an array in JSON, you have to specify all steps to get there. Put it in >> (things||"" << to post empty space if there is no data.
        // Full list of suported endpoints: https://developers.google.com/admin-sdk/directory/v1/reference/chromeosdevices

        AUTO_ChromeOS.getRange(index + lastRow + i, 1).setValue(data[i].orgUnitPath);
        var model = (data[i] && data[i].model)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 2).setValue(model);
        var annotatedAssetId = (data[i] && data[i].annotatedAssetId)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 3).setValue(annotatedAssetId);
        var annotatedLocation = (data[i] && data[i].annotatedLocation)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 4).setValue(annotatedLocation);
        var annotatedUser = (data[i] && data[i].annotatedUser)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 5).setValue(annotatedUser);
        var recentUsersEmail_0 = (data[i] && data[i].recentUsers && data[i].recentUsers[0] && data[i].recentUsers[0].email)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 6).setValue(recentUsersEmail_0);
        var recentUsersEmail_1 = (data[i] && data[i].recentUsers && data[i].recentUsers[1] && data[i].recentUsers[1].email)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 7).setValue(recentUsersEmail_1);
        var recentUsersEmail_2 = (data[i] && data[i].recentUsers && data[i].recentUsers[2] && data[i].recentUsers[2].email)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 8).setValue(recentUsersEmail_2);
        var bootMode = (data[i] && data[i].bootMode)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 9).setValue(bootMode);
        var kind = (data[i] && data[i].kind)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 10).setValue(kind);
        var osVersion = (data[i] && data[i].osVersion)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 11).setValue(osVersion);
        var platformVersion = (data[i] && data[i].platformVersion)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 12).setValue(platformVersion);
        var serialNumber = (data[i] && data[i].serialNumber)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 13).setValue(serialNumber);
        var status = (data[i] && data[i].status)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 14).setValue(status);
        var supportEndDate = (data[i] && data[i].supportEndDate)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 15).setValue(supportEndDate);
        var lastSync = (data[i] && data[i].lastSync)||""; AUTO_ChromeOS.getRange(index + lastRow + i, 16).setValue(lastSync);
        
        
        //debug >> Full answer
        // AUTO_ChromeOS.getRange(index + lastRow + i, 10).setValue(params);
      }
      index += 50;
    } else {
      Logger.log('No Devices found.');
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  
// This actually posts data when it's ready.
  AUTO_ChromeOS.sort(1);
SpreadsheetApp.flush();
}
    
    
    
    


/*******************************************************************************************************************************************
 * Lists users in a G Suite domain.
 * Create a spreedsheet, name one sheer "AUTO_users" enable API's as needed.
 * You will need to enable at least Direcory API and admin SDK
 */
 
// Pulls User data from G Suite
function downloadUsers() {
  var pageToken;
  var page;
  
  // Position in sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var AUTO_users = SpreadsheetApp.setActiveSheet(ss.getSheetByName('AUTO_users'));
  
  // Clear content except header all the way to "K" column. TODO make it find cells with content and cleare those.
  AUTO_users.getRange('A2:K').clear();
  
  // This decided where to post. Starts after header.
  var lastRow = Math.max(AUTO_users.getRange(2, 1).getLastRow(),1);
  var index = 0;
  
  // Run the reqeust
  do {
    page = AdminDirectory.Users.list({
    customer: 'my_customer',
    maxResults: 50,
    orderBy: 'email',
    pageToken: pageToken
  });


//************************
// Assemble User's data
  var params = JSON.stringify(page.users);
  var data = JSON.parse(params);
  
  // Populate sheet
    if (data) {
      for(var i = 0; i < data.length; i++ ){
        
        // Sheet var name, get last lost + previus content, columnt. Set value based on position in JSON
        AUTO_users.getRange(index + lastRow + i, 1).setValue(data[i].orgUnitPath);
        AUTO_users.getRange(index + lastRow + i, 2).setValue(data[i].name.fullName);
        AUTO_users.getRange(index + lastRow + i, 3).setValue(data[i].primaryEmail);
        
        // This data sit in an array in JSON, you have to specify all steps to get there. Put it in >> (things||"" << to post empty space if there is no data.
        var title = (data[i] && data[i].organizations && data[i].organizations[0] && data[i].organizations[0].title)||""; AUTO_users.getRange(index + lastRow + i, 4).setValue(title);
        var department = (data[i] && data[i].organizations && data[i].organizations[0] && data[i].organizations[0].department)||""; AUTO_users.getRange(index + lastRow + i, 5).setValue(department);
        var phone = (data[i] && data[i].phones && data[i].phones[0] && data[i].phones[0].value)||""; AUTO_users.getRange(index + lastRow + i, 6).setValue(phone);
        var manager = (data[i] && data[i].relations && data[i].relations[0] && data[i].relations[0].value)||""; AUTO_users.getRange(index + lastRow + i, 7).setValue(manager);
        
        AUTO_users.getRange(index + lastRow + i, 8).setValue(data[i].lastLoginTime);
        
        //debug >> Full answer
        //  AUTO_users.getRange(index + lastRow + i, 10).setValue(params);
      }
      index += 50;
    } else {
      Logger.log('No users found.');
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  
// This actually posts data when it's ready.
  AUTO_users.sort(1);
SpreadsheetApp.flush();
}
  




//*******************************************************************************************************************************************
// Pulls Groups data from G Suite
function downloadGroups() {
  var pageToken;
  var page;
  
  // Position in sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var AUTO_groups = SpreadsheetApp.setActiveSheet(ss.getSheetByName('AUTO_groups'));
  
  // Clear content except header all the way to "K" column. TODO make it find cells with content and cleare those.
  AUTO_groups.getRange('A2:K').clear();
  
  // This decided where to post. Starts after header.
  var lastRow = Math.max(AUTO_groups.getRange(2, 1).getLastRow(),1);
  var index = 0;
  do {
    page = AdminDirectory.Groups.list({
      customer: 'my_customer',
      maxResults: 50,
      pageToken: pageToken
    });

    var groups = page.groups;
    if (groups) {
      for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        AUTO_groups.getRange((index + lastRow + i), 1).setValue(group.name);
        AUTO_groups.getRange((index + lastRow + i), 2).setValue(group.email);
        var aliases = (group.aliases)||""; AUTO_groups.getRange((index + lastRow + i), 3).setValue(aliases);  // TODO fix to show all aliases
        
      }
      index += 50;
    } else {
      Logger.log('No groups found.');
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  
  AUTO_groups.sort(1);
SpreadsheetApp.flush();
}
