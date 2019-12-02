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
  }
                ];
  sheet.addMenu('Download', entries);
};


/**
 * Lists users in a G Suite domain.
 * Create a spreedsheet, name one sheer "AUTO_users" enable API's as needed.
 * You will need at least Direcory API and admin SDK
 */
 
// Pulls data from G Suite
function downloadUsers() {
    var optionalArgs = {
    customer: 'my_customer',
    maxResults: 500,
    orderBy: 'email'
  };
  
  // TODO: add paging per 100 users, https://developers.google.com/apps-script/advanced/admin-sdk-directory

// Assemble data
  var response = AdminDirectory.Users.list(optionalArgs);
  var params = JSON.stringify(response.users);
  var data = JSON.parse(params);
  
    
  // Position in sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var AUTO_users = SpreadsheetApp.setActiveSheet(ss.getSheetByName('AUTO_users'));
  
  // Clear content except header all the way to "K" column. TODO make it find cells with content and cleare those.
  AUTO_users.getRange('A2:K').clear();
  
  // This decided where to post. Starts after header.
  var lastRow = Math.max(AUTO_users.getRange(2, 1).getLastRow(),1);
    
  // Populate sheet
  for(var i = 0; i < data.length; i++ )
  {
    // Sheet var name, get last lost + previus content, columnt. Set value based on position in JSON
    AUTO_users.getRange(lastRow + i, 1).setValue(data[i].orgUnitPath);
    AUTO_users.getRange(lastRow + i, 2).setValue(data[i].name.fullName);
    AUTO_users.getRange(lastRow + i, 3).setValue(data[i].primaryEmail);
    
    // This data sit in an array in JSON, you have to specify all steps to get there. Put it in >> (things||"" << to post empty space if there is no data.
    var title = (data[i] && data[i].organizations && data[i].organizations[0] && data[i].organizations[0].title)||""; AUTO_users.getRange(lastRow + i, 4).setValue(title);
    var department = (data[i] && data[i].organizations && data[i].organizations[0] && data[i].organizations[0].department)||""; AUTO_users.getRange(lastRow + i, 5).setValue(department);
    var phone = (data[i] && data[i].phones && data[i].phones[0] && data[i].phones[0].value)||""; AUTO_users.getRange(lastRow + i, 6).setValue(phone);
    var manager = (data[i] && data[i].relations && data[i].relations[0] && data[i].relations[0].value)||""; AUTO_users.getRange(lastRow + i, 7).setValue(manager);
    
    AUTO_users.getRange(lastRow + i, 8).setValue(data[i].lastLoginTime);
    
    //debug >> Full answer
   //  AUTO_users.getRange(lastRow + i, 10).setValue(params);

  };
  
// This actually posts data when it's ready.
  AUTO_users.sort(1);
SpreadsheetApp.flush();
}

// Pulls data from G Suite
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
  var index = 0
  do {
    page = AdminDirectory.Groups.list({
      customer: 'my_customer',
      maxResults: 50,
      pageToken: pageToken
    });
        Logger.log('----')

    Logger.log(pageToken);
    Logger.log(page.nextPageToken)
        Logger.log('----')

    var groups = page.groups;
    if (groups) {
      for (var i = 0; i < groups.length; i++) {
        var group = groups[i];
        AUTO_groups.getRange((index + lastRow + i), 1).setValue(group.name);
        AUTO_groups.getRange((index + lastRow + i), 2).setValue(group.email);
        
       // Logger.log('%s (%s)', group.name, group.email);
      }
      index += 50
    } else {
      Logger.log('No groups found.');
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  
  AUTO_groups.sort(1);
SpreadsheetApp.flush();
}
