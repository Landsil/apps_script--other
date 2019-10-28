// Add menu to sheet
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : 'Users',
    functionName : 'downloadUsers'
  }];
  sheet.addMenu('Download', entries);
}


/**
 * Lists users in a G Suite domain.
 * Currently BROKEN
 * When fixed will work over a limit of 500
 */
 
// Pulls user data from G Suite and posts to "AUTO_users" sheet
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
  
  // Actuall Action
  do {
    page = AdminDirectory.Users.list({
      customer: 'my_customer',
      orderBy: 'email',
      maxResults: 100,
      pageToken: pageToken
    });
    
    var params = JSON.stringify(page);
    var data = JSON.parse(params);
    
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
        
        //debug >> Full answer
        //  AUTO_users.getRange(lastRow + i, 10).setValue(params);
        
      }

    pageToken = page.nextPageToken;
  } while (pageToken);
  // This actually posts data when it's ready.
SpreadsheetApp.flush();
}
