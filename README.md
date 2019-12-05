# apps script
Those are relatively simple ready to use scripts that everyone can use just by copy pasting to google sheets apps script.
Most will have full comments explaining code blocks and prerequisites.

Most assume basic knowledge about Apps Script
[Link to example video on how to enable some things.](https://www.youtube.com/watch?v=tJ4_w2596KI)

#### [Custom Functions for Google Sheets](https://github.com/Landsil/apps_script/blob/master/custom_functions.gs)
 - Adds custom menu to sheets `OnOpen`
 - Add different SHA function with instructions and description
 - Indentation function

#### [Web app that will receive POST, ingest JSON and post data to sheet.](https://github.com/Landsil/apps_script/blob/master/ingest_JSON_post.gs)
 - Public Web App
 - Receive Webhook stream
 - Simple JSON manipulation
 - Find first empty row and add data to correct cells

#### [Different G Suite things with google API](https://github.com/Landsil/apps_script/blob/master/google_api.gs)
 - Adds menu `onOpen`
##### Get user data thru Admin directory API
 - Get user data thru admin directory API
 - Read JSON
 - Look thru data and post to sheets in rows
##### Get list of groups thru directory API
 - Get ulist of groups thru admin directory API
 - Loop thru pages
 - Look thru data and post to sheets in rows

#### [Pull employee data from peopleHR](https://github.com/Landsil/apps_script/blob/master/download_PeopleHR.gs)
 - Adds menu `onOpen`
 - `PropertiesService` to access token there
 - Parse JSON
 - Loop thru data and post in rows
