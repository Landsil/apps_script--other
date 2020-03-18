# apps script
Those are relatively simple ready to use scripts that everyone can use just by copy pasting to google sheets apps script.
Most will have full comments explaining code blocks and prerequisites.

Most assume basic knowledge about Apps Script

[Link to example video on how to enable some things.](https://www.youtube.com/watch?v=tJ4_w2596KI)

#### 1. [Custom Functions for Google Sheets](https://github.com/Landsil/apps_script/blob/master/custom_functions.gs)
 - Adds custom menu to sheets `OnOpen`
 - Add different SHA function with instructions and description
 - Indentation function

#### 2. [Web app that will receive POST, ingest JSON and post data to sheet.](https://github.com/Landsil/apps_script/blob/master/ingest_JSON_post.gs)
 - Public Web App
 - Receive Webhook stream
 - Simple JSON manipulation
 - Find first empty row and add data to correct cells

#### 3. [Different G Suite things with google API](https://github.com/Landsil/apps_script/blob/master/google_api.gs)
 - Adds menu `onOpen`
 
##### Get user data thru Admin directory API
 - Get user data thru admin directory API
 - Read JSON
 - Look thru data and post to sheets in rows
##### Get ChromeOS data thru Admin directory API
 - Get device data thru admin directory API
 - Read JSON
 - Look thru data and post to sheets in rows
##### Get list of groups thru directory API
 - Get list of groups thru admin directory API
 - Loop thru pages
 - Look thru data and post to sheets in rows

#### 4. [Pull employee data from peopleHR](https://github.com/Landsil/apps_script/blob/master/download_PeopleHR.gs)
 - Adds menu `onOpen`
 - `PropertiesService` to access token there
 - Parse JSON
 - Loop thru data and post in rows
 
 #### 5. [Send SMS for sheet using Twilio](https://github.com/Landsil/apps_script/blob/master/twilio_api.gs)
Mostly updated code that Twilio is showing on their page.
Uses `PropertiesService` to access tokens there instad of puting them in code.
