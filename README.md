# Spreadsheet Scripts (for Google Sheets & Microsoft Excel)

Welcome to the Spreadsheet Scripts in collaboration with <a href="https://developer.aircall.io/api-references/" target="_blank">Aircall</a>. To integrate Aircall functionalities with Google Sheets or Microsoft Excel, we offer the following functions and pages.
<br><br>

## Introduction</a>

Aircall is a cloud based telephony provider with a Softphone that can be used by employees on any device and from any location. This means that employees:
<br>👉 want to call but not necessarily have access to all the data to make the call
<br>👉 want to review the calls they made
<br>👉 want to help out other teams in handling inbound call volume

For these use cases, please see all the available Google App & Microsoft Excel Scripts and automate processes with Aircall from within a spreadsheet.
<br>

## Google Sheets

### <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Google%20Sheets/baseProperties.js" target="_blank">Base Properties</a>

We recommend you to work with <a href="https://developers.google.com/apps-script/guides/properties" target="_blank">Google's script properties</a> to provide a bit of security in the scripts. For the Aircall connection, a set of <a href="https://support.aircall.io/hc/en-gb/articles/10375354348829-Integrations-and-API" target="_blank">API ID and Secret</a> needs to created in the Aircall Dashboard.

```javascript
const scriptProperties = PropertiesService.getScriptProperties();
const apiId = scriptProperties.getProperty("apiId");
const apiToken = scriptProperties.getProperty("apiToken");
```

Additionally, you can add active or particular sheets to the base properties to reference in the other functions.
<br>

### <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Google%20Sheets/supportFunctions.js" target="_blank">Support Functions</a>

The support functions incorporate:
<br>👉 generic functions for CRUD actions
<br>👉 user functions to start CRUD actions with the <a href="https://developer.aircall.io/api-references/#user" target="_blank">Aircall Users API</a>
<br>👉 call functions to get and edit call data via the <a href="https://developer.aircall.io/api-references/#call" target="_blank">Aircall Calls API</a>
<br>👉 tag functions for CRUD actions with the <a href="https://developer.aircall.io/api-references/#tag" target="_blank">Aircall </a>Tags API
<br>👉 contact functions for CRUD actions with the <a href="https://developer.aircall.io/api-references/#contact" target="_blank">Aircall Contacts API</a>
<br>👉 example of a <a href="https://developers.google.com/apps-script/guides/menus" target="_blank">Google custom menu</a> and <a href="https://developers.google.com/apps-script/guides/dialogs" target="_blank">Google sidebar</a> functions
<br>

### <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Google%20Sheets/softphone.html" target="_blank">HTML Sidebars</a>

Google Sheets offers the ability for small <a href="https://developers.google.com/apps-script/guides/dialogs" target="_blank">HTML sidebars</a> that can interact with the spreadsheet. This can be used as HTML form to control import and export of data but also to embed the Aircall Softphone like:

```html
<iframe class="softphone" allow="microphone; autoplay; clipboard-read; clipboard-write; hid" src="https://phone.aircall.io?integration=generic"> </iframe>
```

The sidebars can be opened and closed with functions such as:

```javascript
function showAircallSoftphone() {
  const aircallSoftphone = HtmlService.createHtmlOutputFromFile("Sidebar");
  aircallSoftphone.setTitle("Aircall Phone");
  SpreadsheetApp.getUi().showSidebar(aircallSoftphone);
}
```

That function can be connected to a custom menu to allow the user to open the sidebar on click but it is also possible to add it to the `onOpen()` function to open up immediately.
<br><br>

## <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Google%20Sheets/clickToCallDial.js" target="_blank">Click To Dial / Call</a>

The functions for click to dial and click to call work with the following Aircall APIs:
<br>👉 dial the phone number in the Softphone using the <a href="https://developer.aircall.io/api-references/#dial-a-phone-number-in-the-phone" target="_blank">Aircall Dial a Number API</a>
<br>👉 call the phone number with the Softphone using the <a href="https://developer.aircall.io/api-references/#start-an-outbound-call" target="_blank">Aircall Start an outbound Call API</a>
<br><br>

## <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Google%20Sheets/dialerCampaign.js" target="_blank">Dialer Campaigns</a>

To load a list of phone numbers to call one after the other, a Dialer Campaign is availabe with the Aircall Softphone. The functions will take a selection of Google Sheet cells and add all numbers in that order into an Aircall Dialer Campaign.
To start the copy over into a Dialer Campaign, an example of a custom menu is added:

```javascript
function onOpen() {
  ui.createMenu("🚀 Aircall 🚀").addItem("Aircall Softphone", "showAircallSoftphone").addSeparator().addItem("Create Dialer Campaign", "createDialerCampaign").addSubMenu(ui.createMenu("Delegate Dialer Campaign").addItem("Name1", "delegateDialerCampaignName1").addItem("Name2", "delegateDialerCampaignName2").addItem("Name3", "delegateDialerCampaignName3")).addToUi();
}
```

A button or cell change to trigger the creation of the Dialer Campaign is also possible.
To delegate Dialer Campaigns to other team members, it is needed to set up a small dictionary with relevent Aircall information:

```javascript
const userDict = [
  {
    name: 'name 1',
    email: 'email 1',
    aircall_id: 'id 1',
    number_id: 'number id 1'
  },
...
];
```

<br>

## <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Google%20Sheets/userImporter.js" target="_blank">User Importer</a>

When many users need to be created, this script will allow you to:
<br>👉 make a CSV file
<br>👉 edit anything in bulk
<br>👉 create all the users via the <a href="https://developer.aircall.io/api-references/#user" target="_blank">Aircall Users API</a>
Please review the code that configure the settings on each user.

```javascript
const payloadBody = {
  email: record["email"],
  first_name: record["firstName"],
  last_name: record["lastName"],
  is_admin: false,
  role_ids: ["agent"],
};
```

Additional settings such as `role_ids`, `is_admin` and `availability_status` can be added to the Google Sheet as columns to further customise the creation of each user.
<br><br>

## <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Google%20Sheets/userImporter.js" target="_blank">Contact Importer</a>

When many contacts need to be created, this script will allow you to:
<br>👉 make/import a CSV file
<br>👉 check all contacts in bulk
<br>👉 create all the contacts via the <a href="https://developer.aircall.io/api-references/#contact" target="_blank">Aircall Contacts API</a>
Please review the code that configure the settings on each user.

```javascript
let payloadBody = {
  phone_numbers: [record["phone1"], record["phone2"]],
};
if (record["email"] != null) payloadBody.email = record["email"];
if (record["firstName"] != null) payloadBody.first_name = record["firstName"];
if (record["lastName"] != null) payloadBody.last_name = record["lastName"];
if (record["company"] != null) payloadBody.company_name = record["company"];
```

Additional settings such as `role_ids`, `is_admin` and `availability_status` can be added to the Google Sheet as columns to further customise the creation of each user.
<br><br>

## <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Google%20Sheets/callImporter.js" target="_blank">Call Importer</a>

Reporting on calls with Aircall can be done from within the Aircall Dashboard but if you want more flexibility on reports, you can download the Aircall Call data via the <a href="https://developer.aircall.io/api-references/#search-calls" target="_blank">Aircall Search Calls API</a>.
The call data will follow the data structure of the Aircall call data and additional columns that are derived from the Aircall data can be added to the spreadsheet like:

```javascript
// additional Google Sheet column headerss
printCols.push(["dateTime"]);
printCols.push(["date"]);
printCols.push(["week"]);
printCols.push(["month"]);
printCols.push(["talkingTime"]);
printCols.push(["waitingTime"]);
printCols.push(["number"]);
printCols.push(["numberCountryCode"]);
printCols.push(["tags"]);
printCols.push(["agent"]);
printCols.push(["agentEmail"]);
printCols.push(["agentNumber"]);
printCols.push(["team"]);
// Additional Google Sheet formulas or specific data for the additional columns
recordRow.push(["=EPOCHTODATE(F:F)"]);
recordRow.push(["=LEFT(AA:AA;10)"]);
recordRow.push(["=WEEKNUM(AB:AB)"]);
recordRow.push(["=MONTH(AB:AB)"]);
recordRow.push(['=IF(D:D<>"done";0;IF(G:G<>"";H:H-G:G;0))']);
recordRow.push(['=IF(D:D<>"done";0;IF(G:G<>"";G:G-F:F;I:I))']);
recordRow.push(objectRecords[r]["number"]["name"]);
recordRow.push(objectRecords[r]["number"] != null && objectRecords[r]["number"]["country"] != null ? objectRecords[r]["number"]["country"] : "");
let tags = "";
if (objectRecords[r]["tags"].length != 0) {
  for (let t = 0; t < objectRecords[r]["tags"].length; t++) {
    if (t === 0) tags = objectRecords[r]["tags"][t]["name"];
    else tags = tags + ", " + objectRecords[r]["tags"][t]["name"];
  }
}
recordRow.push([tags]);
let agentName = "";
let agentEmail = "";
if (objectRecords[r]["user"] != null) {
  // console.log(objectRecords[r]["user"]);
  agentEmail = objectRecords[r]["user"]["email"];
  agentName = objectRecords[r]["user"]["name"];
}
recordRow.push([agentName]);
recordRow.push([agentEmail]);
recordRow.push(objectRecords[r]["number"] != null && objectRecords[r]["number"]["digits"] != null ? objectRecords[r]["number"]["digits"] : "");
let teamName = "";
if (objectRecords[r]["teams"].length > 0) {
  // console.log(objectRecords[r]["teams"]);
  teamName = objectRecords[r]["teams"][0]["name"];
}
recordRow.push([teamName]);
```

When configuration about the column headers and the row data is complete, <a href="https://developers.google.com/apps-script/guides/triggers/installable" target="_blank" >Google Apps triggers</a> can be added to download the Aircall data. With a limit of 10000 calls per API request, it is important to GET the Aircall call data regurlarly. It is recommend to get the call data once a day, the script has an example to download the data each hour and clear the data each Monday at 00:00 hours again:

```javascript
/* get interval data */
let dateTimeNowInSeconds = Math.floor(Date.now() / 1000);
const dateTimeNowMinus1DayInSeconds = dateTimeNowInSeconds - 60 * 60;
// Logger.log('Now: '+dateTimeNowInSeconds+' minus 1 hour: '+dateTimeNowMinus1HourInSeconds);
const dateTimeNowFormat = new Date(dateTimeNowInSeconds);
const dateTimeNowDay = dateTimeNowFormat.getDay();
const dateTimeNowHour = dateTimeNowFormat.getHours();
if (dateTimeNowDay == 1 && dateTimeNowHour == 0) getAll("calls", dateTimeNowMinus1DayInSeconds, dateTimeNowInSeconds, true);
else getAll("calls", dateTimeNowMinus1DayInSeconds, dateTimeNowInSeconds, false);
```

This function will set the `FROM` and `TO` in the Aircall Search Calls API to download call data from the last hour then combined with an hourly Google App Scripts trigger. This is the most frequent time trigger that can be set up and will make the Google Sheet maximum 1 hour delayed.
<br><br>

## <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Google%20Sheets/teamManagement.js" target="_blank">Team Management</a>

Aircall allows supervisors and admins to change the agents within the ringing teams. If each agent can easily change whether to be part of ringing groups without opening up all the permissions in Aircall, a Google Sheet with team management can be created.

The Google Sheet will allow agents to edit whether they are part of a team with a simple cell change. The Google Sheet will execute this change in Aircall using:
<br>👉 add a users via the <a href="https://developer.aircall.io/api-references/#add-a-user-to-a-team" target="_blank">Aircall Add a User to a Team API</a>
<br>👉 remove a user via the <a href="https://developer.aircall.io/api-references/#remove-a-user-from-a-team" target="_blank">Aircall Remove a User from a Team API</a>
For this is needed to make a list of users and teams so that the Aircall API knows which Aircall User ID needs to be removed from which Aircall Team ID.
Additionally, it is needed to have a table overview with users and teams on individual axes and have a specific cell value to use as trigger.
Google Apps Scripts allows to use `onEdit()` function but it has a limitation: it cant execute Fetch API requests in that function. Because it is needed to send the data to the Aircall API, a different function name is selected and a Google Apps Script trigger based upon a change in Spreadsheet.
In this particular script, the cell value needs to be `Logged In` or `Logged Out` to make the individual agent be added or removed from the team. With a cell value changing to `Inactive` all agents in the team are removed.

## Microsoft Excel

## <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/tree/baseline/Microsoft%20Excel" target="_blank">Team Management</a>

Aircall allows supervisors and admins to change the agents within the ringing teams. If each agent can easily change whether to be part of ringing groups without opening up all the permissions in Aircall, a Excel Workbook with team management can be created.

The Excel Workbook will allow agents to edit whether they are part of a team with a simple cell change. The Excel Workbook will execute this change in Aircall using:
<br>👉 add a users via the <a href="https://developer.aircall.io/api-references/#add-a-user-to-a-team" target="_blank">Aircall Add a User to a Team API</a>
<br>👉 remove a user via the <a href="https://developer.aircall.io/api-references/#remove-a-user-from-a-team" target="_blank">Aircall Remove a User from a Team API</a>
For this is needed to make a list of users and teams so that the Aircall API knows which Aircall User ID needs to be removed from which Aircall Team ID.
Additionally, it is needed to have a table overview with users and teams on individual axes and have a specific cell value to use as trigger.

The steps to create this Team Management Worksheet are as follows:
<br>👉 <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Microsoft%20Excel/syncUsers.ts" target="_blank">sync all Aircall users</a> via the <a href="https://developer.aircall.io/api-references/#list-all-users" target="_blank">Aircall List all Users API</a> to a worksheet
<br>👉 <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Microsoft%20Excel/syncTeams.ts" target="_blank">sync all Aircall teams</a> via the <a href="https://developer.aircall.io/api-references/#list-all-teams" target="_blank">Aircall List all Teams API</a> to a worksheet
<br>👉 <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Microsoft%20Excel/createTeamManagement.ts" target="_blank">create the team plan</a>, the table with the overview of all teams and users to a worksheet
<br>👉 <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Microsoft%20Excel/syncTeamManagement.ts" target="_blank">sync all the Aircall users per team</a> via the <a href="https://developer.aircall.io/api-references/#retrieve-a-team" target="_blank">Aircall Retrieve a Team API</a> to the team plan worksheet

Microsoft Excel does not support cell change as a trigger for a script. The recommended solution is to connect <a href="https://github.com/aircallSolutionsEngineering/spreadsheets/blob/baseline/Microsoft%20Excel/logIn_logOut.ts" target="_blank">the logIn_logOut script</a> with a button in the worksheet.
In this particular script, the cell value needs to be `Logged In` or `Logged Out` and a button click to make the script trigger change in Aircall.
