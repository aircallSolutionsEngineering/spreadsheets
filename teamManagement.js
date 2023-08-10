// additional base properties
const reParenthesis = /\(([^)]+)\)/g;
const aircallColor = "#00BD82";

// add Aircall Menu
function onOpen() {
  ui.createMenu("🚀 Aircall 🚀").addItem("Sync User List", "syncUsers").addItem("Sync Team List", "syncTeams").addSeparator().addItem("Create Team Management", "createTeamManagement").addItem("Sync Team Management", "syncTeamManagement").addToUi();
}

// get all users or teams or numbers or contacts
async function listRecords(object) {
  if (object != "users" && object != "teams" && object != "numbers" && object != "contacts") ui.alert("incorrect object: " + object + " is not part of Aircall APIs");
  else {
    let records = [];
    try {
      let req = await UrlFetchApp.fetch(baseUrl + object + "?per_page=50", {
        method: "GET", // *GET, POST, PUT, DELETE, etc.
        headers: {
          Authorization: "Basic " + Utilities.base64Encode(PropertiesService.getScriptProperties().getProperty("apiId") + ":" + PropertiesService.getScriptProperties().getProperty("apiToken")), // authorization header
          "Content-Type": "application/json", // sending JSON data
        },
        muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
        //payload: JSON.stringify(data) // body data type must match 'Content-Type' header
      });
      // Logger.log('existing dialer campaign: '+res.getResponseCode());
      if (req.getResponseCode() !== 200) ui.alert("👎👎👎Error👎👎👎\r\nCant grab all the " + object + "\r\n\r\n" + req.getContentText());
      else {
        // ui.alert('👍👍👍Success👍👍👍\r\nAll '+objects);
        let res = JSON.parse(req.getContentText());
        records = res[object];
        // Logger.log(res.meta);
        if (res.meta["next_page_link"] != null) {
          for (let p = 2; p < Math.ceil(res.meta["total"] / 50); p++) {
            req = await UrlFetchApp.fetch(baseUrl + object + "?per_page=50&page=" + p, {
              method: "GET", // *GET, POST, PUT, DELETE, etc.
              headers: {
                Authorization: "Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
                "Content-Type": "application/json", // sending JSON data
              },
              muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
              //payload: JSON.stringify(data) // body data type must match 'Content-Type' header
            });
            res = JSON.parse(req.getContentText());
            records = records.concat(res[object]);
          }
        }
      }
      return records;
    } catch (error) {
      ui.alert("👎👎👎Error👎👎👎\r\nCant create the " + object + "\r\n\r\n" + error);
      // deal with any errors
      // Logger.log(error);
    }
  }
}

// add all users or teams or numbers or contacts
async function addRecords(object, data) {
  if (object != "users" && object != "teams" && object != "numbers" && object != "contacts") ui.alert("incorrect object: " + object + " is not part of Aircall APIs");
  else {
    if (SpreadsheetApp.getActive().getSheetByName(object) == null) SpreadsheetApp.getActive().insertSheet(object);
    SpreadsheetApp.getActive().getSheetByName(object).clear();
    const objectRecords = data;
    if (objectRecords.length === 0) ui.alert("No " + object + " available");
    // calculate the number of rows and columns needed
    var numRows = objectRecords.length;
    const cols = Object.keys(objectRecords[0]);
    const printCols = [];
    cols.forEach(function (col) {
      printCols.push(col);
    });
    // set column headers
    if (cols.length > 0) {
      SpreadsheetApp.getActive().getSheetByName(object).getRange(1, 1, 1, cols.length).setValues([printCols]);
    }
    // set rows
    const finalRecords = [];
    for (r = 0; r < objectRecords.length; r++) {
      let finalRow = [];
      const recordFields = Object.values(objectRecords[r]);
      recordFields.forEach(function (recordField) {
        // Logger.log(recordField)
        if (Array.isArray(recordField) == true) finalRow.push(JSON.stringify(recordField));
        else finalRow.push(recordField);
      });
      // Logger.log(finalRow);
      finalRecords.push(finalRow);
    }
    // Logger.log(finalRecords);
    // add sheet if not available
    SpreadsheetApp.getActive().getSheetByName(object).getRange(2, 1, objectRecords.length, cols.length).setValues(finalRecords);
  }
}

// sync users of Aircall API with user list in spreadsheet
async function syncUsers() {
  const userList = await listRecords("users");
  await addRecords("users", userList);
}

// sync teams of Aircall API
async function syncTeams() {
  // get all teams from Aircall User API
  const teamList = await listRecords("teams");
  await addRecords("teams", teamList);
}

// create / update team management
async function createTeamManagement() {
  const users = SpreadsheetApp.getActive()
    .getSheetByName("users")
    .getSheetValues(2, 1, SpreadsheetApp.getActive().getSheetByName("users").getLastRow() - 1, 4);
  const teams = SpreadsheetApp.getActive()
    .getSheetByName("teams")
    .getSheetValues(2, 1, SpreadsheetApp.getActive().getSheetByName("teams").getLastRow() - 1, 2);
  if (SpreadsheetApp.getActive().getSheetByName("team plan") == null) SpreadsheetApp.getActive().insertSheet("team plan");
  const teamPlanTab = SpreadsheetApp.getActive().getSheetByName("team plan");
  teamPlanTab.clear();
  // prepare team and user data
  let userData = [];
  for (let u = 0; u < users.length; u++) {
    const userRow = users[u][2] + " (" + users[u][0] + ")";
    userData.push([userRow]);
  }
  teamPlanTab.getRange(2, 1, users.length, 1).setValues(userData);
  let teamData = [];
  for (let t = 0; t < teams.length; t++) {
    const teamRow = teams[t][1] + " (" + teams[t][0] + ")";
    teamData.push([teamRow]);
  }
  teamPlanTab.getRange(1, 2, 1, teams.length).setValues([teamData]);
  // create complete sheet with log in / log out
  const logInLogOutRule = SpreadsheetApp.newDataValidation().requireValueInList(["Logged In", "Logged Out"], true).build();
  teamPlanTab.getRange(2, 2, users.length, teams.length).setDataValidation(logInLogOutRule);
}

// sync current team structure into sheet
async function syncTeamManagement() {
  const teamPlanTab = SpreadsheetApp.getActive().getSheetByName("team plan");
  // clean log in/out
  if (teamPlanTab != null) {
    teamPlanTab.getRange(2, 2, teamPlanTab.getLastRow(), teamPlanTab.getLastColumn()).clear();
  }
  const users = SpreadsheetApp.getActive()
    .getSheetByName("users")
    .getSheetValues(2, 1, SpreadsheetApp.getActive().getSheetByName("users").getLastRow() - 1, 3);
  const teams = SpreadsheetApp.getActive()
    .getSheetByName("teams")
    .getSheetValues(2, 1, SpreadsheetApp.getActive().getSheetByName("teams").getLastRow() - 1, 5);
  const teamsPlan = SpreadsheetApp.getActive()
    .getSheetByName("team plan")
    .getSheetValues(1, 2, 1, SpreadsheetApp.getActive().getSheetByName("team plan").getLastColumn() - 1);
  const usersPlan = SpreadsheetApp.getActive()
    .getSheetByName("team plan")
    .getSheetValues(2, 1, SpreadsheetApp.getActive().getSheetByName("team plan").getLastRow() - 1, 1);
  // loop through each team by ID to get the users
  for (let tp = 0; tp < teamsPlan[0].length; tp++) {
    // Logger.log("going through team: "+teamsPlan[0][tp]);
    const parenthesisTeamData = [...teamsPlan[0][tp].matchAll(reParenthesis)].flat();
    const teamId = parenthesisTeamData.slice(-1)[0];
    // get team users
    let teamUsers = [];
    if (teamId == teams[tp][0]) teamUsers = JSON.parse(teams[tp][4]);
    else {
      for (let t = 0; t < teams.length; t++) {
        if (teamId == teams[t][0]) {
          teamUsers = JSON.parse(teams[tp][4]);
          break;
        }
      }
    }
    // Logger.log(teamUsers);
    if (teamUsers.length > 0) {
      for (let tu = 0; tu < teamUsers.length; tu++) {
        // Logger.log(teamUsers[tu]);
        for (let up = 0; up < usersPlan.length; up++) {
          const parenthesisUserData = [...usersPlan[up][0].matchAll(reParenthesis)].flat();
          const userId = parenthesisUserData.slice(-1)[0];
          if (teamUsers[tu]["id"] == userId) {
            // Logger.log("bingo! comparing user id: "+JSON.stringify(teamUsers[tu])+" with users plan id: "+userId);
            SpreadsheetApp.getActive()
              .getSheetByName("team plan")
              .getRange(up + 1, tp + 2)
              .setValue("Logged In")
              .setBackground(aircallColor);
            break;
          }
        }
      }
    }
  }
}

async function changeTeam(teamId, userId, method) {
  let req = await UrlFetchApp.fetch(baseUrl + "teams/" + teamId + "/users/" + userId, {
    method: method, // *GET, POST, PUT, DELETE, etc.
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
      "Content-Type": "application/json", // sending JSON data
    },
    muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
    //payload: JSON.stringify(data) // body data type must match 'Content-Type' header
  });
  if (req.getResponseCode() !== 200 && req.getResponseCode() !== 201) ui.alert("👎👎👎Error👎👎👎\r\nIssue with " + method + " to a team\r\n\r\n" + req.getContentText());
}

async function cleanTeam(teamId) {
  let req = await UrlFetchApp.fetch(baseUrl + "teams/" + teamId, {
    method: "GET", // *GET, POST, PUT, DELETE, etc.
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
      "Content-Type": "application/json", // sending JSON data
    },
    muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
    //payload: JSON.stringify(data) // body data type must match 'Content-Type' header
  });
  if (req.getResponseCode() !== 200 && req.getResponseCode() !== 201) {
    ui.alert("👎👎👎Error👎👎👎\r\nIssue with " + method + " to a team\r\n\r\n" + req.getContentText());
  } else {
    // get all users from the matched team
    const teamUsers = req.getContentText()["users"];
    // Logger.log(teamUsers);
    // remove users from team
    if (teamUsers.length > 0) {
      for (let tu = 0; tu < teamUsers.length; tu++) {
        let req = await UrlFetchApp.fetch(baseUrl + "teams/" + aircallTeams[t]["id"] + "/users/" + teamUsers[tu]["id"], {
          method: "DELETE", // *GET, POST, PUT, DELETE, etc.
          headers: {
            Authorization: "Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
            "Content-Type": "application/json", // sending JSON data
          },
          muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
          //payload: JSON.stringify(data) // body data type must match 'Content-Type' header
        });
        if (req.getResponseCode() !== 200 && req.getResponseCode() !== 201) {
          ui.alert("👎👎👎Error👎👎👎\r\nIssue with removing all users from team" + teamName + "\r\n\r\n" + req.getContentText());
        }
      }
    }
  }
}

// onEdit function that is triggered using a scheduled onEdit trigger
function cellChange() {
  // get value of the cell
  const range = ss.getCurrentCell();
  const cellValue = range.getValue();
  // range.setNote("test: " + cellValue);
  // if user sets itself to log in or log out
  if (cellValue === "Logged Out" || cellValue === "Logged In") {
    // ss.getRange(20,1).setValue("test: "+ ss.getRange(2,range.getColumn()).getValue());
    const cellTeam = ss.getRange(1, range.getColumn()).getValue();
    const cellUser = ss.getRange(range.getRow(), 1).getValue();
    const aircallUserId = [...cellUser.matchAll(reParenthesis)].flat().slice(-1)[0];
    const aircallTeamId = [...cellTeam.matchAll(reParenthesis)].flat().slice(-1)[0];
    if (aircallUserId == "") ui.alert("User Name and ID is not correctly formatted. Please create and sync the team plan again");
    // add user to team
    else if (cellValue === "Logged In") {
      changeTeam(aircallTeamId, aircallUserId, "POST");
      range.setBackground(aircallColor);
    }
    // remove user from team
    else if (cellValue === "Logged Out") {
      changeTeam(aircallTeamId, aircallUserId, "DELETE");
      range.setBackground("red");
    }
  }
  if (cellValue === "Active" || cellValue === "Inactive") {
    const cellTeam = ss.getRange(range.getRow(), 2).getValue();
    if (cellValue === "Inactive") {
      // remove all users from team
      cleanTeam(aircallTeamId);
      // find the corresponding tab to set all users to logged out in spreadsheet
      const tab = cellTeam.substring(cellTeam.indexOf(" ") + 1, cellTeam.indexOf(" ") + 3);
      // Logger.log(tab);
      const languageTab = SpreadsheetApp.getActive().getSheetByName(tab);
      // find the relevant team column
      const allLanguageTeams = languageTab.getRange(2, 4, 1, languageTab.getLastColumn() - 4).getValues();
      let teamColumn;
      for (tc = 0; tc < allLanguageTeams[0].length; tc++) {
        if (allLanguageTeams[0][tc] == cellTeam) {
          teamColumn = tc;
          break;
        }
      }
      // Logger.log(teamColumn);
      languageTab.getRange(3, teamColumn + 4, languageTab.getLastRow() - 3, 1).setValue("Logged Out");
    }
  }
}
