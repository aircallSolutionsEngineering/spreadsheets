// base information
const baseUrl = "https://api.aircall.io/v1/";
const apiId = "<Aircall API ID>";
const apiToken = "<Aircall API Token>";
const auth = btoa(apiId + ":" + apiToken);
// check if API Token are working correctly
// console.log("apiId: " + apiId + " Token: " + apiToken + "\n" + auth);
// additional base properties
const aircallColor = "#00BD82";

async function main(workbook: ExcelScript.Workbook) {
  const teamPlanTab = workbook.getWorksheet("team plan");
  // clean log in/out
  if (teamPlanTab != null) {
    teamPlanTab.getRangeByIndexes(1, 2, teamPlanTab.getUsedRange().getRowCount(), teamPlanTab.getUsedRange().getColumnCount()).clear();
  }
  const teamsPlan = workbook
    .getWorksheet("team plan")
    .getRangeByIndexes(0, 2, 1, workbook.getWorksheet("team plan").getUsedRange().getColumnCount() - 1)
    .getValues();
  const usersPlan = workbook
    .getWorksheet("team plan")
    .getRangeByIndexes(1, 0, workbook.getWorksheet("team plan").getUsedRange().getRowCount() - 1, 1)
    .getValues();
  // loop through each team by ID to get the users
  for (let tp = 0; tp < teamsPlan[0].length; tp++) {
    // console.log("going through team: "+teamsPlan[0][tp]);
    const teamName: string = teamsPlan[0][tp];
    const teamId: string = teamName.substring(teamName.lastIndexOf("(") + 1, teamName.lastIndexOf(")"));
    // get team users
    let teamUsers: [] = [];
    try {
      const req = await fetch(baseUrl + "teams/" + teamId, {
        method: "GET", // *GET, POST, PUT, DELETE, etc.
        headers: {
          Authorization: "Basic " + auth, // authorization header
          "Content-Type": "application/json", // sending JSON data
        },
        //payload: JSON.stringify(data) // body data type must match 'Content-Type' header
      });
      const res: object = await req.json();
      // console.log(req.status+" data: "+res);
      if (req.status !== 200) console.log("👎👎👎Error👎👎👎\r\nCant grab all the team users\r\n\r\n" + req.body);
      else {
        teamUsers = res["team"]["users"];
      }
    } catch (error) {
      console.log("👎👎👎Error👎👎👎\r\nCant list the team\r\n\r\n" + error);
    }
    // console.log(teamUsers);
    if (teamUsers.length > 0) {
      for (let tu = 0; tu < teamUsers.length; tu++) {
        // console.log(teamUsers[tu]);
        for (let up = 0; up < usersPlan.length; up++) {
          const userName: string = usersPlan[up][0];
          // console.log(userName);
          const userId: string = userName.substring(userName.lastIndexOf("(") + 1, userName.lastIndexOf(")"));
          // console.log(userId);
          if (teamUsers[tu]["id"] == userId) {
            // console.log("bingo! comparing user id: "+JSON.stringify(teamUsers[tu])+" with users plan id: "+userId);
            workbook
              .getWorksheet("team plan")
              .getRangeByIndexes(up + 1, tp + 2, 1, 1)
              .setValue("Logged In");
            workbook
              .getWorksheet("team plan")
              .getRangeByIndexes(up + 1, tp + 2, 1, 1)
              .getFormat()
              .getFill()
              .setColor(aircallColor);
            break;
          }
        }
      }
    }
  }
  const logInLogOutCriteria: ExcelScript.ListDataValidation = {
    inCellDropDown: true,
    source: "Logged In,Logged Out",
  };
  const logInLogOutRule: ExcelScript.DataValidationRule = {
    list: logInLogOutCriteria,
  };
  teamPlanTab.getRangeByIndexes(1, 1, usersPlan.length, teamsPlan[0].length).getDataValidation().setRule(logInLogOutRule);
}
