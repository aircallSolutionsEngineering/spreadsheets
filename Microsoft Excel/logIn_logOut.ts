// base information
const baseUrl = "https://api.aircall.io/v1/";
const apiId = "<Aircall API ID>";
const apiToken = "<Aircall API Token>";
const auth = btoa(apiId + ":" + apiToken);
// check if API Token are working correctly
console.log("apiId: " + apiId + " Token: " + apiToken + "\n" + auth);

const aircallColor = "#00BD82";

async function changeTeam(teamId: string, userId: string, method: string) {
  let req = await fetch(baseUrl + "teams/" + teamId + "/users/" + userId, {
    method: method, // *GET, POST, PUT, DELETE, etc.
    headers: {
      Authorization: "Basic " + btoa(apiId + ":" + apiToken), // authorization header
      "Content-Type": "application/json", // sending JSON data
    },
    //payload: JSON.stringify(data) // body data type must match 'Content-Type' header
  });
  if (req.status !== 200 && req.status !== 201) console.log("👎👎👎Error👎👎👎\r\nIssue with " + method + " to a team\r\n\r\n" + req.body);
}

function main(workbook: ExcelScript.Workbook) {
  // onEdit function that is triggered using a scheduled onEdit trigger
  // get value of the cell
  const range = workbook.getActiveCell();
  const cellValue = range.getValue();
  // range.setNote("test: " + cellValue);
  // if user sets itself to log in or log out
  if (cellValue === "Logged Out" || cellValue === "Logged In") {
    // ss.getRangeByIndexes(20,1).setValue("test: "+ ss.getRangeByIndexes(2,range.getColumn()).getValue());
    const cellTeam: string = workbook.getActiveWorksheet().getRangeByIndexes(0, range.getColumnCount(), 1, 1).getValue();
    const cellUser: string = workbook.getActiveWorksheet().getRangeByIndexes(range.getRowCount(), 0, 1, 1).getValue();
    const aircallUserId = cellUser.substring(cellUser.lastIndexOf("(") + 1, cellUser.lastIndexOf(")"));
    const aircallTeamId = cellTeam.substring(cellTeam.lastIndexOf("(") + 1, cellTeam.lastIndexOf(")"));
    if (aircallUserId == "") console.log("User Name and ID is not correctly formatted. Please create and sync the team plan again");
    // add user to team
    else if (cellValue === "Logged In") {
      changeTeam(aircallTeamId, aircallUserId, "POST");
      range.getFormat().getFill().setColor(aircallColor);
    }
    // remove user from team
    else if (cellValue === "Logged Out") {
      changeTeam(aircallTeamId, aircallUserId, "DELETE");
      range.getFormat().getFill().setColor("red");
    }
  }
}
