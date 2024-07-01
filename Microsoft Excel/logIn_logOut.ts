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
  // const responseBody: string = await req.text();
  // if (req.status !== 200 && req.status !== 201) console.log("👎👎👎Error👎👎👎\r\nIssue with " + method + " to a team\r\n\r\n" + responseBody);
  return req;
}

async function main(workbook: ExcelScript.Workbook) {
  // onEdit function that is triggered using a scheduled onEdit trigger
  // get value of the cell
  const range = workbook.getSelectedRanges();
  // console.log(range.getAddress());
  let rangeCells: string[] = range.getAddress().split(",");
  // console.log(rangeCells);
  // if selecting a large range
  if (rangeCells[0].includes(":") == true) {
    rangeCells = [];
    const numberOfColumns: number = workbook.getSelectedRange().getColumnCount();
    const numberOfRows: number = workbook.getSelectedRange().getRowCount();
    const startingCell: string = workbook.getSelectedRange().getAddress().substring(0, range.getAddress().indexOf(":"));
    // console.log(startingCell);
    for (let cl = 0; cl < numberOfColumns; cl++) {
      for (let r = 0; r < numberOfRows; r++) {
        const nextRowCell: string = workbook.getActiveWorksheet().getRange(startingCell).getOffsetRange(r, cl).getAddress();
        rangeCells.push(nextRowCell);
      }
    }
  }
  // console.log(rangeCells);
  // if user sets itself to log in or log out
  for (let c = 0; c < rangeCells.length; c++) {
    const cellValue = workbook.getActiveWorksheet().getRange(rangeCells[c]).getValue();
    if (cellValue === "Logged Out" || cellValue === "Logged In") {
      const cellRange = workbook.getActiveWorksheet().getRange(rangeCells[c]);
      // ss.getRangeByIndexes(20,1).setValue("test: "+ ss.getRangeByIndexes(2,range.getColumn()).getValue());
      let cellTeam: string | number | boolean = workbook.getActiveWorksheet().getRangeByIndexes(0, cellRange.getColumnIndex(), 1, 1).getValue();
      let cellUser: string | number | boolean = workbook.getActiveWorksheet().getRangeByIndexes(cellRange.getRowIndex(), 0, 1, 1).getValue();
      cellTeam = String(cellTeam);
      cellUser = String(cellUser);
      const aircallUserId = cellUser.substring(cellUser.lastIndexOf("(") + 1, cellUser.lastIndexOf(")"));
      const aircallTeamId = cellTeam.substring(cellTeam.lastIndexOf("(") + 1, cellTeam.lastIndexOf(")"));
      // console.log(aircallUserId,aircallTeamId);
      if (aircallUserId == "") console.log("User Name and ID is not correctly formatted. Please create and sync the team plan again");
      // add user to team
      else if (cellValue === "Logged In") {
        const res = await changeTeam(aircallTeamId, aircallUserId, "POST");
        if (res.status == 201) cellRange.getFormat().getFill().setColor(aircallColor);
        else console.log("Already logged in");
      }
      // remove user from team
      else if (cellValue === "Logged Out") {
        const res = await changeTeam(aircallTeamId, aircallUserId, "DELETE");
        if (res.status == 200) cellRange.getFormat().getFill().setColor("red");
        else console.log("Already logged out");
      }
    }
  }
}
