function main(workbook: ExcelScript.Workbook) {
  const users = workbook
    .getWorksheet("users")
    .getRangeByIndexes(1, 0, workbook.getWorksheet("users").getUsedRange().getRowCount() - 1, 3)
    .getValues();
  const teams = workbook
    .getWorksheet("teams")
    .getRangeByIndexes(1, 0, workbook.getWorksheet("teams").getUsedRange().getRowCount() - 1, 2)
    .getValues();
  if (workbook.getWorksheet("team plan") == null) workbook.addWorksheet("team plan");
  const teamPlanTab = workbook.getWorksheet("team plan");
  teamPlanTab.getRange().clear();
  // prepare team and user data
  let userData: string[][] = [];
  for (let u = 0; u < users.length; u++) {
    const userRow = users[u][2] + " (" + users[u][0] + ")";
    userData.push([userRow]);
  }
  teamPlanTab.getRangeByIndexes(1, 0, users.length, 1).setValues(userData);
  let teamData: string[] = [];
  for (let t = 0; t < teams.length; t++) {
    const teamRow = teams[t][1] + " (" + teams[t][0] + ")";
    teamData.push(teamRow);
  }
  // console.log("size: " + teamData.length + " teams: " + teamData);
  teamPlanTab.getRangeByIndexes(0, 1, 1, 1).setValue("status");
  teamPlanTab.getRangeByIndexes(0, 2, 1, teams.length).setValues([teamData]);
  // create complete sheet with log in / log out

  const logInLogOutCriteria: ExcelScript.ListDataValidation = {
    inCellDropDown: true,
    source: "Logged In,Logged Out",
  };
  const logInLogOutRule: ExcelScript.DataValidationRule = {
    list: logInLogOutCriteria,
  };
  teamPlanTab.getRangeByIndexes(1, 2, users.length, teams.length).getDataValidation().setRule(logInLogOutRule);
}
