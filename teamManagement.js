// user lists
const userListTab = SpreadsheetApp.getActive().getSheetByName("User List");
const existingUsers = userListTab
	.getRange(1, 1, userListTab.getLastRow(), userListTab.getLastColumn())
	.getValues();

// add Aircall Menu
function onOpen() {
	ui.createMenu("🚀 Aircall 🚀")
		.addItem("Sync User List", "syncUsers")
		.addToUi();
}

// get all users or teams or numbers or contacts
async function listRecords(object) {
	if (
		object != "users" &&
		object != "teams" &&
		object != "numbers" &&
		object != "contacts"
	)
		ui.alert("incorrect object: " + object + " is not part of Aircall APIs");
	else {
		let records = [];
		try {
			let req = await UrlFetchApp.fetch(baseUrl + object + "?per_page=50", {
				method: "GET", // *GET, POST, PUT, DELETE, etc.
				headers: {
					Authorization:
						"Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
					"Content-Type": "application/json", // sending JSON data
				},
				muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
				//payload: JSON.stringify(data) // body data type must match 'Content-Type' header
			});
			// Logger.log(object + ' api response: '+req.getResponseCode());
			if (req.getResponseCode() !== 200)
				ui.alert(
					"👎👎👎Error👎👎👎\r\nCant grab all the " +
						object +
						"\r\n\r\n" +
						res.getContentText()
				);
			else {
				// ui.alert('👍👍👍Success👍👍👍\r\nAll '+object);
				let res = JSON.parse(req.getContentText());
				records = res[object];
				// Logger.log(res.meta);
				if (res.meta.next_page_link != null) {
					for (p = 2; p < Math.ceil(res.meta.count / res.meta.count); p++) {
						req = await UrlFetchApp.fetch(
							baseUrl + object + "?per_page=50&page=" + p,
							{
								method: "GET", // *GET, POST, PUT, DELETE, etc.
								headers: {
									Authorization:
										"Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
									"Content-Type": "application/json", // sending JSON data
								},
								muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
								//payload: JSON.stringify(data) // body data type must match 'Content-Type' header
							}
						);
						records += records.push(JSON.parse(req.getContextText())[object]);
					}
				}
			}
			return records;
		} catch (error) {
			ui.alert(
				"👎👎👎Error👎👎👎\r\nCant create the " + object + "\r\n\r\n" + error
			);
			// deal with any errors
			// Logger.log(error);
		}
	}
}

// sync users of Aircall API with user list in spreadsheet
async function syncUsers() {
	// get all users from Aircall User API
	const userList = await listRecords("users");
	// Logger.log(existingUsers);
	const emailIndex = existingUsers[0].findIndex((c) => c == "Email");
	const aircallIdIndex = existingUsers[0].findIndex((c) => c == "Aircall ID");
	// Logger.log('email column: '+ emailIndex + ' aircall ID column: '+aircallIdIndex);
	// match each user based on email
	for (let r = 1; r < existingUsers.length; r++) {
		const email = existingUsers[r][emailIndex];
		for (let u = 0; u < userList.length; u++) {
			// Logger.log(email +" : "+ userList[u]["email"])
			// set Aircall user ID in column Aircall ID
			if (userList[u]["email"] === email) {
				userListTab
					.getRange(r + 1, aircallIdIndex + 1)
					.setValue(userList[u]["id"]);
				break;
			}
		}
	}
}

async function changeTeam(teamName, userId, method) {
	const aircallTeams = await listRecords("teams");
	// match teams based on name
	for (let t = 0; t < aircallTeams.length; t++) {
		// Logger.log(aircallTeams[t]["name"]+" compared to: "+ teamName);
		// add or remove user from the team
		if (aircallTeams[t]["name"] === teamName) {
			let req = await UrlFetchApp.fetch(
				baseUrl + "teams/" + aircallTeams[t]["id"] + "/users/" + userId,
				{
					method: method, // *GET, POST, PUT, DELETE, etc.
					headers: {
						Authorization:
							"Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
						"Content-Type": "application/json", // sending JSON data
					},
					muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
					//payload: JSON.stringify(data) // body data type must match 'Content-Type' header
				}
			);
			if (req.getResponseCode() !== 200 && req.getResponseCode() !== 201)
				ui.alert(
					"👎👎👎Error👎👎👎\r\nIssue with " +
						method +
						" to a team\r\n\r\n" +
						req.getContentText()
				);
			break;
		}
	}
}

async function cleanTeam(teamName) {
	const aircallTeams = await listRecords("teams");
	// find team based on name match
	for (let t = 0; t < aircallTeams.length; t++) {
		// Logger.log(aircallTeams[t]["name"]+" compared to: "+ teamName);
		if (aircallTeams[t]["name"] === teamName) {
			// get all users from the matched team
			const teamUsers = aircallTeams[t]["users"];
			// Logger.log(teamUsers);
			// remove users from team
			if (teamUsers.length > 0) {
				for (let tu = 0; tu < teamUsers.length; tu++) {
					let req = await UrlFetchApp.fetch(
						baseUrl +
							"teams/" +
							aircallTeams[t]["id"] +
							"/users/" +
							teamUsers[tu]["id"],
						{
							method: "DELETE", // *GET, POST, PUT, DELETE, etc.
							headers: {
								Authorization:
									"Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
								"Content-Type": "application/json", // sending JSON data
							},
							muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
							//payload: JSON.stringify(data) // body data type must match 'Content-Type' header
						}
					);
					if (req.getResponseCode() !== 200 && req.getResponseCode() !== 201)
						ui.alert(
							"👎👎👎Error👎👎👎\r\nIssue with removing all users from team" +
								teamName +
								"\r\n\r\n" +
								req.getContentText()
						);
				}
			}
			break;
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
		const cellTeam = ss.getRange(2, range.getColumn()).getValue();
		const cellUser = ss.getRange(range.getRow(), 3).getValue();
		const userEmails = userListTab
			.getRange(2, 3, userListTab.getLastRow(), 1)
			.getValues();
		// find aircall ID from users list based on email
		const aircallUserId =
			existingUsers[userEmails.findIndex((u) => u == cellUser) + 1][9];
		// Logger.log(aircallUserId);
		// ss.getRange(20,1).setValue("test: "+ aircallUserId);
		if (aircallUserId == "")
			ui.alert("Please sync the users to get the Aircall User ID");
		// add user to team
		else if (cellValue === "Logged In")
			changeTeam(cellTeam, aircallUserId, "POST");
		// remove user from team
		else if (cellValue === "Logged Out")
			changeTeam(cellTeam, aircallUserId, "DELETE");
	}
	if (cellValue === "Active" || cellValue === "Inactive") {
		const cellTeam = ss.getRange(range.getRow(), 2).getValue();
		if (cellValue === "Inactive") {
			// remove all users from team
			cleanTeam(cellTeam);
			// find the corresponding tab to set all users to logged out in spreadsheet
			const tab = cellTeam.substring(
				cellTeam.indexOf(" ") + 1,
				cellTeam.indexOf(" ") + 3
			);
			// Logger.log(tab);
			const languageTab = SpreadsheetApp.getActive().getSheetByName(tab);
			// find the relevant team column
			const allLanguageTeams = languageTab
				.getRange(2, 4, 1, languageTab.getLastColumn() - 4)
				.getValues();
			let teamColumn;
			for (tc = 0; tc < allLanguageTeams[0].length; tc++) {
				if (allLanguageTeams[0][tc] == cellTeam) {
					teamColumn = tc;
					break;
				}
			}
			// Logger.log(teamColumn);
			languageTab
				.getRange(3, teamColumn + 4, languageTab.getLastRow() - 3, 1)
				.setValue("Logged Out");
		};
	};
};
