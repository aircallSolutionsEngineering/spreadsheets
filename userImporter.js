async function getUserData() {
	const userTab = SpreadsheetApp.getActive().getSheetByName("Users");
	const userData = userTab
		.getRange(2, 1, userTab.getLastRow() - 1, 3)
		.getValues();
	// Logger.log(userData);
	for (let u = 0; u < userData.length; u++) {
		const userRecord = {
			firstName: userData[u][0],
			lastName: userData[u][1],
			email: userData[u][2],
		};
		const createUser = await createRecord("user", userRecord);
		Logger.log("created: " + createUser);
		Utilities.sleep(700);
	}
}

async function createRecord(object, record) {
	if (object != "user" && object != "tag" && object != "contact")
		ui.alert("please provide correct object. " + object + " is not valid");
	else {
		const payloadBody = {
			email: record["email"],
			first_name: record["firstName"],
			last_name: record["lastName"],
			is_admin: false,
			role_ids: ["agent"],
		};
		const res = await UrlFetchApp.fetch(baseUrl + "users", {
			method: "POST", // *GET, POST, PUT, DELETE, etc.
			contentType: "application/json", // sending JSON data
			muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
			headers: {
				Authorization:
					"Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
			},
			payload: JSON.stringify(payloadBody), // body data type must match "Content-Type" header
		});
		if (res.getResponseCode() != 201)
			ui.alert("Error in creating user: " + res.getContentText());
		else {
			return res.getContentText();
		}
	}
}
