// base information
const ss = SpreadsheetApp.getActiveSheet();
const ui = SpreadsheetApp.getUi();
const userTab = SpreadsheetApp.getActive().getSheetByName("Users");
const baseUrl = "https://api.aircall.io/v1/";
const scriptProperties = PropertiesService.getScriptProperties();
const apiId =
	scriptProperties.getProperty("apiId") == null
		? userTab.getRange(1, 2).getValue()
		: scriptProperties.getProperty("apiId");
const apiToken =
	scriptProperties.getProperty("apiToken") == null
		? userTab.getRange(1, 4).getValue()
		: scriptProperties.getProperty("apiToken");
// check if API Token are working correctly
//Logger.log('apiId: '+apiId+' Token: '+apiToken+'\n'+Utilities.base64Encode(apiId+':'+apiToken));
let activeUserEmail = Session.getActiveUser();

function onOpen() {
	ui.createMenu("🚀 Aircall 🚀")
		.addItem("Check Company", "checkCompanyOverview")
		.addItem("Upload Users", "getUserData")
		.addToUi();
}

// helper functions
async function containsNumbers(str) {
	return /\d/.test(str);
}
async function validateEmail(input) {
	const validateEmailRegex = /^\S+@\S+\.\S+$/;
	return validateEmailRegex.test(input);
}

// check if API credentials and Aircall Company is correct
async function checkCompanyOverview() {
	let errorMessage = "";
	if (apiId.length < 10)
		errorMessage +=
			"Please provide API ID in the Google Apps Script Properties or cell A2\n";
	if (apiToken.length < 10)
		errorMessage +=
			"Please provide API Secret in the Google Apps Script Properties or cell A4";
	// Show alert and dont progress
	if (errorMessage.length > 0) ui.alert(errorMessage);
	else {
		const res = await UrlFetchApp.fetch(baseUrl + "company", {
			method: "GET", // *GET, POST, PUT, DELETE, etc.
			contentType: "application/json", // sending JSON data
			muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
			headers: {
				Authorization:
					"Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
			},
			// payload: JSON.stringify(payloadBody), // body data type must match "Content-Type" header
		});
		if (res.getResponseCode() == 403)
			ui.alert(
				"Please provide correct API credentials in the Google Apps Script Properties or cells A2 and A4"
			);
		else if (res.getResponseCode() != 200) ui.alert(res.getContentText());
		else {
			const aircallCompanyData = JSON.parse(res.getContentText());
			userTab
				.getRange(1, 6)
				.setValue(
					aircallCompanyData["company"]["name"] +
						" (users:" +
						aircallCompanyData["company"]["users_count"] +
						" numbers:" +
						aircallCompanyData["company"]["numbers_count"] +
						")"
				);
		}
	}
}

// grab user data from cells
async function getUserData() {
	userTab.getRange(2, 7, userTab.getLastRow(), 1).clear();
	userTab.getRange(2, 1, userTab.getLastRow(), 7).clearFormat();
	// check everything is fine with credentials
	await checkCompanyOverview();
	const userHeaders = userTab.getRange(2, 1, 1, 6).getValues()[0];
	const userData = userTab
		.getRange(3, 1, userTab.getLastRow() - 1, 6)
		.getValues();
	// Logger.log(userHeaders);
	// Logger.log(userData);
	let errorMessage = "";
	// check columns
	if (userHeaders.includes("First Name") == false)
		errorMessage += "First Name column missing\n";
	if (userHeaders.includes("Last Name") == false)
		errorMessage += "Last Name column missing\n";
	if (userHeaders.includes("Email") == false)
		errorMessage += "Email column missing\n";
	if (userHeaders.includes("Admin") == false)
		errorMessage += "Admin column missing\n";
	if (userHeaders.includes("Owner") == false)
		errorMessage += "Owner column missing\n";
	if (userHeaders.includes("Supervisor") == false)
		errorMessage += "Supervisor column missing";
	if (errorMessage.length > 0) ui.alert(errorMessage);
	else {
		// create position mapping to the columns
		let headersPosition = [];
		for (let h = 0; h < userHeaders.length; h++) {
			// Logger.log(userHeaders[h]);
			if (userHeaders[h] === "First Name")
				headersPosition.push({ column: "firstName", position: h });
			if (userHeaders[h] === "Last Name")
				headersPosition.push({ column: "lastName", position: h });
			if (userHeaders[h] === "Email")
				headersPosition.push({ column: "email", position: h });
			if (userHeaders[h] === "Admin")
				headersPosition.push({ column: "admin", position: h });
			if (userHeaders[h] === "Owner")
				headersPosition.push({ column: "owner", position: h });
			if (userHeaders[h] === "Supervisor")
				headersPosition.push({ column: "supervisor", position: h });
		}
		if (userData.length < 1)
			ui.alert("No user data provided. Please add user data from the rows 3");
		else {
			// first do a check before creating
			let invalidUserData = 0;
			const headersPositionFirstName = headersPosition.find(
				(uh) => uh.column == "firstName"
			)["position"];
			const headersPositionLastName = headersPosition.find(
				(uh) => uh.column == "lastName"
			)["position"];
			const headersPositionEmail = headersPosition.find(
				(uh) => uh.column == "email"
			)["position"];
			const headersPositionAdmin = headersPosition.find(
				(uh) => uh.column == "admin"
			)["position"];
			const headersPositionOwner = headersPosition.find(
				(uh) => uh.column == "owner"
			)["position"];
			const headersPositionSupervisor = headersPosition.find(
				(uh) => uh.column == "supervisor"
			)["position"];
			for (let u = 0; u < userData.length - 1; u++) {
				if (
					userData[u][headersPositionFirstName].length < 1 ||
					containsNumbers(userData[u][headersPositionFirstName]) == true
				) {
					userTab
						.getRange(u + 3, headersPositionFirstName + 1)
						.setBackground("red");
					invalidUserData += 1;
				} else
					userTab.getRange(u + 3, headersPositionFirstName + 1).clearFormat();
				if (
					userData[u][headersPositionLastName].length < 1 ||
					containsNumbers(userData[u][headersPositionLastName]) == true
				) {
					userTab
						.getRange(u + 3, headersPositionLastName + 1)
						.setBackground("red");
					invalidUserData += 1;
				} else
					userTab.getRange(u + 3, headersPositionLastName + 1).clearFormat();
				Logger.log(await validateEmail(userData[u][headersPositionEmail]));
				if (
					userData[u][headersPositionEmail].length < 1 ||
					(await validateEmail(userData[u][headersPositionEmail])) == false
				) {
					userTab
						.getRange(u + 3, headersPositionEmail + 1)
						.setBackground("red");
					invalidUserData += 1;
				} else userTab.getRange(u + 3, headersPositionEmail + 1).clearFormat();
				if (
					userData[u][headersPositionAdmin] != false &&
					userData[u][headersPositionAdmin] != true
				) {
					userTab
						.getRange(u + 3, headersPositionAdmin + 1)
						.setBackground("red");
					invalidUserData += 1;
				} else userTab.getRange(u + 3, headersPositionAdmin + 1).clearFormat();
				if (
					userData[u][headersPositionOwner] != false &&
					userData[u][headersPositionOwner] != true
				) {
					userTab
						.getRange(u + 3, headersPositionOwner + 1)
						.setBackground("red");
					invalidUserData += 1;
				} else userTab.getRange(u + 3, headersPositionOwner + 1).clearFormat();
				if (
					userData[u][headersPositionSupervisor] != false &&
					userData[u][headersPositionSupervisor] != true
				) {
					userTab
						.getRange(u + 3, headersPositionSupervisor + 1)
						.setBackground("red");
					invalidUserData += 1;
				} else
					userTab.getRange(u + 3, headersPositionSupervisor + 1).clearFormat();
			}
			if (invalidUserData > 0)
				ui.alert(
					"There are " +
						invalidUserData +
						" invalid cells highlighted with a red background that need to be corrected before the upload"
				);
			else {
				userTab.getRange(2, 7).setValue("Result");
				// create users by API
				for (let u = 0; u < userData.length - 1; u++) {
					let roles = ["agent"];
					if (userData[u][headersPositionAdmin] == true) roles.push("admin");
					Logger.log(roles);
					if (userData[u][headersPositionOwner] == true) roles.push("owner");
					if (userData[u][headersPositionSupervisor] == true)
						roles.push("supervisor");
					const userRecord = {
						firstName: userData[u][headersPositionFirstName],
						lastName: userData[u][headersPositionLastName],
						email: userData[u][headersPositionEmail],
						role: roles,
						is_admin: roles.includes("admin") === true ? true : false,
					};
					const createUser = await createRecord("user", userRecord);
					if (createUser.getResponseCode != 201) {
						userTab.getRange(u + 3, 7).setValue(createUser.getContentText());
						userTab.getRange(u + 3, 7).setBackground("red");
					} else {
						userTab.getRange(u + 3, 7).setValue("created");
						userTab.getRange(u + 3, 7).setBackground("green");
					}
					Logger.log("created: " + createUser);
					Utilities.sleep(500);
				}
			}
		}
	}
}

// create users in Aircall via API
async function createRecord(object, record) {
	if (apiId == null)
		ui.alert(
			"please add the Aircall API ID to the Google Apps Script Properties"
		);
	if (apiToken == null)
		ui.alert(
			"please add the Aircall API Token to the Google Apps Script Properties"
		);
	if (object != "user" && object != "tag" && object != "contact")
		ui.alert("please provide correct object. " + object + " is not valid");
	else {
		const payloadBody = {
			email: record["email"],
			first_name: record["firstName"],
			last_name: record["lastName"],
			role_ids: record["role"],
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
		return res;
	}
}
