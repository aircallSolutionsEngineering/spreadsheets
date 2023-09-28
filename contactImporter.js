// base information
const ss = SpreadsheetApp.getActiveSheet();
const ui = SpreadsheetApp.getUi();
const contactsTab = SpreadsheetApp.getActive().getSheetByName("Contacts");
const baseUrl = "https://api.aircall.io/v1/";
const scriptProperties = PropertiesService.getScriptProperties();
const apiId = scriptProperties.getProperty("apiId") == null ? contactsTab.getRange(1, 2).getValue() : scriptProperties.getProperty("apiId");
const apiToken = scriptProperties.getProperty("apiToken") == null ? contactsTab.getRange(1, 4).getValue() : scriptProperties.getProperty("apiToken");
// check if API Token are working correctly
//Logger.log('apiId: '+apiId+' Token: '+apiToken+'\n'+Utilities.base64Encode(apiId+':'+apiToken));
let activeUserEmail = Session.getActiveUser();

function onOpen() {
  ui.createMenu("🚀 Aircall 🚀").addItem("Check Company", "checkCompanyOverview").addItem("Check Contacts", "checkContactsData").addItem("Upload Contacts", "uploadContactsData").addToUi();
}

// helper functions
async function containsNumbers(str) {
  return /\d/.test(str);
}
async function isPhoneNumber(str) {
  const validatePhoneRegex = /^\+[1-9]\d{10,14}$/;
  Logger.log(str + " return: " + validatePhoneRegex.test(str));
  return validatePhoneRegex.test(str);
}
async function validateEmail(input) {
  const validateEmailRegex = /^\S+@\S+\.\S+$/;
  return validateEmailRegex.test(input);
}

// check if API credentials and Aircall Company is correct
async function checkCompanyOverview() {
  let errorMessage = "";
  if (apiId.length < 10) errorMessage += "Please provide API ID in the Google Apps Script Properties or cell A2\n";
  if (apiToken.length < 10) errorMessage += "Please provide API Secret in the Google Apps Script Properties or cell A4";
  // Show alert and dont progress
  if (errorMessage.length > 0) ui.alert(errorMessage);
  else {
    const res = await UrlFetchApp.fetch(baseUrl + "company", {
      method: "GET", // *GET, POST, PUT, DELETE, etc.
      contentType: "application/json", // sending JSON data
      muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
      headers: {
        Authorization: "Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
      },
      // payload: JSON.stringify(payloadBody), // body data type must match "Content-Type" header
    });
    if (res.getResponseCode() == 403) ui.alert("Please provide correct API credentials in the Google Apps Script Properties or cells A2 and A4");
    else if (res.getResponseCode() != 200) ui.alert(res.getContentText());
    else {
      const aircallCompanyData = JSON.parse(res.getContentText());
      contactsTab.getRange(1, 6).setValue(aircallCompanyData["company"]["name"] + " (users:" + aircallCompanyData["company"]["users_count"] + " numbers:" + aircallCompanyData["company"]["numbers_count"] + ")");
    }
  }
}

// grab user data from cells and check against validations
async function checkContactsData() {
  contactsTab.getRange(2, 7, contactsTab.getLastRow(), 1).clear();
  contactsTab.getRange(2, 1, contactsTab.getLastRow(), 7).clearFormat();
  // check everything is fine with credentials
  await checkCompanyOverview();
  const contactsHeaders = contactsTab.getRange(2, 1, 1, 6).getValues()[0];
  const contactsData = contactsTab.getRange(3, 1, contactsTab.getLastRow() - 1, 6).getValues();
  // Logger.log(contactsHeaders);
  // Logger.log(contactsData);
  let errorMessage = "";
  // check columns
  if (contactsHeaders.includes("First Name") == false) errorMessage += "First Name column missing\n";
  if (contactsHeaders.includes("Last Name") == false) errorMessage += "Last Name column missing\n";
  if (contactsHeaders.includes("Email") == false) errorMessage += "Email column missing\n";
  if (contactsHeaders.includes("Phone1") == false) errorMessage += "Phone1 column missing\n";
  if (contactsHeaders.includes("Phone2") == false) errorMessage += "Phone2 column missing\n";
  if (contactsHeaders.includes("Company") == false) errorMessage += "Company column missing";
  if (errorMessage.length > 0) ui.alert(errorMessage);
  else {
    // create position mapping to the columns
    let headersPosition = [];
    for (let h = 0; h < contactsHeaders.length; h++) {
      // Logger.log(contactsHeaders[h]);
      if (contactsHeaders[h] === "First Name") headersPosition.push({ column: "firstName", position: h });
      if (contactsHeaders[h] === "Last Name") headersPosition.push({ column: "lastName", position: h });
      if (contactsHeaders[h] === "Email") headersPosition.push({ column: "email", position: h });
      if (contactsHeaders[h] === "Phone1") headersPosition.push({ column: "phone1", position: h });
      if (contactsHeaders[h] === "Phone2") headersPosition.push({ column: "phone2", position: h });
      if (contactsHeaders[h] === "Company") headersPosition.push({ column: "company", position: h });
    }
    if (contactsData.length < 1) ui.alert("No user data provided. Please add user data from the rows 3");
    else {
      // first do a check before creating
      let invalidcontactsData = 0;
      const headersPositionFirstName = headersPosition.find((uh) => uh.column == "firstName")["position"];
      const headersPositionLastName = headersPosition.find((uh) => uh.column == "lastName")["position"];
      const headersPositionEmail = headersPosition.find((uh) => uh.column == "email")["position"];
      const headersPositionPhone1 = headersPosition.find((uh) => uh.column == "phone1")["position"];
      const headersPositionPhone2 = headersPosition.find((uh) => uh.column == "phone2")["position"];
      const headersPositionCompany = headersPosition.find((uh) => uh.column == "company")["position"];
      for (let u = 0; u < contactsData.length - 1; u++) {
        if (containsNumbers(contactsData[u][headersPositionFirstName]) == true) {
          contactsTab.getRange(u + 3, headersPositionFirstName + 1).setBackground("red");
          invalidcontactsData += 1;
        } else contactsTab.getRange(u + 3, headersPositionFirstName + 1).clearFormat();
        if (containsNumbers(contactsData[u][headersPositionLastName]) == true) {
          contactsTab.getRange(u + 3, headersPositionLastName + 1).setBackground("red");
          invalidcontactsData += 1;
        } else contactsTab.getRange(u + 3, headersPositionLastName + 1).clearFormat();
        // Logger.log(await validateEmail(contactsData[u][headersPositionEmail]));
        if (contactsData[u][headersPositionEmail].length > 0 && (await validateEmail(contactsData[u][headersPositionEmail])) == false) {
          contactsTab.getRange(u + 3, headersPositionEmail + 1).setBackground("red");
          invalidcontactsData += 1;
        } else contactsTab.getRange(u + 3, headersPositionEmail + 1).clearFormat();
        if (contactsData[u][headersPositionPhone1].length < 1 || (await isPhoneNumber(contactsData[u][headersPositionPhone1])) === false) {
          contactsTab.getRange(u + 3, headersPositionPhone1 + 1).setBackground("red");
          invalidcontactsData += 1;
        } else contactsTab.getRange(u + 3, headersPositionPhone1 + 1).clearFormat();
        if (contactsData[u][headersPositionPhone2].length > 0 && (await isPhoneNumber(contactsData[u][headersPositionPhone2])) === false) {
          contactsTab.getRange(u + 3, headersPositionPhone2 + 1).setBackground("red");
          invalidcontactsData += 1;
        } else contactsTab.getRange(u + 3, headersPositionPhone2 + 1).clearFormat();
        if (containsNumbers(contactsData[u][headersPositionCompany]) == true) {
          contactsTab.getRange(u + 3, headersPositionCompany + 1).setBackground("red");
          invalidcontactsData += 1;
        } else contactsTab.getRange(u + 3, headersPositionCompany + 1).clearFormat();
      }
      ui.alert("There are " + invalidcontactsData + " invalid cells highlighted with a red background that need to be corrected before the upload");
      return invalidcontactsData, contactsData;
    }
  }
}

// check and upload contacts
async function uploadContactsData() {
  const contactsCheck = await checkContactsData();
  const invalidcontactsData = contactsCheck[0];
  const contactsData = contactsCheck[1];
  if (invalidcontactsData > 0) ui.alert("There are " + invalidcontactsData + " invalid cells highlighted with a red background that need to be corrected before the upload");
  else {
    contactsTab.getRange(2, 7).setValue("Result");
    // create users by API
    for (let u = 0; u < contactsData.length - 1; u++) {
      const contactRecord = {
        firstName: contactsData[u][headersPositionFirstName],
        lastName: contactsData[u][headersPositionLastName],
        email: contactsData[u][headersPositionEmail],
        phone1: contactsData[u][headersPositionPhone1],
        phone2: contactsData[u][headersPositionPhone2],
        company: contactsData[u][headersPositionCompany],
      };
      const createContact = await createRecord("contact", contactRecord);
      if (createContact.getResponseCode != 201) {
        contactsTab.getRange(u + 3, 7).setValue(createContact.getContentText());
        contactsTab.getRange(u + 3, 7).setBackground("red");
      } else {
        contactsTab.getRange(u + 3, 7).setValue("created");
        contactsTab.getRange(u + 3, 7).setBackground("green");
      }
      Logger.log("created: " + createContact);
      Utilities.sleep(500);
    }
  }
}

// create users in Aircall via API
async function createRecord(object, record) {
  if (apiId == null) ui.alert("please add the Aircall API ID to the Google Apps Script Properties");
  if (apiToken == null) ui.alert("please add the Aircall API Token to the Google Apps Script Properties");
  if (object != "user" && object != "tag" && object != "contact") ui.alert("please provide correct object. " + object + " is not valid");
  else {
    let payloadBody = {
      phone_numbers: [record["phone1"], record["phone2"]],
    };
    if (record["email"] != null) payloadBody.email = record["email"];
    if (record["firstName"] != null) payloadBody.first_name = record["firstName"];
    if (record["lastName"] != null) payloadBody.last_name = record["lastName"];
    if (record["company"] != null) payloadBody.company_name = record["company"];
    // Logger.log(payloadBody);
    const res = await UrlFetchApp.fetch(baseUrl + "contacts", {
      method: "POST", // *GET, POST, PUT, DELETE, etc.
      contentType: "application/json", // sending JSON data
      muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
      headers: {
        Authorization: "Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
      },
      payload: JSON.stringify(payloadBody), // body data type must match "Content-Type" header
    });
    return res;
  }
}
