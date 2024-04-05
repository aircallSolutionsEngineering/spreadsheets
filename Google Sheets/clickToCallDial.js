const lastRow = ss.getLastRow();
//Logger.log(lastRow);
// grab all phone numbers fixed to all rows
const allPhoneNumbers = ss.getRange("B3:B" + lastRow).getDisplayValues();
// grab only selected phone numbers
const selectedPhoneNumbers = ss.getActiveRange().getDisplayValues();
//Logger.log(allPhoneNumbers);
let activeUserEmail = Session.getActiveUser();
let activeUserAircallId = "";

// add Aircall Menu
function onOpen() {
  ui.createMenu('🚀 Aircall 🚀')
    .addItem('Aircall Softphone','showAircallSoftphone')
    .addSeparator()
    .addItem('Start Outbound Call','startOutboundCall')
    .addItem('Click To Dial','clickToDial')
    .addToUi();
};

// add Aircall phone app in sidebar
function showAircallSoftphone() {
  const aircallSoftphone = HtmlService.createHtmlOutputFromFile('Sidebar');
  aircallSoftphone.setTitle('Aircall Phone');
  SpreadsheetApp.getUi().showSidebar(aircallSoftphone);
};

// click to call configuration
function aircallConfiguration() {
	return (userDict = [
		{
			name: "name 1",
			email: "email 1",
			aircall_id: "id 1",
			number_id: "number id 1",
		},
		{
			name: "name 2",
			email: "email 2",
			aircall_id: "id 2",
			number_id: "number id 2",
		},
		{
			name: "name 3",
			email: "email 3",
			aircall_id: "id 3",
			number_id: "number id 3",
		}
	]);
};

// start outbound call
async function startOutboundCall() {
  try {
    // grab active user Aircall ID
    if(activeUserAircallId === '') {
      const allUsers = await getUsers();
      // Logger.log(allUsers);
      for(u = 0; u < allUsers.length; u++) {
        if(allUsers[u].email === activeUserEmail) {
          activeUserAircallId = parseInt(allUsers[u].id);
          break;
        };
      };
    };
    // grab user Aircall number
    const userConfiguration = aircallConfiguration();
    // Logger.log(userConfiguration.find(u => u.email.toString() === activeUserEmail.toString()).number_id);
    const phoneTo = (ss.getActiveCell().getDisplayValue().includes("+") == true) ? ss.getActiveCell().getDisplayValue() : '+' + ss.getActiveCell().getDisplayValue();
    const data = {
      number_id: userConfiguration.find(u => u.email.toString() === activeUserEmail.toString()).number_id,
      to: phoneTo
    };
    // Logger.log(baseUrl + 'users/'+activeUserAircallId+'/calls');
    // Logger.log(JSON.stringify(data));
    let req = await UrlFetchApp.fetch(baseUrl + 'users/'+activeUserAircallId+'/calls', {
      method: 'POST', // *GET, POST, PUT, DELETE, etc.
      headers: {
        'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
        'Content-Type': 'application/json', // sending JSON data
      },
      'muteHttpExceptions': true, // prevent Google alerts with 400 / 500 status codes
      payload: JSON.stringify(data) // body data type must match 'Content-Type' header
    });
    if(req.getResponseCode() != 204) ui.alert('👎👎👎Error👎👎👎\r\nCant start Outbound Call\r\n\r\n'+req.getContentText());
    // else ui.alert('👍👍👍Success👍👍👍\r\nOutbound call started');
  } catch (e){
    ui.alert('👎👎👎Error👎👎👎\r\nCant start Outbound Call\r\n\r\n'+e);
  };
};

// click to dial
async function clickToDial() {
  try {
    // grab active user Aircall ID
    if(activeUserAircallId === '') {
      const allUsers = await getUsers();
      // Logger.log(allUsers);
      activeUserEmail = 'koen.verduijn+demo@aircall.io'; // !!!remove only test account !!! 
      for(u = 0; u < allUsers.length; u++) {
        if(allUsers[u].email === activeUserEmail) {
          activeUserAircallId = parseInt(allUsers[u].id);
          break;
        };
      };
    };
    const phoneTo = (ss.getActiveCell().getDisplayValue().includes("+") == true) ? ss.getActiveCell().getDisplayValue() : '+' + ss.getActiveCell().getDisplayValue();
    // grab user Aircall number
    const data = {
      to: phoneTo
    };
    let req = await UrlFetchApp.fetch(baseUrl + 'users/'+activeUserAircallId+'/dial', {
      method: 'POST', // *GET, POST, PUT, DELETE, etc.
      headers: {
        'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
        'Content-Type': 'application/json', // sending JSON data
      },
      'muteHttpExceptions': true, // prevent Google alerts with 400 / 500 status codes
      payload: JSON.stringify(data) // body data type must match 'Content-Type' header
    });
    if(req.getResponseCode() != 204) ui.alert('👎👎👎Error👎👎👎\r\nCant dial Outbound Call\r\n\r\n'+req.getContentText());
    // else ui.alert('👍👍👍Success👍👍👍\r\nOutbound call dialed');
  } catch (e){
    ui.alert('👎👎👎Error👎👎👎\r\nCant dial Outbound Call\r\n\r\n'+e);
  };
};