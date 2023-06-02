// add Aircall Menu
function onOpen() {
	ui.createMenu("🚀 Aircall 🚀")
		.addItem("Aircall Softphone", "showAircallSoftphone")
		.addSeparator()
		.addItem("Create Dialer Campaign", "createDialerCampaign")
		.addSubMenu(
			ui
				.createMenu("Delegate Dialer Campaign")
				.addItem("Name1", "delegateDialerCampaignName1")
				.addItem("Name2", "delegateDialerCampaignName2")
				.addItem("Name3", "delegateDialerCampaignName3")
		)
		.addToUi();
}

// add Aircall phone app in sidebar
function showAircallSoftphone() {
	const aircallSoftphone = HtmlService.createHtmlOutputFromFile("Sidebar");
	aircallSoftphone.setTitle("Aircall Phone");
	SpreadsheetApp.getUi().showSidebar(aircallSoftphone);
};

// create different dialer campaigns
function delegateDialerCampaignName1() {
  delegateDialerCampaign('Name 1');
};
function delegateDialerCampaignName2() {
  delegateDialerCampaign('Name 2');
};
function delegateDialerCampaignName3() {
  delegateDialerCampaign('Name 3');
};
// create Dialer Campaign for logged in user
async function createDialerCampaign() {
  let cleanedPhoneNumbers = [];
  for(p = 0;p < selectedPhoneNumbers.length; p++) {
    if(selectedPhoneNumbers[p][0] !== '') {
      let cleanPhoneNumber = selectedPhoneNumbers[p];
      cleanPhoneNumber = cleanPhoneNumber[0].toString();
      Logger.log(cleanPhoneNumber);
      if(cleanPhoneNumber.includes("+") != true) cleanPhoneNumber = '+' + cleanPhoneNumber;
      //Logger.log(cleanPhoneNumber);
      cleanedPhoneNumbers.push(cleanPhoneNumber);
    }
  };
  // Logger.log(cleanedPhoneNumbers);
  try {
    const allUsers = await getUsers();
    // Logger.log(allUsers);
    for(u = 0; u < allUsers.length; u++) {
      // Logger.log(allUsers[u].email);
      if(allUsers[u].email.toString() === activeUserEmail.toString()) {
        activeUserAircallId = parseInt(allUsers[u].id);
        break;
      };
    };
    if(activeUserAircallId === '') ui.alert('👎👎👎Error👎👎👎\r\nCant create the Dialer Campaign\r\n\r\n'+'No Aircall user found with Google user: '+activeUserEmail);
    else {
      // Logger.log(activeUserAircallId);
      // check if existing dialer campaign for user
      res = await UrlFetchApp.fetch(baseUrl + 'users/'+activeUserAircallId+'/dialer_campaign', {
        method: 'GET', // *GET, POST, PUT, DELETE, etc.
        headers: {
          'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
          'Content-Type': 'application/json', // sending JSON data
        },
        muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
        //payload: JSON.stringify(data) // body data type must match 'Content-Type' header
      });
      // Logger.log('existing dialer campaign: '+res.getResponseCode());
      if(res.getResponseCode() !== 404) {
        // delete existing dialer campaign for user
        let res = await UrlFetchApp.fetch(baseUrl + 'users/'+activeUserAircallId+'/dialer_campaign', {
          method: 'DELETE', // *GET, POST, PUT, DELETE, etc.
          headers: {
            'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
            'Content-Type': 'application/json', // sending JSON data
          },
          'muteHttpExceptions': true, // prevent Google alerts with 400 / 500 status codes
          //payload: JSON.stringify(data) // body data type must match 'Content-Type' header
        });
        // Logger.log('delete dialer campaign: '+res.getResponseCode());
      };
      // create new dialer campaign
      const data = {
        'phone_numbers': cleanedPhoneNumbers
      };
      //Logger.log(JSON.stringify(data));
      res = await UrlFetchApp.fetch(baseUrl + 'users/'+activeUserAircallId+'/dialer_campaign', {
        method: 'POST', // *GET, POST, PUT, DELETE, etc.
        headers: {
          'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
          'Content-Type': 'application/json', // sending JSON data
        },
        'muteHttpExceptions': true, // prevent Google alerts with 400 / 500 status codes
        payload: JSON.stringify(data) // body data type must match 'Content-Type' header
      });
      Logger.log('Aircall API response: '+res.getResponseCode()+' and JSON response: '+res.getContentText());
      if(res.getResponseCode() == 204 || res.getResponseCode() == 422) {
        // ui.alert('👍👍👍Success👍👍👍\r\nDialer Campaign created!');
      } else ui.alert('👎👎👎Error👎👎👎\r\nCant create the Dialer Campaign\r\n\r\n'+res.getContentText());
    }
  } catch (error) {
    ui.alert('👎👎👎Error👎👎👎\r\nCant create the Dialer Campaign\r\n\r\n'+error);
    // deal with any errors
    // Logger.log(error);
  };
};
// user dictionary for other agents
const userDict = [
  {
    name: 'name 1',
    email: 'email 1',
    aircall_id: 'id 1',
    number_id: 'number id 1'
  },
  {
    name: 'name 2',
    email: 'email 2',
    aircall_id: 'id 2',
    number_id: 'number id 2'
  },
  {
    name: 'name 3',
    email: 'email 3',
    aircall_id: 'id 3',
    number_id: 'number id 3'
  }
];

// give dialer campaign to someone else
async function delegateDialerCampaign(agent) {
  // Logger.log(selectedPhoneNumbers);
  // Logger.log(agent);
  let cleanedPhoneNumbers = [];
  for(p = 0;p < selectedPhoneNumbers.length; p++) {
    if(selectedPhoneNumbers[p][0] !== '') {
      let cleanPhoneNumber = selectedPhoneNumbers[p];
      cleanPhoneNumber = cleanPhoneNumber[0];
      if(cleanPhoneNumber.includes("+") != true) cleanPhoneNumber = '+' + cleanPhoneNumber;
      //Logger.log(cleanPhoneNumber);
      cleanedPhoneNumbers.push(cleanPhoneNumber);
    }
  };
  // Logger.log(cleanedPhoneNumbers);
  try {
    // check if existing dialer campaign for user
    let res = await UrlFetchApp.fetch(baseUrl + 'users/'+userDict.find(u => u.name === agent).aircall_id+'/dialer_campaign', {
      method: 'GET', // *GET, POST, PUT, DELETE, etc.
      headers: {
        'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
        'Content-Type': 'application/json', // sending JSON data
      },
      muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
      //payload: JSON.stringify(data) // body data type must match 'Content-Type' header
    });
    // Logger.log('existing dialer campaign: '+res.getResponseCode());
    if(res.getResponseCode() !== 404) {
      // delete existing dialer campaign for user
      let res = await UrlFetchApp.fetch(baseUrl + 'users/'+userDict.find(u => u.name === agent).aircall_id+'/dialer_campaign', {
        method: 'DELETE', // *GET, POST, PUT, DELETE, etc.
        headers: {
          'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
          'Content-Type': 'application/json', // sending JSON data
        },
        'muteHttpExceptions': true, // prevent Google alerts with 400 / 500 status codes
        //payload: JSON.stringify(data) // body data type must match 'Content-Type' header
      });
      // Logger.log('delete dialer campaign: '+res.getResponseCode());
    };
    // create new dialer campaign
    const data = {
      'phone_numbers': cleanedPhoneNumbers
    };
    //Logger.log(JSON.stringify(data));
    res = await UrlFetchApp.fetch(baseUrl + 'users/'+userDict.find(u => u.name === agent).aircall_id+'/dialer_campaign', {
      method: 'POST', // *GET, POST, PUT, DELETE, etc.
      headers: {
        'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
        'Content-Type': 'application/json', // sending JSON data
      },
      'muteHttpExceptions': true, // prevent Google alerts with 400 / 500 status codes
      payload: JSON.stringify(data) // body data type must match 'Content-Type' header
    });
    //Logger.log('Aircall API response: '+res.getResponseCode()+' and JSON response: '+res.getContentText());
    if(res.getResponseCode() !== 204) ui.alert('👎👎👎Error👎👎👎\r\nCant create the Dialer Campaign\r\n\r\n'+res.getContentText());
    // else ui.alert('👍👍👍Success👍👍👍\r\nDialer Campaign created!');
  } catch (error) {
    ui.alert('👎👎👎Error👎👎👎\r\nCant create the Dialer Campaign\r\n\r\n'+error);
    // deal with any errors
    // Logger.log(error);
  };
};
