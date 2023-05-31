// get all users
async function getUsers() {
	let users = [];
	try {
		let req = await UrlFetchApp.fetch(
			PropertiesService.getScriptProperties().getProperty("baseUrl") +
				"users?per_page=50",
			{
				method: "GET", // *GET, POST, PUT, DELETE, etc.
				headers: {
					Authorization:
						"Basic " +
						Utilities.base64Encode(
							PropertiesService.getScriptProperties().getProperty("apiId") +
								":" +
								PropertiesService.getScriptProperties().getProperty("apiToken")
						), // authorization header
					"Content-Type": "application/json", // sending JSON data
				},
				muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
				//payload: JSON.stringify(data) // body data type must match 'Content-Type' header
			}
		);
		// Logger.log('existing dialer campaign: '+res.getResponseCode());
		if (req.getResponseCode() !== 200)
			ui.alert(
				"👎👎👎Error👎👎👎\r\nCant grab all the Users\r\n\r\n" +
					res.getContentText()
			);
		else {
			// ui.alert('👍👍👍Success👍👍👍\r\nAll Users');
			let res = JSON.parse(req.getContentText());
			users = res.users;
			// Logger.log(res.meta);
			if (res.meta.next_page_link != null) {
				for (p = 2; p < Math.ceil(res.meta.count / res.meta.count); p++) {
					req = await UrlFetchApp.fetch(
						PropertiesService.getScriptProperties().getProperty("baseUrl") +
							"users?per_page=50&page=" +
							p,
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
					users += users.push(JSON.parse(req.getContextText()).users);
				}
			}
		}
		return users;
	} catch (error) {
		ui.alert("👎👎👎Error👎👎👎\r\nCant create the Users\r\n\r\n" + error);
		// deal with any errors
		// Logger.log(error);
	}
}

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
		},
		{
			name: "Koen Verduijn",
			email: "koen.verduijn+demo@aircall.io",
			aircall_id: "731416",
			number_id: "397375",
		},
	]);
};

async function getAllCalls() {
  getAll('calls');
};
async function tagCall() {
  let callId = ss.getCurrentCell();
  if(callId.getColumn() != 1 || callId.getRow() < 9) ui.prompt('no call ID selected');
  else {
    try {
      callId = callId.getValue();
      const result = ui.prompt("Please add tag(s):");
      //Get the button that the user pressed.
      const button = result.getSelectedButton();
      if (button === ui.Button.OK) {
        const data = { tags: [String(result.getResponseText())]};
        const res = await UrlFetchApp.fetch(baseUrl + 'calls/'+String(callId)+'/tags', {
          method: 'POST', // *GET, POST, PUT, DELETE, etc.
          headers: {
            'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
            'Content-Type': 'application/json', // sending JSON data
            'muteHttpExceptions': false, // prevent Google alerts with 400 / 500 status codes
          },
          payload: JSON.stringify(data) // body data type must match "Content-Type" header
        });
        Logger.log(res.getResponseCode());
        if(res.getResponseCode() != 201) ui.alert('Tag could not be added')
        else ui.alert('Tag; '+result.getResponseText()+' was added to call: '+callId);
        Logger.log(ret);
      }
    } catch (error) {
      // deal with any errors
      Logger.log(error);
    };
  };
};
async function commentCall() {
  let callId = ss.getCurrentCell();
  if(callId.getColumn() != 1 || callId.getRow() < 9) ui.prompt('no call ID selected');
  else {
    try {
      callId = callId.getValue();
      const result = ui.prompt("Please add comment:");
      //Get the button that the user pressed.
      const button = result.getSelectedButton();
      if (button === ui.Button.OK) {
        const data = { content: String(result.getResponseText())};
        const res = await UrlFetchApp.fetch(baseUrl + 'calls/'+String(callId)+'/comments', {
          method: 'POST', // *GET, POST, PUT, DELETE, etc.
          headers: {
            'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
            'Content-Type': 'application/json', // sending JSON data
            'muteHttpExceptions': false, // prevent Google alerts with 400 / 500 status codes
          },
          payload: JSON.stringify(data) // body data type must match "Content-Type" header
        });
        Logger.log(res.getResponseCode());
        if(res.getResponseCode() != 201) ui.alert('Tag could not be added')
        else ui.alert('Comment; '+result.getResponseText()+' was added to call: '+callId);
        Logger.log(ret);
      }
    } catch (error) {
      // deal with any errors
      Logger.log(error);
    };
  };
};

async function getAllNumbers() {
  getAll("numbers");
}

async function getAllTags() {
  getAll("tags");
};
async function createTag() {
  const checkTag = await createRecord("tag");
  const options = {
    method: 'POST',
    muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
    'Content-Type': 'application/json',
    headers: {
      'Authorization': 'Basic '+Utilities.base64Encode(apiId+':'+apiToken),
    },
    payload: JSON.stringify(checkTag) // body data type must match "Content-Type" header
  }
  const sendPost = await UrlFetchApp.fetch(baseUrl+'tags',options);
  Logger.log(sendPost.getContentText());
  if(sendPost.getResponseCode() != 201) {
    ss.getRange('K4').setValue('status');
    ss.getRange('L4').setValue(JSON.stringify(sendPost.getContentText()));
  } else {
    ss.getRange('K4').setValue('status');
    ss.getRange('L4').setValue('created');
  };
};
async function updateTag() {
  const objectId = Math.floor(ss.getRange(ss.getCurrentCell().getRow(),1,1,1).getValue());
  const payloadBody = await updateRecord("tag");
  Logger.log(payloadBody);
  const res = await UrlFetchApp.fetch(baseUrl+'tags/'+String(objectId), {
      method: 'PUT', // *GET, POST, PUT, DELETE, etc.
      contentType: 'application/json', // sending JSON data
      muteHttpExceptions: false, // prevent Google alerts with 400 / 500 status codes
      headers: {
        'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
      },
      payload: JSON.stringify(payloadBody) // body data type must match "Content-Type" header
    });
    if(res.getResponseCode() != 200) ui.prompt('Error in updating tag: '+objectId);
    else {
      ss.getRange(ss.getCurrentCell().getRow(),1,1,1).setBackground('green');
    };
};
async function deleteTag() {
  if(ss.getCurrentCell().getColumn() != 1 || ss.getCurrentCell().getRow() < 9) ui.prompt('Please select an ID from column A');
  else deleteRecord("tag");
};

async function getAll(object) {
  const lastRow = ss.getLastRow();
  if(lastRow != 0) ss.getRange(8,1,lastRow,25).clear();
  try {
    const from = ss.getRange('L5').isBlank() == false ? '&from='+(ss.getRange('L2').getValue().getTime() / 1000) : '';
    const to = ss.getRange('L6').isBlank() == false ? '&to='+(ss.getRange('L3').getValue().getTime() / 1000) : '';
    // Default options are marked with *
    const res = await UrlFetchApp.fetch(baseUrl + object+'?per_page=50&order=desc'+from+to, {
      method: 'GET', // *GET, POST, PUT, DELETE, etc.
      headers: {
        'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
        'Content-Type': 'application/json', // sending JSON data
        'muteHttpExceptions': true, // prevent Google alerts with 400 / 500 status codes
      },
      //payload: JSON.stringify(data) // body data type must match "Content-Type" header
    });
    const data = JSON.parse(res);
    if(data[object] == null) ui.alert('Trouble with connecting to the Aircall '+object+' API, status code: '+res.status+' and body: '+res.data);
    else {
      const objectRecords = data[object];
      if(objectRecords.length === 0) ui.alert('No '+object+' available');
      // calculate the number of rows and columns needed
      var numRows = objectRecords.length;
      const cols = Object.keys(objectRecords[0]);
      const printCols = []
      cols.forEach(function(col) { printCols.push(col);});
      // set column headers
      if (cols.length > 0) {
        ss.getRange(8, 1, 1, cols.length).setValues([printCols]);
      };
      // set rows
      for(r=0;r<objectRecords.length;r++) {
        const finalRecords = [];
        const recordFields = Object.values(objectRecords[r]);
        recordFields.forEach(function(recordField) { finalRecords.push([recordField])});
        ss.getRange(9+r,1,1,recordFields.length).setValues([finalRecords]);
      }
    }
  } catch (error) {
    // deal with any errors
    Logger.log(error);
  };
};
const createDict = [
  {
    "name": "tag",
    "fields": ["name","color","description"]
  },
  {
    "name": "user",
    "fields": ["email","first_name","last_name","availability","admin","time_zone","language","wrap_up_time"]
  },
  {
    "name": "contact",
    "fields": ["first_name","last_name","company_name","information","emails.label","emails.value","phone_numbers.label","phone_numbers.value"]
  }
];
async function createRecord(object) {
  if(object != "user" && object != "tag" && object != "contact") ui.alert('please provide correct object. '+object+' is not valid');
  else {
    const objectFields = createDict.find(i => i.name === object).fields;
    if(ss.getRange(5,1,1,objectFields.length).isBlank() == true) ss.getRange(5,1,1,objectFields.length).setValues([objectFields]);
    else {
      const objectValues = ss.getRange(6,1,1,objectFields.length).getValues();
      if(ss.getRange(6,1,1,objectFields.length).isBlank() === true) ui.prompt('Please provide data to create the '+object);
      else{
        const payloadBody = {};
        for(x=0;x<objectFields.length;x++) {
          if(objectValues[0][x] == '') ui.prompt('Please provide data for the '+object+' in the field '+objectFields[x]);
          else {
            const payloadName = objectFields[x];
            payloadBody[payloadName] = objectValues[0][x]
          }
        };
        return payloadBody;
      };
    }
  }
};
async function updateRecord(object) {
  if(object != "user" && object != "tag" && object != "contact") ui.alert('please provide correct object. '+object+' is not valid');
  else {
    const objectRow = ss.getCurrentCell().getRow();
    ss.getRange(objectRow,1,1,1).clearFormat();
    const objectValues = ss.getRange(objectRow,1,1,ss.getLastColumn()).getValues();
    const objectHeaders = ss.getRange(8,1,1,ss.getLastColumn()).getValues();
    const payloadBody = {};
    for(x=0;x<objectHeaders[0].length;x++) {
      if(objectHeaders[0][x].length < 2) { break };
      const payloadName = objectHeaders[0][x];
      if(object =='user' & payloadName == 'name') {
        Logger.log(objectValues[0][x].split(' ')[0]);
        Logger.log(objectValues[0][x].substring(objectValues[0][x].indexOf(' ')+1));
        payloadBody['first_name'] = objectValues[0][x].split(' ')[0];
        payloadBody['last_name'] = objectValues[0][x].substring(objectValues[0][x].indexOf(' ')+1);
      } else{
        payloadBody[payloadName] = objectValues[0][x];
      };
    };
    return payloadBody;
  };
};
async function deleteRecord(object) {
  ss.getRange('L4').clear();
  if(object != "user" && object != "tag" && object != "contact") ui.alert('please provide correct object. '+object+' is not valid');
  else {
    const recordId = ss.getCurrentCell().getValue();
    const res = await UrlFetchApp.fetch(baseUrl + object+'s/'+recordId, {
      method: 'DELETE', // *GET, POST, PUT, DELETE, etc.
      headers: {
        'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
        'Content-Type': 'application/json', // sending JSON data
        'muteHttpExceptions': true, // prevent Google alerts with 400 / 500 status codes
      },
      //payload: JSON.stringify(data) // body data type must match "Content-Type" header
    });
    if(res.getResponseCode() != 204) ui.prompt('Cant delete the '+object);
    else {
      ss.getRange('K4').setValue('status');
      ss.getRange('L4').setValue('deleted');
    };      
  };
};

function onOpen() {
  ui
    .createMenu('🚀 Aircall 🚀')
    .addItem('Create', 'showSidebar')
    .addToUi();
};

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Select Object');
  ui.showSidebar(html);
};

async function getAllUsers() {
  getAll("users");
};
async function createUser() {
  createRecord("user");
};
async function deleteUser() {
  if(ss.getCurrentCell().getColumn() != 1 || ss.getCurrentCell().getRow() < 9) ui.prompt('Please select an ID from column A');
  else deleteRecord("user");
};
async function updateUser() {
  const objectId = Math.floor(ss.getRange(ss.getCurrentCell().getRow(),1,1,1).getValue());
  const payloadBody = await updateRecord("user");
  const res = await UrlFetchApp.fetch(baseUrl+'users/'+String(objectId), {
      method: 'PUT', // *GET, POST, PUT, DELETE, etc.
      contentType: 'application/json', // sending JSON data
      muteHttpExceptions: false, // prevent Google alerts with 400 / 500 status codes
      headers: {
        'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
      },
      payload: JSON.stringify(payloadBody) // body data type must match "Content-Type" header
    });
    if(res.getResponseCode() != 200) ui.prompt('Error in updating user: '+objectId);
    else {
      ss.getRange(ss.getCurrentCell().getRow(),1,1,1).setBackground('green');
    };
};

function getAllContacts() {
  getAll("contacts");
};
function createContact() {
  createRecord("contact");
};
async function deleteTag() {
  if(ss.getCurrentCell().getColumn() != 1 || ss.getCurrentCell().getRow() < 9) ui.prompt('Please select an ID from column A');
  else deleteRecord("contact");
};
async function updateContact() {
  const objectId = Math.floor(ss.getRange(ss.getCurrentCell().getRow(),1,1,1).getValue());
  const payloadBody = await updateRecord("contact");
  const res = await UrlFetchApp.fetch(baseUrl+'contacts/'+String(objectId), {
      method: 'POST', // *GET, POST, PUT, DELETE, etc.
      contentType: 'application/json', // sending JSON data
      muteHttpExceptions: false, // prevent Google alerts with 400 / 500 status codes
      headers: {
        'Authorization': 'Basic '+ Utilities.base64Encode(apiId+':'+apiToken), // authorization header
      },
      payload: JSON.stringify(payloadBody) // body data type must match "Content-Type" header
    });
    if(res.getResponseCode() != 200) ui.prompt('Error in updating contact: '+objectId);
    else {
      ss.getRange(ss.getCurrentCell().getRow(),1,1,1).setBackground('green');
    };
};