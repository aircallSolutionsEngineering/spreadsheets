// base information
// const ss = SpreadsheetApp.getActiveSheet();
// const ui = SpreadsheetApp.getUi();
const baseUrl = "https://api.aircall.io/v1/";
const scriptProperties = PropertiesService.getScriptProperties();
const apiId = scriptProperties.getProperty("apiId");
const apiToken = scriptProperties.getProperty("apiToken");
// check if API Token are working correctly
//Logger.log('apiId: '+apiId+' Token: '+apiToken+'\n'+Utilities.base64Encode(apiId+':'+apiToken));
let activeUserEmail = Session.getActiveUser();
