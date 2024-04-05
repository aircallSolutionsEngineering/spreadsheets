// base information
// const ss = SpreadsheetApp.getActive();
// const ui = SpreadsheetApp.getUi();
const baseUrl = "https://api.aircall.io/v1/";
const scriptProperties = PropertiesService.getScriptProperties();
const apiId = scriptProperties.getProperty("apiId");
const apiToken = scriptProperties.getProperty("apiToken");
// check if API Token are working correctly
//Logger.log('apiId: '+apiId+' Token: '+apiToken+'\n'+Utilities.base64Encode(apiId+':'+apiToken));
// let activeUserEmail = Session.getActiveUser();

// scheduler to work with the Apps trigger
function triggerAircallCallData() {
  /* get interval data */
  const dateTimeNowInSeconds = Math.floor(Date.now());
  const dateTimeNowFormat = new Date(dateTimeNowInSeconds);
  const dateTimeNowDay = dateTimeNowFormat.getDay();
  const dateTimeNowHour = dateTimeNowFormat.getHours();
  const dateTimeNowMinute = dateTimeNowFormat.getMinutes();
  // Logger.log(dateTimeNowInSeconds);
  const dateTimeAPIQueryTo = Math.floor(dateTimeNowInSeconds / 1000);
  const dateTimeAPIQueryFrom = dateTimeAPIQueryTo - 5 * 60;
  // Logger.log("day of the week: "+dateTimeNowDay+" and the hour: "+ dateTimeNowHour);
  if (dateTimeNowDay == 1 && dateTimeNowHour == 0 && dateTimeNowMinute < 6) getAll("calls", dateTimeAPIQueryFrom, dateTimeAPIQueryTo, true);
  else getAll("calls", dateTimeAPIQueryFrom, dateTimeAPIQueryTo, false);
}

// to get as much historic data into sheet at once
async function historicAircallCallDataImporter() {
  /* get specific date range */
  const dateFrom = new Date("2024-02-19").getTime() / 1000;
  const dateTo = new Date("2024-02-21").getTime() / 1000;
  for (let d = dateFrom; d <= dateTo; d = d + 60 * 60 * 24) {
    await getAll("calls", d, d + 60 * 60 * 24, false);
  }
}

// generic API GET request to all Aircall objects
async function getAircallData(apiUrl, object) {
  const res = await UrlFetchApp.fetch(apiUrl, {
    method: "GET", // *GET, POST, PUT, DELETE, etc.
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(apiId + ":" + apiToken), // authorization header
      "Content-Type": "application/json", // sending JSON data
      muteHttpExceptions: true, // prevent Google alerts with 400 / 500 status codes
    },
    //payload: JSON.stringify(data) // body data type must match "Content-Type" header
  });
  data = JSON.parse(res);
  if (data[object] == null) Logger.log("Trouble with connecting to the Aircall " + object + " API, status code: " + res.status + " and body: " + res.data);
  return data;
}

// generic data request handler to point to clean sheet and place data in the sheet
async function getAll(object, dateFrom, dateTo, dataOverwrite) {
  const objectName = object.charAt(0).toUpperCase() + object.substring(1, object.length);
  let dataSheet = SpreadsheetApp.getActive().getSheetByName(object + " data");
  try {
    if (!dataSheet) {
      SpreadsheetApp.getActive().insertSheet(object + " data");
      dataSheet = SpreadsheetApp.getActive().getSheetByName(object + " data");
      Logger.log("New Data Sheet created: " + object + " data");
    }
    if (dataOverwrite === true) {
      Logger.log("Removing all " + object + " data");
      dataSheet.clear();
      if (dataSheet.getMaxRows() > 100) dataSheet.deleteRows(100, dataSheet.getMaxRows() - 100);
    }
    let dateFilter = "";
    if (object === "calls") dateFilter = "&from=" + dateFrom + "&to=" + dateTo;
    let page = 1;
    let objectRecords = [];
    let nextUrl = "";
    let data = null;
    while (nextUrl != null) {
      let apiUrl = baseUrl + object + "?per_page=50&order=asc&page=" + page + dateFilter;
      data = await getAircallData(apiUrl, object);
      objectRecords = data[object];
      const pages = Math.ceil(data["meta"]["total"] / 50);
      if (data["meta"]["total"] > 9999) {
        Logger.log("More than " + data["meta"]["total"] + " " + object + " data. Please reduce the time period");
        break;
      }
      if (objectRecords.length === 0) {
        Logger.log("No " + object + " available in (" + new Date(dateFrom * 1000).toLocaleString() + " - " + new Date(dateTo * 1000).toLocaleString() + ")");
        break;
      }
      if (page === 1) Logger.log("Importing " + data["meta"]["total"] + " " + object + " data (" + new Date(dateFrom * 1000).toLocaleString() + " - " + new Date(dateTo * 1000).toLocaleString() + ") data from " + pages + (pages > 1 ? " pages." : " page.") + "\nThis will take roughly " + Math.ceil((2 * pages) / 60) + " minutes");
      Logger.log("collecting " + data["meta"]["total"] + " " + object + " from " + page + "/" + pages);
      let finalRecords = [];
      if (
        (dataOverwrite === true && page === 1) ||
        SpreadsheetApp.getActive()
          .getSheetByName(object + " data")
          .getLastRow() === 0
      ) {
        const cols = Object.keys(objectRecords[0]);
        const printCols = [];
        cols.forEach(function (col) {
          printCols.push([col]);
        });
        // additional Google Sheet column headerss
        printCols.push(["dateTime"]);
        printCols.push(["date"]);
        printCols.push(["week"]);
        printCols.push(["month"]);
        printCols.push(["talkingTime"]);
        printCols.push(["waitingTime"]);
        printCols.push(["number"]);
        printCols.push(["numberCountryCode"]);
        printCols.push(["tags"]);
        printCols.push(["agent"]);
        printCols.push(["agentEmail"]);
        printCols.push(["agentNumber"]);
        printCols.push(["team"]);
        printCols.push(["localTime"]);
        // set column headers
        if (cols.length > 0) {
          finalRecords.push(printCols);
        }
      }
      // determine column length for number of spreadsheet columns
      let maxRecordColumns = finalRecords[0] != null ? finalRecords[0].length : 0;
      for (r = 0; r < objectRecords.length; r++) {
        let recordRow = [];
        const recordFields = Object.values(objectRecords[r]);
        for (let rf = 0; rf < 26; rf++) recordRow.push([recordFields[rf]]);
        // Additional Google Sheet formulas or specific data for the additional columns
        recordRow.push(["=EPOCHTODATE(F:F)"]);
        recordRow.push(['=LEFT(AA:AA,(FIND(" ",AA:AA,1)-1))']);
        recordRow.push(["=WEEKNUM(AB:AB)"]);
        recordRow.push(["=MONTH(AB:AB)"]);
        recordRow.push(['=IF(D:D<>"done",0,IF(G:G<>"",H:H-G:G,0))']);
        recordRow.push(['=IF(D:D<>"done",0,IF(G:G<>"",G:G-F:F,I:I))']);
        recordRow.push(objectRecords[r]["number"]["name"]);
        recordRow.push(objectRecords[r]["number"] != null && objectRecords[r]["number"]["country"] != null ? objectRecords[r]["number"]["country"] : "");
        let tags = "";
        if (objectRecords[r]["tags"].length != 0) {
          for (let t = 0; t < objectRecords[r]["tags"].length; t++) {
            if (t === 0) tags = objectRecords[r]["tags"][t]["name"];
            else tags = tags + ", " + objectRecords[r]["tags"][t]["name"];
          }
        }
        recordRow.push([tags]);
        let agentName = "";
        let agentEmail = "";
        if (objectRecords[r]["user"] != null) {
          // console.log(objectRecords[r]["user"]);
          agentEmail = objectRecords[r]["user"]["email"];
          agentName = objectRecords[r]["user"]["name"];
        }
        recordRow.push([agentName]);
        recordRow.push([agentEmail]);
        recordRow.push(objectRecords[r]["number"] != null && objectRecords[r]["number"]["digits"] != null ? objectRecords[r]["number"]["digits"] : "");
        let teamName = "";
        if (objectRecords[r]["teams"].length > 0) {
          // console.log(objectRecords[r]["teams"]);
          teamName = objectRecords[r]["teams"][0]["name"];
        }
        recordRow.push([teamName]);
        const localTime = Utilities.formatDate(new Date(objectRecords[r]["started_at"] * 1000), "America/Chicago", "yyyy/MM/dd G 'at' HH:mm:ss z");
        recordRow.push(localTime);
        finalRecords.push(recordRow);
        if (recordRow.length > maxRecordColumns) maxRecordColumns = recordRow.length;
      }
      const numberOfRows =
        SpreadsheetApp.getActive()
          .getSheetByName(object + " data")
          .getLastRow() === 0
          ? 1
          : SpreadsheetApp.getActive()
              .getSheetByName(object + " data")
              .getLastRow() + 1;
      // Logger.log('writing from row: '+numberOfRows);
      /* send out notifications if getting to the 10 million cell per Workbook limit
      if(numberOfRows > (9500000 / maxRecordColumns)) {
        SpreadsheetApp.getActive().getSheetByName(object + " data").setTabColor("ff0000");
        MailApp.sendEmail({
          to: '',
          subject: 'Google Sheet with Aircall data: call data is close to limit',
          htlmBody: 'Hi,\nthe Google Sheet with the Aircall call data: https://docs.google.com/spreadsheets/d/ is reaching more than 9.5 million cells. Please start a new Worksheet.'
        });
      };
      */
      SpreadsheetApp.getActive()
        .getSheetByName(object + " data")
        .getRange(numberOfRows, 1, finalRecords.length, maxRecordColumns)
        .setValues(finalRecords);
      page += 1;
      nextUrl = data["meta"]["next_page_link"];
    }
    if (objectRecords.length > 1) Logger.log("Importing " + object + " data complete!");
  } catch (error) {
    // deal with any errors
    Logger.log(error);
  }
}
