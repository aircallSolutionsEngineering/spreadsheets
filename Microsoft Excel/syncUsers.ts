// base information
const baseUrl = "https://api.aircall.io/v1/";
const apiId = "<Aircall API ID>";
const apiToken = "<Aircall API Token>";
const auth = btoa(apiId + ":" + apiToken);
// check if API Token are working correctly
// console.log("apiId: " + apiId + " Token: " + apiToken + "\n" + auth);

async function main(workbook: ExcelScript.Workbook) {
  const userList: string[] = await listRecords("users");
  // console.log(userList);
  await addRecords("users", userList, workbook);
}

/// get all users or teams or numbers or contacts
async function listRecords(object: string) {
  if (object != "users" && object != "teams" && object != "numbers" && object != "contacts") console.log("incorrect object: " + object + " is not part of Aircall APIs");
  else {
    let records: string[] = [];
    try {
      let threshold = 1;
      for (let p = 1; p <= threshold; p++) {
        let req = await fetch(baseUrl + object + "?per_page=50&page=" + p, {
          method: "GET", // *GET, POST, PUT, DELETE, etc.
          headers: {
            Authorization: "Basic " + auth, // authorization header
            "Content-Type": "application/json", // sending JSON data
          },
          //payload: JSON.stringify(data) // body data type must match 'Content-Type' header
        });
        const res: object = await req.json();
        // console.log(res);
        if (req.status !== 200) console.log("👎👎👎Error👎👎👎\r\nCant grab all the " + object + "\r\n\r\n" + req.body);
        else {
          // console.log('👍👍👍Success👍👍👍\r\nAll '+objects);
          // console.log("test"+res);
          threshold = Math.ceil(res["meta"]["total"] / 50);
          records = records.concat(res[object]);
          // console.log(records.length);
        }
      }
      return records;
    } catch (error) {
      console.log("👎👎👎Error👎👎👎\r\nCant list the " + object + "\r\n\r\n" + error);
    }
  }
}

// add all users or teams or numbers or contacts
async function addRecords(object: string, data: string[], workbook: ExcelScript.Workbook) {
  if (object != "users" && object != "teams" && object != "numbers" && object != "contacts") console.log("incorrect object: " + object + " is not part of Aircall APIs");
  else {
    if (workbook.getWorksheet(object) == null) workbook.addWorksheet(object);
    let objectWorksheet = workbook.getWorksheet(object);
    objectWorksheet.getRange().clear();
    const objectRecords = data;
    if (objectRecords.length === 0) console.log("No " + object + " available");
    // calculate the number of rows and columns needed
    var numRows = objectRecords.length;
    const cols: string[] = Object.keys(objectRecords[0]);
    const printCols: string[] = [];
    for (let c: number = 0; c < cols.length; c++) printCols.push(cols[c]);
    // set column headers
    if (cols.length > 0) {
      objectWorksheet.getRangeByIndexes(0, 0, 1, cols.length).setValues([printCols]);
    }
    // set rows
    const finalRecords: string[][] = [];
    for (let r: number = 0; r < objectRecords.length; r++) {
      let finalRow: string[] = [];
      const recordFields: string[] = Object.values(objectRecords[r]);
      for (let rf: number = 0; rf < recordFields.length; rf++) {
        // console.log(recordField)
        if (Array.isArray(recordFields[rf]) == true) finalRow.push(JSON.stringify(recordFields[rf]));
        else finalRow.push(recordFields[rf]);
      }
      // console.log(finalRow);
      finalRecords.push(finalRow);
    }
    // console.log(finalRecords);
    // add sheet if not available
    objectWorksheet.getRangeByIndexes(1, 0, objectRecords.length, cols.length).setValues(finalRecords);
  }
}
