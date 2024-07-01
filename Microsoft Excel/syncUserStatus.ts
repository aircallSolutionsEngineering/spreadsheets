// base information
const baseUrl = "https://api.aircall.io/v1/";
const apiId = "<Aircall API ID>";
const apiToken = "<Aircall API Token>";
const auth = btoa(apiId + ":" + apiToken);
// check if API Token are working correctly
// console.log("apiId: " + apiId + " Token: " + apiToken + "\n" + auth);

async function main(workbook: ExcelScript.Workbook) {
  const userList: string[] | undefined = await listRecords("users");
  // console.log(userList);
  const usersPlan: string[] = workbook
    .getWorksheet("team plan")
    .getRangeByIndexes(1, 0, workbook.getWorksheet("team plan").getUsedRange().getRowCount() - 1, 1)
    .getValues();
  for (let up = 0; up < usersPlan.length; up++) {
    const userName: string = usersPlan[up][0];
    // console.log(userName);
    const userId: string = userName.substring(userName.lastIndexOf("(") + 1, userName.lastIndexOf(")"));
    // console.log(userId);
    for (let ul = 0; ul < (userList as string[]).length; ul++) {
      if ((userList as string[])[ul]["id"] == userId) {
        // console.log("bingo! comparing user id: "+JSON.stringify(userList[ul])+" with users plan id: "+userId);
        if ((userList as string[])[ul]["available"] === false) {
          workbook
            .getWorksheet("team plan")
            .getRangeByIndexes(up + 1, 1, 1, 1)
            .setValue("unavailable");
        } else {
          workbook
            .getWorksheet("team plan")
            .getRangeByIndexes(up + 1, 1, 1, 1)
            .setValue("available");
        }
        break;
      }
    }
  }
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
