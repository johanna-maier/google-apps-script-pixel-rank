// Data for SEO username, password & endpoint
// Note: Pixel rank continues on SERP 2

const username = "johanna.maier@deptagency.com";
const pw = "c00526b67e539b1d";
const endpoint =
  "https://api.dataforseo.com/v3/serp/google/organic/live/advanced";


// user input from settings tab
const keywordIn = SpreadsheetApp.getActive().getSheetByName('Settings').getRange("A7").getValue();
const keywordClean = keywordIn.toLocaleLowerCase().trim().replace(/ +(?= )/g," ");
const domainIn = SpreadsheetApp.getActive().getSheetByName('Settings').getRange("B4").getValue();

// HTTP request headers
const headers = {
  Authorization: "Basic " + Utilities.base64Encode(`${username}:${pw}`),
  "Content-Type": "application/json",
};

// HTTP request options
// https://docs_v3.dataforseo.com/v3/keywords-data-se-locations/
// https://docs_v3.dataforseo.com/v3/keywords-data-se-languages/

const options = {
  method: "POST",
  headers: headers,
  payload: JSON.stringify([
    {
      keyword: keywordClean,
      calculate_rectangles: true,
      language_code: SpreadsheetApp.getActive().getSheetByName('Settings').getRange("C3").getValue(),
      location_code: SpreadsheetApp.getActive().getSheetByName('Settings').getRange("C2").getValue(),
    },
  ]),
  muteHttpExceptions: true,
};

// Function to write output data to sheet
function writeDataToSheet(outputData) {
  // If it doesn't exist yet, add a sheet where rankings are placed.
  let rankingSheet = SpreadsheetApp.getActive().getSheetByName('SERP Rankings');
  if(!rankingSheet) {
    rankingSheet = SpreadsheetApp.getActive().insertSheet("SERP Rankings");
    // rankingSheet.appendRow(["Keyword","Position","URL","Title","Meta Description"]);
  }
  SpreadsheetApp.getActive().getSheetByName('SERP Rankings').getRange("A1").setValue(outputData);
}

// Now we define an asynch function that waits for information of API.

async function httpRequest() {
  // If things go wrong, we catch the error and do sth with it.
  try {
    // We create an URL fetch request to endpoint with options sent in body of request.
    const response = await UrlFetchApp.fetch(endpoint, options);
    const data = JSON.parse(response);
    const allSerpResults = data["tasks"][0]["result"][0]["items"];
    writeDataToSheet(allSerpResults);


    // const allOrganicResults = allSerpResults.filter((element) => {
    //   return element["type"] === "organic";
    // });


    // const pixelsTopTen = allOrganicResults.slice(0, 10).map((element) => {
    //   const pixelRank = element["rectangle"]["y"];
    //   return pixelRank;
    // });

    // const totalPixels = pixelsTopTen.reduce((previousValue, currentValue) => {
    //   return previousValue + currentValue;
    // }, 0);
    // console.log(allOrganicResults);
    // console.log(pixelsTopTen);
    // console.log(totalPixels);

    // fs.writeFileSync(
    //   `./output/${keywordClean}.json`,
    //   JSON.stringify(rankingDomainIn)
    // );



  } catch (error) {
    console.log(error);
  }
}



// TO-DOs with data: Fetch current ranking position and pixel if property ranks in top 10 - maybe 100. (Done)
// Learn how to deploy node.js with webpack to Apps Script: https://medium.com/geekculture/the-ultimate-guide-to-npm-modules-in-google-apps-script-a84545c3f57c
// Learn how to write Google Apps code locally: https://medium.com/geekculture/how-to-write-google-apps-script-code-locally-in-vs-code-and-deploy-it-with-clasp-9a4273e2d018
