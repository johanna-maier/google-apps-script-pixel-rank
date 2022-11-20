// Function to write test output to JSON file into my Google Drive
// https://stackoverflow.com/questions/52777524/how-to-save-a-json-file-to-google-drive-using-google-apps-script
// https://developers.google.com/apps-script/reference/drive/folder#createfileblob
// https://developers.google.com/apps-script/reference/base/blob#setnamename

function saveAsJSON(json, keyword) {
  // Creates a file in the users selected Google Drive folder

  const folder = DriveApp.getFolderById("18hNXSQ4Se-YH4pvPx0uGwZTaKrSPL4QO");
  const blob = Utilities.newBlob(JSON.stringify(json), "application/vnd.google-apps.script+json");
  blob.setContentType('application/json').setName(`serp_output_${keyword.trim().replace(/ /g,"_")}.json`)
  const file = folder.createFile(blob);
  Logger.log('ID: %s, File size (bytes): %s, type: %s', file.id, file.fileSize, file.mimeType);
}

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
      location_code: SpreadsheetApp.getActive().getSheetByName('Settings').getRange("C2").getValue()
    },
  ]),
  muteHttpExceptions: true,
};

// Function to write output data to sheet
// function writeDataToSheet(outputData) {
//   // If it doesn't exist yet, add a sheet where rankings are placed.
//   let rankingSheet = SpreadsheetApp.getActive().getSheetByName("SERP Rankings");
//   if(!rankingSheet) {
//     rankingSheet = SpreadsheetApp.getActive().insertSheet("SERP Rankings");
//     // rankingSheet.appendRow(["Keyword","Position","URL","Title","Meta Description"]);
//   }
//   SpreadsheetApp.getActive().getSheetByName('SERP Rankings').getRange("A1").setValue(outputData);
// }

function writeOverviewDataToSheet(keyword, language, location, device, serpUrl, serpItemTypes, pixelTen, pixelTotal) {
  const serpLayoutOverviewRow = [keyword, language, location, device, serpUrl, serpItemTypes, pixelTen, pixelTotal];

  let serpLayoutOverviewSheet = SpreadsheetApp.getActive().getSheetByName("SERP Layout Overview");
  if(!serpLayoutOverviewSheet) {
    serpLayoutOverviewSheet = SpreadsheetApp.getActive().insertSheet("SERP Layout Overview");
    serpLayoutOverviewSheet.appendRow(["Language Used" , "Location Used" , "Keyword Used" , "Device Used" , "SERP URL Used" , "SERP Item Types", "Pixel Top 10 Organic", "Sum Pixel Top 10 Organic"]);
  }
   serpLayoutOverviewSheet.appendRow(serpLayoutOverviewRow);
}

// Now we define an asynch function that waits for information of API.

async function httpRequest() {
  // If things go wrong, we catch the error and do sth with it.
  try {
    // We create an URL fetch request to endpoint with options sent in body of request.
    const response = await UrlFetchApp.fetch(endpoint, options);
    const data = JSON.parse(response);
    // saveAsJSON(data, keywordClean);
    const arrayAllSerpResults = data["tasks"][0]["result"][0]["items"];

    const organicResultTen = arrayAllSerpResults.find( (element) => {
      return element["type"] === "organic" && element["rank_group"] === 10;
    });

    const organicResultTenAbsoluteRank = organicResultTen["rank_absolute"];

    const arrayAllSerpResultsToOrganicTen = arrayAllSerpResults.filter((element) => {
      return element["rank_absolute"] <= organicResultTenAbsoluteRank;
    });

    // saveAsJSON(arrayAllSerpResultsToOrganicTen, keywordClean);

    const serpItemTypesToOrganicTopTen = arrayAllSerpResultsToOrganicTen.map ((element) => {
      return element["type"];
    });

    const topTenOrganicResults = arrayAllSerpResultsToOrganicTen.filter( (element) => {
      return element["type"] === "organic";
    });

    const pixelsTopTenOrganic = topTenOrganicResults.map((element) => {
      const pixelRank = element["rectangle"]["y"];
      return pixelRank;
    });

    const totalPixelsOrganic = pixelsTopTenOrganic.reduce((previousValue, currentValue) => {
     return previousValue + currentValue;
    }, 0);

    // const serpItemTypesTopHundred = data["tasks"][0]["result"][0]["item_types"].toString();


    const languageUsed = data["tasks"][0]["data"]["language_code"];
    const locationUsed = data["tasks"][0]["data"]["location_code"];
    const keywordUsed = data["tasks"][0]["data"]["keyword"];
    const deviceUsed = data["tasks"][0]["data"]["device"];
    const serpUrl = data["tasks"][0]["result"][0]["check_url"];
    const serpItemTypes = [...new Set(serpItemTypesToOrganicTopTen.flat(1))].toString();


   writeOverviewDataToSheet(languageUsed, locationUsed, keywordUsed, deviceUsed, serpUrl, serpItemTypes, pixelsTopTenOrganic.toString(), totalPixelsOrganic);

   // Loop through top ten organic data and write it to details sheet
   // Values: Position, URL, Meta, Title, PixelRank
   // Add yes/no filter options for all SERP item types as well as 10 columns for pixel rank values
   // Add setting to filter out paid items since API is not reliable
   // Compare value consistency for non paid version
   // Add info about domain ranking (position and pixel rank)
   // Include triggers for big keyword sets and failsafe option
   // Idea: List outputs for keywords that were run and keywords that need to be rerun
   // Own button to rerun missing keywords 

   // writeDataToSheet(data);


  } catch (error) {
    console.log(error);
  }
}



// TO-DOs with data: Fetch current ranking position and pixel if property ranks in top 10 - maybe 100. (Done)
// Learn how to deploy node.js with webpack to Apps Script: https://medium.com/geekculture/the-ultimate-guide-to-npm-modules-in-google-apps-script-a84545c3f57c
// Learn how to write Google Apps code locally: https://medium.com/geekculture/how-to-write-google-apps-script-code-locally-in-vs-code-and-deploy-it-with-clasp-9a4273e2d018


    // fs.writeFileSync(
    //   `./output/${keywordClean}.json`,
    //   JSON.stringify(rankingDomainIn)
    // );

