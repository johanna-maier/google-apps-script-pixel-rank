// Function to write test output to JSON file into my Google Drive
// https://stackoverflow.com/questions/52777524/how-to-save-a-json-file-to-google-drive-using-google-apps-script
// https://developers.google.com/apps-script/reference/drive/folder#createfileblob
// https://developers.google.com/apps-script/reference/base/blob#setnamename

// function saveAsJSON(json, keyword) {
//   // Creates a file in the users selected Google Drive folder

//   const folder = DriveApp.getFolderById(driveFolderID.toString());
//   const blob = Utilities.newBlob(JSON.stringify(json), "application/vnd.google-apps.script+json");
//   blob.setContentType('application/json').setName(`serp_output_${keyword.trim().replace(/ /g,"_")}.json`);
//   const file = folder.createFile(blob);
//   Logger.log('ID: %s, File size (bytes): %s, type: %s', file.id, file.fileSize, file.mimeType);
// }

// Check if triggers are running to show in Settings tab
function checkTriggers() {
  let allTriggers = ScriptApp.getProjectTriggers();
  if(allTriggers.length > 0) {
    return true
  } else {
    return false
  }
}

// Delete all active triggers 

function deleteActiveTrigger() {
  let allTriggers = ScriptApp.getProjectTriggers();
  if(allTriggers.length > 0) {
    ScriptApp.deleteTrigger(allTriggers[0]);
  }
} 


// Timer class that will be used to track how long the script is already running.
// https://medium.com/geekculture/bypassing-the-maximum-script-runtime-in-google-apps-script-e510aa9ae6da
class Timer {
  start() {
    this.start = Date.now();
  }

  getDuration() {
    return Date.now() - this.start;
  }
}

// Get DataForSEO username, password & endpoint
// Get other user input from settings tab

const username = SpreadsheetApp.getActive()
  .getSheetByName("Settings")
  .getRange("F3")
  .getValue();
const pw = SpreadsheetApp.getActive()
  .getSheetByName("Settings")
  .getRange("F4")
  .getValue();

const domainIn = SpreadsheetApp.getActive()
  .getSheetByName("Settings")
  .getRange("B4")
  .getValue()
  .toString();
const paidResultsOn = SpreadsheetApp.getActive()
  .getSheetByName("Settings")
  .getRange("B5")
  .getValue();

// Get Drive Folder ID (not in use in production sheet)
// const driveFolderID = SpreadsheetApp.getActive()
//   .getSheetByName("Settings")
//   .getRange("F4")
//   .getValue();

// Add relevant functions to UI
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("🔍 DataForSEO")
    .addItem("📐 Analyse SERPs", "getPixelRanks")
    .addItem("🗑️ Delete Trigger", "deleteActiveTrigger")
    .addItem("💲 Get CPC Data", "getAdwordsData")
    .addItem("🔗 Connect CPC Data to SERP Overview Sheet", "connectAdwordsData")
    .addItem("🗑️ Reset Processed Keywords", "removeProcessedKeywords")
    .addToUi();
}

// Function to reset processed keywords
function removeProcessedKeywords() {
  const processedKeywordRange = SpreadsheetApp.getActive()
    .getSheetByName("Settings")
    .getRange("D8:D");
  processedKeywordRange.clearContent();
}

// Functions to  get row of a keyowrd 
function rowOfKeyword(keyword) {
  let data = SpreadsheetApp.getActive()
    .getSheetByName("Settings")
    .getRange("A8:A")
    .getValues()
    .flat()
    .filter((r) => r != "");
  let current_keyword = keyword;

  for (let i = 0; i < data.length; i++) {
    if (data[i] == current_keyword) {
      return i + 8; // + 8 because data list starts at row 8
    }
  }
}

// Before running the SERP Analysis, we need to clean the list for any duplicates or whitespaces
// Otherwise we could run into an infinite loop that keeps on querying the the API

function cleanKeywordInput() {
    const userKeywordInput = SpreadsheetApp.getActive()
    .getSheetByName("Settings")
    .getRange("A8:A");

  const input = userKeywordInput.getValues().flat().map((keyword) => keyword.trim());
  const userKeywordInputUnique = [...new Set(input)];
  console.log(userKeywordInputUnique)
  var toAddArray = [];
  for (i = 0; i < userKeywordInputUnique.length; ++i){
      toAddArray.push([userKeywordInputUnique[i]]);
  }
  userKeywordInput.clearContent();
  Utilities.sleep(5000); 

  SpreadsheetApp.getActive()
    .getSheetByName("Settings")
    .getRange(8, 1, userKeywordInputUnique.length, 1)
    .setValues(toAddArray);
}

// Functions to write data to Google Sheets
function writeOverviewDataToSheet(serpLayoutOverviewRow) {
  let serpLayoutOverviewSheet = SpreadsheetApp.getActive().getSheetByName(
    "SERP Layout Overview"
  );
  if (!serpLayoutOverviewSheet) {
    serpLayoutOverviewSheet = SpreadsheetApp.getActive().insertSheet(
      "SERP Layout Overview"
    );
    serpLayoutOverviewSheet.appendRow([
      "Language Used",
      "Location Used",
      "Keyword Used",
      "Device Used",
      "SERP URL Used",
      "SERP Item Types",
      "Pixel Top 10 Organic",
      "Sum Pixel Top 10 Organic",
      "Sum Pixel Top 3 Organic",
      "Domain Position",
      "Domain Pixel Rank",
      "CPC",
      "Organic Plus",
      "Verticals",
      "Knowledge Graph",
      "Paid",
      "Pixel To First Relevant Rank",
      "Pixel To First Rank",
      "PR 2",
      "PR 3",
      "PR 4",
      "PR 5",
      "PR 6",
      "PR 7",
      "PR 8",
      "PR 9",
      "PR 10",
      "Answer Box",
      "App",
      "Carousel",
      "Multi Carousel",
      "Featured Snippet",
      "Google Flights",
      "Google Reviews",
      "Google Posts",
      "Images",
      "Jobs",
      "Knowledge Graph",
      "Local Pack",
      "Hotels Pack",
      "Map",
      "Organic",
      "Paid",
      "People Also Ask",
      "Related Searches",
      "People Also Search",
      "Shopping",
      "Top Stories",
      "Twitter",
      "Video",
      "Events",
      "Mention Carousel",
      "Recipes",
      "Top Sights",
      "Scholarly Articles",
      "Popular Products",
      "Podcasts",
      "Questions And Answers",
      "Find Results On",
      "Stocks Box",
      "Visual Stories",
      "Commercial Units",
      "Local Services",
      "Google Hotels",
      "Math Solver",
      "Currency Box",
      "Found On The Web",
      "Short Videos",
      "Product Considerations",
    ]);
    serpLayoutOverviewSheet.setTabColor('#000000');
    serpLayoutOverviewSheet.setFrozenRows(1);
    setColors(serpLayoutOverviewSheet);
    serpLayoutOverviewSheet.getDataRange().createFilter();
    serpLayoutOverviewSheet
      .getDataRange()
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  serpLayoutOverviewSheet.appendRow(serpLayoutOverviewRow);
  const lastRow = serpLayoutOverviewSheet.getLastRow();
  const lastColumn = serpLayoutOverviewSheet.getLastColumn();
  const rangeLastRow = serpLayoutOverviewSheet.getRange(
    lastRow,
    1,
    1,
    lastColumn
  );

  rangeLastRow.copyTo(rangeLastRow, { contentsOnly: true });
}

function writeDetailDataToSheet(organicTopTen) {
  const topTen = organicTopTen;
  let featuredSnippet = false;
  if (topTen[0]["type"] === "featured_snippet") {
    featuredSnippet = true;
  }

  let serpLayoutDetailsSheet = SpreadsheetApp.getActive().getSheetByName(
    "SERP Layout Details"
  );
  if (!serpLayoutDetailsSheet) {
    serpLayoutDetailsSheet = SpreadsheetApp.getActive().insertSheet(
      "SERP Layout Details"
    );
    serpLayoutDetailsSheet.appendRow([
      "Keyword",
      "URL",
      "Position",
      "Pixel Rank",
      "Title",
      "Meta Description",
    ]);
    serpLayoutDetailsSheet.setFrozenRows(1);
    serpLayoutDetailsSheet.setTabColor('#000000');
    setColors(serpLayoutDetailsSheet);
    serpLayoutDetailsSheet.getDataRange().createFilter();
    serpLayoutDetailsSheet
      .getDataRange()
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  topTen.forEach((ranking) => {
    let position;
    if (featuredSnippet === true && ranking["type"] === "featured_snippet") {
      position = ranking["rank_group"];
    } else if (
      featuredSnippet === true &&
      ranking["type"] != "featured_snippet"
    ) {
      position = ranking["rank_group"] + 1;
    } else {
      position = ranking["rank_group"];
    }

    const url = ranking["url"];
    const title = ranking["title"];
    const description = ranking["description"];
    const pixelRank = Math.round(ranking["rectangle"]["y"]);

    serpLayoutDetailsSheet.appendRow([
      current_keyword,
      url,
      position,
      pixelRank,
      title,
      description,
    ]);
  });
}

// Function to set alternate formatting for newly created sheets
function setColors(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  const range = sheet.getRange(lastRow, 1, 1, lastColumn);

  // first remove any existing alternating colors in range to prevent error "Exception: You cannot add alternating background colors to a range that already has alternating background colors."
  range.getBandings().forEach((banding) => banding.remove());
  // apply alternate background colors
  range.applyRowBanding();
}

// Regular Expression that matches all unwanted characters which can break the request for CPC data. 
const adwordsRegex = new RegExp(
  "/[\u0000-\u001F]|[\u0021]|[\u0025]|[\u0028-\u002A]|[\u002C]|[\u003B-\u0040]|[\u005C]|[\u005E]|[\u0060]|[\u007B-\u009F]|[\u00A1-\u00A2]|[\u00A4-\u00A9]|[\u00AB-\u00B4]|[\u00B6]|[\u00B8-\u00B9]|[\u00BB-\u00BF]|[\u00D7]|[\u00F7]|[\u0250-\u0258]|[\u025A-\u02AF]|[\u02C2-\u02C5]|[\u02D2-\u02DF]|[\u02E5-\u02EB]|[\u02ED]|[\u02EF-\u02FF]|[\u0375]|[\u037E]|[\u0384-\u0385]|[\u0387]|[\u03F6]|[\u0482]|[\u0488-\u0489]|[\u055A-\u0560]|[\u0588-\u058F]|[\u05BE]|[\u05C0]|[\u05C3]|[\u05C6]|[\u05EF]|[\u05F3-\u060F]|[\u061B-\u061F]|[\u066A-\u066D]|[\u06D4]|[\u06DD-\u06DE]|[\u06E9]|[\u06FD-\u06FE]|[\u0700-\u070F]|[\u07F6-\u07F9]|[\u07FD-\u07FF]|[\u0830-\u083E]|[\u085E]|[\u0870-\u089F]|[\u08B5]|[\u08BE-\u08D3]|[\u08E2]|[\u0964-\u0965]|[\u0970]|[\u09F2-\u09FB]|[\u09FD-\u09FE]|[\u0A76]|[\u0AF0-\u0AF1]|[\u0B55]|[\u0B70]|[\u0B72-\u0B77]|[\u0BF0-\u0BFA]|[\u0C04]|[\u0C3C]|[\u0C5D]|[\u0C77-\u0C7F]|[\u0C84]|[\u0CDD]|[\u0D04]|[\u0D4F]|[\u0D58-\u0D5E]|[\u0D70-\u0D79]|[\u0D81]|[\u0DF4]|[\u0E3F]|[\u0E4F]|[\u0E5A-\u0E5B]|[\u0E86]|[\u0E89]|[\u0E8C]|[\u0E8E-\u0E93]|[\u0E98]|[\u0EA0]|[\u0EA8-\u0EA9]|[\u0EAC]|[\u0EBA]|[\u0F01-\u0F17]|[\u0F1A-\u0F1F]|[\u0F2A-\u0F34]|[\u0F36]|[\u0F38]|[\u0F3A-\u0F3D]|[\u0F85]|[\u0FBE-\u0FC5]|[\u0FC7-\u0FDA]|[\u104A-\u104F]|[\u109E-\u109F]|[\u10FB]|[\u1360-\u137C]|[\u1390-\u1399]|[\u1400]|[\u166D-\u166E]|[\u169B-\u169C]|[\u16EB-\u16ED]|[\u170D]|[\u1715-\u171F]|[\u1735-\u1736]|[\u17D4-\u17D6]|[\u17D8-\u17DB]|[\u17F0-\u180A]|[\u180E-\u180F]|[\u1878]|[\u1940-\u1945]|[\u19DA-\u19FF]|[\u1A1E-\u1A1F]|[\u1AA0-\u1AA6]|[\u1AA8-\u1AAD]|[\u1ABE-\u1ACE]|[\u1B4C]|[\u1B5A-\u1B6A]|[\u1B74-\u1B7E]|[\u1BFC-\u1BFF]|[\u1C3B-\u1C3F]|[\u1C7E-\u1C7F]|[\u1C90-\u1CC7]|[\u1CD3]|[\u1CFA]|[\u1DFA]|[\u1FBD]|[\u1FBF-\u1FC1]|[\u1FCD-\u1FCF]|[\u1FDD-\u1FDF]|[\u1FED-\u1FEF]|[\u1FFD-\u1FFE]|[\u200B-\u2027]|[\u202A-\u202E]|[\u2030-\u203E]|[\u2041-\u2053]|[\u2055-\u205E]|[\u2060-\u2070]|[\u2074-\u207E]|[\u2080-\u208E]|[\u20A0-\u20AB]|[\u20AD-\u20C0]|[\u20DD-\u20E0]|[\u20E2-\u20E4]|[\u2100-\u2101]|[\u2103-\u2106]|[\u2108-\u2109]|[\u2114]|[\u2116-\u2118]|[\u211E-\u2123]|[\u2125]|[\u2127]|[\u2129]|[\u212E]|[\u213A-\u213B]|[\u2140-\u2144]|[\u214A-\u214D]|[\u214F-\u2169]|[\u2170-\u2179]|[\u2189-\u2BFF]|[\u2C2F]|[\u2C5F]|[\u2CE5-\u2CEA]|[\u2CF9-\u2CFF]|[\u2D70]|[\u2E00-\u2E2E]|[\u2E30-\u2FFB]|[\u3001-\u3004]|[\u3006-\u3020]|[\u3030]|[\u3036-\u3037]|[\u303D-\u303F]|[\u309B-\u309C]|[\u30A0]|[\u30FD-\u30FE]|[\u312F]|[\u3190-\u319F]|[\u31BB-\u31E3]|[\u3200-\u33FF]|[\u4DBF-\u4DFF]|[\u4E28]|[\u4EDD]|[\u4F00]|[\u4F03]|[\u4F39]|[\u4F56]|[\u4F92]|[\u4F94]|[\u4F9A]|[\u4FC9]|[\u4FFF]|[\u5040]|[\u5042]|[\u5046]|[\u5094]|[\u50D8]|[\u50F4]|[\u514A]|[\u5164]|[\u519D]|[\u51BE]|[\u51EC]|[\u529C]|[\u52AF]|[\u5307]|[\u5324]|[\u53DD]|[\u548A]|[\u54FF]|[\u5759]|[\u5765]|[\u57AC]|[\u57C7-\u57C8]|[\u58B2]|[\u590B]|[\u595B]|[\u595D]|[\u5963]|[\u5CA6]|[\u5CF5]|[\u5D42]|[\u5D53]|[\u5DD0]|[\u5F21]|[\u5F34]|[\u5F45]|[\u608A]|[\u60DE]|[\u6111]|[\u6130]|[\u6198]|[\u6213]|[\u62A6]|[\u63F5]|[\u6460]|[\u649D]|[\u661E]|[\u6624]|[\u662E]|[\u6659]|[\u6699]|[\u66A0]|[\u66B2]|[\u66BF]|[\u66FA-\u66FB]|[\u670E]|[\u6766]|[\u6801]|[\u6852]|[\u68C8]|[\u68CF]|[\u6998]|[\u6A30]|[\u6A46]|[\u6A73]|[\u6A7E]|[\u6AE2]|[\u6AE4]|[\u6C6F]|[\u6C86]|[\u6D96]|[\u6DCF]|[\u6DF2]|[\u6EBF]|[\u6FB5]|[\u7007]|[\u7104]|[\u710F]|[\u7146]|[\u72B1]|[\u72BE]|[\u7324]|[\u73BD]|[\u73F5]|[\u7429]|[\u769C]|[\u7821]|[\u7864]|[\u7994]|[\u799B]|[\u7AE7]|[\u7B9E]|[\u7D48]|[\u7E8A]|[\u8362]|[\u837F]|[\u83F6]|[\u84DC]|[\u856B]|[\u8807]|[\u88F5]|[\u891C]|[\u8A37]|[\u8AA7]|[\u8ABE]|[\u8ADF]|[\u8B53]|[\u8B7F]|[\u8CF0]|[\u8D12]|[\u9067]|[\u91DA]|[\u91DE]|[\u91E4]|[\u91EE]|[\u9206]|[\u923C]|[\u924E]|[\u9259]|[\u9288]|[\u92A7]|[\u92D3]|[\u92D5]|[\u92D7]|[\u92E0]|[\u92E7]|[\u92FF]|[\u9302]|[\u931D]|[\u9325]|[\u93A4]|[\u93C6]|[\u93F8]|[\u9431]|[\u9445]|[\u969D]|[\u96AF]|[\u9733]|[\u9743]|[\u974D]|[\u974F]|[\u9755]|[\u9857]|[\u9927]|[\u9ADC]|[\u9B72]|[\u9B75]|[\u9B8F]|[\u9BB1]|[\u9BBB]|[\u9C00]|[\u9E19]|[\u9FEB-\u9FFF]|[\uA490-\uA4C6]|[\uA4FE-\uA4FF]|[\uA60D-\uA60F]|[\uA670-\uA673]|[\uA67E]|[\uA6F2-\uA716]|[\uA720-\uA721]|[\uA789-\uA78A]|[\uA7AF]|[\uA7B8-\uA7F6]|[\uA828-\uA839]|[\uA874-\uA877]|[\uA8CE-\uA8CF]|[\uA8F8-\uA8FA]|[\uA8FC]|[\uA8FE-\uA8FF]|[\uA92E-\uA92F]|[\uA95F]|[\uA9C1-\uA9CD]|[\uA9DE-\uA9DF]|[\uAA5C-\uAA5F]|[\uAA77-\uAA79]|[\uAADE-\uAADF]|[\uAAF0-\uAAF1]|[\uAB5B]|[\uAB66-\uAB6B]|[\uABEB]|[\uE000-\uF8FF]|[\uFA0E-\uFA0F]|[\uFA11]|[\uFA13-\uFA15]|[\uFA1F-\uFA21]|[\uFA23-\uFA24]|[\uFA27-\uFA29]|[\uFB00-\uFB06]|[\uFB29]|[\uFBB2-\uFBC2]|[\uFD3E-\uFD4F]|[\uFDCF]|[\uFDFC-\uFDFF]|[\uFE10-\uFE19]|[\uFE30-\uFE32]|[\uFE35-\uFE4C]|[\uFE50-\uFE6B]|[\uFEFF-\uFF05]|[\uFF07-\uFF0A]|[\uFF0C-\uFF0F]|[\uFF1A-\uFF20]|[\uFF3B-\uFF3E]|[\uFF40]|[\uFF5B-\uFFFF]/",
  "g"
);

// Make all SERP analysis with JSON data received from API
function getSerpToOrganicTopTen(data) {
  const arrayAllSerpResults = data["tasks"][0]["result"][0]["items"];

  // Which element is the 10th organic result?
  const organicResultTen = arrayAllSerpResults.find((element) => {
    return element["type"] === "organic" && element["rank_group"] === 10;
  });

  // What is the absolute rank of the 10th organic result?
  const organicResultTenAbsoluteRank = organicResultTen["rank_absolute"];

  // Grabs all SERP elements that appear before the 10th organic result.
  const arrayAllSerpResultsToOrganicTenWithPaid = arrayAllSerpResults.filter(
    (element) => {
      return element["rank_absolute"] <= organicResultTenAbsoluteRank;
    }
  );

  // Remove paid & shopping serpItems including pixel values from item array
  // Iterate through all results and check if item type falls in "paid" category that might be affected by Google blocks
  const paidItemTypes = ["shopping", "paid", "commercial_units"];
  let paidHeightCounter = 0;
  // https://www.freecodecamp.org/news/how-to-clone-an-array-in-javascript-1d3183468f6a/
  // https://stackoverflow.com/questions/35922429/why-does-a-js-map-on-an-array-modify-the-original-array
  // Original issue: Map updated values of objects also in "with paid" array.

  const deepCopyArrayAllSerpResultsToOrganicTenWithoutPaid = JSON.parse(
    JSON.stringify(arrayAllSerpResultsToOrganicTenWithPaid)
  );

  const arrayAllSerpResultsToOrganicTenWithoutPaid =
    deepCopyArrayAllSerpResultsToOrganicTenWithoutPaid.map((element) => {
      // if element is paid, add height pixels to counter and return
      if (paidItemTypes.includes(element["type"])) {
        paidHeightCounter = paidHeightCounter + element["rectangle"]["height"];
        return {};
      } else {
        // if element is not paid, deduct paidPixelCounter from pixel rank, overwrite it in element and return with new pixel rank
        const currentPixelRank = Math.round(element["rectangle"]["y"]);
        const newPixelRank = currentPixelRank - paidHeightCounter;
        element["rectangle"]["y"] = newPixelRank;
        return element;
      }
    });

  let arrayAllSerpResultsToOrganicTen;
  if (paidResultsOn === true) {
    arrayAllSerpResultsToOrganicTen = arrayAllSerpResultsToOrganicTenWithPaid;
  } else {
    arrayAllSerpResultsToOrganicTen =
      arrayAllSerpResultsToOrganicTenWithoutPaid;
  }

  // saveAsJSON(arrayAllSerpResultsToOrganicTen, current_keyword);

  return arrayAllSerpResultsToOrganicTen;
}

// Prepare overview data row from items until top 10  organic result
function prepareOverviewData(arrayAllSerpResultsToOrganicTen,topTenOrganicResults,data) {
  // Extracts item types of all all SERP elements that appear before the 10th organic result.
  const serpItemTypesToOrganicTopTen = arrayAllSerpResultsToOrganicTen
    .filter((element) => element["type"] != undefined)
    .map((element) => {
      return element["type"];
    });

  const pixelsTopTenOrganic = topTenOrganicResults.map((element) => {
    const pixelRank = Math.round(element["rectangle"]["y"]);
    return pixelRank;
  });

  const domainResult = topTenOrganicResults.filter((element) => {
    return element["domain"] === domainIn;
  });

  const totalPixelsOrganic = pixelsTopTenOrganic.reduce(
    (previousValue, currentValue) => {
      return previousValue + currentValue;
    },
    0
  );

  const totalPixelsTopThree = pixelsTopTenOrganic
    .slice(0, 3)
    .reduce((previousValue, currentValue) => {
      return previousValue + currentValue;
    }, 0);

  const arrayIrrelevantDomains = SpreadsheetApp.getActive()
    .getSheetByName("Settings")
    .getRange("G8:G")
    .getValues()
    .flat()
    .filter((r) => r != "");

  const arrayRelevantElements = topTenOrganicResults.filter((element) => {
    const domainCurrentElement = element["domain"];
    return !arrayIrrelevantDomains.includes(domainCurrentElement);
  });

  const pixelToFirstRelevantRanking = arrayRelevantElements[0]["rectangle"]["y"];

  // const serpItemTypesTopHundred = data["tasks"][0]["result"][0]["item_types"].toString();

  const languageUsed = SpreadsheetApp.getActive()
    .getSheetByName("Settings")
    .getRange("B2")
    .getValue();
  const locationUsed = SpreadsheetApp.getActive()
    .getSheetByName("Settings")
    .getRange("B3")
    .getValue();
  const keywordUsed = data["tasks"][0]["data"]["keyword"];
  const deviceUsed = data["tasks"][0]["data"]["device"];
  const serpUrl = data["tasks"][0]["result"][0]["check_url"];

  const serpItemTypes = [
    ...new Set(serpItemTypesToOrganicTopTen.flat(1)),
  ].toString();
  const formulaPixelRanks = '=split(INDIRECT("R[0]C7";false);",")';

  let domainPosition;
  let domainPixelRank;
  if (domainResult.length > 0) {
    domainPosition = domainResult[0]["rank_group"];
    domainPixelRank = Math.round(domainResult[0]["rectangle"]["y"]);
  } else {
    domainPosition = "not in top 10";
    domainPixelRank = "not in top 10";
  }
  const formulaSerpItems =
    '=if(regexmatch(INDIRECT("R[0]C6";false);index(' +
    "'" +
    "SERP Feature Details" +
    "'" +
    '!$A:$A; match(INDIRECT("R1C[0]";false);' +
    "'" +
    "SERP Feature Details" +
    "'" +
    "!$B:$B;0)));1;0)";
  const serpItemArray = Array(42).fill(formulaSerpItems);

  const formulaOrganicPlus =
    '=sum(arrayformula(if(regexmatch(INDIRECT("R[0]C6";false);filter(' +
    "'" +
    "SERP Feature Details" +
    "'" +
    "!A:A;" +
    "'" +
    "SERP Feature Details" +
    "'" +
    '!C:C="organic plus"));1;0)))';
  const formulaVerticals =
    '=sum(arrayformula(if(regexmatch(INDIRECT("R[0]C6";false);filter(' +
    "'" +
    "SERP Feature Details" +
    "'" +
    "!A:A;" +
    "'" +
    "SERP Feature Details" +
    "'" +
    '!C:C="vertical"));1;0)))';
  const formulaKnowledgeGraph =
    '=sum(arrayformula(if(regexmatch(INDIRECT("R[0]C6";false);filter(' +
    "'" +
    "SERP Feature Details" +
    "'" +
    "!A:A;" +
    "'" +
    "SERP Feature Details" +
    "'" +
    '!C:C="knowledge graph"));1;0)))';
  const formulaPaid =
    '=sum(arrayformula(if(regexmatch(INDIRECT("R[0]C6";false);filter(' +
    "'" +
    "SERP Feature Details" +
    "'" +
    "!A:A;" +
    "'" +
    "SERP Feature Details" +
    "'" +
    '!C:C="paid"));1;0)))';

  const formulaCpc =
    '=iferror(index(CPC!B:B;match(INDIRECT("R[0]C3";false);CPC!A:A;0));"no CPC data")';

  const serpLayoutOverviewRow = [
    languageUsed,
    locationUsed,
    keywordUsed,
    deviceUsed,
    serpUrl,
    serpItemTypes,
    pixelsTopTenOrganic.toString(),
    totalPixelsOrganic,
    totalPixelsTopThree,
    domainPosition,
    domainPixelRank,
    formulaCpc,
    formulaOrganicPlus,
    formulaVerticals,
    formulaKnowledgeGraph,
    formulaPaid,
    pixelToFirstRelevantRanking,
    formulaPixelRanks,
    ,
    ,
    ,
    ,
    ,
    ,
    ,
    ,
    ,
  ].concat(serpItemArray);

  return serpLayoutOverviewRow;
}



// Write AdWords Data from Response into Sheet
// Refactor to add array instead of each row itself

function writeKeywordDataToSheet(keywordData) {
  const keywordDataArray = [];
  const keywordDataObject = keywordData; 

  keywordDataObject["tasks"][0]["result"].forEach((result) => {
    const keyword = result.keyword;
    const avgCPC = result.average_cpc;
    
    if (avgCPC == null) {
      keywordDataArray.push([keyword, "no CPC data"]);
    } else {
      keywordDataArray.push([keyword, avgCPC]);
    }
  });
  console.log(keywordDataArray);

  let keywordDataSheet = SpreadsheetApp.getActive().getSheetByName("CPC");

  if (!keywordDataSheet) {
    keywordDataSheet = SpreadsheetApp.getActive().insertSheet("CPC");
    keywordDataSheet.setTabColor('#000000');
    keywordDataSheet.setFrozenRows(1);
    keywordDataSheet.appendRow(["Keyword", "Average CPC"]);
    setColors(keywordDataSheet);
    keywordDataSheet.getDataRange().createFilter();
    keywordDataSheet
      .getDataRange()
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  var lastRow = keywordDataSheet.getLastRow();
  keywordDataSheet.getRange(lastRow + 1,1,keywordDataArray.length, keywordDataArray[0].length).setValues(keywordDataArray);

}

// HTTP request headers to access DataForSEO and endpoints
const headers = {
  Authorization: "Basic " + Utilities.base64Encode(`${username}:${pw}`),
  "Content-Type": "application/json",
};

const endpointSerp =
  "https://api.dataforseo.com/v3/serp/google/organic/live/advanced";
const endpointAdwords =
  "https://api.dataforseo.com/v3/keywords_data/google_ads/ad_traffic_by_keywords/live";


// Asynch function to get pixel ranks & SERP data from DataForSEO SERP API.
async function getPixelRanks() {
  if (username == "" || pw == "") {
    SpreadsheetApp.getUi().alert(
      "No Data for SEO credentials. Please add them in settings."
    );
    throw new Error(
      "No Data for SEO credentials. Please add them in settings."
    );
  }
  
  cleanKeywordInput();
    
  // Start timer to avoid execution limits.
  let timer = new Timer();
  timer.start();
  let timeThreshold = 4 * 60 * 1000; // 4000 miliseconds = 4 minutes

  // A1notation gets all cells under "Remaining Keywords" and filters out empty ones.
  // This way the function can rerun after being interrupted by time-limit.
  // https://stackoverflow.com/questions/64286235/how-to-get-values-in-a-column-and-put-them-into-an-array-google-sheets
  let searchTermList = SpreadsheetApp.getActive()
    .getSheetByName("Settings")
    .getRange("E8:E")
    .getValues()
    .flat()
    .filter((r) => r != "");

  if (searchTermList.length == 0) {
    // If we proccessed all terms, we can stop the time-based trigger.
    let allTriggers = ScriptApp.getProjectTriggers();
    if(allTriggers.length > 0) {
      ScriptApp.deleteTrigger(allTriggers[0]);
    }
    SpreadsheetApp.getUi().alert("All search terms processed. Please select 'Remove Processed Keywords' or add additional keywords to column A.");
    throw new Error("All search terms processed. Please select 'Remove Processed Keywords' or add additional keywords to column A.");
  }

  // For each keyword in the list, request API with different term and get organic_results.
  // Log current keyword row to create list of analysed keywords.
  // For of loop to be able to use break statements in if conditions inside loop.
  // Otherwise: https://bobbyhadz.com/blog/javascript-illegal-break-statement
  for (const keyword of searchTermList) {
    // First check if timer is beyond 4 minutes and get out of function if necessary.
    if (timeThreshold < timer.getDuration()) {
      // How many keywords are still on the unchecked list?
      let remainingSearchTermList = SpreadsheetApp.getActive()
        .getSheetByName("Settings")
        .getRange("E8:E")
        .getValues()
        .flat()
        .filter((r) => r != "");

      // Do we already have a project trigger active?
      let allTriggers = ScriptApp.getProjectTriggers();

      if (remainingSearchTermList.length == 0) {
        // If we proccessed all terms, we can stop the time-based trigger.
        ScriptApp.deleteTrigger(allTriggers[0]);
        Logger.log("Search list empty");
        SpreadsheetApp.getUi().alert("All search terms processed.");
        break;
      } else {
        // If we still need to process keywords, we either return and keep the already active trigger running
        // or we create a new trigger.
        if (allTriggers.length > 0) {
          break;
        } else {
          ScriptApp.newTrigger("getPixelRanks")
            .timeBased()
            .everyMinutes(5)
            .create();
          Logger.log("New trigger created");
          break;
        }
      }
    }
    current_keyword_row = rowOfKeyword(keyword);

    const keyword_cell_string = "A" + current_keyword_row.toString();

    const current_keyword = SpreadsheetApp.getActive()
      .getSheetByName("Settings")
      .getRange(keyword_cell_string)
      .getValue()
      .toLocaleLowerCase()
      .trim()
      .replace(/ +(?= )/g, " ");

    const options = {
      method: "POST",
      headers: headers,
      payload: JSON.stringify([
        {
          keyword: current_keyword,
          calculate_rectangles: true,
          search_param: SpreadsheetApp.getActive()
            .getSheetByName("Settings")
            .getRange("C5")
            .getValue(),
          language_code: SpreadsheetApp.getActive()
            .getSheetByName("Settings")
            .getRange("C3")
            .getValue(),
          location_code: SpreadsheetApp.getActive()
            .getSheetByName("Settings")
            .getRange("C2")
            .getValue(),
          device: SpreadsheetApp.getActive()
            .getSheetByName("Settings")
            .getRange("F5")
            .getValue(),
        },
      ]),
      muteHttpExceptions: true,
    };

    try {
      // We create an URL fetch request to endpoint with options sent in body of request.
      const response = await UrlFetchApp.fetch(endpointSerp, options);
      const data = JSON.parse(response);
      // Save raw data output to Google Drive folder as reference (not implemented in production version)
      // saveAsJSON(data, current_keyword); (Saving to Drive folder not implemented in production version)
      if(data["tasks_error"] === 1) {
        SpreadsheetApp.getUi().alert(
          "Error connecting to DataForSEO Google SERP API. Check the response message below, the dashboard (https://app.dataforseo.com/api-errors) and contact support. \n\n Response details: \n\n" + response
        );
        throw new Error("Error connecting to DataForSEO Google SERP API. Response details: " + response );
      }
      const arrayAllSerpResultsToOrganicTen = getSerpToOrganicTopTen(data);

      const topTenOrganicResults = arrayAllSerpResultsToOrganicTen
        .filter((element) => {
          // Featured snippet is reagarded as organic ranking position.
          return (
            element["type"] === "organic" ||
            element["type"] === "featured_snippet"
          );
        })
        .slice(0, 10);

      // saveAsJSON(topTenOrganicResults, current_keyword);

      writeDetailDataToSheet(topTenOrganicResults);

      const serpLayoutOverviewRow = prepareOverviewData(
        arrayAllSerpResultsToOrganicTen,
        topTenOrganicResults,
        data
      );
      writeOverviewDataToSheet(serpLayoutOverviewRow);
    } catch (error) {
      console.log(error);
      break

    }
    
    SpreadsheetApp.getActive()
      .getSheetByName("Settings")
      .getRange("D" + current_keyword_row.toString())
      .setValue(current_keyword);
  }
  // Here we are out of the keyword loop
  // If things go wrong, we catch the error and do sth with it.
}




// Functions to query CPC data for lists of 1000 keywords at a time
function getAdwordsData() {
  const keywordList = SpreadsheetApp.getActive()
    .getSheetByName("Settings")
    .getRange("A8:A")
    .getValues()
    .flat()
    .filter((r) => r != "").filter( (r) => {
      return r.match(adwordsRegex) === null;
    }).filter( (r) => {
      return r.length < 80 && r.split(" ").length < 11
    });

  // Split list into array of arrays of under 1000: https://stackoverflow.com/questions/11318680/split-array-into-chunks-of-n-length
  const arrayMaxSize = 1000;
  let arraysKeywordLists = [];

  while (keywordList.length > 0) {
    arraysKeywordLists.push(keywordList.splice(0, arrayMaxSize));
  }

  // Query keyword data for each array of 1000 keywords via asynch API call
  arraysKeywordLists.forEach(async (keywordArray) => {
    const options = {
      method: "POST",
      headers: headers,
      payload: JSON.stringify([
        {
          keywords: keywordArray,
          match: "exact",
          bid: 999,
          language_code: SpreadsheetApp.getActive()
            .getSheetByName("Settings")
            .getRange("C3")
            .getValue(),
          location_code: SpreadsheetApp.getActive()
            .getSheetByName("Settings")
            .getRange("C2")
            .getValue()
        },
      ]),
      muteHttpExceptions: true,
    };

    
    try {
      const response = await UrlFetchApp.fetch(endpointAdwords, options);
      const keywordData = JSON.parse(response);
      if(keywordData["tasks_error"] === 1) {
        SpreadsheetApp.getUi().alert(
          "Error connecting to DataForSEO Google Ads API. Check the response message below, the dashboard (https://app.dataforseo.com/api-errors) and contact support. \n\n Response details: \n\n" + response
        );
        throw new Error("Error connecting to DataForSEO Google Ads API. Response details: " + response );
      }
      writeKeywordDataToSheet(keywordData);
    } catch (error) {
      console.log(error);     
    }
  });
}

// If the Adwords data gets collected after the SERP analysis, it the SERP sheets needs to be updated
function connectAdwordsData() {

  const serpLayoutOverviewSheet = SpreadsheetApp.getActive().getSheetByName(
    "SERP Layout Overview"
  );
  if (!serpLayoutOverviewSheet) {
    SpreadsheetApp.getUi().alert(
      "Overview sheet missing. No SERP analysis executed yet?"
    );
    throw new Error("Overview sheet missing. No SERP analysis executed yet?");
  }

  const CpcSheet = SpreadsheetApp.getActive().getSheetByName(
      "CPC"
  );
  if (!CpcSheet) {
    SpreadsheetApp.getUi().alert(
      "CPC sheet missing. No CPC data queried yet?"
    );
    throw new Error("CPC sheet missing. No CPC data queried yet?");
  }

  const formulaCpc =
    '=iferror(index(CPC!B:B;match(INDIRECT("R[0]C3";false);CPC!A:A;0));"no CPC data")';

  const lastRow = serpLayoutOverviewSheet.getLastRow();
  const rangeCpcColumn = serpLayoutOverviewSheet.getRange(2, 12, lastRow, 1);
  const valuesCpcColumn = rangeCpcColumn.getValues();
  for (let i = 0; i < valuesCpcColumn.length; i++) {
    // Get the cell in the current row
    let cell = rangeCpcColumn.getCell(i+1, 1);
    //Add formula to the cell
    cell.setFormula(formulaCpc);
  }
  SpreadsheetApp.flush();

  Utilities.sleep(5000); 
  rangeCpcColumn.copyTo(rangeCpcColumn, { contentsOnly: true });
}




// Add logic to query an entire keyword list.
// https://script.google.com/home/projects/15_4qhZHbL1DA1u0S1zrfW3mZ7HxZMs0jbvtAG1Cs7GhhT9TIoAWlElUv/edit

// TO-DOs with data: Fetch current ranking position and pixel if property ranks in top 10 - maybe 100. (Done)
// Learn how to deploy node.js with webpack to Apps Script: https://medium.com/geekculture/the-ultimate-guide-to-npm-modules-in-google-apps-script-a84545c3f57c
// Learn how to write Google Apps code locally: https://medium.com/geekculture/how-to-write-google-apps-script-code-locally-in-vs-code-and-deploy-it-with-clasp-9a4273e2d018

// Loop through top ten organic data and write it to details sheet
// Values: Position, URL, Meta, Title, PixelRank
// Add yes/no filter options for all SERP item types as well as 10 columns for pixel rank values
// Add setting to filter out paid items since API is not reliable
// Compare value consistency for non paid version
// Add info about domain ranking (position and pixel rank)
// Include triggers for big keyword sets and failsafe option
// Idea: List outputs for keywords that were run and keywords that need to be rerun
// Own button to rerun missing keywords

// fs.writeFileSync(
//   `./output/${current_keyword}.json`,
//   JSON.stringify(rankingDomainIn)
// );
