// Data for SEO username & password
// Note: Pixel rank continues on SERP 2

const endpoint =
  "https://api.dataforseo.com/v3/serp/google/organic/live/advanced";

// user input
const keywordIn = "Drabiny przystawne";
const keyword = keywordIn.toLocaleLowerCase();
const keywordClean = keyword.replace(/\s/g, "-"); // replace spaces - all of them > globally
const domainIn = "www.canyon.com";

// HTTP request headers
const headers = {
  Authorization: "Basic " + Utilities.base64Encode(`${username}:${pw}`),
  "Content-Type": "application/json",
};

// HTTP request options
const options = {
  method: "POST",
  headers: headers,
  payload: JSON.stringify([
    {
      keyword: keyword,
      calculate_rectangles: true,
      language_code: "pl",
      location_code: 2616,
    },
  ]),
  muteHttpExceptions: true,
};

// Now we define an asynch function that waits for information of API.

async function httpRequest() {
  // If things go wrong, we catch the error and do sth with it.
  try {
    // We create an URL fetch request to endpoint with options sent in body of request.
    const response = await UrlFetchApp.fetch(endpoint, options);

    const data = JSON.parse(response);

    const allSerpResults = data["tasks"][0]["result"][0]["items"];

    // console.log(typeof allSerpResults);

    const allOrganicResults = allSerpResults.filter((element) => {
      return element["type"] === "organic";
    });

    const pixelsTopTen = allOrganicResults.slice(0, 10).map((element) => {
      const pixelRank = element["rectangle"]["y"];
      return pixelRank;
    });

    const totalPixels = pixelsTopTen.reduce((previousValue, currentValue) => {
      return previousValue + currentValue;
    }, 0);
    console.log(allOrganicResults);
    console.log(pixelsTopTen);
    console.log(totalPixels);

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
