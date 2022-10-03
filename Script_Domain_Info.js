// Data for SEO username & password
const username = "...";
const pw = "...";
const endpoint =
  "https://api.dataforseo.com/v3/serp/google/organic/live/advanced";

// user input
const keywordIn = "bike";
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
      language_code: "en",
      location_code: 2480,
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

    const organicInfoPerRanking = allOrganicResults.map((element) => {
      const position = element["rank_group"];
      const pixelRank = element["rectangle"]["y"];
      const url = element["url"];
      const domain = element["domain"];
      const title = element["title"];
      const description = element["description"];

      return {
        position: position,
        pixelRank: pixelRank,
        url: url,
        domain: domain,
        title: title,
        description: description,
      };
    });

    const rankingDomainIn = organicInfoPerRanking.filter((element) => {
      return element["domain"] === domainIn;
    });
    console.log(rankingDomainIn);

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
