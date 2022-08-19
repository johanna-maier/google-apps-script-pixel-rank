require("dotenv").config();
// fs = standard node library variable "file system"
// importing package fs
const fs = require("fs");
// package axios > makes HTTP requests incl. POST
// posting data to API = sending config of response that we want
const axios = require("axios");

const keywordIn = "adult bikes";
const keyword = keywordIn.toLocaleLowerCase();
const keywordClean = keyword.replace(/\s/g, "-"); // replace spaces - all of them > globally
const endpoint =
  "https://api.dataforseo.com/v3/serp/google/organic/live/advanced";

// When posting to an endpoint, you often need to specify endpoint and the options.
const options = {
  method: "post",
  auth: {
    username: process.env.USERNAME,
    password: process.env.PASSWORD,
  },
  data: [
    {
      keyword: encodeURIComponent(keyword),
      calculate_rectangles: true,
      language_code: "en",
      location_code: 2480,
    },
  ],
  headers: { content_type: "application/json" }, // Axios knows what type of content to expect in request.
};

// Now we define an asynch function that waits for information of API.

async function httpRequest() {
  // If things go wrong, we catch the error and do sth with it.
  try {
    // We create an axios request to endpoint with options sent in body of request.
    const response = await axios(endpoint, options);
    // console.log(response);
    // console.log(response.data);
    // console.log(JSON.stringify(response.data, null, 2));
    // Everything is going to wait for the data to be saved.
    fs.writeFileSync(
      `./output/${keywordClean}.json`,
      JSON.stringify(response.data)
    );
  } catch (error) {
    console.log(error);
  }
}

httpRequest();
