import * as Constants from "./Constants.js";
import * as Conf from "../Configurations.js";

import axios from "axios";

export const getFilterData = async (filterID) => {
  let response = await axios({
    url: `${Constants.api_url}/sr/jira.issueviews:searchrequest-csv-current-fields/${filterID}/SearchRequest-${filterID}.csv?delimiter=,`,
    method: "GET",
    headers: {
      Authorization: `Basic ${Conf.Api_key}`,
    },
  });
  return response.data;
};
