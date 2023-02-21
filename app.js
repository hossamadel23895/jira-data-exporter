import * as Conf from "./Configurations.js";
import * as Helpers from "./Helpers/Helpers.js";
import * as Requests from "./Helpers/Requests.js";
import * as Constants from "./Helpers/Constants.js";

import ExcelJS from "exceljs";
import dayjs from "dayjs";
import utc from "dayjs/plugin/utc.js";
dayjs.extend(utc);

// App entry
(async () => {
  console.info("---------------------------------------------------------");
  console.info("Welcome to Jira Reports Exporter");
  while (true) {
    try {
      console.info("---------------------------------------------------------");

      for (const filter of Conf.Filters) {
        Helpers.logMsg(`---------- File ${filter.File_Name} ----------`);

        Helpers.logMsg(`Getting data from jira for "${filter.File_Name}"...`);

        let filterData = await Requests.getFilterData(filter.Filter_ID);

        Helpers.logMsg(
          `Finished getting data from jira for "${filter.File_Name}"...`
        );

        let filterDataArray = filterData
          .split("\n") // split string to lines
          .map((e) => e.trim()) // remove white spaces for each line
          .map((e) => e.split(",").map((e) => e.trim())); // split each line to array

        // Remove first "headers" and last "empty arr" elements after formatting
        filterDataArray.shift();
        filterDataArray.pop();

        // Add Date String to each row
        filterDataArray.map((e) => e.unshift(Helpers.dateStringJira()));

        // Add new data rows to the excel sheet
        Helpers.logMsg(`Adding new data to "${filter.File_Name}"...`);

        let workbook = new ExcelJS.Workbook();
        workbook = await workbook.xlsx.readFile(
          `${filter.File_path}/${filter.File_Name}`
        );
        let worksheet = workbook.getWorksheet(filter.Sheet_name);
        worksheet.addRows(filterDataArray);
        await workbook.xlsx.writeFile(
          `${filter.File_path}/${filter.File_Name}`
        );

        Helpers.logMsg(`Data added successfully to "${filter.File_Name}"`);
        Helpers.logMsg(`-----------------------------------------------------`);
      }

      Helpers.logMsg(
        `Finished getting all new data, Updating after ${Conf.refresh_time_in_hours} hours ...`
      );

      await Helpers.sleep(Conf.refresh_time_in_hours * 60 * 60 * 1000);
    } catch (error) {
      Helpers.logMsg(error);

      Helpers.logMsg(
        `Application encountered an error, retrying in ${Constants.Retry_time_in_mins} min ...`
      );

      await Helpers.sleep(Constants.Retry_time_in_mins * 60 * 1000);
    }
  }
})();
