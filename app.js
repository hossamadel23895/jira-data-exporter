import * as Conf from "./Configurations.js";
import * as Helpers from "./Helpers/Helpers.js";
import * as Requests from "./Helpers/Requests.js";
import * as Constants from "./Helpers/Constants.js";

import util from "util";
import * as child from "child_process";
const exec = util.promisify(child.exec);

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

      let filtersData = [];

      for (const filter of Conf.Filters) {
        Helpers.logMsg(`Getting data from jira for "${filter.File_Name}"...`);

        let filterData = await Requests.getFilterData(filter.Filter_ID);
        filtersData.push(filterData);

        Helpers.logMsg(
          `Finished getting data from jira for "${filter.File_Name}"...`
        );
        Helpers.logMsg(`----------------------------`);
      }

      Helpers.logMsg(`-----------------------------------------------------`);

      for (const [index, filter] of Conf.Filters.entries()) {
        let filterDataArray = filtersData[index]
          .split("\n") // split string to lines
          .map((e) => e.trim()) // remove white spaces for each line
          .map((e) => e.split(",").map((e) => e.trim())); // split each line to array

        // Remove first "headers" and last "empty arr" elements after formatting
        filterDataArray.shift();
        filterDataArray.pop();

        // Removing unneeded columns
        filterDataArray = filterDataArray.map((dataRow) => {
          if (filter.Columns_to_keep.includes("all")) {
            return dataRow;
          } else {
            let filteredDataRow = dataRow.filter((dataCell, index) => {
              return filter.Columns_to_keep.includes(index + 1);
            });
            return filteredDataRow;
          }
        });

        // Converting numbers strings to numbers
        filterDataArray = filterDataArray.map((dataRow) => {
          let correctedDataRow = dataRow.map((dataCell) => {
            if (!isNaN(dataCell) && !isNaN(parseFloat(dataCell))) {
              dataCell = parseFloat(dataCell);
            }
            return dataCell;
          });
          return correctedDataRow;
        });

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
        Helpers.logMsg(`----------------------------`);
      }

      // Synchronizing drive
      Helpers.logMsg(`Synchronizing Drive files...`);
      // const synchronizeCommand = "onedrive --synchronize";
      const synchronizeCommand = "cd /flfl";
      const { stdout, stderr } = await exec(synchronizeCommand);
      if (stderr) throw new Error(stderr);
      console.log(stdout);

      Helpers.logMsg(`Finished Synchronizing Drive files`);
      Helpers.logMsg(`----------------------------`);

	  // Before finishing
      Helpers.logMsg(
        `Finished getting all new data, Updating after ${Conf.refresh_time_in_hours} hours ...`
      );
      Helpers.logMsg(`-----------------------------------------------------`);

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
