import { promises as promisesFs } from "fs";
import fs from "fs";
import path from "path";
import * as Conf from "../Configurations.js";
import * as Constants from "./Constants.js";

import dayjs from "dayjs";
import utc from "dayjs/plugin/utc.js";
dayjs.extend(utc);

// Adding Json version of the error for logging
if (!("toJSON" in Error.prototype))
  Object.defineProperty(Error.prototype, "toJSON", {
    value: function () {
      var alt = {};

      Object.getOwnPropertyNames(this).forEach(function (key) {
        alt[key] = this[key];
      }, this);

      return alt;
    },
    configurable: true,
    writable: true,
  });

export const dateStringJira = () => {
  let stampObj = dayjs().utc();
  let year = stampObj.$y;
  let month = ("0" + (stampObj.$M + 1)).slice(-2);
  let day = ("0" + stampObj.$D).slice(-2);
  return `${day}/${month}/${year}`;
};

export const sleep = (milliseconds) => {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
};

export const logMsg = async (msg) => {
  // Pushing message to terminal
  console.info(addTimeStamp(), msg);

  // Logging to a log file
  let normalizedPath = path.normalize(Constants.Logs_files_path);

  // Adding log line to logs
  if (msg instanceof Error) {
    await promisesFs.appendFile(
      `${normalizedPath}/${createLogsFileName()}.${
        Constants.Logs_file_extension
      }`,
      `${addTimeStamp()} ${JSON.stringify(msg)} \n`
    );
  } else {
    await promisesFs.appendFile(
      `${normalizedPath}/${createLogsFileName()}.${
        Constants.Logs_file_extension
      }`,
      `${addTimeStamp()} ${msg} \n`
    );
  }
};

export const createLogsFileName = () => {
  let stampObj = dayjs().utc();
  let year = stampObj.$y;
  let month = ("0" + (stampObj.$M + 1)).slice(-2);
  let day = ("0" + stampObj.$D).slice(-2);
  return `${year}-${month}-${day}`;
};

export const addTimeStamp = () => {
  let stampObj = dayjs().utc();
  let hour = ("0" + stampObj.$H).slice(-2);
  let min = ("0" + stampObj.$m).slice(-2);
  let second = ("0" + stampObj.$s).slice(-2);
  return `(${hour}:${min}:${second})`;
};
