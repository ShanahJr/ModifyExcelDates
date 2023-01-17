const readXlsxFile = require("read-excel-file/node");
const moment = require("moment");
const writeXlsxFile = require("write-excel-file");
const XLSX = require("xlsx");
const fs = require("fs");
const { exit } = require("process");

// ToDo: Create schema when working with different excel files
// const schema = [
//   {
//     column: "Old Date",
//     type: String,
//     value: (object) => object.oldDate,
//   },
//   {
//     column: "New Date",
//     type: String,
//     value: (object) => object.newDate,
//   },
// ];

// let objects = [];
// let dateObject = {};

readXlsxFile("./file.xlsx").then((rows) => {
  let counter = 0;

  let dateList = [];

  rows.forEach((row) => {
    counter++;

    //If string
    if (typeof row[0] === "string") {
      const containsSlash = row[0].includes("/");
      const containsDash = row[0].includes("-");

      let dateBits = [];
      let date = new Date();

      if (containsDash) {
        dateBits = row[0].split("-");
        date = new Date(
          parseInt(dateBits[2].trim()),
          parseInt(dateBits[1].trim()) - 1,
          parseInt(dateBits[0].trim())
        );
      } else if (containsSlash) {
        dateBits = row[0].split("/");

        date = new Date(
          parseInt(dateBits[0].trim()),
          parseInt(dateBits[1].trim()) - 1,
          parseInt(dateBits[2].trim())
        );
      } else if (row[0] === "Date" || row[0] === "date") {
      } else {
        console.error(
          "Date is a string but does not contain a slash or a dash",
          date[0]
        );
        exit(1);
      }

      const formattedDate = moment(date).format("DD-MMM-YY");

      dateList.push(formattedDate);
    } else if (typeof row[0] === "object") {
      const formattedDate = moment(row[0]).format("DD-MM-YY");
      dateList.push(formattedDate);
    } else if (typeof row[0] === null) {
      dateList.push(" ");
    } else {
      console.error(
        `Data type is unknown, value of ${
          row[0]
        } has a data type of ${typeof row[0]}`
      );
    }
  }); //End of For Each

  console.log(dateList);

  writeToTextFileDateList(dateList);
});

const writeToTextFileDateList = (objects) => {
  let file = fs.createWriteStream("output.txt");
  file.on("error", function (err) {
    console.error("There was an eror writing the file");
  });

  objects.forEach((object) => {
    file.write(object + "\n");
  });

  console.log("The job is done");
};

// const writeToTextFile = (objects) => {
//   let file = fs.createWriteStream("output.txt");
//   file.on("error", function (err) {
//     console.error("There was an eror writing the file");
//   });

//   objects.forEach((object) => {
//     file.write(object.newDate + "\n");
//   });

//   console.log("The job is done");
// };
