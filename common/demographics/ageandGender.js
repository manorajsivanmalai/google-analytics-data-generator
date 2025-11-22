const path = require('path');
const fetchGaData = require('../../config/ga-connection');
const getLastMonthDateRange = require("../../utils/monthDateRange");
const { applyHeaderStyle, applyCellStyle } = require('../../utils/excelStyles')

let dateRanges = getLastMonthDateRange();

//dimention and metrics
let dimension = [{ name: 'userAgeBracket' }];
let metrics = [{ name: 'sessions', type: 'INTEGER' }]

const ageGender = async (workbook) => {

  try {

    const sheet = workbook.sheet("age&gender");

    if (!sheet) {
      sheet = workbook.addSheet("age&gender");
    }


    const age = await fetchGaData(dateRanges, dimension, metrics);
    applyHeaderStyle(sheet, "A1", age[0].dimensionHeaders[0].name);
    applyHeaderStyle(sheet, "B1", age[0].metricHeaders[0].name);
    appendExcel(sheet, age[0].rows, 2)

    // dimension reinitializing
    dimension = [{ name: "userGender" }]

    const gender = await fetchGaData(dateRanges, dimension, metrics)
    applyHeaderStyle(sheet, "A30", gender[0].dimensionHeaders[0].name);
    applyHeaderStyle(sheet, "B30", gender[0].metricHeaders[0].name);
    appendExcel(sheet, gender[0].rows, 31)

    await workbook.toFileAsync(path.resolve(__dirname, '..', '..', 'TKM-Monthly-Report-Workbook.xlsx'));

    console.log("age and gender sheet has been written successfully...");

  }
  catch (error) {
    console.log(error, "error occurred in", __filename);
  };

}


const appendExcel = (sheet, arr, startingPoint) => {
  arr.map((data, index) => {
    applyCellStyle(sheet, String.fromCharCode(65) + (index + startingPoint), data.dimensionValues[0].value)
    applyCellStyle(sheet, String.fromCharCode(66) + (index + startingPoint), parseInt(data.metricValues[0].value))
  })
}

module.exports = ageGender;