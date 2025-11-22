const path = require('path');
const gaData = require('../../config/ga-connection');
const getMonthDateRange = require("../../utils/monthDateRange");

let dateRanges = getMonthDateRange();

let dimension = [{ name: 'browser' }];
let metrics = [
  {
    name: 'sessions',
    type: 'INTEGER'
  },
  {
    name: "newUsers",
    type: 'INTEGER'
  }
]


const browserandOs = async (workbook) => {

  try {
    const sheet = workbook.sheet("browser");

    const browser = await gaData(dateRanges, dimension, metrics, 15)

    browser[0].rows.map((data, index) => {
      sheet.cell(String.fromCharCode(65) + (index + 2)).value(data.dimensionValues[0].value)
      sheet.cell(String.fromCharCode(66) + (index + 2)).value(parseInt(data.metricValues[0].value))
      sheet.cell(String.fromCharCode(68) + (index + 2)).value(parseInt(data.metricValues[1].value))
    })

    dimension = [{ name: "operatingSystem" }]

    const sheet1 = workbook.sheet("operatingSystem");

    const os = await gaData(dateRanges, dimension, metrics, 15)

    os[0].rows.map((data, index) => {
      sheet1.cell(String.fromCharCode(65) + (index + 2)).value(data.dimensionValues[0].value)
      sheet1.cell(String.fromCharCode(66) + (index + 2)).value(parseInt(data.metricValues[0].value))
      sheet1.cell(String.fromCharCode(68) + (index + 2)).value(parseInt(data.metricValues[1].value))
    })

    await workbook.toFileAsync(path.resolve(__dirname, '..', '..', 'TKM-Monthly-Report-Workbook.xlsx'));
    console.log("browser and os sheet has been written..");

  } catch (error) {
    console.error("Error:", error);
  }
}


module.exports = browserandOs;