const path = require('path');
const fetchGaData = require('../../config/ga-connection');
const getMonthDateRange = require("../../utils/monthDateRange")


let dateRanges = getMonthDateRange();
let dimension = [{ name: 'landingPage' }];
let metrics = [
  {
    name: 'sessions',
    type: 'INTEGER'
  },
  {
    name: "screenPageViews",
    type: 'INTEGER'
  }
]

let dimensionFilter = {
  filter: {
    fieldName: 'deviceCategory',
    stringFilter: {
      value: 'mobile',
      matchType: 'EXACT'
    }
  }
}

const landingAndExits = async (workbook) => {

  try {

    let sheet = workbook.sheet("Landing&Exits");

    if (!sheet) {
      sheet = workbook.addSheet("Landing&Exits");
    }

    const landing = await fetchGaData(dateRanges, dimension, metrics, 15)
    appendExcel(sheet, landing[0].rows, 2);

    //adding dimension for filtering only mobile device data
    dimension.push({ name: "deviceCategory" })

    const landingMobile = await fetchGaData(dateRanges, dimension, metrics, dimensionFilter, 15)
    appendExcel(sheet, landingMobile[0].rows, 21);

    await workbook.toFileAsync(path.resolve(__dirname, '..', '..', 'TKM-Monthly-Report-Workbook.xlsx'));

  } catch (error) {
    console.log(error, "error occurred in", __filename);
  };
}


const appendExcel = (sheet, rows, startPoint) => {

  rows.map((data, index) => {
    sheet.cell(String.fromCharCode(65) + (index + startPoint)).value(data.dimensionValues[0].value)
    sheet.cell(String.fromCharCode(66) + (index + startPoint)).value(parseInt(data.metricValues[0].value))
    sheet.cell(String.fromCharCode(67) + (index + startPoint)).value(parseInt(data.metricValues[1].value) / parseInt(data.metricValues[0].value))
    sheet.cell(String.fromCharCode(68) + (index + startPoint)).value(parseInt(data.metricValues[1].value))
  })
}

module.exports = landingAndExits;