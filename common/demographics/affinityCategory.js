const path = require('path');
const fetchGaData = require('../../config/ga-connection');
const getLastMonthDateRange = require("../../utils/monthDateRange");
const { applyCellStyle, applyHeaderStyle } = require("../../utils/excelStyles")


let dateRanges = getLastMonthDateRange();
let dimension = [{ name: 'brandingInterest' }];
let metrics = [{ name: 'sessions', type: 'INTEGER' }]

const affinityCategoryinterest = async (workbook) => {

   let sheet = workbook.sheet("sesionIntrest");

   if (!sheet) {
      sheet = workbook.addSheet("sesionIntrest");
   }

   try {

      const interestac = await fetchGaData(dateRanges, dimension, metrics, 15);

      //headers
      applyHeaderStyle(sheet, "A1", interestac[0].dimensionHeaders[0].name.toUpperCase());
      applyHeaderStyle(sheet, "B1", interestac[0].metricHeaders[0].name.toLocaleUpperCase());

      //values
      interestac[0].rows.map((data, index) => {
         applyCellStyle(sheet, String.fromCharCode(65) + (index + 2), data.dimensionValues[0].value, false)
         applyCellStyle(sheet, String.fromCharCode(66) + (index + 2), parseInt(data.metricValues[0].value), true)
      })

      await workbook.toFileAsync(path.resolve(__dirname, '..', '..', 'TKM-Monthly-Report-Workbook.xlsx'));

      console.log("The affinity category interest sheet has been written successfully.");

   } catch (err) {
      console.log(err, "error occurred in", __filename);
   }

}


module.exports = affinityCategoryinterest;