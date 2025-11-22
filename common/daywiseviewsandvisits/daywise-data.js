
const path = require('path');
const fetchGaData = require('../../config/ga-connection');
const getLastMonthDateRange = require("../../utils/monthDateRange");

//date Ranges
const dateRanges = getLastMonthDateRange()

// Dimensions and metrics setup
const dimension = [{ name: 'date' }];
const dimensionM = [{ name: 'date' }, { name: 'deviceCategory' }];
const metrics = [
    { name: 'screenPageViews', type: 'INTEGER' },
    { name: 'newUsers', type: 'INTEGER' }
];

// Dimension filter for 'mobile' device category
const dimensionFilter = {
    filter: {
        fieldName: 'deviceCategory',
        stringFilter: {
            value: 'mobile',
            matchType: 'CONTAINS'
        }
    }
};

// Helper function to format date from YYYYMMDD to DD-MM-YYYY
const formatDate = (dateString) => {
    const year = dateString.substring(0, 4);
    const month = dateString.substring(4, 6);
    const day = dateString.substring(6, 8);
    return `${day}-${month}-${year}`;
};

// Function to sort rows by date
const sortRowsByDate = (data) => {
    return data.sort((a, b) => {
        if (a.dimensionValues[0].value < b.dimensionValues[0].value) return -1;
        if (a.dimensionValues[0].value > b.dimensionValues[0].value) return 1;
        return 0;
    });
};


const daywisedata = async (workbook) => {


    const sheet = workbook.sheet("Site - Month Graph");

    if (!sheet) {
        sheet = workbook.addSheet("Site - Month Graph");
    }

    try {
        const daywisedata = await fetchGaData(dateRanges, dimension, metrics)
        const daywisedatamobile = await fetchGaData(dateRanges, dimensionM, metrics, dimensionFilter)


        const sortedDWD = sortRowsByDate(daywisedata[0].rows)
        const sortedDWDM = sortRowsByDate(daywisedatamobile[0].rows)

        sortedDWD.map((data, index) => {
            //date
            sheet.cell(String.fromCharCode(67) + (index + 4)).value(formatDate(data.dimensionValues[0].value))
            sheet.cell(String.fromCharCode(67) + (index + 45)).value(formatDate(data.dimensionValues[0].value))
            //views
            sheet.cell(String.fromCharCode(68) + (index + 4)).value(parseInt(data.metricValues[0].value))
            //new user
            sheet.cell(String.fromCharCode(68) + (index + 45)).value(parseInt(data.metricValues[1].value))
        })

        sortedDWDM.map((data, index) => {
            // mob views
            sheet.cell(String.fromCharCode(69) + (index + 4)).value(parseInt(data.metricValues[0].value))
            //mob users
            sheet.cell(String.fromCharCode(69) + (index + 45)).value(parseInt(data.metricValues[1].value))
        })

        await workbook.toFileAsync(path.resolve(__dirname, '..', '..', 'TKM-Monthly-Report-Workbook.xlsx'));

        console.log("day wise visitor sheet has been written successfully...");

    } catch (err) {
        console.log(err, "error occurred in", __filename);
    }


}


module.exports = daywisedata;