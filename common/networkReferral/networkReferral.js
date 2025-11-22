const path = require('path');
const fetchGaData = require('../../config/ga-connection');
const getMonthDateRange = require("../../utils/monthDateRange")
const { applyCellStyle } = require("../../utils/excelStyles");

let dateRanges = getMonthDateRange();
let dimension = [{ name: 'sessionPrimaryChannelGroup' }, { name: "sessionSourceMedium" }];
let metrics = [{ name: 'sessions', type: 'INTEGER' }, { name: "screenPageViews", type: "INTEGER" }]

let dimensionFilter = {
    andGroup: {
        expressions: [
            {

                filter: {
                    fieldName: 'sessionPrimaryChannelGroup',
                    stringFilter: {
                        value: 'Organic Social',
                        matchType: 'EXACT'
                    }
                }
            },
            {

                filter: {
                    fieldName: 'sessionSourceMedium',
                    stringFilter: {
                        value: 'referral',
                        matchType: 'CONTAINS'
                    }
                }
            }
        ]
    }
}

const platformMap = {
    facebook: ['facebook', 'm.facebook', 'lm.facebook', 'l.facebook'],
    instagram: ['instagram', 'l.instagram'],
    pinterest: ['pinterest'],
    linkedin: ['linkedin', 'lnkd.in'],
    quora: ['quora'],
    reddit: ['reddit'],
    google: ['sites.google']
};

function getPlatformCategory(referralSource) {
    for (let platform in platformMap) {
        if (platformMap[platform].some(source => referralSource.toLowerCase().includes(source))) {
            return platform;
        }
    }
    return referralSource;
}

const netWorkReferral = async (workbook) => {

    try {
        let sheet = workbook.sheet("networkReferral");

        if (!sheet) {
            workbook.addSheet("networkReferral");
        }

        const networkRefferal = await fetchGaData(dateRanges, dimension, metrics, dimensionFilter, 15)

        const aggregatedResults = networkRefferal[0].rows.reduce((acc, item) => {
            const referralSource = item.dimensionValues[1].value;
            const platform = getPlatformCategory(referralSource);
            const metrics = item.metricValues.map(metric => parseInt(metric.value));

            if (!acc[platform]) {
                acc[platform] = { platform, metrics: [...metrics] };
            } else {
                // Add the metric values together for the same platform
                acc[platform].metrics = acc[platform].metrics.map((value, index) => value + metrics[index]);
            }

            return acc;
        }, {});


        // Convert the aggregated results into an array of objects
        const result = Object.values(aggregatedResults);

        result.map((data, index) => {
            applyCellStyle(sheet, String.fromCharCode(65) + (index + 2), data.platform)
            applyCellStyle(sheet, String.fromCharCode(66) + (index + 2), parseInt(data.metrics[0]))
            applyCellStyle(sheet, String.fromCharCode(67) + (index + 2), parseInt(data.metrics[1]))
            applyCellStyle(sheet, String.fromCharCode(68) + (index + 2), parseInt(data.metrics[1]) / parseInt(data.metrics[0]))
        })


        await workbook.toFileAsync(path.resolve(__dirname, '..', '..', 'TKM-Monthly-Report-Workbook.xlsx'));

        console.log("The networkrefferal sheet has been written successfully.");

    } catch (error) {
        console.log(err, "error occurred in", __filename);
    };
}



module.exports = netWorkReferral;