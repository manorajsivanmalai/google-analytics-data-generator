const XlsxPopulate = require('xlsx-populate');
require('dotenv').config();
const path = require('path');

// Common imports
const affinityCategoryinterest = require("./common/demographics/affinityCategory");
const ageGender = require("./common/demographics/ageandGender");
const landingAndExits = require("./common/landingpageAndExitpage/landingExiting");
const netWorkReferral = require("./common/networkReferral/networkReferral");
const daywisedata = require("./common/daywiseviewsandvisits/daywise-data");
const daywisevisits = require("./common/daywisevisits/daywisevisits");
const { noOfvisitorsMonthwise } = require("./common/noofvisitors/noOfVisitors-last-six-Month");
const refferingSites = require("./common/refferingSite/refferingSites");
const secperformance = require("./common/sectionperformance/sectionperformance");
const browserandOs = require("./common/browserandOs/browser-and-os");
const traficSourceOP = require("./common/traficSource/traficsource");

// TKM-based imports
const allCarChepterBounceRate = require("./tkm-based/ga4-Data/carchapterBouncerate/bounceRate");
const kpiReport = require("./tkm-based/ga4-Data/month-kpi/monthly-kpi");
const brochureDetails = require("./tkm-based/ga4-Data/showroom-brochure/brochureDetails");
const tfsinpageTKM = require("./tkm-based/ga4-Data/tfsinkpi/tfsinkpi");
const virtualShowroom = require("./tkm-based/ga4-Data/virtualshowroom/virtualshowroom");
const onlineRequestTrend = require("./tkm-based/onlinerequesttrend/online-request-trend");
const sheetTransactionLeads = require("./tkm-based/tbltransactionleads/sheetUpdateTransactionsLeads");

const start = async () => {
    try {
        const workbook = await XlsxPopulate.fromFileAsync(
            path.resolve(__dirname, 'TKM-Monthly-Report-Workbook.xlsx')
        );

        // Array of tasks to process sequentially
        const tasks = [
            affinityCategoryinterest,
            daywisedata,
            ageGender,
            landingAndExits,
            netWorkReferral,
            noOfvisitorsMonthwise,
            refferingSites,
            secperformance,
            allCarChepterBounceRate,
            kpiReport,
            brochureDetails,
            tfsinpageTKM,
            virtualShowroom,
            browserandOs,
            traficSourceOP,
            onlineRequestTrend,
            sheetTransactionLeads,
            // daywisevisits // Uncomment if needed
        ];

        // Sequential execution
        for (const task of tasks) {
            try {
                await task(workbook);
            } catch (err) {
                console.error(`Error executing task ${task.name}:`, err);
            }
        }

        console.log("Workbook processing completed successfully.");

    } catch (error) {
        console.error("Error loading workbook:", error);
    }
};

start();
