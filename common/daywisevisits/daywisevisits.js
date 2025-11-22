
const path = require('path');
const gaData =require('../../config/ga-connection');
const getMonthDateRange = require("../../utils/monthDateRange");
const { log } = require('console');
const { type } = require('os');

let dateRanges =getMonthDateRange();
dateRanges.endDate ="2024-08-31"
dateRanges.startDate ="2024-01-01"
let dimension =[{name:"searchTerm"}, { name : "month" }];

let metrics = [
  { name: "totalUsers", type: "INTEGER" },
  { name: "sessions", type: "INTEGER" },
  { name: "averageSessionDuration", type: "FLOAT" },
  { name: "newUsers", type: "INTEGER" },
  { name: "screenPageViews", type: "INTEGER" }
];

//  let dimensionFilter ={
//                 filter: {
//                  fieldName: 'unifiedPagePathScreen',
//                  stringFilter: {
//                     value: '/showroom/vellfire/',
//                     matchType: 'EXACT'
//                 }
//         }
//     }  

const daywisevisits  = async (workbook) =>{
    let sheet = workbook.sheet("daywiseOneMonthVisits");

    if(!sheet){
        workbook.addSheet("daywiseOneMonthVisits");
    }

    const daywisedata = await gaData(dateRanges,dimension,metrics)

  
     
    //  daywisedata[0].rows.map((data,index)=>{
    //     console.log(data);
        
    //     sheet.cell(String.fromCharCode(67)+(index+2)).value(data.dimensionValues[0].value)
    //      sheet.cell(String.fromCharCode(68)+(index+2)).value(parseInt(data.dimensionValues[1].value))
    //     //  console.log(+data.metricValues[1);
         
    //     sheet.cell(String.fromCharCode(69)+(index+2)).value(parseInt(data.metricValues[0].value))
    //  })

    // First, group by date
        // const groupedData = {};

        // daywisedata[0].rows.forEach((data) => {
        // const date = data.dimensionValues[1].value; 
    
        // if (!groupedData[date]) {
        //     groupedData[date] = {
        //     dimensionValues: data,
        //     metricValue: parseInt(data.metricValues[0]?.value)
        //     };
        // } else {
        //     groupedData[date].metricValue += parseInt(data.metricValues[0]?.value); 
        // }
        // });
       
        
        
        // // Now write grouped data to sheet
        // Object.values(groupedData).forEach((item, index) => {
    
        // sheet.cell(String.fromCharCode(67) + (index + 2)).value(item.dimensionValues.dimensionValues[0]?.value);
        // sheet.cell(String.fromCharCode(68) + (index + 2)).value(parseInt(item.dimensionValues.dimensionValues[1]?.value));
        // sheet.cell(String.fromCharCode(69) + (index + 2)).value(formatSecondsToMinSec(item.metricValue));
        // });
var target = 1;
         daywisedata[0].rows.map((data,index)=>{
          console.log(data);
          
              sheet.cell(String.fromCharCode(66)+(target++)).value(data.dimensionValues[1].value);
              data.metricValues.map((matricMomData,index1)=>{
                sheet.cell(String.fromCharCode(66)+(target++)).value(matricMomData.value);
              })
            target+=3;
         })
   
          
        await workbook.toFileAsync(path.resolve(__dirname, '..','..','TKM-Monthly-Report-Workbook.xlsx'));
    
}

function formatSecondsToMinSec(totalSeconds) {
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = Math.floor(totalSeconds % 60);
  return `${minutes} min ${seconds} sec`;
}




module.exports = daywisevisits;