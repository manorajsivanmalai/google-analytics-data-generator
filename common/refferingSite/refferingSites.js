const path = require('path');
const gaData =require('../../config/ga-connection');
const getMonthDateRange =require("../../utils/monthDateRange")


let dateRanges = getMonthDateRange();
let dimension =[{ name: 'sessionPrimaryChannelGroup' },{name:"sessionSourceMedium"} ];
let metrics =[ {  name:'sessions',  type:'INTEGER' }, { name : "screenPageViews", type : "INTEGER" } ];
let dimensionFilter = {
     filter: {
        fieldName: 'sessionPrimaryChannelGroup',
        stringFilter: {
            value: 'Referral',
            matchType: 'EXACT'
        }
    }
  };

 const  refferingSites = async (workbook)=>{

 try{

    const sheet = workbook.sheet("referringSites"); 

    const refferel = await gaData(dateRanges,dimension,metrics,dimensionFilter,15);

    refferel[0].rows.map((data,index)=>{
        sheet.cell(String.fromCharCode(65)+(index+2)).value(data.dimensionValues[1].value);
        sheet.cell(String.fromCharCode(66)+(index+2)).value(parseInt(data.metricValues[0].value));
        sheet.cell(String.fromCharCode(68)+(index+2)).value(parseInt(data.metricValues[1].value));
    })         
        
    await workbook.toFileAsync(path.resolve(__dirname, '..','..','TKM-Monthly-Report-Workbook.xlsx'));
  
  } catch(error) {
     console.error("Error:", error);
  }

}


module.exports = refferingSites;