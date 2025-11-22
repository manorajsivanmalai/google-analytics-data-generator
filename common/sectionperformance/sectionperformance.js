const path = require('path');
const fetchGaData =require('../../config/ga-connection');
const getMonthDateRange = require("../../utils/monthDateRange");

let dateRanges = getMonthDateRange();

let dimension =[{ name: 'pagePath' }];

let metrics = [{ name:'screenPageViews', type:'INTEGER' },{ name:'newUsers', type:'INTEGER'}]


 const  secperformance = async (workbook)=>{


  try{

    let sheet = workbook.sheet("Section Performance"); 

    if(!sheet){
        sheet = workbook.addSheet("Section Performance"); 
    }

    let targetCell = 0;
    
    for (let i = 0; i < 2; i++) {

     if(i==1){
            targetCell=24;
            dateRanges = { 
                 startDate: `${(new Date().getFullYear())}-01-01`, 
                 endDate: dateRanges.endDate
                 };
             }
            
         const sectionperformance = await fetchGaData(dateRanges,dimension,metrics,10);
         sectionperformance[0].rows.map((data,index)=>{ 
                 sheet.cell(String.fromCharCode(66)+(targetCell+(index+4))).value(data.dimensionValues[0].value)
                 sheet.cell(String.fromCharCode(67)+(targetCell+(index+4))).value(parseInt(data.metricValues[0].value))
                 sheet.cell(String.fromCharCode(69)+(targetCell+(index+4))).value(parseInt(data.metricValues[1].value ))
         });
     
          //  total
         const totalColumns = await fetchGaData(dateRanges,null,metrics) 
         sheet.cell(String.fromCharCode(67)+(targetCell+14)).value(parseInt(totalColumns[0].rows[0].metricValues[0].value ))
         sheet.cell(String.fromCharCode(69)+(targetCell+14)).value(parseInt(totalColumns[0].rows[0].metricValues[1].value ))  
     
    }

       await workbook.toFileAsync(path.resolve(__dirname, '..','..','TKM-Monthly-Report-Workbook.xlsx'));
      console.log( "section perdfromacnce");

  }catch(error){
    console.error("Error:", error, __filename);
  }
}


module.exports = secperformance;

 