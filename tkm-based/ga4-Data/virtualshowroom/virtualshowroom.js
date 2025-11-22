const path = require('path');
const gaData =require('../../../config/ga-connection');
const getMonthDateRange = require("../../../utils/monthDateRange")


let dateRanges = getMonthDateRange();

const dimension =[{ name: 'pagePath' }];

const metrics =[
                {
                    name:'sessions',
                    type:'INTEGER'
                },
                {
                    name:'screenPageViews',
                    type:'INTEGER'
                },
                {
                    name:'averageSessionDuration',
                    type:'STRING'
                }
            ]

 const dimensionFilter ={
                filter: {
                 fieldName: 'pagePath',
                 stringFilter: {
                    value: '/virtual-showroom/',
                    matchType: 'CONTAINS'
                }
            
        }
    }  


 const  virtualShowroom = async (workbook)=>{

  try{

    let sheet = workbook.sheet("virtualshowroom"); 

    if(!sheet){
      sheet = workbook.addSheet("virtualshowroom"); 
    }

    let targetCell = 0;

    for (let i = 0; i < 2; i++) {

        if(i==1){
                     targetCell=19;
                     dateRanges.startDate=`${new Date().getFullYear()}-01-01`;
                }

        const virtualShowroomPerfomance = await gaData(dateRanges,dimension,metrics,dimensionFilter,15);

        virtualShowroomPerfomance[0].rows.map((data,index)=>{
             sheet.cell(String.fromCharCode(65)+(targetCell+(index+2))).value(data.dimensionValues[0].value)
             sheet.cell(String.fromCharCode(66)+(targetCell+(index+2))).value(parseInt(data.metricValues[0].value))
             sheet.cell(String.fromCharCode(67)+(targetCell+(index+2))).value(parseInt(data.metricValues[1].value ))
             sheet.cell(String.fromCharCode(68)+(targetCell+(index+2))).value((Math.floor(data.metricValues[2].value / 60)>0?Math.floor(data.metricValues[2].value / 60)+"m ":"")+(Math.floor(data.metricValues[2].value % 60).toString().padStart(2,"0"))+"s")
      });

    }
    

     await workbook.toFileAsync(path.resolve(__dirname, '..','..','..','TKM-Monthly-Report-Workbook.xlsx'));
     console.log("virtualshowroom sheet has been written..");
   
  }catch(error) {
    console.error("Error:", error);
  };
}

 module.exports = virtualShowroom;