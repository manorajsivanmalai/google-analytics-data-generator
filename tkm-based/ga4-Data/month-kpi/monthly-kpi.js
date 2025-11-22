const path = require('path');
const fetchGaData =require('../../../config/ga-connection');
const getMonthDateRange = require("../../../utils/monthDateRange");

const today = new Date();
const lastMonth = new Date(today.getFullYear(), today.getMonth(), 0); 
const startOfYear = new Date(today.getFullYear(), 0, 0); 
const diff = lastMonth - startOfYear; 
const dayNumberLastMonthEnd = Math.floor(diff / (1000 * 60 * 60 * 24)); 


let dateRanges = getMonthDateRange();
let dimension =[{ name: 'pagePath' }];
let metrics =[
                {
                    name:"bounceRate",
                    type:'DOUBLE'
                },
                {
                    name:"averageSessionDuration",
                    type:"DOUBLE"
                },
                 {
                    name:'screenPageViews',
                    type:'INTEGER'
                },
                {
                    name:"activeUsers",
                    type:"INTEGER"
                }
            ]

let dimensionFilter = {
  orGroup: {
    expressions: [
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value:'/showroom/camry/',
            matchType: 'EXACT'
          }
        }
      },
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value: '/showroom/fortuner/index-fortuner.html', 
            matchType: 'EXACT'
          }
        }
      },
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value: '/showroom/fortuner/index-legender.html', 
            matchType: 'EXACT'
          }
        }
      },
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value: '/showroom/glanza/'.toLowerCase(),
            matchType: 'EXACTLY'
          }
        }
      },
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value:  '/showroom/hilux/', 
            matchType: 'EXACT'
          }
        }
      },
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value: '/showroom/innova/',
            matchType: 'EXACT'
          }
        }
      },
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value:  '/showroom/innova-crysta/',
            matchType: 'EXACT'
          }
        }
      },
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value:  '/showroom/lc300/',
            matchType: 'EXACT'
          }
        }
      },
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value:  '/showroom/rumion/',
            matchType: 'EXACT'
          }
        }
      },
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value:   '/showroom/urbancruiser-hyryder/',
            matchType: 'EXACT'
          }
        }
      },
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value:    '/showroom/urbancruiser-taisor/',
            matchType: 'EXACT'
          }
        }
      },
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value:   '/showroom/vellfire/',
            matchType: 'EXACT'
          }
        }
      },
      {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value: '/',
            matchType: 'EXACT'
          }
        }
      },
        {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            value: '/mobility-solutions/',
            matchType: 'EXACT'
          }
        }
      },
    ]
  }
};

function convertSecondsToTimeFormat(seconds) {
    const hours = Math.floor(seconds / 3600);
    const minutes = Math.floor((seconds % 3600) / 60);
    const remainingSeconds = Math.floor(seconds % 60);

    // Pad single-digit minutes and seconds with leading zeros
    const timeFormatted = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${remainingSeconds.toString().padStart(2, '0')}`;
    return timeFormatted;
}

function sortAndFilterConflictValues(rows){
  return rows.sort((a, b) => {
        if (a.dimensionValues[0].value.replace("/showroom/","") < b.dimensionValues[0].value.replace("/showroom/","")) return -1;
        if (a.dimensionValues[0].value.replace("/showroom/","")  > b.dimensionValues[0].value.replace("/showroom/","")) return 1;
        return 0;
    }).filter(dt=> dt.dimensionValues[0].value !="/showroom/Glanza/" && dt.dimensionValues[0].value !="/showroom/innova-Crysta/");
}

function appendExcel(sheet,data,targetCell,mobilityPostion,totalDays,isYTD=false) {

    let temp=targetCell;

     data.map((data,index)=>{
        if(data.dimensionValues[0].value=="/mobility-solutions/"){
            targetCell= mobilityPostion;
        }else{
             targetCell=temp;
        }
        sheet.cell(String.fromCharCode(isYTD?74-1:74)+(index+targetCell)).value(data.dimensionValues[0].value)   
        sheet.cell(String.fromCharCode(isYTD?68-1:68)+(index+targetCell)).value(parseFloat(data.metricValues[0].value))   
        sheet.cell(String.fromCharCode(isYTD?69-1:69)+(index+targetCell)).value(convertSecondsToTimeFormat(data.metricValues[1].value))   
        sheet.cell(String.fromCharCode(isYTD?70-1:70)+(index+targetCell)).value(parseInt(data.metricValues[2].value))  
        sheet.cell(String.fromCharCode(isYTD?71-1:71)+(index+targetCell)).value(parseInt(data.metricValues[2].value)/totalDays)   
        sheet.cell(String.fromCharCode(isYTD?72-1:72)+(index+targetCell)).value(parseInt(data.metricValues[3].value))   
        sheet.cell(String.fromCharCode(isYTD?73-1:73)+(index+targetCell)).value(parseInt(data.metricValues[3].value)/totalDays)   
     })

  }

 const  kpiReport = async (workbook)=>{

  const totalDaysinMonth =dateRanges.endDate.split("-")[2];

  try{

    let sheet = workbook.sheet("kpimonthly"); 

    if(!sheet){
      sheet =  workbook.addSheet("kpimonthly"); 
    }

   //current month
    const kpimom = await fetchGaData(dateRanges,dimension,metrics,dimensionFilter);
    const sortByName =await sortAndFilterConflictValues(kpimom[0].rows);

    let target=3;
    let mobilityPostion = 16;

    appendExcel(sheet,sortByName,target,mobilityPostion,totalDaysinMonth)
    
    //same data year to till date 
    dateRanges.startDate = `${new Date().getFullYear()}-01-01`;

    const kpimomYTD = await fetchGaData(dateRanges,dimension,metrics,dimensionFilter);

    const filteredValues = kpimomYTD[0].rows.filter((data) => {
          return dimensionFilter.orGroup.expressions.some((filterData) => {
              return data.dimensionValues[0].value === filterData.filter.stringFilter.value;
          });
      });

    const sortByNameYTD =await sortAndFilterConflictValues(filteredValues);

     let targetYTD=32;
     let mobilityPostionYTD = 45;
     appendExcel(sheet,sortByNameYTD,targetYTD,mobilityPostionYTD,dayNumberLastMonthEnd,true);

     await workbook.toFileAsync(path.resolve(__dirname, '..','..','..','TKM-Monthly-Report-Workbook.xlsx'));  
     console.log("monthly kpi sheet written successfuly");
    
  }catch(error) {
    console.error("Error:", error,__filename);
  };
}


module.exports = kpiReport;