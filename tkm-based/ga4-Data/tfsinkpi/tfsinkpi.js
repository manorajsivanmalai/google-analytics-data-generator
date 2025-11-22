const path = require('path');
const fetchGaData =require('../../../config/ga-connection');
const getMonthDateRange = require("../../../utils/monthDateRange")

let dateRanges = getMonthDateRange();

let dimension =[{ name: 'sessionSource' }];
let metrics =[
                
                {
                    name:'sessions',
                    type:'INTEGER'
                }
            ]   


  let dimensionFilter = {
     filter: {
        fieldName: 'sessionSource',
        stringFilter: {
            value: 'toyotabharat',
            matchType: 'CONTAINS'
        }
    }
  }          

 const  tfsinpageTKM = async (workbook)=>{

 try{

      const sheet = workbook.sheet("tfsinKpi"); 

      const tfsinKpiLndPage = await fetchGaData(dateRanges,dimension,metrics,dimensionFilter,200,"386271876")
      let total = 0;
      tfsinKpiLndPage[0].rows.map((data)=>{
              total += parseInt(data.metricValues[0].value);
          })
          
      sheet.cell(String.fromCharCode(66)+(3)).value(total)


    dimensionFilter = {
        filter: {
            fieldName: 'customEvent:event_label',
            stringFilter: {
                value: 'EMI Calculator',
                matchType: 'CONTAINS'
            }
        }
      }

    const tfsinKpiEMIclc = await fetchGaData(dateRanges,[{name:"customEvent:event_label"}],[ { name: 'eventCount' }],dimensionFilter)
    const emiCl = tfsinKpiEMIclc[0].rows.map((data,index)=>{
          return data.metricValues[0].value;
      })

    sheet.cell(String.fromCharCode(66)+(4)).value(emiCl[0])  

    dimensionFilter = {
        andGroup :{
        expressions:[
            {
            filter: {
              fieldName: 'customEvent:event_label',
              stringFilter: {
                  value: 'Toyota Finance',
                  matchType: 'CONTAINS'
              }
          }
        },{
            filter: {
                  fieldName: 'customEvent:event_category',
                  stringFilter: {
                      value: 'Car Chapters',
                      matchType: 'CONTAINS'
                  }
              }
        }
        ]
        }
      }

   const tfsinKpiApplyLoan= await fetchGaData(dateRanges,[{name:"customEvent:event_label",name:"customEvent:event_category"}],[ { name: 'eventCount' }],dimensionFilter)
   const aploanCount =  tfsinKpiApplyLoan[0].rows.map((data,index)=>{
        return data.metricValues[0].value;
     })
   sheet.cell(String.fromCharCode(66)+(5)).value(aploanCount[0])
     

    //tfsin footerlogo clicks
      dimensionFilter.andGroup.expressions[1].filter.stringFilter.value ="Logo";
      let dimensionFootLogo =[{name:"customEvent:event_label",name:"customEvent:event_category"}];
      let metricsFootLogo =[ { name: 'eventCount' }];

      const tfsinKpifooterLogo= await fetchGaData(dateRanges,dimensionFootLogo,metricsFootLogo,dimensionFilter);
      const footlogo = tfsinKpifooterLogo[0].rows.map((data,index)=>{
        return data.metricValues[0].value;
     })
     
   sheet.cell(String.fromCharCode(66)+(6)).value(footlogo[0])
     
           
   await workbook.toFileAsync(path.resolve(__dirname, '..','..', '..','TKM-Monthly-Report-Workbook.xlsx'));
   console.log("tfsin kpi has been written successfully..");
  
  }catch(error) {
    console.error("Error:", error);
  };
}


 module.exports = tfsinpageTKM;