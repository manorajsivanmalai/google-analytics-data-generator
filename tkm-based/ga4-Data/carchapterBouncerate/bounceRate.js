const path = require('path');
const gaData = require('../../../config/ga-connection');
const getMonthDateRange = require("../../../utils/monthDateRange");

let dateRanges = getMonthDateRange();
let dimension =[{ name: 'pagePath' }];
let metrics =[ { name:'bounceRate', type:'DOUBLE' },{ name:'screenPageViews', type:'INTEGER'}]
let dimensionFilter ={
  filter: {
     fieldName: 'pagePath',
     inListFilter: {
          values: [
              '/showroom/camry/',
              '/showroom/fortuner/index-fortuner.html', 
              '/showroom/fortuner/index-legender.html',
              '/showroom/glanza/',
              '/showroom/hilux/', 
              '/showroom/innova/',
              '/showroom/innova-crysta/',
              '/showroom/lc300/',
              '/showroom/rumion/',
              '/showroom/urbancruiser-hyryder/',
              '/showroom/urbancruiser-taisor/',
              '/showroom/vellfire/',
           ]
    }   
            
  }
}  


 const  allCarChepterBounceRate = async (workbook)=>{

  try{

    const sheet = workbook.sheet("productsbouncerate"); 

    if(!sheet){
        sheet = workbook.addSheet("productsbouncerate")
    }

    const bounceRate = await gaData(dateRanges,dimension,metrics,dimensionFilter);

    const sortedBR = bounceRate[0].rows.sort((a, b) => {
                if (a.dimensionValues[0].value < b.dimensionValues[0].value) {
                    return -1; 
                }
                if (a.dimensionValues[0].value > b.dimensionValues[0].value) {
                    return 1; 
                }
                return 0; 
            });

    sortedBR.map((data,index)=>{
         sheet.cell(String.fromCharCode(65)+(index+2)).value(data.dimensionValues[0].value)
         sheet.cell(String.fromCharCode(66)+(index+2)).value((data.metricValues[0].value * 100).toFixed(2) +"%")
         sheet.cell(String.fromCharCode(67)+(index+2)).value(parseInt(data.metricValues[1].value))
    })
            
                
     await workbook.toFileAsync(path.resolve(__dirname, '..', 'TKM-Monthly-Report-Workbook.xlsx'));
     console.log("bounceRate written successfully");
   
  }catch(error) {
    console.error("Error:", error);
  };
}

 module.exports = allCarChepterBounceRate;