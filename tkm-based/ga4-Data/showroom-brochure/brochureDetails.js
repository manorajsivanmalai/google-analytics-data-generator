const path = require('path');
const gaData =require('../../../config/ga-connection');
const getMonthDateRange = require("../../../utils/monthDateRange");

const dateRanges = getMonthDateRange();
dateRanges.startDate =`${new Date().getFullYear()}-01-01`;
        
const brochureDownloadedClick = async (sheet)=>{

    const dimensions = [
      { name: 'eventName' },         
      { name: 'customEvent:event_label' }    
    ];

    const metrics1 = [
      { name: 'eventCount' }        
    ];


    const dimensionFilter = {
        filter  : {
            fieldName: 'customEvent:event_label',
            inListFilter: {
              values: [
                "Download Brochure - Glanza",
                "Download Brochure - Urban Cruiser Taisor",
                "Download Brochure - Rumion",
                "Download Brochure - Urban Cruiser Hyryder",
                "Download Brochure - Innova Crysta",
                "Download Brochure - Innova Hycross",
                "Download Brochure - Hilux",
                "Download Brochure - Fortuner",
                "Download Brochure - Legender",
                "Download Brochure - Camry",
                "Download Brochure - Vellfire",
                "Download Brochure - Land Cruiser 300"
              ]
            }
          }
          
        };
        
    const flipBookClickandViews  = await flipBookClick();

    // total
    const total = flipBookClickandViews[0].rows.filter(d=>d.dimensionValues[0].value !='session_start').reduce((prev, data, i) => {
                    if (i === 0) {
                        return parseInt(data.metricValues[0].value); 
                    }
                    return prev + parseInt(data.metricValues[0].value);
                }, 0);
 
     
    
    const carchapterViewBrochure= flipBookClickandViews[0].rows.filter(d=>d.dimensionValues[0].value !='session_start'&& d.dimensionValues[0].value !='Download')  

    const brochureClicks= await gaData(dateRanges,dimensions,metrics1,dimensionFilter)
    
    var index=0;

    //Flipbook Clicks
    sheet.cell(String.fromCharCode(68)+(index+2)).value(total);

    for (let i = 0; i < brochureClicks[0].rowCount; i++) {

    if(brochureClicks[0].rows[i].dimensionValues[0].value==="Download"){
       
       const onlyBrochureAvailable = carchapterViewBrochure.filter((d)=>d.dimensionValues[2].value.includes((brochureClicks[0].rows[i].dimensionValues[1].value).replace("Download Brochure - ","")))
       const onlyBrochureAvailableCount = onlyBrochureAvailable[0]?.metricValues[0]?.value? onlyBrochureAvailable[0]?.metricValues[0]?.value : "-";
         sheet.cell(String.fromCharCode(65)+(index+2)).value(brochureClicks[0].rows[i].dimensionValues[1].value)
         sheet.cell(String.fromCharCode(66)+(index+2)).value(brochureClicks[0].rows[i].metricValues[0].value)
         sheet.cell(String.fromCharCode(67)+(index+2)).value(onlyBrochureAvailableCount)
        
         index++;
    }
    
   }

  }       

const downLoadBrochurePageViews =async (sheet)=>{
  const dimension =[{ name: 'pagePath' }];
  const metrics =[ {name:'screenPageViews',type:'INTEGER'}]
  const dimensionFilter ={
                  filter: {
                  fieldName: 'pagePath',
                  stringFilter: {
                      value: '/brochure/download-brochure/',
                      matchType: 'CONTAINS'
                  } 
              
          }
      }  

    const brochurePageViews= await gaData(dateRanges,dimension,metrics,dimensionFilter);
    sheet.cell(String.fromCharCode(69)+(2)).value(brochurePageViews[0].rows[0].metricValues[0].value)

} 


const flipBookClick = async ()=>{
  
    const dimensions = [
      { name: 'eventName'},         
      { name: 'customEvent:event_label'},
      { name: 'customEvent:event_category'}   
    ];

    const metrics1 = [
      { name: 'eventCount' }        
    ];


    const dimensionFilter ={
                  filter: {
                  fieldName: 'customEvent:event_label',
                  stringFilter: {
                      value: 'View B',
                      matchType: 'CONTAINS'
                  } 
              
          }
      }  

    const flipBookClickandViews= await gaData(dateRanges,dimensions,metrics1,dimensionFilter); 

  return flipBookClickandViews;  
}

 const  brochureDetails = async (workbook)=>{

  try{

    let sheet = workbook.sheet("Brochure And FlipBook Counts"); 

           await downLoadBrochurePageViews(sheet);
           await brochureDownloadedClick(sheet);

      await workbook.toFileAsync(path.resolve(__dirname, '..',"..","..",'TKM-Monthly-Report-Workbook.xlsx'));

    console.log("Brochure And FlipBook Counts");
    
  }catch(error ) {
    console.error("Error:", error);
  };
}



module.exports = brochureDetails;

