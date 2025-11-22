const path = require('path');
const gaData =require('../../config/ga-connection');
const getMonthDateRange = require("../../utils/monthDateRange")

let dateRanges =getMonthDateRange();

let dimension =[{ name: 'sessionPrimaryChannelGroup' }];

let metrics = [ 
                {
                    name:'sessions',
                    type:'INTEGER'
                }
            ]

 let dimensionFilter ={
                filter: {
                 fieldName: 'deviceCategory',
                 stringFilter: {
                    value: 'mobile',
                    matchType: 'EXACT'
                }
        }
    }  


 const  traficSourceOP = async (workbook)=>{

 try{

    const sheet = workbook.sheet("Traffic source"); 
    let targetCell = 0;
    let traficSource;

   await  sameSlidetwoTable1(sheet);
   await  sameSlidetwoTable2(sheet);

    for (let i = 0; i < 2; i++) {

             if(i==1){
                     targetCell=21;
                     dimension.push({name:"deviceCategory"})
                     traficSource = await gaData(dateRanges,dimension,metrics,dimensionFilter)
                    
                }else{
                     traficSource = await gaData(dateRanges,dimension,metrics)
                }
    
        traficSource[0].rows.map((data,index)=>{
                  sheet.cell(String.fromCharCode(65)+(targetCell+(index+2))).value(data.dimensionValues[0].value)
                  sheet.cell(String.fromCharCode(66)+(targetCell+(index+2))).value(parseInt(data.metricValues[0].value))
        });
     
    }
      total = await gaData(dateRanges,dimension=null,metrics)
      total1 = await gaData(dateRanges,dimension=null,metrics,dimensionFilter)
      sheet.cell(String.fromCharCode(66)+(19)).value(parseInt(total[0].rows[0].metricValues[0].value))
      sheet.cell(String.fromCharCode(66)+(39)).value(parseInt(total1[0].rows[0].metricValues[0].value))
      
      
      await workbook.toFileAsync(path.resolve(__dirname, '..','..','TKM-Monthly-Report-Workbook.xlsx'));
      console.log(" trafic source ");

  }catch(error) {
    console.error("Error:", error);
  };
}


const sameSlidetwoTable1 =async (sheet) =>{

    let dimension =[{ name: 'sessionSourceMedium' }];

    let metrics = [ 
                    {
                        name:'sessions',
                        type:'INTEGER'
                    }
                ]

    let dimensionFilter ={
                    filter: {
                    fieldName: 'deviceCategory',
                    stringFilter: {
                        value: 'mobile',
                        matchType: 'EXACT'
                    }
            }
        }
     
    
   const traficTables =  await gaData(dateRanges,dimension,metrics,10)
   traficTables[0].rows.map((data,index)=>{
      sheet.cell(String.fromCharCode(65)+(index+47)).value(data.dimensionValues[0].value)
      sheet.cell(String.fromCharCode(66)+(index+47)).value(parseInt(data.metricValues[0].value))
   })
   dimension.push({name:"deviceCategory"});

  const traficTablesmobile =  await gaData(dateRanges,dimension,metrics,dimensionFilter,10)
   traficTablesmobile[0].rows.map((data,index)=>{
      sheet.cell(String.fromCharCode(71)+(index+47)).value(data.dimensionValues[0].value)
      sheet.cell(String.fromCharCode(72)+(index+47)).value(parseInt(data.metricValues[0].value))
   })
       

}

const sameSlidetwoTable2 = async (sheet) =>{
       let dimension =[{ name: 'sessionSourceMedium' },{name:"deviceCategory"},];

    let metrics = [ 
                    {
                        name:'sessions',
                        type:'INTEGER'
                    },{
                        name:"screenPageViews",
                        type:'INTEGER'
                    }
                ]

let dimensionFilter = {
  andGroup: {
    expressions: [
      {
        filter: {
          fieldName: 'sessionSourceMedium',
          stringFilter: {
            value: 'organic',
            matchType: 'CONTAINS'
          }
        }
      },
      {
        filter: {
          fieldName: 'deviceCategory',
          stringFilter: {
            value: 'mobile',
            matchType: 'EXACT'
          }
        }
      }
    ]
  }
};

     const traficTables2 =  await gaData(dateRanges,dimension,metrics,dimensionFilter,10)
  
    traficTables2[0].rows.map((data,index)=>{
          sheet.cell(String.fromCharCode(71)+(index+68)).value(data.dimensionValues[0].value)
         sheet.cell(String.fromCharCode(72)+(index+68)).value(parseInt(data.metricValues[0].value))
         sheet.cell(String.fromCharCode(74)+(index+68)).value(parseInt(data.metricValues[1].value))
   })

    dimension.pop({name:"deviceCategory"});
    dimensionFilter ={
        filter: {
            fieldName: 'sessionSourceMedium',
            stringFilter: {
                value: 'organic',
                matchType: 'CONTAINS'
            }
            }
        }
   
     const traficTablesmobile2 =  await gaData(dateRanges,dimension,metrics,dimensionFilter,10)
     
    traficTablesmobile2[0].rows.map((data,index)=>{
      sheet.cell(String.fromCharCode(65)+(index+68)).value(data.dimensionValues[0].value)
      sheet.cell(String.fromCharCode(66)+(index+68)).value(parseInt(data.metricValues[0].value))
       sheet.cell(String.fromCharCode(68)+(index+68)).value(parseInt(data.metricValues[1].value))
   })
}


 module.exports = traficSourceOP;

 