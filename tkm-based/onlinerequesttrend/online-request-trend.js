const dbConnection =require('../../config/db-connection');
const datePick = require('../../utils/datePick');
const path = require('path');
const { formatMonthData } =require("../../common/noofvisitors/noOfVisitors-last-six-Month");

const onlineRequestTrend = async (workbook) => {

 try{

    const pool = await dbConnection();
    
    const { firstDayOfCurrentMonth , firstDayOf7MonthsAgo } = datePick(6);
    
      const result = await pool.request().query`SELECT 
        FORMAT(t.tra_TransactDateTime, 'MM-yyyy') AS Month_Year,
        trt.trt_Type,
        COUNT(t.tra_ID) AS Lead
    FROM [tb].[dbo].[tblTransaction] t
    INNER JOIN [tb].[dbo].[tblTransactionType] trt 
        ON trt.trt_ID = t.trt_ID
    WHERE t.tra_TransactDateTime BETWEEN ${firstDayOf7MonthsAgo.toString()} AND ${firstDayOfCurrentMonth.toString()}
    GROUP BY 
        FORMAT(t.tra_TransactDateTime, 'MM-yyyy'),
        trt.trt_Type
    ORDER BY 
        FORMAT(t.tra_TransactDateTime, 'MM-yyyy'), 
        trt.trt_Type;
    `;




    const sheet = workbook.sheet("Online Requests Trends"); 

    const sortedByMonth = result.recordset.sort((a, b) => {
        const [monthA, yearA] = a.Month_Year.split('-').map(Number);
        const [monthB, yearB] = b.Month_Year.split('-').map(Number);

        if (yearA < yearB) return -1;
        if (yearA > yearB) return 1;

        if (monthA < monthB) return -1;
        if (monthA > monthB) return 1;

        return 0;
    });

     let index =0;
     for (let i = 0; i < sortedByMonth.length;i++) {
          let total =0; let sameMY=0;
         for (let j = i; j < sortedByMonth.length; j++) {
             if(sortedByMonth[i].Month_Year == sortedByMonth[j].Month_Year){
                 total  += sortedByMonth[j].Lead;
                 sameMY++;
                 if(i==20){
                        sheet.cell(String.fromCharCode(68)+((j-i)+3)).value(sortedByMonth[j].trt_Type);
                        sheet.cell(String.fromCharCode(69)+((j-i)+3)).value(parseInt(sortedByMonth[j].Lead));
                     }
             }else{
                 break;
             } 
         }
          i+=sameMY-1; 
          sheet.cell(String.fromCharCode(65)+(index+2)).value(formatMonthData(sortedByMonth[i].Month_Year.split('-')[0],sortedByMonth[i].Month_Year.split('-')[1]))
          sheet.cell(String.fromCharCode(66)+(index+2)).value(parseInt(total))
          index++;
        
     }

    await workbook.toFileAsync(path.resolve(__dirname, '..','..','TKM-Monthly-Report-Workbook.xlsx'));

    console.log("online request trend sheet has been written successfully...");
  
  }catch(error){
    console.error("Error:", error);
  };
}

module.exports = onlineRequestTrend;


