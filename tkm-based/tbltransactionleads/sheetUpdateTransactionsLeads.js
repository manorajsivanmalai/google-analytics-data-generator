const path = require('path');
const tblTransactionLeads = require('../tbltransactionleads/tbltransactionleads');
// Load the existing workbook

const  sheetTransactionLeads = async (workbook)=>{

  try{

//    tbltransaction-Paid leads
    const sheet = workbook.sheet("tbltransaction-paid"); 
    const paidLeads = await tblTransactionLeads(false);
  
    paidLeads.map((data,index)=>{
    const monthdata=Object.values(data)
        monthdata.map((val,i)=>{
          sheet.cell(String.fromCharCode(65+i)+(index+2)).value(i==0?val:parseInt(val)); 
        })
    })

//    tbltransaction-organic leads
    const sheet1 = workbook.sheet("tbltransaction-organic");  
    const organicLeads = await tblTransactionLeads(true);
    
    organicLeads.map((data,index)=>{
    const monthdata=Object.values(data)
        monthdata.map((val,i)=>{
          sheet1.cell(String.fromCharCode(65+i)+(index+2)).value(i==0?val:parseInt(val)); 
        })
    })

    await workbook.toFileAsync(path.resolve(__dirname, '..','..', 'TKM-Monthly-Report-Workbook.xlsx'));
    console.log("TBL Transaction organic and paid lead written successfully...");
  } catch(error) {
    console.error("Error:", error);
  }
}

module.exports = sheetTransactionLeads;