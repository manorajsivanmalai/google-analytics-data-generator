const path = require('path');
const fetchGaData = require('../../config/ga-connection');

const currentDate = new Date();
const currentMonth = currentDate.getMonth() + 1;
const nextMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(),1);
const lastDayOfCurrentMonth = new Date(nextMonth - 1);
const daysInCurrentMonth = lastDayOfCurrentMonth.getDate();

let startMonth = currentMonth - 6;
if (startMonth <= 0) {
    startMonth = 12 + startMonth; 
}

let endMonth = currentMonth - 1;  
if (endMonth == 0) {
    endMonth = 12;
}



// Date range for the current month
const dateRanges = {
    startDate: `${currentDate.getFullYear()}-${startMonth.toString().padStart(2, '0')}-01`,
    endDate: `${currentDate.getFullYear()}-${endMonth.toString().padStart(2, '0')}-${daysInCurrentMonth}`
};

// Dimensions and metrics setup
const dimension = [{ name: 'month' }];
const dimensionM = [{ name: 'month' }, { name: 'deviceCategory' }];
const metrics = [{ name: 'sessions', type: 'INTEGER' }];

// Dimension filter for 'mobile' device category
const dimensionFilter = {
    filter: {
        fieldName: 'deviceCategory',
        stringFilter: {
            value: 'mobile',
            matchType: 'CONTAINS'
        }
    }
};


// Function to sort rows by date
const sortRowsByDate = (data) => {
    return data.sort((a, b) => {
        if (a.dimensionValues[0].value < b.dimensionValues[0].value) return -1;
        if (a.dimensionValues[0].value > b.dimensionValues[0].value) return 1;
        return 0;
    });
};

function formatMonthData(monthValue, year) {
    
    const monthNames = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", 
        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
    ];

    const monthIndex = parseInt(monthValue, 10) - 1; 
    const monthName = monthNames[monthIndex];
    // Return the formatted date as "DD-MMM-YYYY"
    return `${monthName}-${year}`;
}

 const  noOfvisitorsMonthwise = async (workbook)=>{

  try{

    let sheet = workbook.sheet("No of Visitors"); 

    if(!sheet){
      sheet = workbook.addSheet("No of Visitors");
    }

    const daywisedata = await fetchGaData(dateRanges,dimension,metrics)
    const daywisedatamobile = await fetchGaData(dateRanges,dimensionM,metrics,dimensionFilter)

    //sorting the data monthwise
    const sortedDWD = sortRowsByDate(daywisedata[0].rows)
    const sortedDWDM = sortRowsByDate(daywisedatamobile[0].rows)
   
    sortedDWD.map((data,index)=>{
         sheet.cell(String.fromCharCode(65)+(index+2)).value(formatMonthData(data.dimensionValues[0].value,currentDate.getFullYear()));
         sheet.cell(String.fromCharCode(66)+(index+2)).value(parseInt(data.metricValues[0].value));
    })

    sortedDWDM.map((data,index)=>{
         sheet.cell(String.fromCharCode(67)+(index+2)).value(formatMonthData(data.dimensionValues[0].value,currentDate.getFullYear()));
         sheet.cell(String.fromCharCode(68)+(index+2)).value(parseInt(data.metricValues[0].value));
    })

    await workbook.toFileAsync(path.resolve(__dirname, '..','..','TKM-Monthly-Report-Workbook.xlsx'));

    console.log("The No Of Visitors sheet has been written successfully.");  

  }catch(error) {
   console.log(err, "error occurred in", __filename);
  }

}


module.exports = { noOfvisitorsMonthwise , formatMonthData };