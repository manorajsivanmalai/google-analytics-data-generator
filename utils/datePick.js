const moment = require('moment');

const datePick = (monthperiod)=>{
     // Get current date
        let currentDate = moment();

        // Get first day of current month
        let firstDayOfCurrentMonth = currentDate.startOf('month').format('DD-MMM-YYYY');

        // Get first day of the month 7 months ago
        let firstDayOf7MonthsAgo = currentDate.subtract(monthperiod, 'months').startOf('month').format('DD-MMM-YYYY'); 
     
        return {firstDayOfCurrentMonth,firstDayOf7MonthsAgo}
}


module.exports =datePick