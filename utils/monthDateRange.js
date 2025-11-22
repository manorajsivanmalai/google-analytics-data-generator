const getMonthDateRange = () => {
    
    const currentDate = new Date();
    const currentMonth = currentDate.getMonth() + 1;
    const nextMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
    const lastDayOfCurrentMonth = new Date(nextMonth - 1);
    const daysInCurrentMonth = lastDayOfCurrentMonth.getDate();

    let month = currentMonth - 1;
    if (month == 0) {
        month = 12;
    }

    return {
        startDate: `${currentDate.getFullYear()}-${month.toString().padStart(2, '0')}-01`,
        endDate: `${currentDate.getFullYear()}-${month.toString().padStart(2, '0')}-${daysInCurrentMonth}`
    };
};



module.exports = getMonthDateRange;