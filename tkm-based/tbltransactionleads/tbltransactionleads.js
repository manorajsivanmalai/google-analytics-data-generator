const dbConnection =require('../../config/db-connection');
const datePick = require('../../utils/datePick');

const tblTransactionLeads = async (isorganic) => {
 
    const pool =  await dbConnection()
    const { firstDayOfCurrentMonth , firstDayOf7MonthsAgo } = datePick(7);

    // Example query
    const result = await pool.request().query`
        DECLARE @fromDate DATETIME = ${firstDayOf7MonthsAgo}, @toDate DATETIME = ${firstDayOfCurrentMonth};
        SELECT [Display Month], ISNULL([Innova Crysta], 0) AS [Innova Crysta], 
               ISNULL([Camry], 0) AS [Camry], ISNULL([Fortuner], 0) AS [Fortuner],
               ISNULL([Legender], 0) AS [Legender], ISNULL([Vellfire], 0) AS [Vellfire],
               ISNULL([Innova Hycross], 0) AS [Innova Hycross], ISNULL([Hilux], 0) AS [Hilux],
               ISNULL([Glanza], 0) AS [Glanza], ISNULL([Urban Cruiser Hyryder], 0) AS [Urban Cruiser Hyryder],
               ISNULL([Urban Cruiser Taisor], 0) AS [Urban Cruiser Taisor], 
               ISNULL([Land Cruiser 300], 0) AS [Land Cruiser 300], ISNULL([Rumion], 0) AS [Rumion]
        FROM (
            SELECT cast(convert(varchar, tra_TransactDateTime, 23) AS varchar(7)) AS [Month],
                   RIGHT(convert(varchar, tra_TransactDateTime, 106), 8) AS [Display Month], 
                   var_Name AS [Variant], Count([tra_ID]) AS [Count]
            FROM [tb].[dbo].[tblTransaction]
            INNER JOIN [tb].[dbo].tblVariant ON [tb].[dbo].[tblTransaction].var_ID = tblVariant.var_ID
            WHERE trt_ID != 4 
              AND tra_TransactDateTime BETWEEN @fromDate AND @toDate 
              AND tra_IsCampaignLead = ${isorganic?0:1}
            GROUP BY cast(convert(varchar, tra_TransactDateTime, 23) AS varchar(7)),
                     RIGHT(convert(varchar, tra_TransactDateTime, 106), 8),
                     var_Name
        ) AS M
        PIVOT (
            MAX(COUNT) FOR [Variant] IN (
                [Innova Crysta], [Camry], [Fortuner], [Legender], [Vellfire],
                [Innova Hycross], [Hilux], [Glanza], [Urban Cruiser Hyryder],
                [Urban Cruiser Taisor], [Land Cruiser 300], [Rumion]
            )
        ) AS pvt
        ORDER BY [MONTH];
    `;

    return result.recordset;
}

module.exports= tblTransactionLeads;