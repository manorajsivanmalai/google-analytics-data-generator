// db.js
const sql = require('mssql');

// Define a global connection variable
let poolPromise = null;

// Function to establish or return the existing connection pool
const dbConnection = async () => {
    if (!poolPromise) {
        poolPromise = sql.connect({
            user: process.env.SQL_USER_NAME,  
            password:  process.env.SQL_PASSWORD, 
            server: process.env.SQL_SERVER,
            database:  process.env.SQL_DB_NAME, 
            options: {
                encrypt: false, 
                enableArithAbort: true,
                connectTimeout: 30000 
            }
        });
    }
    return poolPromise;
};

module.exports = dbConnection;
