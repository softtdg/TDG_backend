const sql = require('mssql');
const config = require('../config/config');

async function getDbPool(databaseName) {
  const sqlConfig = {
    server: 'localhost',
    port: 1433,
    database: databaseName,
    user: config.user,
    password: config.password,
    options: {
      encrypt: false,
      trustServerCertificate: true,
    },
  };

  // Create a new connection pool for each database (or implement caching if needed)
  const pool = await sql.connect(sqlConfig);
  return pool;
}

module.exports = getDbPool;
