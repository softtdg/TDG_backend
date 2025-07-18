const sql = require('mssql');

async function getDbPool(databaseName) {
  const config = {
    server: 'localhost',
    database: databaseName,
    user: 'tejas',
    password: 'Tejas@123',
    options: {
      encrypt: false,
      trustServerCertificate: true,
    },
  };

  // Create a new connection pool for each database (or implement caching if needed)
  const pool = await sql.connect(config);
  return pool;
}

module.exports = getDbPool;
