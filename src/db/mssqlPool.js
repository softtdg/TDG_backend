const sql = require('mssql');
const config = require('../config/config');

sql.on('error', err => {
  console.error('SQL global error:', err);
});

async function getDbPool(databaseName) {
  const sqlConfig = {
    server: '127.0.0.1',
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
  try {
    const pool = await sql.connect(sqlConfig);
    return pool;
  } catch (err) {
    console.error('Connection failed:', err);
    throw err;
  }
}

module.exports = getDbPool;
