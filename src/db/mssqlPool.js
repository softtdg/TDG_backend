const sql = require("mssql");
const config = require("../config/config");

sql.on("error", (err) => {
  console.error("SQL global error:", err);
});

async function getDbPool(databaseName) {
  const sqlConfig = {
    server: config.sqlServer,
    port: config.sqlPort,
    database: databaseName,
    user: config.user,
    password: config.password,
    options: {
      encrypt: false,
      trustServerCertificate: true,
      requestTimeout: 60000, // 60 seconds timeout
      connectionTimeout: 30000, // 30 seconds connection timeout
      pool: {
        max: 10, // Maximum number of connections in pool
        min: 0, // Minimum number of connections in pool
        idleTimeoutMillis: 30000, // Close idle connections after 30 seconds
      },
    },
  };

  // Create a new connection pool for each database (or implement caching if needed)
  try {
    const pool = await sql.connect(sqlConfig);
    return pool;
  } catch (err) {
    console.error("Connection failed:", err);
    throw err;
  }
}

module.exports = getDbPool;
