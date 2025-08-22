const sql = require("mssql");
const config = require("../config/config");

sql.on("error", (err) => console.error("SQL global error:", err));

// Store pools per database
const databasePools = {};

// Initialize all database connections
async function initializeAllConnections() {
  for (const dbName of Object.keys(config.databases)) {
    try {
      const pool = await new sql.ConnectionPool(config.getDatabaseConfig(dbName)).connect();
      databasePools[dbName] = pool;
      console.log(`✅ Connected to ${dbName}`);
    } catch (err) {
      console.error(`❌ Failed to connect ${dbName}:`, err.message);
      databasePools[dbName] = null;
    }
  }
}

// Run query on a given DB
async function query(dbName, queryStr, params = {}) {
  const pool = databasePools[dbName];
  if (!pool) throw new Error(`No active pool for DB: ${dbName}`);

  const request = pool.request();
  for (const [key, value] of Object.entries(params)) {
    request.input(key, value);
  }

  const result = await request.query(queryStr);
  return result.recordset;
}

// Close all DB connections
async function closeAllConnections() {
  for (const [dbName, pool] of Object.entries(databasePools)) {
    if (pool) {
      try {
        await pool.close();
        console.log(`✅ Closed ${dbName}`);
      } catch (err) {
        console.error(`❌ Error closing ${dbName}:`, err.message);
      }
      delete databasePools[dbName]; // remove from map
    }
  }
}

module.exports = { initializeAllConnections, query, closeAllConnections, databasePools };
