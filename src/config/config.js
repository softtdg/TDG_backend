const config = {
  MODE: process.env.SERVER,
  PORT: process.env.PORT || 3000,
  baseUri: process.env.baseURL,
  prefix: '/api',
  mongoUri: process.env.mongoUri,
  inventoryDomain: process.env.inventoryDomain,
  mongoDbName: process.env.mongoDbName,
  // Base SQL Server configuration
  sqlBaseConfig: {
    server: process.env.sqlServer,
    port: parseInt(process.env.sqlPort),
    user: process.env.sqlUser,
    password: process.env.sqlPassword,
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
  },

  // Database names mapping
  databases: {
    sop: process.env.sop_database,
    qa: process.env.qa_database,
    design: process.env.design_database,
    electrical: process.env.electrical_database,
    overview: process.env.overview_database,
    purchasing: process.env.purchasing_database,
    tdg: process.env.tdg_database,
  },

  // Generate database configurations dynamically
  getDatabaseConfig: function (dbName) {
    const database = this.databases[dbName];
    if (!database) {
      throw new Error(`Database configuration for '${dbName}' not found`);
    }
    return {
      ...this.sqlBaseConfig,
      database: database,
    };
  },

  // Legacy configs for backward compatibility
  get sop_config() {
    return this.getDatabaseConfig('sop');
  },
  get qa_config() {
    return this.getDatabaseConfig('qa');
  },
  get design_config() {
    return this.getDatabaseConfig('design');
  },
  get electrical_config() {
    return this.getDatabaseConfig('electrical');
  },
  get overview_config() {
    return this.getDatabaseConfig('overview');
  },
  get purchasing_config() {
    return this.getDatabaseConfig('purchasing');
  },
  get tdg_config() {
    return this.getDatabaseConfig('tdg');
  },
};
module.exports = config;
