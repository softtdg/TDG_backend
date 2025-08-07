const config = {
  MODE: process.env.SERVER,
  PORT: process.env.PORT || 3000,
  HOST: process.env.HOST,
  baseUri: process.env.baseURL,
  prefix:"/api",
  mongoUri: process.env.mongoUri,
  user: process.env.sqlUser,
  password: process.env.sqlPassword,
  sqlServer: process.env.sqlServer,
  sqlPort: process.env.sqlPort,
};
module.exports = config;
