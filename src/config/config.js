const config = {
  MODE: process.env.SERVER,
  PORT: process.env.PORT || 4000,
  HOST: process.env.HOST,
  baseUri: process.env.baseURL,
  prefix:"/api",
};
module.exports = config;
