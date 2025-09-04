const express = require('express');
const path = require('path');
const bodyParser = require('body-parser');
var cors = require('cors');
const http = require('http');

const app = express();

const envPath = path.join(__dirname, `src/config/.env_${process.env.NODE_ENV}`);
require('dotenv').config({ path: envPath, debug: false });

require('./src/db/conn');
const routes = require('./src/routes/index');
const { badRequest } = require('./src/utils/messages');
app.use(require('./src/utils/responseHandler'));
const config = require('./src/config/config');

app.use(cors());

app.use(express.static(path.join(__dirname, 'public')));

app.use(bodyParser.json({ limit: '5mb' }));
app.use(
  express.urlencoded({
    limit: '100mb',
    extended: true,
  }),
);

app.use(function (req, res, next) {
  const origin = req.headers.origin ? req.headers.origin : '*';
  res.setHeader('Access-Control-Allow-Origin', origin);
  res.setHeader(
    'Access-Control-Allow-Methods',
    'GET, POST, OPTIONS, PUT, PATCH, DELETE',
  );
  res.setHeader(
    'Access-Control-Allow-Headers',
    'Authorization, Authentication, Content-Type, origin,action, accept, token,withCredentials',
  );
  res.setHeader('Access-Control-Expose-Headers', 'security_token');
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Cache-Control', 'no-store,max-age=0');
  next();
});

app.use(config.prefix, routes);

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'download.html'));
});

app.use((error, req, res, next) => {
  if (error instanceof SyntaxError) {
    let logData = { error: error };
    logFunction.createLogDb(config.TABLE_ERROR_LOG, {
      store_client_id: 0,
      type: 'Invalid_JSON',
      log: JSON.stringify(logData),
      tech_type: 1,
    });
    return badRequest({ message: 'Invalid Json Formate...!' }, res);
  }
  next();
});

const server = http.createServer(app);
server.listen(config.PORT, () => {
  console.log(`Server Running At Port : ${config.PORT}`);
});
