const { MongoClient } = require('mongodb');
const config = require('../config/config');
const uri = config.mongoUri;
const client = new MongoClient(uri);

let db;

async function connectDB() {
  if (!db) {
    await client.connect();
    db = client.db('BOMs'); // Always connects to BOMs
  }
  return db;
}

module.exports = connectDB;
