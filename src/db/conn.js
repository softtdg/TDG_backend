// src/db/conn.js
const { MongoClient } = require('mongodb');
const config = require('../config/config');

const uri = config.mongoUri;
const client = new MongoClient(uri, {
  serverSelectionTimeoutMS: 5000, // optional: quick failure
});

let db;

async function connectDB() {
  if (!db) {
    try {
      await client.connect();
      db = client.db(config.mongoDbName); // change DB name if needed
      console.log('✅ MongoDB connected');
    } catch (err) {
      console.error('❌ MongoDB connection failed:', err.message);
      throw err;
    }
  }
  return db;
}

async function closeDB() {
  if (client && client.isConnected()) {
    await client.close();
    console.log('✅ MongoDB connection closed');
  }
}

module.exports = { connectDB, closeDB, client };
