const { MongoClient } = require('mongodb');
const config = require('../config/config');
const uri = config.mongoUri;
const client = new MongoClient(uri,{
  serverSelectionTimeoutMS: 5000 // optional: short timeout for quicker failure
});

let db;

async function connectDB() {
  if (!db) {
    try {
      await client.connect();
      db = client.db('BOMs');
      console.log('✅ MongoDB connected');
    } catch (err) {
      console.error('❌ MongoDB connection failed:', err.message);
      throw new Error('MongoDB connection failed');
    }
  }
  return db;
}

module.exports = connectDB;
