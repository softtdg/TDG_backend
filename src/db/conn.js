const { MongoClient } = require('mongodb');
const uri = 'mongodb://localhost:27017';
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
