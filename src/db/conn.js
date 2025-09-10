// src/db/conn.js
const { MongoClient } = require('mongodb');
const mongoose = require('mongoose');
const config = require('../config/config');

const uri = config.mongoUri;
const uriTDG = config.mongoUriTDG;

// Create separate clients for each MongoDB instance
const client = new MongoClient(uri, {
  serverSelectionTimeoutMS: 5000,
});

const clientTDG = new MongoClient(uriTDG, {
  serverSelectionTimeoutMS: 5000,
});

let databases = {};

async function connectDB() {
  // Connect Mongoose for models (using BOMs database)
  if (mongoose.connection.readyState === 0) {
    try {
      await mongoose.connect(uri, {
        serverSelectionTimeoutMS: 5000,
        bufferCommands: false,
      });
      console.log('✅ Mongoose connected to MongoDB (BOMs)');
    } catch (err) {
      console.error('❌ Mongoose connection failed:', err.message);
      throw err;
    }
  }

  // Connect native MongoDB client for BOMs database
  if (!client.isConnected?.() && !client.topology?.isConnected()) {
    try {
      await client.connect();
      console.log('✅ MongoDB client connected (BOMs)');
    } catch (err) {
      console.error('❌ MongoDB connection failed (BOMs):', err.message);
      throw err;
    }
  }

  // Connect native MongoDB client for TDG database
  if (!clientTDG.isConnected?.() && !clientTDG.topology?.isConnected()) {
    try {
      await clientTDG.connect();
      console.log('✅ MongoDB client connected (TDG)');
    } catch (err) {
      console.error('❌ MongoDB connection failed (TDG):', err.message);
      throw err;
    }
  }

  // Setup BOMs database
  if (!databases.BOMs) {
    databases.BOMs = client.db(config.mongoDbName);
    console.log('✅ BOMs database connected');
  }

  // Setup TDG database
  if (!databases.TDG) {
    databases.TDG = clientTDG.db(config.mongoDbNameTDG);
    console.log('✅ TDG database connected');
  }

  return databases;
}

async function closeDB() {
  // Close Mongoose connection
  if (mongoose.connection.readyState !== 0) {
    await mongoose.connection.close();
    console.log('✅ Mongoose connection closed');
  }

  // Close BOMs MongoDB client
  if (client && client.isConnected?.()) {
    await client.close();
    console.log('✅ MongoDB connection closed (BOMs)');
  }

  // Close TDG MongoDB client
  if (clientTDG && clientTDG.isConnected?.()) {
    await clientTDG.close();
    console.log('✅ MongoDB connection closed (TDG)');
  }
}

module.exports = { connectDB, closeDB, client, clientTDG };
