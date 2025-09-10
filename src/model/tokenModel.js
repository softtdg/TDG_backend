/**
 * authTokensModel.js
 * @description :: model of a database collection authTokens
 */

const mongoose = require('mongoose');
const mongoosePaginate = require('mongoose-paginate-v2');
const { paginatorCustomLabels } = require('../db/config');
const config = require('../config/config');

mongoosePaginate.paginate.options = { customLabels: paginatorCustomLabels };
const { Schema } = mongoose;
const schema = new Schema(
  {
    userId: {
      type: String,
    },

    token: { type: String },
  },
  {
    timestamps: {
      createdAt: 'createdAt',
      updatedAt: 'updatedAt',
    },
  },
);

schema.method('toJSON', () => {
  const { _id, __v, ...object } = this.toObject({ virtuals: true });
  object.token = undefined;
  return object;
});

schema.plugin(mongoosePaginate);

// Create a connection to the TDG database for tokens
const tdgConnection = mongoose.createConnection(
  `${config.mongoUriTDG}/${config.mongoDbNameTDG}`,
);
const TokensModel = tdgConnection.model('token', schema, 'token');
module.exports = TokensModel;
