const jwt = require('jsonwebtoken');
const config = require('../config/config');

const jwtExpiry = config.JWTEXPIRY;
const secretKey = config.SECRET;

/**
 * @description : service to generate JWT token for authentication.
 * @param {String} userName : id of the user.
 * @return {string}  : returns JWT token.
 */
exports.generateToken = async (userName) => {
  let token = jwt.sign(
    {
      userName,
    },
    secretKey,
    { expiresIn: jwtExpiry },
  );

  token = Buffer.from(token).toString('base64');

  return { token };
};
