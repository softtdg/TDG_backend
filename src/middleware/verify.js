const jwt = require('jsonwebtoken');
const config = require('../config/config');
const { query } = require('../db/mssqlPool');
const tokensModel = require('../model/tokenModel');

const checkTokenValidity = async (token) =>
  new Promise((res, rej) => {
    (async () => {
      try {
        const decoded = jwt.verify(token, config.SECRET);

        const tokenData = await tokensModel
          .findOne({ token })
          .sort({ createdAt: -1 });

        if (tokenData && decoded) {
          res({
            user: decoded.userName,
            userId: tokenData.userId,
          });
        } else {
          rej('Your Token Is Expire.....!!');
        }
      } catch (err) {
        rej('Your Token Is Expire.....!');
      }
    })();
  });

module.exports = async (req, res, next) => {
  let token = req.headers.authentication;

  if (!token) {
    return res.unAuthorizedRequest({ message: 'Token Not Found....' });
  }

  checkTokenValidity(token)
    .then(async (result) => {
      const userData = await query(
        'overview',
        `
        SELECT *
        FROM AspNetUsers
        WHERE UserName = @username
      `,
        { username: result.user },
      );

      if (!userData.length) {
        return res.badRequest({ message: 'User Not Found..!' });
      }

      if (userData.length) {
        res.locals.userNAme = userData[0].UserName;
        return next();
      }

      return res.unAuthorizedRequest({
        message: 'Only Authorized User Access This API..!',
      });
    })
    .catch((err) => res.unAuthorizedRequest({ message: err }));

  return undefined;
};
