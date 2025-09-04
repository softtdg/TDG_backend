const crypto = require('crypto');
const sql = require('mssql');
const getDbPool = require('../db/mssqlPool');

const findByUsername = async (username) => {
  try {
    const pool = await getDbPool('OVERVIEW');
    const request = pool.request();
    request.input('username', sql.NVarChar, username);

    const result = await request.query(`
        SELECT *
       FROM [dbo].[AspNetUsers] 
        WHERE UserName = @username
      `);

    return result.recordset[0] || null;
  } catch (error) {
    console.error('Error fetching user:', error);
    throw error;
  }
};

// ASP.NET Identity v3+ password verification
const verifyPassword = (password, passwordHash) => {
  try {
    const hashBytes = Buffer.from(passwordHash, 'base64');

    if (hashBytes.length < 13) {
      console.log('Invalid hash length');
      return false;
    }

    const version = hashBytes[0]; // 0x01
    const prf = hashBytes.readUInt32BE(1); // <-- use BE
    const iterations = hashBytes.readUInt32BE(5); // <-- use BE
    const saltLength = hashBytes.readUInt32BE(9); // <-- use BE

    const salt = hashBytes.slice(13, 13 + saltLength);
    const storedSubkey = hashBytes.slice(13 + saltLength);

    console.log({ version, prf, iterations, saltLength });
    console.log('Salt:', salt.toString('hex'));
    console.log('StoredSubkey:', storedSubkey.toString('hex'));

    if (version !== 1 || prf !== 1) {
      console.log('Unsupported hash format');
      return false;
    }

    const derivedKey = crypto.pbkdf2Sync(
      password,
      salt,
      iterations,
      storedSubkey.length,
      'sha256',
    );

    console.log('DerivedKey:', derivedKey.toString('hex'));

    return crypto.timingSafeEqual(derivedKey, storedSubkey);
  } catch (err) {
    console.error('Password verification error:', err);
    return false;
  }
};

// Fetch user roles (read-only)
const getUserRoles = async (userId) => {
  try {
    const pool = await getDbPool('OVERVIEW');
    const request = pool.request();
    request.input('userId', sql.NVarChar, userId);

    const result = await request.query(`
        SELECT r.Name as RoleName
        FROM AspNetUserRoles ur
        INNER JOIN AspNetRoles r ON ur.RoleId = r.Id
        WHERE ur.UserId = @userId
      `);

    return result.recordset.map((record) => record.RoleName);
  } catch (error) {
    console.error('Error fetching user roles:', error);
    throw error;
  }
};

exports.login = async (req, res) => {
  try {
    const { UserName, Password } = req.body;

    // Validation
    if (!UserName || !Password) {
      return res.badRequest({
        message: 'Username and password are required',
      });
    }

    // Fetch user data from database
    const user = await findByUsername(UserName);
    if (!user) {
      return res.badRequest({
        message: 'Username/password not found',
      });
    }

    // Verify password using fetched PasswordHash
    const isPasswordValid = verifyPassword(String(Password), user.PasswordHash);
    if (!isPasswordValid) {
      return res.badRequest({
        message: 'Username/password not found',
      });
    }

    // Password correct - create session (if session middleware is available)
    if (req.session) {
      req.session.user = {
        id: user.Id,
        username: user.UserName,
        email: user.Email,
      };
    }

    // Fetch user roles and redirect
    const roles = await getUserRoles(user.Id);

    return res.ok({
      message: 'Login successful',
      data: {
        roles: roles || [],
      },
    });
  } catch (error) {
    console.error('Login POST error:', error);
    return res.failureResponse({
      message: 'Server error occurred',
      error: error,
    });
  }
};
