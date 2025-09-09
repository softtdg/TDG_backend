/**
 * index.js
 * @description :: index route of platforms
 */

const express = require('express');

const router = express.Router();
const verify = require('../middleware/verify');

router.use('/sopSearch', verify, require('./sopSearch.routes'));
router.use('/auth', require('./login.routes'));

module.exports = router;
