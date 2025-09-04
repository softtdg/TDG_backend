/**
 * index.js
 * @description :: index route of platforms
 */

const express = require('express');

const router = express.Router();

router.use('/sopSearch', require('./sopSearch.routes'));
router.use('/auth', require('./login.routes'));

module.exports = router;
