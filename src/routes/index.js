const express = require('express');

const router = express.Router();

const { testing, generatePickLists } = require('../controller/index');

router.get('/testing', testing);
router.post('/generatePickLists', generatePickLists);

module.exports = router;
