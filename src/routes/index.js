const express = require('express');

const router = express.Router();

const { testing, generatePickLists, SOPSerchService } = require('../controller/index');

router.get('/testing', testing);
router.get('/SOPSerchService', SOPSerchService);
router.post('/generatePickLists', generatePickLists);

module.exports = router;
