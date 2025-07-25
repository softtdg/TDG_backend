const express = require('express');

const router = express.Router();

const { testing, generatePickLists, SOPSerchService, fixtureDetails } = require('../controller/index');

router.get('/testing', testing);
router.get('/SOPSerchService', SOPSerchService);
router.get('/fixtureDetails', fixtureDetails);
router.post('/generatePickLists', generatePickLists);

module.exports = router;
