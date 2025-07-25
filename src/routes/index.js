const express = require('express');

const router = express.Router();

const { testing, generatePickLists, SOPSerchService, fixtureDetails, downloadPickList } = require('../controller/index');

router.get('/testing', testing);
router.get('/SOPSerchService', SOPSerchService);
router.get('/fixtureDetails', fixtureDetails);
router.get('/downloadPickList', downloadPickList);
// router.post('/generatePickLists', generatePickLists);

module.exports = router;
