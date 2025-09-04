const express = require('express');

const router = express.Router();

const {
  testing,
  generatePickLists,
  SOPSerchService,
  fixtureDetails,
  downloadPickList,
  getSheetsBomData,
  downloadupdatedDataSheets,
} = require('../controller/sopSearchController');

router.get('/testing', testing);
router.get('/SOPSerchService', SOPSerchService);
router.get('/fixtureDetails', fixtureDetails);
router.get('/downloadPickList', downloadPickList);
// router.post('/generatePickLists', generatePickLists);
router.get('/getPickListData', getSheetsBomData);
router.post('/downloadupdatedDataSheets', downloadupdatedDataSheets);

module.exports = router;
