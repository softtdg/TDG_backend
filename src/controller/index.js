const sql = require('mssql');
const ExcelJS = require('exceljs');
const path = require('path');
const getDbPool = require('../db/mssqlPool');
const connectDB = require('../db/conn');
const axios = require('axios');
const os = require('os');
const fs = require('fs');
const fsp = require('fs/promises');

exports.testing = async (req, res) => {
  try {
    const pool = await getDbPool('Purchasing');
    const result = await pool
      .request()
      .query(`SELECT * FROM [dbo].[PurchasingOrders]`);
    res.json({ connected: true, result: result.recordset });

    const db = await connectDB('BOMs');
    const collection = db.collection('ExcelFixture');

    // Demo: Get all documents
    const data = await collection.find({}).limit(10).toArray();

    res.json({ success: true, data });
  } catch (err) {
    console.log('err', err);
    res.status(500).json({ connected: false, error: err.message });
  }
};

const getMasterList = async () => {
  try {
    const pool = await getDbPool('Design');

    const result = await pool.request().query(`
      SELECT 
        ML.MasterListEntryId,
        ML.TDGPN,
        ML.Description,
        ML.Vendor,
        ML.VendorPN,
        ML.Material,
        ML.Finish,
        ML.Size,
        ML.Path,
        ML.Config,
        ML.VariableSize,
        ML.Comments,
        ML.CreatedBy,
        ML.CreatedOn,
        ML.Customer,
        ML.Project,
        ML.Requester,

        -- Foreign Key Relationships
        G.Name AS GroupingName,
        UOM.UOM AS UnitOfMeasure

      FROM [Design].[dbo].[MasterList] ML
      LEFT JOIN [Design].[dbo].[Groupings] G 
        ON G.GroupEntryId = ML.GroupingGroupEntryId
      LEFT JOIN [Design].[dbo].[UOMs] UOM 
        ON UOM.UOMEntryId = ML.UnitOfMeasureUOMEntryId
    `);

    return result.recordset;
  } catch (error) {
    console.error('❌ Error fetching MasterList:', error);
  }
};

const getLeadHandEntry = async (SOPLeadHandEntryId) => {
  try {
    const pool = await getDbPool('SOP');

    const result = await pool
      .request()
      .input('SOPLeadHandEntryId', sql.Int, SOPLeadHandEntryId).query(`
        SELECT TOP 1 *
        FROM [SOP].[dbo].[SOPLeadHandEntries]
        WHERE SOPLeadHandEntryId = @SOPLeadHandEntryId
      `);

    if (result.recordset.length === 0) {
      return null;
    }

    return result.recordset[0];
  } catch (error) {
    console.error('❌ Error fetching LeadHandEntry:', error);
    return null;
  }
};

const fixFixtureName = (fixture) => {
  if (!fixture) {
    return '';
  }

  let fixtureString = fixture.toUpperCase();
  fixtureString = fixtureString.replace(/-?WAR/g, '');
  fixtureString = fixtureString.replace(/-?RPR/g, '');
  fixtureString = fixtureString.replace(/-?EVAL/g, '');

  return fixtureString;
};

const getExplodedBOM = async (fixtureName, db) => {
  const Fixtures = db.collection('Fixture');
  const PDMSubAssemblies = db.collection('PDMSubAssembly');

  // Find matching fixture by name
  const tempResult = await Fixtures.find({ Name: fixtureName }).toArray();

  if (tempResult.length > 0) {
    const item = tempResult[0];

    // Load all fixtures and subassemblies
    const fixtures = await Fixtures.find().toArray();
    const subAssemblies = await PDMSubAssemblies.find().toArray();

    // Add subassemblies to fixtures list
    fixtures.push(...subAssemblies);

    // Explode if item has components
    if (item.Components && item.Components.length > 0) {
      explodeFixture(item, fixtures);
    }

    return item;
  } else {
    return {}; // Equivalent to new Fixture() in C#
  }
};

const getChildren = (parent, componentPool) => {
  const children = [];
  const parentSplit = parent.Level.split('.');
  const parentSplitCount = parentSplit.length;

  for (const item of componentPool) {
    const split = item.Level.split('.');
    const splitCount = split.length;

    if (splitCount > 1 && parentSplitCount < splitCount) {
      if (item.Level.startsWith(parent.Level + '.')) {
        children.push(item);
      }
    }
  }

  return children;
};

const getForceMakeBool = async (TDGPN) => {
  if (!TDGPN) return false;

  try {
    const pool = await getDbPool('Purchasing');

    const result = await pool
      .request()
      .input('TDGPN', sql.VarChar, TDGPN.toLowerCase()).query(`
        SELECT 1 
        FROM [Purchasing].[dbo].[MakePartNumbers]
        WHERE LOWER(TDGPN) = @TDGPN
      `);

    return result.recordset.length > 0;
  } catch (error) {
    console.error('❌ Error in getForceMakeBool:', error);
    return false;
  }
};

const getForceBuyBool = async (TDGPN) => {
  if (!TDGPN) return false;

  try {
    const pool = await getDbPool('Purchasing');

    const result = await pool
      .request()
      .input('TDGPN', sql.VarChar, TDGPN.toLowerCase()).query(`
        SELECT 1
        FROM [Purchasing].[dbo].[BuyPartNumbers]
        WHERE LOWER(TDGPN) = @TDGPN
      `);

    return result.recordset.length > 0;
  } catch (error) {
    console.error('❌ Error in getForceBuyBool:', error);
    return false;
  }
};

/**
 * Checks whether a fixture in the pool should be bought based on path match.
 * @param {string} targetPath - The input path to match.
 * @param {Array<Object>} fixturePool - Array of Fixture objects.
 * @returns {boolean}
 */
const getBuyBool = (targetPath, fixturePool) => {
  if (!targetPath || !Array.isArray(fixturePool)) return false;

  const lowerTarget = targetPath.toLowerCase();

  const match = fixturePool.find((fixture) => {
    if (!fixture.Path) return false;
    const fileName = path.basename(fixture.Path).toLowerCase(); // gets the last part after '\'
    return fileName === lowerTarget;
  });

  return match ? Boolean(match.Buy) : false;
};

const explodeFixture = (item, fixturePool) => {
  for (let i = item.Components.length - 1; i >= 0; i--) {
    const component = item.Components[i];

    if (component.Type === 'P') {
      const group = component.Group;

      if (component.Quantity !== 0) {
        const potentialFamily = item.Components.filter(
          (x) =>
            x.Level.includes(component.Level) &&
            x.Level.split('.').length >= component.Level.split('.').length,
        );

        const potentialChildren = getChildren(component, item.Components);
        const tempComponent = component;

        if (potentialFamily.length > 1 && potentialChildren.length >= 1) {
          if (group === 'MetalPart' || group === 'PCB') {
            const forceMake = getForceMakeBool(component.TDGPN);
            if (!forceMake) {
              for (const child of potentialChildren) {
                child.Quantity = 0;
              }
            } else {
              if (potentialChildren.filter((x) => x.Quantity > 0).length > 0) {
                component.Quantity = 0;
              }
            }
          } else if (group === 'PlasticPart') {
            const forceBuy = getForceBuyBool(component.TDGPN);
            if (forceBuy) {
              for (const child of potentialChildren) {
                child.Quantity = 0;
              }
            } else {
              if (potentialChildren.filter((x) => x.Quantity > 0).length > 0) {
                component.Quantity = 0;
              }
            }
          } else {
            let skipCondition = false;
            for (const child of potentialChildren) {
              if (child.Quantity !== 0) {
                skipCondition = true;
              }
            }
            if (skipCondition) {
              component.Quantity = 0;
            }
          }
        }
      }
    } else if (component.Type === 'S' && component.Group !== 'PCB') {
      const buy = getBuyBool(component.PathName, fixturePool);
      const levelString = component.Level + '.';
      const potentialChildren = getChildren(component, item.Components);

      if (buy) {
        const potentialFamily = item.Components.filter(
          (x) =>
            x.Level.includes(component.Level) &&
            x.Level.split('.').length >= component.Level.split('.').length,
        );

        if (potentialFamily.length > 1 && potentialChildren.length >= 1) {
          for (const child of potentialChildren) {
            if (child.Level !== component.Level) {
              child.Quantity = 0;
            }
          }
        }
      } else {
        const forceMake = getForceMakeBool(component.TDGPN);
        if (!forceMake) {
          let quantityFound = false;
          for (const child of potentialChildren) {
            if (child.Quantity !== 0) {
              quantityFound = true;
            }
          }
          if (quantityFound) {
            component.Quantity = 0;
          }
        } else {
          component.Quantity = 0;
        }
      }
    }
  }
};

const getStoredFixture = async (fixtureNumber, db) => {
  if (!fixtureNumber) return null;

  const collection = db.collection('Fixture'); // Adjust name if needed

  const result = await collection
    .find({ Name: fixtureNumber })
    .limit(1)
    .toArray();

  return result.length > 0 ? result[0] : null;
};

/**
 * @param {string} TDGPN
 * @returns {Promise<InventoryEntry[]>}
 */
const GetInventoryLocations = async (TDGPN) => {
  if (!TDGPN || TDGPN.trim() === '') {
    return [];
  }

  try {
    const response = await axios.get(`http://192.168.2.175:62625/api/inventory/getlocations`, {
      params: { tdgpn: TDGPN },
    });
    return response.data;
  } catch (err) {
    console.error('Error fetching inventory locations:', err.message);
    return [];
  }
};

/**
 * @param {string} TDGPN
 * @returns {Promise<InventoryEntry[]>}
 */
const GetINTLInventoryLocations = async (TDGPN) => {
  if (!TDGPN || TDGPN.trim() === '') {
    return [];
  }

  try {
    const response = await axios.get(`http://192.168.2.175:62625/api/inventory/getintllocations`, {
      params: { tdgpn: TDGPN },
    });
    return response.data;
  } catch (err) {
    console.error('Error fetching INTL inventory locations:', err.message);
    return [];
  }
};

/**
 * @param {string} TDGPN
 * @param {boolean} INTL
 * @returns {Promise<{ location: string, type: boolean, quantity: number }>}
 */
const GetInventoryTuple = async (TDGPN, INTL = false) => {
  let returnString = '';
  let counter = 0;
  let returnQuantity = 0;
  let applyConsumableOrVMI = false;

  const inventoryLocations = INTL
    ? await GetINTLInventoryLocations(TDGPN)
    : await GetInventoryLocations(TDGPN);

  for (const item of inventoryLocations) {
    if (counter++ === 2) break;

    const { ConsumableType, Location, Quantity } = item;
    const newline = os.EOL; // Cross-platform new line

    if (returnString === '') {
      if (ConsumableType === 'CONSUMABLE') {
        returnString += 'CONSUMABLE' + newline + Location;
        applyConsumableOrVMI = true;
      } else if (ConsumableType === 'INHOUSE') {
        returnString += 'INHOUSE' + newline + Location;
        applyConsumableOrVMI = true;
      } else if (ConsumableType === 'VMI') {
        returnString += Location;
        applyConsumableOrVMI = true;
      } else {
        returnQuantity += Quantity;
        returnString += `${Location} (${Math.floor(Quantity)})`;
      }
    } else {
      if (Location && Location !== '') {
        if (ConsumableType === 'VMI') {
          returnQuantity += Quantity;
          returnString += newline + Location;
        } else {
          returnQuantity += Quantity;
          returnString += newline + `${Location} (${Math.floor(Quantity)})`;
        }
      }
    }
  }

  return {
    location: returnString,
    type: applyConsumableOrVMI,
    quantity: returnQuantity,
  };
};

const getUsersInRole = async (roleName) => {
  try {
    const pool = await getDbPool('OVERVIEW');

    const result = await pool
      .request()
      .input('roleName', sql.NVarChar, roleName).query(`
      SELECT u.*
    FROM [OVERVIEW].[dbo].[AspNetUsers] AS u
    INNER JOIN [OVERVIEW].[dbo].[AspNetUserRoles] AS ur ON u.Id = ur.UserId
    INNER JOIN [OVERVIEW].[dbo].[AspNetRoles] AS r ON ur.RoleId = r.Id
    WHERE r.[Name] = @roleName
    `);

    return result.recordset;
  } catch (error) {
    console.error('❌ Error in getUsersInRole:', error);
    return false;
  }
};

const getUserByUsername = async (username) => {
  try {
    const pool = await getDbPool('OVERVIEW');
    const result = await pool
      .request()
      .input('username', sql.NVarChar, username).query(`
        SELECT * FROM [OVERVIEW].[dbo].[AspNetUsers]
        WHERE [UserName] = @username
      `);

    return result.recordset[0] || null; // return user or null if not found
  } catch (err) {
    console.error('Error fetching user by username:', err);
    throw err;
  }
};

function formatDateDMmmYY(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  const day = date.getDate();
  const month = date.toLocaleString('en-US', { month: 'short' });
  const year = date.getFullYear().toString().slice(-2);
  return `${day}-${month}-${year}`;
}

const getINTLInventoryLocations = async (TDGPN) => {  
  if (!TDGPN || TDGPN.trim() === "") {
      return [];
  }

  const baseUri = 'http://192.168.2.175:62625/api/'; // Replace with your actual baseUri

  try {
      const response = await axios.get(`${baseUri}inventory/getintllocations`, {
          params: { tdgpn: TDGPN }
      });
      return response.data; // Assuming response is already a JSON array of InventoryEntry objects
  } catch (error) {
      console.error('Error fetching inventory locations:', error.message);
      return [];
  }
}

async function getInventoryLocations(TDGPN) {
    if (!TDGPN || TDGPN.trim() === "") {
        return [];
    }

    const baseUri = 'http://192.168.2.175:62625/api/'; // Replace with your actual baseUri

    try {
        const response = await axios.get(`${baseUri}inventory/getlocations`, {
            params: { tdgpn: TDGPN }
        });
        return response.data; // Expected to be an array of InventoryEntry objects
    } catch (error) {
        console.error('Error fetching inventory locations:', error.message);
        return [];
    }
}

async function addSOP(
  components,
  fixtureDescription,
  SOP,
  workbook,
  Project,
  Fixture,
  Quantity,
  RequiredDate,
) {
  const worksheet = workbook.addWorksheet(SOP, {
    pageSetup: {
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    },
  });

  // Insert header rows (Row 1-5)
  worksheet.insertRows(1, [
    [
      'SOP #',
      SOP,
      `PICK LIST #${SOP}`,
      '',
      '',
      '',
      'PICK LIST PRINTED ON',
      '',
      new Date().toLocaleDateString(),
    ],
    ['PROJECT', Project, '', '', '', '', 'PICK LIST LOG NUMBER', '', ''],
    ['FIXTURE', Fixture, fixtureDescription, '', '', '', 'DATE PICKED', '', ''],
    ['QUANTITY', Quantity, '', '', '', '', 'LEAD HAND SIGN OFF', '', ''],
    [
      'REQUIRED ON',
      RequiredDate.toLocaleDateString('en-CA'),
      '',
      '',
      '',
      '',
      '',
      '',
      '',
    ],
  ]);

  // Merge cells (ExcelJS uses "A1:B1" format)
  worksheet.mergeCells('C1:F1');
  worksheet.mergeCells('G1:H1');
  worksheet.mergeCells('G2:H2');
  worksheet.mergeCells('C3:F5');
  worksheet.mergeCells('G3:H3');
  worksheet.mergeCells('G4:H5');
  worksheet.mergeCells('I4:I5');

  // Column widths
  const colWidths = [22.28, 69, 22.42, 27, 11.57, 13.28, 20.42, 21.57, 31.57];
  colWidths.forEach((w, i) => (worksheet.getColumn(i + 1).width = w));

  // Table header (Row 7)
  worksheet.getRow(7).values = [
    'TDG PART NO',
    'DESCRIPTION',
    'VENDOR',
    'VENDOR P/N',
    'PER FIX QTY.',
    'TOTAL QTY NEEDED',
    'ACTUAL QTY PICKED',
    'LOCATION/ PURCHASING COMMENTS',
    'LEAD HAND COMMENTS',
  ];
  worksheet.getRow(7).font = { bold: true };
  worksheet.getRow(7).alignment = {
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true,
  };

  // Set row height to auto
  worksheet.getRow(7).height = undefined; // Let Excel auto-adjust

  // Data rows (starting from Row 8)
  components.forEach((comp, idx) => {
    const rowIdx = idx + 8;
    const descParts = (comp.Description || '').split('<line>');
    const goesInto = descParts[0] ? `GOES INTO ${descParts[0]}` : '';
    const restDesc = descParts[1] || '';
    const fullDesc = goesInto ? `${goesInto}\n${restDesc}` : restDesc;

    worksheet.getRow(rowIdx).values = [
      comp.TDGPN,
      fullDesc,
      comp.Vendor,
      comp.VendorPN,
      comp.QuantityPerFixture,
      { formula: `\$B\$4*E${rowIdx}`, result: comp.QuantityNeeded || 0 },
      '',
      comp.Location,
      comp.LeadHandComments,
    ];

    // Alignment & font styling
    ['A', 'C', 'D', 'E', 'F'].forEach((col) => {
      worksheet.getCell(`${col}${rowIdx}`).alignment = { horizontal: 'center' };
    });

    worksheet.getCell(`B${rowIdx}`).alignment = { wrapText: true };
    worksheet.getCell(`A${rowIdx}`).font = { bold: true };
    ['E', 'F', 'G', 'H'].forEach((col) => {
      worksheet.getCell(`${col}${rowIdx}`).font = { bold: true };
    });

    // // Red fill if short
    const quantity = comp.QuantityNeeded || 0;
    const available = comp.QuantityAvailable || 0;

    // Gray fill for INHOUSE/VMI/etc
    const loc = (comp.Location || '').toUpperCase();
    const isGray =
      loc.includes('INHOUSE') ||
      loc.includes('CONSUMABLE') ||
      (loc.includes('V') && !loc.includes('HV')) ||
      quantity === 0;
    if (isGray) {
      for (let col = 1; col <= 9; col++) {
        worksheet.getRow(rowIdx).getCell(col).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD3D3D3' },
        };
      }
    }
  });

  const startRow = 8;
  const endRow = startRow + components.length - 1;

  // Apply gray background to header row (Row 7)
  worksheet.getRow(7).eachCell((cell) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFD3D3D3' }, // Light gray (same used for INHOUSE/etc.)
    };
  });

  // Heading styles
  worksheet.getCell('C1').font = { bold: true, size: 18 };
  worksheet.getCell('C3').font = { size: 18 };
  worksheet.getCell('C1').alignment = { horizontal: 'center', vertical: 'top' };

  // Header alignment
  ['A1', 'A2', 'A3', 'A4', 'A5'].forEach((cell) => {
    worksheet.getCell(cell).alignment = { horizontal: 'left' };
  });
  ['B1', 'B2', 'B3', 'B4', 'B5'].forEach((cell) => {
    worksheet.getCell(cell).alignment = { horizontal: 'right' };
  });
  [
    'G1',
    'G2',
    'G3',
    'G4',
    'G5',
    'H1',
    'H2',
    'H3',
    'H4',
    'H5',
    'I4',
    'I5',
  ].forEach((cell) => {
    worksheet.getCell(cell).alignment = { horizontal: 'center' };
  });
  ['C1', 'C2', 'C3', 'C4', 'C5'].forEach((cell) => {
    worksheet.getCell(cell).alignment = {
      horizontal: 'center',
      vertical: 'top',
    };
  });
}

exports.generatePickLists = async (req, res) => {
  try {
    const vm = req.body.vm; // { LHREntries: [..] }
    const user = req.body.user || null;
    let fixture = req.body.fixture || null;

    let sopNum = '-';
    const ml = await getMasterList();

    const workbook = new ExcelJS.Workbook();

    for (const LHREntryId of vm.LHREntries) {
      let tempSOP = null;
      let tempQuantity = 1;

      if (LHREntryId !== 0) {
        const tempLHREntry = await getLeadHandEntry(LHREntryId);
        fixture = fixFixtureName(tempLHREntry.FixtureNumber);

        // testing
        tempSOP = {
          Program: { Name: '' },
          ODD: new Date(0), // DateTime.MinValue equivalent
        };

        // tempSOP = tempLHREntry.SOP;
        // sopNum = tempSOP.SOPNum;
        tempQuantity = tempLHREntry.Quantity;
      } else {
        tempSOP = {
          Program: { Name: '' },
          ODD: new Date(0), // DateTime.MinValue equivalent
        };
      }

      const db = await connectDB('BOMs');
      const tempFixture = await getExplodedBOM(fixture, db);
      const refFixture = await getStoredFixture(fixture, db);
      const tempComponents = [];

      for (const comp of tempFixture.Components) {
        const split = comp.Level.split('.');
        const parentLevel = split.slice(0, -1).join('.');
        let parent = '';

        const tempParent = tempFixture.Components.find(
          (x) => x.Level === parentLevel,
        );
        if (tempParent) parent = tempParent.TDGPN;

        let quantityPerFixture = Math.round(comp.Quantity);
        const refComp = refFixture.Components.find(
          (x) => x.Level === comp.Level,
        );
        if (refComp) quantityPerFixture = Math.round(refComp.Quantity);

        const tempComp = {
          Description: `${parent}<line>${comp.Description}`,
          TDGPN: comp.TDGPN,
          QuantityPerFixture: quantityPerFixture,
          QuantityNeeded: quantityPerFixture * tempQuantity,
          Vendor: comp.Vendor,
          VendorPN: comp.VendorPN,
        };

        const isDieGroup = ml.find(
          (x) => x.TDGPN === comp.TDGPN && x.GroupingName === 'Die',
        );
        if (isDieGroup) {
          tempComp.QuantityNeeded = 0;
          tempComp.QuantityPerFixture = 0;
        }

        let isIntlUser = false;
        const intlUsers = await getUsersInRole('INTL');
        const currentUser = await getUserByUsername(user);

        if (currentUser && intlUsers) {
          // Check if the current user's ID exists in the INTL users array
          isIntlUser = intlUsers.some(
            (intlUser) => intlUser.Id === currentUser.Id,
          );
        }
        const inventory = await GetInventoryTuple(comp.TDGPN, isIntlUser);
        tempComp.Location = inventory.location;
        tempComp.QuantityAvailable = inventory.quantity;
        tempComp.ConsumableOrVMI = inventory.type;

        tempComponents.push(tempComp);
      }

      console.log('tempSOP-----------------------------------', tempSOP);

      // await addSOP(tempComponents, tempFixture.Description, sopNum, workbook, tempSOP.Program.Name, fixture, tempQuantity, tempSOP.ODD);
      await addSOP(
        tempComponents,
        tempFixture.Description,
        'SOP-123',
        workbook,
        tempSOP.Program.Name,
        fixture,
        tempQuantity,
        tempSOP.ODD,
      );
    }

    // const fileName = `${sopNum} (${fixture}) ${new Date().toLocaleDateString('en-US', { month: 'short', day: '2-digit' })}.xlsx`;
    // const filePath = path.join(__dirname, 'exports', fileName);

    // // Ensure the exports directory exists
    // const exportsDir = path.join(__dirname, 'exports');
    // if (!fs.existsSync(exportsDir)) {
    //   await fsp.mkdir(exportsDir, { recursive: true });
    // }

    // await workbook.xlsx.writeFile(filePath);

    // res.download(filePath, fileName);

    // ✅ Place these lines at the end of your controller function:
    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      'Content-Disposition',
      'attachment; filename="PickList.xlsx"',
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.send(buffer);
    res.end();
  } catch (err) {
    console.error('Error generating picklists:', err);
    res.status(500).send('Error generating picklists');
  }
};
