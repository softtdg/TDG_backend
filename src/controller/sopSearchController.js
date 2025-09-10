const ExcelJS = require('exceljs');
const path = require('path');
const { connectDB } = require('../db/conn');
const axios = require('axios');
const os = require('os');
const config = require('../config/config');
const { query } = require('../db/mssqlPool');

exports.testing = async (req, res) => {
  try {
    const result = await query('sop', 'SELECT TOP 2 * FROM SOPs');

    const db = await connectDB();
    const collection = db.BOMs.collection('Fixture');

    // Demo: Get all documents
    const datas = await collection.find({}).limit(10).toArray();

    res.json({ success: true, data: { result, datas } });
  } catch (err) {
    console.log('err', err);
    res.status(500).json({ connected: false, error: err.message });
  }
};

const getMasterList = async () => {
  try {
    const result = await query(
      'design',
      `
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

      FROM MasterList ML
      LEFT JOIN Groupings G 
        ON G.GroupEntryId = ML.GroupingGroupEntryId
      LEFT JOIN UOMs UOM 
        ON UOM.UOMEntryId = ML.UnitOfMeasureUOMEntryId
    `,
    );

    return result;
  } catch (error) {
    console.error('❌ Error fetching MasterList:', error);
  }
};

const getAllDieAndLabelList = async () => {
  try {
    const result = await query(
      'design',
      `
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

      FROM MasterList ML
      LEFT JOIN Groupings G 
        ON G.GroupEntryId = ML.GroupingGroupEntryId
      LEFT JOIN UOMs UOM 
        ON UOM.UOMEntryId = ML.UnitOfMeasureUOMEntryId
      WHERE G.Name IN ('Die', 'Label');
    `,
    );

    return result;
  } catch (error) {
    console.error('❌ Error fetching MasterList:', error);
  }
};

const getLeadHandEntry = async (SOPLeadHandEntryId) => {
  try {
    const result = await query(
      'sop',
      `
        SELECT TOP 1 *
        FROM SOPLeadHandEntries
        WHERE SOPLeadHandEntryId = @SOPLeadHandEntryId
      `,
      { SOPLeadHandEntryId },
    );

    if (result.length === 0) {
      return null;
    }

    return result[0];
  } catch (error) {
    console.error('❌ Error fetching LeadHandEntry:', error);
    return null;
  }
};

const getSOPIdtoSopData = async (SOPId) => {
  try {
    const result = await query(
      'sop',
      `
        SELECT TOP 1 *
        FROM SOPs
        WHERE SOPId = @SOPId
      `,
      { SOPId },
    );

    if (result.length === 0) {
      return null;
    }

    return result[0];
  } catch (error) {
    console.error('❌ Error fetching SOPIdtoSopData:', error);
    return null;
  }
};

const getProgramNameByProgramId = async (SOPProgramId) => {
  try {
    const result = await query(
      'sop',
      `
        SELECT TOP 1 *
        FROM SOPPrograms
        WHERE SOPProgramId = @SOPProgramId
      `,
      { SOPProgramId },
    );

    if (result.length === 0) {
      return null;
    }

    return result[0];
  } catch (error) {
    console.error('❌ Error fetching SOPIdtoSopData:', error);
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
  const Fixtures = db.BOMs.collection('Fixture');
  const PDMSubAssemblies = db.BOMs.collection('PDMSubAssembly');

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
    const result = await query(
      'purchasing',
      `
        SELECT 1 
        FROM MakePartNumbers
        WHERE LOWER(TDGPN) = @TDGPN
      `,
      { TDGPN: TDGPN.toLowerCase() },
    );

    return result.length > 0;
  } catch (error) {
    console.error('❌ Error in getForceMakeBool:', error);
    return false;
  }
};

const getForceBuyBool = async (TDGPN) => {
  if (!TDGPN) return false;

  try {
    const result = await query(
      'purchasing',
      `
        SELECT 1
        FROM BuyPartNumbers
        WHERE LOWER(TDGPN) = @TDGPN
      `,
      { TDGPN: TDGPN.toLowerCase() },
    );

    return result.length > 0;
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

  const collection = db.BOMs.collection('Fixture'); // Adjust name if needed

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
    const response = await axios.get(
      `${config.inventoryDomain}/api/inventory/getlocations`,
      {
        params: { tdgpn: TDGPN },
      },
    );
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
    const response = await axios.get(
      `${config.inventoryDomain}/api/inventory/getintllocations`,
      {
        params: { tdgpn: TDGPN },
      },
    );
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
    const sqlQuery = `
      SELECT u.*
      FROM [OVERVIEW].[dbo].[AspNetUsers] AS u
      INNER JOIN [OVERVIEW].[dbo].[AspNetUserRoles] AS ur ON u.Id = ur.UserId
      INNER JOIN [OVERVIEW].[dbo].[AspNetRoles] AS r ON ur.RoleId = r.Id
      WHERE r.[Name] = @roleName
    `;

    // params object matches input names in the query
    const params = { roleName };

    const users = await query('overview', sqlQuery, params);

    return users;
  } catch (error) {
    console.error('❌ Error in getUsersInRole:', error);
    return false;
  }
};

const getUserByUsername = async (username) => {
  try {
    const sqlQuery = `
        SELECT * FROM AspNetUsers
        WHERE [UserName] = @username
    `;

    const params = { username };

    const user = await query('overview', sqlQuery, params);

    return user[0] || null; // return user or null if not found
  } catch (err) {
    console.error('Error fetching user by username:', err);
    throw err;
  }
};

const isValidNumberLocationFormat = (location) => {
  return /^\d+-\d+-\d+$/.test(location);
};

function getCellText(cell) {
  if (!cell || cell.value == null) return '';

  if (typeof cell.value === 'string' || typeof cell.value === 'number') {
    return cell.value.toString();
  }

  // RichText
  if (cell.value.richText) {
    return cell.value.richText.map((rt) => rt.text).join('');
  }

  // Hyperlink or formula
  if (cell.value.text) return cell.value.text;
  if (cell.value.result) return cell.value.result.toString();

  return cell.value.toString();
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
  // Validate and sanitize worksheet name
  let worksheetName = SOP;
  if (!worksheetName || worksheetName.trim() === '') {
    worksheetName = 'Sheet1'; // Default fallback name
  } else {
    // Sanitize the name to remove invalid characters
    worksheetName = worksheetName.replace(/[*?:\\/[\]]/g, '_');
    worksheetName = worksheetName.trim();
    if (worksheetName.length > 31) {
      worksheetName = worksheetName.substring(0, 31);
    }
  }

  const worksheet = workbook.addWorksheet(worksheetName, {
    pageSetup: {
      orientation: 'landscape',
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    },
  });

  // Format the current date to "Month Day, Year"
  const formattedPrintedDate = new Date().toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });

  // Format the required date to "D-MMM-YY" (e.g., 1-Jan-01)
  const formattedRequiredDate = RequiredDate
    ? new Date(RequiredDate)
        .toLocaleDateString('en-GB', {
          day: 'numeric',
          month: 'short',
          year: '2-digit',
        })
        .replace(/ /g, '-')
    : '';

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
      formattedPrintedDate,
    ],
    ['PROJECT', Project, '', '', '', '', 'PICK LIST LOG NUMBER', '', ''],
    ['FIXTURE', Fixture, fixtureDescription, '', '', '', 'DATE PICKED', '', ''],
    ['QUANTITY', Quantity, '', '', '', '', 'LEAD HAND SIGN OFF', '', ''],
    ['REQUIRED ON', formattedRequiredDate, '', '', '', '', '', '', ''],
  ]);

  // Merges
  [
    'C1:F1',
    'C2:F2',
    'G1:H1',
    'G2:H2',
    'C3:F5',
    'G3:H3',
    'G4:H5',
    'I4:I5',
  ].forEach((range) => worksheet.mergeCells(range));

  // Column widths
  [22.28, 69, 22.42, 27, 11.57, 14, 20.42, 21.57, 31.57].forEach(
    (w, i) => (worksheet.getColumn(i + 1).width = w),
  );

  // Row 7 headers
  const headers = [
    'TDG PART NO',
    'DESCRIPTION',
    'VENDOR',
    'VENDOR P/N',
    'PER FIX QTY.',
    'TOTAL QTY NEEDED',
    'ACTUAL QTY TO BE PICKED',
    'LOCATION/ PURCHASING COMMENTS',
    'LEAD HAND COMMENTS',
  ];
  worksheet.getRow(7).values = headers;
  worksheet.getRow(7).height = undefined;
  worksheet.getRow(7).eachCell((cell) => {
    cell.font = { bold: true, size: 18, name: 'Calibri' };
    cell.alignment = {
      horizontal: 'center',
      vertical: 'middle',
      wrapText: true,
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'D9E1F2' },
    };
  });

  // Apply borders and fonts to header rows
  for (let r = 1; r <= 5; r++) {
    for (let c = 1; c <= 9; c++) {
      const cell = worksheet.getRow(r).getCell(c);
      cell.border = {
        top: { style: 'medium' },
        bottom: { style: 'medium' },
        left: { style: 'medium' },
        right: { style: 'medium' },
      };
    }
  }

  // Apply bold borders to second table (starting from Row 7)
  for (let row = 7; row <= worksheet.rowCount; row++) {
    for (let col = 1; col <= 9; col++) {
      const cell = worksheet.getRow(row).getCell(col);
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
      };
    }
  }

  // Font + fill for header info cells
  const headerFontCells = ['A', 'B', 'G'].flatMap((col) =>
    [1, 2, 3, 4].map((row) => `${col}${row}`),
  );
  headerFontCells
    .concat(['A5', 'B5', 'I1', 'I2', 'I3', 'I4', 'I5'])
    .forEach((cell) => {
      const c = worksheet.getCell(cell);
      c.font = { bold: true, size: 18, name: 'Calibri' };
      c.alignment = { horizontal: 'center', vertical: 'middle' };
      if (
        cell.startsWith('A') ||
        cell.startsWith('B') ||
        cell.startsWith('G')
      ) {
        c.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'D9E1F2' },
        };
      }
    });

  for (let r = 1; r <= 5; r++) {
    worksheet.getCell(`I${r}`).alignment = {
      horizontal: 'center',
      vertical: 'middle',
    };
  }

  //   // Data rows
  //   components.forEach((comp, i) => {
  //     const row = worksheet.getRow(i + 8);
  //     const descParts = (comp.Description || "").split("<line>");
  //     const parent = descParts[0];
  //     const isTopLevel = parent && parent.trim() !== "" && parent === comp.TDGPN;

  //     // Create rich text for description with bold "GOES INTO" part
  //     let richText = [];
  //     if (parent) {
  //       richText.push({
  //         text: `GOES INTO ${parent}\n`,
  //         font: { bold: true, size: 18, name: "Calibri" },
  //       });
  //     }
  //     if (descParts[1]) {
  //       richText.push({
  //         text: descParts[1],
  //         font: { bold: false, size: 18, name: "Calibri" },
  //       });
  //     }

  //     // Determine if this is a consumable item and vmi
  //     const isConsumable =
  //       comp.ConsumableOrVMI ||
  //       (comp.Location && (
  //         comp.Location.toUpperCase().includes("CONSUMABLE")
  //         || comp.Location.toUpperCase().includes("VMI")
  //       )) ||
  //       (comp.LeadHandComments && (
  //         comp.LeadHandComments.toUpperCase().includes("CONSUMABLE")
  //         || comp.LeadHandComments.toUpperCase().includes("VMI")
  //       ));

  //     let totalQty = 0;

  //     const isWire =
  //       comp.Description && comp.Description.toUpperCase().includes("WIRE");

  //     if (isConsumable) {
  //       // For consumables, show the per-fixture quantity (no multiplication)
  //       totalQty = comp.QuantityPerFixture || 0;
  //     } else if (comp.TDGPN.includes("LABEL")) {
  //       totalQty = 0;
  //     } else if (isWire) {
  //       totalQty = 0;
  //     } else {
  //       // For regular items, multiply by fixture quantity
  //       totalQty = (comp.QuantityPerFixture || 0) * Quantity;
  //     }

  //     row.values = [
  //       comp.TDGPN,
  //       richText.length > 0 ? { richText } : descParts[1] || "",
  //       comp.Vendor,
  //       comp.VendorPN,
  //       comp.QuantityPerFixture,
  //       totalQty,
  //       "",
  //       comp.Location,
  //       comp.LeadHandComments,
  //     ];

  //     // create table border for description.....
  //     for (let c = 1; c <= 9; c++) {
  //       const cell = row.getCell(c);
  //       // Apply normal font for B, C, D; bold for others
  //       const isNormalFont = c === 2 || c === 3 || c === 4;
  //       cell.font = { size: 18, name: "Calibri", bold: !isNormalFont };
  //       cell.alignment = {
  //         wrapText: true,
  //         vertical: "middle",
  //         ...(c <= 6 && c !== 2 ? { horizontal: "center" } : {}),
  //       };
  //       cell.border = {
  //         top: { style: "thin" },
  //         bottom: { style: "thin" },
  //         left: { style: "thin" },
  //         right: { style: "thin" },
  //       };
  //     }

  //     const qty = totalQty;
  //     const available = comp.QuantityAvailable || 0;
  //     if (qty > available && !isConsumable) {
  //       row.getCell(6).fill = {
  //         type: "pattern",
  //         pattern: "lightTrellis",
  //         fgColor: { argb: "FFFF0000" },
  //       };
  //     }

  //     const loc = (comp.Location || "").toUpperCase();
  //     const locationNumberCheck = comp.Location ? isValidNumberLocationFormat(comp.Location) : false;

  //     const isGray =
  //       !comp.Location ||                                  // 1. Location is missing or empty
  //       isConsumable ||                                    // 2. Item is consumable
  //       loc.includes("INHOUSE") ||                         // 3. Location contains "INHOUSE"
  //       loc.includes("VMI") ||                             // 4. Location contains "VMI"
  //       (loc.includes("V") && !loc.includes("HV")) ||      // 5. Location has "V" but not "HV"
  //       qty === 0
  //       || locationNumberCheck ;                             // 7. Location not in "number-number-number" format
  //     if (isGray) {
  //       for (let c = 1; c <= 9; c++) {
  //         row.getCell(c).fill = {
  //           type: "pattern",
  //           pattern: "solid",
  //           fgColor: { argb: "FFD3D3D3" },
  //         };
  //       }
  //     }
  //   });

  // // Step 1: Build maps
  // const partInfo = new Map();  // partNo → { vendor, hasChildren }

  // worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
  //   if (rowNumber < 8) return;

  //   const tdgpn = (row.getCell(1).value || "").toString().trim().toUpperCase();
  //   const desc = getCellText(row.getCell(2)).toUpperCase().trim();
  //   const vendor = (row.getCell(3).value || "").toString().trim().toUpperCase();

  //   if (!tdgpn) return;

  //   // Save vendor for each part
  //   if (!partInfo.has(tdgpn)) {
  //     partInfo.set(tdgpn, { vendor, hasChildren: false });
  //   }

  //   // If this row is a child ("GOES INTO ..."), mark parent as having children
  //   const match = desc.match(/GOES INTO\s+([A-Z0-9-]+)/);
  //   if (match) {
  //     const parentPart = match[1];
  //     if (partInfo.has(parentPart)) {
  //       partInfo.get(parentPart).hasChildren = true;
  //     } else {
  //       partInfo.set(parentPart, { vendor: "", hasChildren: true });
  //     }
  //   }
  // });

  // // Step 2: Apply coloring logic
  // worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
  //   if (rowNumber < 8) return;

  //   const tdgpn = (row.getCell(1).value || "").toString().trim().toUpperCase();
  //   const desc = getCellText(row.getCell(2)).toUpperCase().trim();

  //   const match = desc.match(/GOES INTO\s+([A-Z0-9-]+)/);
  //   const isChild = !!match;

  //   let shouldGray = false;

  //   if (isChild) {
  //     const parentPart = match[1];
  //     const parentInfo = partInfo.get(parentPart) || { vendor: "" };
  //     const parentVendor = parentInfo.vendor || "";

  //     // GOES INTO row: gray if parent vendor is NOT TDG or FASTENAL
  //     shouldGray = !(parentVendor.includes("TDG") || parentVendor.includes("FASTENAL"));
  //   } else {
  //     const info = partInfo.get(tdgpn) || { vendor: "", hasChildren: false };
  //     const { vendor, hasChildren } = info;

  //     // Main row: gray only if vendor is TDG or FASTENAL and has children
  //     shouldGray = (vendor.includes("TDG") || vendor.includes("FASTENAL")) && hasChildren;
  //   }

  //   // Apply fill color
  //   for (let c = 1; c <= 9; c++) {
  //     row.getCell(c).fill = shouldGray
  //       ? {
  //           type: "pattern",
  //           pattern: "solid",
  //           fgColor: { argb: "FFD3D3D3" }, // light gray
  //         }
  //       : null; // white
  //   }
  // });

  // Step 1: Populate the data first (component loop)
  components.forEach((comp, i) => {
    const row = worksheet.getRow(i + 8);
    const descParts = (comp.Description || '').split('<line>');
    const parent = descParts[0];
    const isTopLevel = parent && parent.trim() !== '' && parent === comp.TDGPN;

    // Create rich text for description with bold "GOES INTO" part
    let richText = [];
    if (parent) {
      richText.push({
        text: `GOES INTO ${parent}\n`,
        font: { bold: true, size: 18, name: 'Calibri' },
      });
    }
    if (descParts[1]) {
      richText.push({
        text: descParts[1],
        font: { bold: false, size: 18, name: 'Calibri' },
      });
    }

    // Determine if this is a consumable item and vmi
    const isConsumable =
      comp.ConsumableOrVMI ||
      (comp.Location &&
        (comp.Location.toUpperCase().includes('CONSUMABLE') ||
          comp.Location.toUpperCase().includes('VMI'))) ||
      (comp.LeadHandComments &&
        (comp.LeadHandComments.toUpperCase().includes('CONSUMABLE') ||
          comp.LeadHandComments.toUpperCase().includes('VMI')));

    let totalQty = 0;
    const isWire =
      comp.Description && comp.Description.toUpperCase().includes('WIRE');

    if (isConsumable) {
      totalQty = comp.QuantityPerFixture || 0;
    } else if (comp.TDGPN.includes('LABEL') || isWire) {
      totalQty = 0;
    } else {
      totalQty = (comp.QuantityPerFixture || 0) * Quantity;
    }

    // Set row values
    row.values = [
      comp.TDGPN,
      richText.length > 0 ? { richText } : descParts[1] || '',
      comp.Vendor,
      comp.VendorPN,
      comp.QuantityPerFixture,
      totalQty,
      '',
      comp.Location,
      comp.LeadHandComments,
    ];

    // Apply normal font, borders, and formatting
    for (let c = 1; c <= 9; c++) {
      const cell = row.getCell(c);
      const isNormalFont = c === 2 || c === 3 || c === 4; // B, C, D columns should have normal font
      cell.font = { size: 18, name: 'Calibri', bold: !isNormalFont };
      cell.alignment = {
        wrapText: true,
        vertical: 'middle',
        ...(c <= 8 && c !== 2 ? { horizontal: 'center' } : {}),
      };
      cell.border = {
        top: { style: 'thin' },
        bottom: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
      };
    }
  });

  // Step 2: Apply color logic after the data is populated (workbook loop)

  const partInfo = new Map(); // partNo → { vendor, hasChildren }

  // Apply vendor-based coloring logic
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber < 8) return;

    const tdgpn = (row.getCell(1).value || '').toString().trim().toUpperCase();
    const desc = getCellText(row.getCell(2)).toUpperCase().trim();
    const vendor = (row.getCell(3).value || '').toString().trim().toUpperCase();

    if (!tdgpn) return;

    // Save vendor for each part
    if (!partInfo.has(tdgpn)) {
      partInfo.set(tdgpn, { vendor, hasChildren: false });
    }

    // If this row is a child ("GOES INTO ..."), mark parent as having children
    const match = desc.match(/GOES INTO\s+([A-Z0-9-]+)/);
    if (match) {
      const parentPart = match[1];
      if (partInfo.has(parentPart)) {
        partInfo.get(parentPart).hasChildren = true;
      } else {
        partInfo.set(parentPart, { vendor: '', hasChildren: true });
      }
    }
  });

  // Apply vendor-based coloring for main and child rows
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber < 8) return;

    const tdgpn = (row.getCell(1).value || '').toString().trim().toUpperCase();
    const desc = getCellText(row.getCell(2)).toUpperCase().trim();

    const match = desc.match(/GOES INTO\s+([A-Z0-9-]+)/);
    const isChild = !!match;

    let shouldGray = false;

    if (isChild) {
      const parentPart = match[1];
      const parentInfo = partInfo.get(parentPart) || { vendor: '' };
      const parentVendor = parentInfo.vendor || '';

      // GOES INTO row: gray if parent vendor is NOT TDG or FASTENAL
      shouldGray = !(
        parentVendor.includes('TDG') || parentVendor.includes('FASTENAL')
      );
    } else {
      const info = partInfo.get(tdgpn) || { vendor: '', hasChildren: false };
      const { vendor, hasChildren } = info;

      // Main row: gray only if vendor is TDG or FASTENAL and has children
      shouldGray =
        (vendor.includes('TDG') || vendor.includes('FASTENAL')) && hasChildren;
    }

    // Apply gray fill (vendor-based coloring)
    for (let c = 1; c <= 9; c++) {
      row.getCell(c).fill = shouldGray
        ? {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD3D3D3' }, // light gray color
          }
        : null; // white
    }
  });

  // --------------------------------------------------------------------------------------------------------------------------
  // NEW LOGIC: Gray main part if:
  // - Main part has NO vendor
  // - Has children (GOES INTO)
  // - ALL children have valid Location
  // → Main part = gray, Children = white

  const childToParentMap = new Map(); // childTDGPN → parentTDGPN
  const parentToChildrenMap = new Map(); // parentTDGPN → array of children (rowNumber, hasLocation)
  const tdgpnRowMap = new Map(); // TDGPN → row

  // First pass: Build parent-child relationships and gather the required information
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber < 8) return;

    const tdgpn = (row.getCell(1).value || '').toString().trim().toUpperCase();
    const desc = getCellText(row.getCell(2)).toUpperCase().trim();
    const location = (row.getCell(8).value || '').toString().trim();
    const vendor = (row.getCell(3).value || '').toString().trim().toUpperCase();

    if (!tdgpn) return;

    tdgpnRowMap.set(tdgpn, { row, vendor });

    const match = desc.match(/GOES INTO\s+([A-Z0-9-]+)/);
    if (match) {
      const parent = match[1];
      childToParentMap.set(tdgpn, parent);

      if (!parentToChildrenMap.has(parent)) {
        parentToChildrenMap.set(parent, []);
      }

      parentToChildrenMap.get(parent).push({
        row,
        hasLocation: location !== '',
      });
    }
  });

  // Step 2: Apply new logic: Gray parent if it has no vendor, has children, and all children have locations
  parentToChildrenMap.forEach((children, parentTDGPN) => {
    const parentData = tdgpnRowMap.get(parentTDGPN);

    if (!parentData) return;

    const { row: parentRow, vendor } = parentData;

    // Check if all children have location
    const allChildrenHaveLocation = children.every((c) => c.hasLocation);

    // Apply the new graying logic if needed
    if (!vendor && allChildrenHaveLocation && children.length > 0) {
      // ✅ Gray the parent
      for (let c = 1; c <= 9; c++) {
        parentRow.getCell(c).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD3D3D3' }, // gray
        };
      }

      // ✅ Ensure children are white (override other gray fills if any)
      children.forEach(({ row }) => {
        for (let c = 1; c <= 9; c++) {
          row.getCell(c).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' }, // white
          };
        }
      });
    }
  });
  //--------------------------------------------------------------------------------------------------------------------

  // Step 3: Apply component-based logic (gray and lightTrellis patterns)
  components.forEach((comp, i) => {
    const row = worksheet.getRow(i + 8);

    // Define and calculate totalQty for each component inside the loop
    let totalQty = 0;
    const isConsumable =
      comp.ConsumableOrVMI ||
      (comp.Location &&
        (comp.Location.toUpperCase().includes('CONSUMABLE') ||
          comp.Location.toUpperCase().includes('VMI'))) ||
      (comp.LeadHandComments &&
        (comp.LeadHandComments.toUpperCase().includes('CONSUMABLE') ||
          comp.LeadHandComments.toUpperCase().includes('VMI')));

    const isWire =
      comp.Description && comp.Description.toUpperCase().includes('WIRE');

    // Calculate totalQty based on component data
    if (isConsumable) {
      totalQty = comp.QuantityPerFixture || 0;
    } else if (comp.TDGPN.includes('LABEL') || isWire) {
      totalQty = 0;
    } else {
      totalQty = (comp.QuantityPerFixture || 0) * comp.Quantity;
    }

    const available = comp.QuantityAvailable || 0;

    // Highlight if quantity exceeds available and not consumable
    if (totalQty > available && !isConsumable) {
      row.getCell(7).fill = {
        type: 'pattern',
        pattern: 'lightTrellis',
        fgColor: { argb: 'FFFF0000' }, // red color for over quantity
      };
    }

    // Location-based coloring logic
    const loc = (comp.Location || '').toUpperCase();
    const locationNumberCheck = comp.Location
      ? isValidNumberLocationFormat(comp.Location)
      : false;

    const isGray =
      !comp.Location || // Location is missing or empty
      isConsumable || // Item is consumable
      loc.includes('INHOUSE') || // Location contains "INHOUSE"
      loc.includes('VMI') || // Location contains "VMI"
      (loc.includes('V') && !loc.includes('HV')) || // Location has "V" but not "HV"
      totalQty === 0 || // No quantity
      locationNumberCheck; // Location in wrong number format

    // here check location exist but not number- number -number format
    // not consumable
    // not VMI
    // not V (but not HV)
    // this all condition is false then set row to white
    let shouldBeWhite;
    if (comp.Location) {
      // Check for conditions that should set row to white
      shouldBeWhite =
        comp.Location && // Location exists
        !isConsumable && // Not consumable
        !loc.includes('VMI') && // Not VMI
        !(loc.includes('V') && !loc.includes('HV')) && // Not V (but not HV)
        !locationNumberCheck; // Location not in number-number-number format
    }

    // Apply gray fill for conditions (based on consumable, missing location, etc.)
    if (isGray) {
      for (let c = 1; c <= 9; c++) {
        row.getCell(c).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD3D3D3' }, // light gray color
        };
      }
    }

    if (shouldBeWhite) {
      for (let c = 1; c <= 9; c++) {
        row.getCell(c).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFFFF' }, // white color
        };
      }
    }
  });

  // last row set gray color
  const finalRow = worksheet.getRow(components.length + 8);
  for (let c = 1; c <= 9; c++) {
    finalRow.getCell(c).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFA9A9A9' },
    };
  }

  // Heading font + alignment
  worksheet.getCell('C1').font = { bold: true, size: 18 };
  worksheet.getCell('C1').alignment = { horizontal: 'center', vertical: 'top' };
  worksheet.getCell('C3').font = { size: 18 };
  ['A', 'B'].forEach((col) =>
    [1, 2, 3, 4, 5].forEach(
      (row) =>
        (worksheet.getCell(`${col}${row}`).alignment = {
          horizontal: col === 'A' ? 'left' : 'right',
        }),
    ),
  );
  ['G', 'H'].forEach((col) =>
    [1, 2, 3, 4, 5].forEach(
      (row) =>
        (worksheet.getCell(`${col}${row}`).alignment = {
          horizontal: 'center',
        }),
    ),
  );
  ['C1', 'C2', 'C3', 'C4', 'C5'].forEach(
    (cell) =>
      (worksheet.getCell(cell).alignment = {
        horizontal: 'center',
        vertical: 'top',
      }),
  );

  worksheet.getCell('C3').alignment = {
    horizontal: 'center',
    vertical: 'top',
    wrapText: true,
  };

  worksheet.views = [{ showGridLines: false }];
}

const generatePickLists = async (vmParam, userParam, fixtureParam, res) => {
  try {
    const vm = vmParam || { LHREntries: [0] };
    const user = userParam || null;
    let fixture = fixtureParam || null;

    let sopNum = '-';
    // const ml = await getMasterList();
    const isDieLists = await getAllDieAndLabelList();

    const workbook = new ExcelJS.Workbook();

    // Fetch user role info only once
    const intlUsers = await getUsersInRole('INTL');
    const currentUser = await getUserByUsername(user);
    const isIntlUser =
      currentUser &&
      intlUsers.some((intlUser) => intlUser.Id === currentUser.Id);

    for (const LHREntryId of vm.LHREntries) {
      let tempSOP = null;
      let tempQuantity = 1;

      if (LHREntryId !== 0) {
        const tempLHREntry = await getLeadHandEntry(LHREntryId);
        fixture = fixFixtureName(tempLHREntry.FixtureNumber);

        tempSOP = tempLHREntry.SOPId;
        tempSOP = await getSOPIdtoSopData(tempSOP);

        const ProgramName = await getProgramNameByProgramId(
          tempSOP.SOPProgramId,
        );
        sopNum = tempSOP.SOPNum;
        tempQuantity = tempLHREntry.Quantity;

        // tempSOP = {
        //   Program: { Name: ProgramName.Name },
        //   ODD: new Date(0), // DateTime.MinValue equivalent
        // };

        tempSOP.Program = { Name: ProgramName.Name };
      } else {
        tempSOP = {
          Program: { Name: '' },
          ODD: '0001-01-01T00:00:00.000Z', // DateTime.MinValue equivalent
        };
      }

      const db = await connectDB();
      const tempFixture = await getExplodedBOM(fixture, db);
      const refFixture = await getStoredFixture(fixture, db);

      // Process components in parallel
      const tempComponents = await Promise.all(
        tempFixture.Components.map(async (comp) => {
          const split = comp.Level.split('.');
          const parentLevel = split.slice(0, -1).join('.');
          const tempParent = tempFixture.Components.find(
            (x) => x.Level === parentLevel,
          );
          const parent = tempParent ? tempParent.TDGPN : '';

          const tempComp = {
            Description: `${parent}<line>${comp.Description}`,
            TDGPN: comp.TDGPN,
          };

          tempComp.QuantityPerFixture = Math.ceil(comp.Quantity); // MidpointRounding.ToPositiveInfinity equivalent
          tempComp.QuantityNeeded = tempComp.QuantityPerFixture * tempQuantity;
          tempComp.QuantityPerFixture = Math.ceil(
            refFixture?.Components?.find((x) => x.Level === comp.Level)
              .Quantity,
          );

          const groupMatch = isDieLists.find(
            (x) =>
              x.TDGPN === comp.TDGPN &&
              (x.GroupingName === 'Die' || x.GroupingName === 'Label'),
          );

          if (groupMatch) {
            if (groupMatch.GroupingName === 'Die') {
              tempComp.QuantityPerFixture = 0;
              tempComp.QuantityNeeded = 0;
            } else if (groupMatch.GroupingName === 'Label') {
              tempComp.QuantityNeeded = 0;
            }
          }

          tempComp.Vendor = comp.Vendor;
          tempComp.VendorPN = comp.VendorPN;

          const inventory = await GetInventoryTuple(comp.TDGPN, isIntlUser);
          tempComp.Location = inventory.location;
          tempComp.QuantityAvailable = inventory.quantity;
          tempComp.ConsumableOrVMI = inventory.type;

          return tempComp;
        }),
      );

      await addSOP(
        tempComponents,
        tempFixture.Description,
        sopNum,
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

    // Format: "-(fixture) Jul-24.xlsx"
    const now = new Date();
    const month = now.toLocaleString('en-US', { month: 'short' });
    const day = String(now.getDate()).padStart(2, '0');
    const dateStr = `${month}-${day}`;
    const safeFixture = fixture.replace(/[\\/:*?"<>|]/g, '-'); // Avoid filename issues
    const fileName = sopNum + ' (' + safeFixture + ') ' + dateStr + '.xlsx';

    res.setHeader(
      'Content-Disposition',
      `attachment; filename=\"${fileName}\"`,
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );

    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');
    res.send(buffer);
    // Do not send a JSON response after sending the file buffer
  } catch (err) {
    return res.failureResponse({ message: err.message });
  }
};
function buildPicklistData(components, fixtureDescription, Quantity) {
  let listData = [];

  // --------------------------------------------------------------------------------
  // STEP 1: Build partInfo (TDGPN → { vendor, hasChildren })
  const partInfo = new Map();

  components.forEach((comp) => {
    const tdgpn = (comp.TDGPN || '').toUpperCase();
    const desc = (comp.Description || '').toUpperCase();

    if (!tdgpn) return;

    if (!partInfo.has(tdgpn)) {
      partInfo.set(tdgpn, {
        vendor: (comp.Vendor || '').toUpperCase(),
        hasChildren: false,
      });
    }

    // If this row is a child ("GOES INTO ..."), mark parent
    const match = desc.match(/GOES INTO\s+([A-Z0-9-]+)/);
    if (match) {
      const parentPart = match[1];
      if (partInfo.has(parentPart)) {
        partInfo.get(parentPart).hasChildren = true;
      } else {
        partInfo.set(parentPart, { vendor: '', hasChildren: true });
      }
    }
  });

  // --------------------------------------------------------------------------------
  // STEP 2: Build parent–child mapping for "all children have location" logic
  const childToParentMap = new Map();
  const parentToChildrenMap = new Map();
  const tdgpnMap = new Map(); // TDGPN → { vendor, comp }

  components.forEach((comp) => {
    const tdgpn = (comp.TDGPN || '').toUpperCase();
    const desc = (comp.Description || '').toUpperCase();
    const location = (comp.Location || '').trim();
    const vendor = (comp.Vendor || '').toUpperCase();

    if (!tdgpn) return;
    tdgpnMap.set(tdgpn, { vendor, comp });

    const match = desc.match(/GOES INTO\s+([A-Z0-9-]+)/);
    if (match) {
      const parent = match[1];
      childToParentMap.set(tdgpn, parent);

      if (!parentToChildrenMap.has(parent)) {
        parentToChildrenMap.set(parent, []);
      }

      parentToChildrenMap.get(parent).push({
        tdgpn,
        hasLocation: location !== '',
      });
    }
  });

  // --------------------------------------------------------------------------------
  // STEP 3: Build listData with flags
  components.forEach((comp) => {
    const tdgpn = (comp.TDGPN || '').toUpperCase();
    const desc = (comp.Description || '').toUpperCase();
    const vendor = (comp.Vendor || '').toUpperCase();

    const qtyPerFixture = comp.QuantityPerFixture || 0;
    const totalQty = qtyPerFixture * Quantity;
    const available = comp.QuantityAvailable || 0;

    // Consumable detection
    const isConsumable =
      comp.ConsumableOrVMI ||
      (comp.Location &&
        (comp.Location.toUpperCase().includes('CONSUMABLE') ||
          comp.Location.toUpperCase().includes('VMI'))) ||
      (comp.LeadHandComments &&
        (comp.LeadHandComments.toUpperCase().includes('CONSUMABLE') ||
          comp.LeadHandComments.toUpperCase().includes('VMI')));

    const isWire = desc.includes('WIRE');

    // --- total qty logic
    let finalTotalQty = 0;
    if (isConsumable) {
      finalTotalQty = qtyPerFixture;
    } else if (tdgpn.includes('LABEL') || isWire) {
      finalTotalQty = 0;
    } else {
      finalTotalQty = qtyPerFixture * Quantity;
    }

    // --- Flags
    let isGray = false;
    let isLightTrellis = false;

    // A. GOES INTO child logic
    const match = desc.match(/GOES INTO\s+([A-Z0-9-]+)/);
    const isChild = !!match;

    if (isChild) {
      const parentPart = match[1];
      const parentInfo = partInfo.get(parentPart) || { vendor: '' };
      const parentVendor = parentInfo.vendor || '';

      // GOES INTO row: gray if parent vendor is NOT TDG or FASTENAL
      isGray = !(
        parentVendor.includes('TDG') || parentVendor.includes('FASTENAL')
      );
    } else {
      const info = partInfo.get(tdgpn) || { vendor: '', hasChildren: false };
      const { vendor: mainVendor, hasChildren } = info;

      // Main row: gray only if vendor is TDG or FASTENAL AND has children
      isGray =
        (mainVendor.includes('TDG') || mainVendor.includes('FASTENAL')) &&
        hasChildren;
    }

    // B. New rule: gray parent with no vendor, has children, and all children have location
    if (parentToChildrenMap.has(tdgpn)) {
      const children = parentToChildrenMap.get(tdgpn);
      const allChildrenHaveLocation = children.every((c) => c.hasLocation);
      if (!vendor && allChildrenHaveLocation && children.length > 0) {
        isGray = true;
      }
    }

    // C. Shortage → lightTrellis
    if (finalTotalQty > available && !isConsumable) {
      isLightTrellis = true;
    }

    // D. Location-based gray
    const loc = (comp.Location || '').toUpperCase();
    const locationNumberCheck = comp.Location
      ? isValidNumberLocationFormat(comp.Location)
      : false;

    if (
      !comp.Location ||
      isConsumable ||
      loc.includes('INHOUSE') ||
      loc.includes('VMI') ||
      (loc.includes('V') && !loc.includes('HV')) ||
      finalTotalQty === 0 ||
      locationNumberCheck
    ) {
      isGray = true;
    }

    if (
      comp.Location && // Location exists
      !isConsumable && // Not consumable
      !loc.includes('VMI') && // Not VMI
      !(loc.includes('V') && !loc.includes('HV')) && // Not V (but not HV)
      !locationNumberCheck
    ) {
      isGray = false;
    }

    listData.push({
      TDGPN: comp.TDGPN,
      Description: comp.Description,
      Vendor: comp.Vendor,
      VendorPN: comp.VendorPN,
      QuantityPerFixture: qtyPerFixture,
      TotalQtyNeeded: finalTotalQty,
      ActualQtyPicked: '',
      Location: comp.Location,
      LeadHandComments: comp.LeadHandComments,
      UnitOfMeasure: comp.UnitOfMeasure,
      QuantityAvailable: comp.QuantityAvailable,
      ConsumableOrVMI: comp.ConsumableOrVMI,
      isGray,
      isLightTrellis,
    });
  });

  return listData;
}

const getPickListData = async (lhrEntryId, user, fixtureParam) => {
  try {
    const listData = [];
    let finalData = [];
    let excelFixtureDetail = {};
    const vm = lhrEntryId || { LHREntries: [0] };
    // const user = user || null;
    let fixture = fixtureParam || null;

    let sopNum = '-';
    // const ml = await getMasterList();
    const isDieLists = await getAllDieAndLabelList();

    const workbook = new ExcelJS.Workbook();

    // Fetch user role info only once
    const intlUsers = await getUsersInRole('INTL');
    const currentUser = await getUserByUsername(user || null);
    const isIntlUser =
      currentUser &&
      intlUsers.some((intlUser) => intlUser.Id === currentUser.Id);

    for (const LHREntryId of vm.LHREntries) {
      let tempSOP = null;
      let tempQuantity = 1;

      if (LHREntryId !== 0) {
        const tempLHREntry = await getLeadHandEntry(LHREntryId);
        fixture = fixFixtureName(tempLHREntry.FixtureNumber);

        tempSOP = tempLHREntry.SOPId;
        tempSOP = await getSOPIdtoSopData(tempSOP);

        const ProgramName = await getProgramNameByProgramId(
          tempSOP.SOPProgramId,
        );
        sopNum = tempSOP.SOPNum;
        tempQuantity = tempLHREntry.Quantity;

        tempSOP.Program = { Name: ProgramName.Name };
      } else {
        tempSOP = {
          Program: { Name: '' },
          ODD: '0001-01-01T00:00:00.000Z', // DateTime.MinValue equivalent
        };
      }

      const databases = await connectDB();
      const tempFixture = await getExplodedBOM(fixture, databases);
      const refFixture = await getStoredFixture(fixture, databases);

      // Process components in parallel
      const tempComponents = await Promise.all(
        tempFixture.Components.map(async (comp) => {
          const split = comp.Level.split('.');
          const parentLevel = split.slice(0, -1).join('.');
          const tempParent = tempFixture.Components.find(
            (x) => x.Level === parentLevel,
          );
          const parent = tempParent ? tempParent.TDGPN : '';

          const tempComp = {
            Description: `${parent}<line>${comp.Description}`,
            TDGPN: comp.TDGPN,
          };

          tempComp.QuantityPerFixture = Math.ceil(comp.Quantity); // MidpointRounding.ToPositiveInfinity equivalent
          tempComp.QuantityNeeded = tempComp.QuantityPerFixture * tempQuantity;
          tempComp.QuantityPerFixture = Math.ceil(
            refFixture?.Components?.find((x) => x.Level === comp.Level)
              .Quantity,
          );

          const isDieGroup = isDieLists.find(
            (x) => x.TDGPN === comp.TDGPN && x.GroupingName === 'Die',
          );

          if (isDieGroup) {
            tempComp.QuantityPerFixture = 0;
            tempComp.QuantityNeeded = 0;
          }

          tempComp.Vendor = comp.Vendor;
          tempComp.VendorPN = comp.VendorPN;
          tempComp.UnitOfMeasure = comp.UnitOfMeasure;

          const inventory = await GetInventoryTuple(comp.TDGPN, isIntlUser);
          tempComp.Location = inventory.location;
          tempComp.QuantityAvailable = inventory.quantity;
          tempComp.ConsumableOrVMI = inventory.type;

          return tempComp;
        }),
      );

      tempComponents.forEach((comp, i) => {
        const descParts = (comp.Description || '').split('<line>');
        const parent = descParts[0];
        const isTopLevel =
          parent && parent.trim() !== '' && parent === comp.TDGPN;

        // Create rich text for description with bold "GOES INTO" part
        let richText = '';
        if (parent) {
          richText += `GOES INTO ${parent}\n`;
        }
        if (descParts[1]) {
          richText += descParts[1];
        }

        // Determine if this is a consumable item
        const isConsumable =
          comp.ConsumableOrVMI ||
          (comp.Location &&
            (comp.Location.toUpperCase().includes('CONSUMABLE') ||
              comp.Location.toUpperCase().includes('VMI'))) ||
          (comp.LeadHandComments &&
            (comp.LeadHandComments.toUpperCase().includes('CONSUMABLE') ||
              comp.LeadHandComments.toUpperCase().includes('VMI')));

        let totalQty = 0;

        const isWire =
          comp.Description && comp.Description.toUpperCase().includes('WIRE');

        if (isConsumable) {
          // For consumables, show the per-fixture quantity (no multiplication)
          totalQty = comp.QuantityPerFixture || 0;
        } else if (comp.TDGPN.includes('LABEL')) {
          totalQty = 0;
        } else if (isWire) {
          totalQty = 0;
        } else {
          // For regular items, multiply by fixture quantity
          totalQty = (comp.QuantityPerFixture || 0) * tempQuantity;
        }

        listData.push({
          TDGPN: comp.TDGPN,
          Description: richText || '',
          Vendor: comp.Vendor,
          VendorPN: comp.VendorPN,
          QuantityPerFixture: comp.QuantityPerFixture,
          TotalQtyNeeded: totalQty,
          ActualQtyPicked: '',
          Location: comp.Location,
          LeadHandComments: '',
          UnitOfMeasure: comp.UnitOfMeasure,
          QuantityAvailable: comp.QuantityAvailable,
          ConsumableOrVMI: comp.ConsumableOrVMI,
        });
      });

      finalData = await buildPicklistData(
        listData,
        tempFixture.Description,
        tempQuantity,
      );

      excelFixtureDetail = {
        description: tempFixture.Description,
        sopNum,
        programName: tempSOP.Program.Name,
        fixture,
        tempQuantity,
        odd: tempSOP.ODD,
      };
    }

    return { excelFixtureDetail, listData: finalData };
  } catch (err) {
    console.log('error', err);
    return (listData = []);
  }
};

const getOpenPickLists = async (fixtureNumber) => {
  try {
    let fixtureLists = [];

    const sqlQuery = `
      SELECT *
      FROM [SOP].[dbo].[SOPLeadHandEntries]
      WHERE FixtureNumber = @FixtureNumber
    `;
    const params = { FixtureNumber: fixtureNumber };
    const retriveAllFixturesResult = await query('sop', sqlQuery, params);

    if (retriveAllFixturesResult.length === 0) {
      return fixtureLists;
    }

    const result = await query(
      'sop',
      `
    SELECT 
        sle.*, 
        s.FinalDeliveryDate,
        s.SOPNum,
        s.ODD, 
        spe.ProductionDateOut,
        she.ShippingDateIn
    FROM [SOP].[dbo].[SOPLeadHandEntries] sle
    JOIN [SOP].[dbo].[SOPs] s 
        ON sle.SOPId = s.SOPId
    JOIN [SOP].[dbo].[SOPProductionEntries] spe 
        ON s.SOPProductionEntryId = spe.SOPProductionEntryId
    JOIN [SOP].[dbo].[SOPShippingEntries] she 
        ON s.SOPShippingEntryId = she.SOPShippingEntryId
    WHERE 
        REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(UPPER(sle.FixtureNumber), '-WAR', ''), 'WAR', ''), '-RPR', ''), 'RPR', ''), '-EVAL', ''), 'EVAL', '') = UPPER(@fixture) AND
        UPPER(s.SOPNum) <> 'CANCELLED' AND
        spe.ProductionDateOut = '0001-01-01' AND
        she.ShippingDateIn = '0001-01-01' AND
        s.FinalDeliveryDate = '0001-01-01'
    ORDER BY s.ODD;
  `,
      { fixture: fixtureNumber.toUpperCase() },
    );

    return result;
  } catch (err) {
    console.log('error:', err);
    return [];
  }
};

const updatedSheetDownload = async (excelFixtureDetail, sheetlistData, res) => {
  try {
    const workbookCreate = new ExcelJS.Workbook();

    // Validate and sanitize worksheet name
    let worksheetName = excelFixtureDetail.sopNum;
    if (!worksheetName || worksheetName.trim() === '') {
      worksheetName = 'Sheet1'; // Default fallback name
    } else {
      // Sanitize the name to remove invalid characters
      worksheetName = worksheetName.replace(/[*?:\\/[\]]/g, '_');
      worksheetName = worksheetName.trim();
      if (worksheetName.length > 31) {
        worksheetName = worksheetName.substring(0, 31);
      }
    }

    const worksheet = workbookCreate.addWorksheet(worksheetName, {
      pageSetup: {
        orientation: 'landscape',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
      },
    });

    // Format the current date to "Month Day, Year"
    const formattedPrintedDate = new Date().toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    });

    // Format the required date to "D-MMM-YY" (e.g., 1-Jan-01)
    const formattedRequiredDate = excelFixtureDetail.odd
      ? new Date(excelFixtureDetail.odd)
          .toLocaleDateString('en-GB', {
            day: 'numeric',
            month: 'short',
            year: '2-digit',
          })
          .replace(/ /g, '-')
      : '';

    // Insert header rows (Row 1-5)
    worksheet.insertRows(1, [
      [
        'SOP #',
        excelFixtureDetail.sopNum,
        `PICK LIST #${excelFixtureDetail.sopNum}`,
        '',
        '',
        '',
        'PICK LIST PRINTED ON',
        '',
        formattedPrintedDate,
        '',
        '',
      ],
      [
        'PROJECT',
        excelFixtureDetail.programName || '',
        '',
        '',
        '',
        '',
        'PICK LIST LOG NUMBER',
        '',
        '',
        '',
        '',
      ],
      [
        'FIXTURE',
        excelFixtureDetail.fixture,
        excelFixtureDetail.description,
        '',
        '',
        '',
        'DATE PICKED',
        '',
        '',
        '',
        '',
      ],
      [
        'QUANTITY',
        excelFixtureDetail.tempQuantity,
        '',
        '',
        '',
        '',
        'LEAD HAND SIGN OFF',
        '',
        '',
        '',
        '',
      ],
      [
        'REQUIRED ON',
        formattedRequiredDate,
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
      ],
    ]);

    // Merges
    [
      'C1:F1',
      'C2:F2',
      'G1:H1',
      'G2:H2',
      'C3:F5',
      'G3:H3',
      'G4:H5',
      'I3:J3',
      'I2:J2',
      'I1:J1',
      'I4:J5',
    ].forEach((range) => worksheet.mergeCells(range));

    // Column widths
    [
      22.28, // TDG PART NO
      69, // DESCRIPTION
      22.42, // VENDOR
      27, // VENDOR P/N
      11.57, // PER FIX QTY.
      14, // ACTUAL QTY TO BE PICKED
      20.42, // TOTAL QTY NEEDED
      21.57, // UNIT OF MEASURE
      25, // LOCATION / PURCHASING COMMENTS (increased)
      31.57, // LEAD HAND COMMENTS (also increased)
    ].forEach((w, i) => (worksheet.getColumn(i + 1).width = w));

    // Row 7 headers
    const headers = [
      'TDG PART NO',
      'DESCRIPTION',
      'VENDOR',
      'VENDOR P/N',
      'PER FIX QTY.',
      'ACTUAL QTY TO BE PICKED',
      'TOTAL QTY NEEDED',
      'UNIT OF MEASURE',
      'LOCATION/ PURCHASING COMMENTS',
      'LEAD HAND COMMENTS',
    ];
    worksheet.getRow(7).values = headers;
    worksheet.getRow(7).height = undefined;
    worksheet.getRow(7).eachCell((cell) => {
      cell.font = { bold: true, size: 18, name: 'Calibri' };
      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true,
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'D9E1F2' },
      };
    });

    // Apply borders and fonts to header rows
    for (let r = 1; r <= 5; r++) {
      for (let c = 1; c <= 9; c++) {
        const cell = worksheet.getRow(r).getCell(c);
        cell.border = {
          top: { style: 'medium' },
          bottom: { style: 'medium' },
          left: { style: 'medium' },
          right: { style: 'medium' },
        };
      }
    }

    // Apply bold borders to second table (starting from Row 7)
    for (let row = 7; row <= worksheet.rowCount; row++) {
      for (let col = 1; col <= 10; col++) {
        const cell = worksheet.getRow(row).getCell(col);
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        };
      }
    }

    // Font + fill for header info cells
    const headerFontCells = ['A', 'B', 'G'].flatMap((col) =>
      [1, 2, 3, 4].map((row) => `${col}${row}`),
    );
    headerFontCells
      .concat(['A5', 'B5', 'I1', 'I2', 'I3', 'I4', 'I5'])
      .forEach((cell) => {
        const c = worksheet.getCell(cell);
        c.font = { bold: true, size: 18, name: 'Calibri' };
        c.alignment = { horizontal: 'center', vertical: 'middle' };
        if (
          cell.startsWith('A') ||
          cell.startsWith('B') ||
          cell.startsWith('G')
        ) {
          c.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'D9E1F2' },
          };
        }
      });

    for (let r = 1; r <= 5; r++) {
      worksheet.getCell(`I${r}`).alignment = {
        horizontal: 'center',
        vertical: 'middle',
      };
    }
    // worksheet.getColumn(9).width = 20;

    // Data rows
    sheetlistData.forEach((comp, i) => {
      const row = worksheet.getRow(i + 8);
      const descParts = (comp.Description || '').includes('GOES INTO')
        ? true
        : false;
      const parent = descParts;
      // const isTopLevel = parent && parent.trim() !== "" && parent === comp.TDGPN;

      // Create rich text for description with bold "GOES INTO" part
      let richText = [];
      if (parent) {
        richText.push({
          text: comp.Description,
          font: { bold: true, size: 18, name: 'Calibri' },
        });
      } else {
        richText.push({
          text: comp.Description,
          font: { bold: false, size: 18, name: 'Calibri' },
        });
      }

      // Determine if this is a consumable item
      const isConsumable =
        comp.ConsumableOrVMI ||
        (comp.Location &&
          (comp.Location.toUpperCase().includes('CONSUMABLE') ||
            comp.Location.toUpperCase().includes('VMI'))) ||
        (comp.LeadHandComments &&
          (comp.LeadHandComments.toUpperCase().includes('CONSUMABLE') ||
            comp.LeadHandComments.toUpperCase().includes('VMI')));

      let totalQty = 0;

      const isWire =
        comp.Description && comp.Description.toUpperCase().includes('WIRE');

      if (isConsumable) {
        // For consumables, show the per-fixture quantity (no multiplication)
        totalQty = comp.QuantityPerFixture || 0;
      } else if (comp.TDGPN.includes('LABEL')) {
        totalQty = 0;
      } else if (isWire) {
        totalQty = 0;
      } else {
        // For regular items, multiply by fixture quantity
        totalQty =
          (comp.QuantityPerFixture || 0) * excelFixtureDetail.tempQuantity;
      }

      row.values = [
        comp.TDGPN,
        richText.length > 0 ? { richText } : comp.Description || '',
        comp.Vendor,
        comp.VendorPN,
        comp.QuantityPerFixture,
        comp.ActualQtyPicked || '',
        totalQty,
        comp.UnitOfMeasure,
        comp.Location,
        comp.LeadHandComments,
      ];

      // create table border for description.....
      for (let c = 1; c <= 10; c++) {
        const cell = row.getCell(c);
        // Apply normal font for B, C, D; bold for others
        const isNormalFont = c === 2 || c === 3 || c === 4;
        cell.font = { size: 18, name: 'Calibri', bold: !isNormalFont };

        cell.alignment = {
          wrapText: true,
          vertical: 'middle',
          horizontal: c === 2 ? 'left' : 'center', // only Description left, everything else centered
        };
        cell.border = {
          top: { style: 'thin' },
          bottom: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
        };
      }
      //-------------------------------------------------------------------------------------------------------------------

      // const qty = totalQty;
      // const available = comp.QuantityAvailable || 0;
      // if (qty > available && !isConsumable) {
      //   row.getCell(7).fill = {
      //     type: "pattern",
      //     pattern: "lightTrellis",
      //     fgColor: { argb: "FFFF0000" },
      //   };
      // }

      // const loc = (comp.Location || "").toUpperCase();
      // const isGray =
      //   isConsumable ||
      //   loc.includes("INHOUSE") ||
      //   (loc.includes("V") && !loc.includes("HV")) ||
      //   qty === 0;
      // if (isGray) {
      //   for (let c = 1; c <= 10; c++) {
      //     row.getCell(c).fill = {
      //       type: "pattern",
      //       pattern: "solid",
      //       fgColor: { argb: "FFD3D3D3" },
      //     };
      //   }
      // }
    });

    // Step 2: Apply color logic after the data is populated (workbook loop)

    const partInfo = new Map(); // partNo → { vendor, hasChildren }

    // Apply vendor-based coloring logic
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber < 8) return;

      const tdgpn = (row.getCell(1).value || '')
        .toString()
        .trim()
        .toUpperCase();
      const desc = getCellText(row.getCell(2)).toUpperCase().trim();
      const vendor = (row.getCell(3).value || '')
        .toString()
        .trim()
        .toUpperCase();

      if (!tdgpn) return;

      // Save vendor for each part
      if (!partInfo.has(tdgpn)) {
        partInfo.set(tdgpn, { vendor, hasChildren: false });
      }

      // If this row is a child ("GOES INTO ..."), mark parent as having children
      const match = desc.match(/GOES INTO\s+([A-Z0-9-]+)/);
      if (match) {
        const parentPart = match[1];
        if (partInfo.has(parentPart)) {
          partInfo.get(parentPart).hasChildren = true;
        } else {
          partInfo.set(parentPart, { vendor: '', hasChildren: true });
        }
      }
    });

    // Apply vendor-based coloring for main and child rows
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber < 8) return;

      const tdgpn = (row.getCell(1).value || '')
        .toString()
        .trim()
        .toUpperCase();
      const desc = getCellText(row.getCell(2)).toUpperCase().trim();

      const match = desc.match(/GOES INTO\s+([A-Z0-9-]+)/);
      const isChild = !!match;

      let shouldGray = false;

      if (isChild) {
        const parentPart = match[1];
        const parentInfo = partInfo.get(parentPart) || { vendor: '' };
        const parentVendor = parentInfo.vendor || '';

        // GOES INTO row: gray if parent vendor is NOT TDG or FASTENAL
        shouldGray = !(
          parentVendor.includes('TDG') || parentVendor.includes('FASTENAL')
        );
      } else {
        const info = partInfo.get(tdgpn) || { vendor: '', hasChildren: false };
        const { vendor, hasChildren } = info;

        // Main row: gray only if vendor is TDG or FASTENAL and has children
        shouldGray =
          (vendor.includes('TDG') || vendor.includes('FASTENAL')) &&
          hasChildren;
      }

      // Apply gray fill (vendor-based coloring)
      for (let c = 1; c <= 9; c++) {
        row.getCell(c).fill = shouldGray
          ? {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFD3D3D3' }, // light gray color
            }
          : null; // white
      }
    });
    // NEW LOGIC: Gray main part if:
    // - Main part has NO vendor
    // - Has children (GOES INTO)
    // - ALL children have valid Location
    // → Main part = gray, Children = white

    const childToParentMap = new Map(); // childTDGPN → parentTDGPN
    const parentToChildrenMap = new Map(); // parentTDGPN → array of children (rowNumber, hasLocation)
    const tdgpnRowMap = new Map(); // TDGPN → row

    // First pass: Build parent-child relationships and gather the required information
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber < 8) return;

      const tdgpn = (row.getCell(1).value || '')
        .toString()
        .trim()
        .toUpperCase();
      const desc = getCellText(row.getCell(2)).toUpperCase().trim();
      const location = (row.getCell(8).value || '').toString().trim();
      const vendor = (row.getCell(3).value || '')
        .toString()
        .trim()
        .toUpperCase();

      if (!tdgpn) return;

      tdgpnRowMap.set(tdgpn, { row, vendor });

      const match = desc.match(/GOES INTO\s+([A-Z0-9-]+)/);
      if (match) {
        const parent = match[1];
        childToParentMap.set(tdgpn, parent);

        if (!parentToChildrenMap.has(parent)) {
          parentToChildrenMap.set(parent, []);
        }

        parentToChildrenMap.get(parent).push({
          row,
          hasLocation: location !== '',
        });
      }
    });

    // Step 2: Apply new logic: Gray parent if it has no vendor, has children, and all children have locations
    parentToChildrenMap.forEach((children, parentTDGPN) => {
      const parentData = tdgpnRowMap.get(parentTDGPN);

      if (!parentData) return;

      const { row: parentRow, vendor } = parentData;

      // Check if all children have location
      const allChildrenHaveLocation = children.every((c) => c.hasLocation);

      // Apply the new graying logic if needed
      if (!vendor && allChildrenHaveLocation && children.length > 0) {
        // ✅ Gray the parent
        for (let c = 1; c <= 9; c++) {
          parentRow.getCell(c).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD3D3D3' }, // gray
          };
        }

        // ✅ Ensure children are white (override other gray fills if any)
        children.forEach(({ row }) => {
          for (let c = 1; c <= 9; c++) {
            row.getCell(c).fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFFFFF' }, // white
            };
          }
        });
      }
    });

    // Data rows
    sheetlistData.forEach((comp, i) => {
      const row = worksheet.getRow(i + 8);

      // Determine if this is a consumable item
      const isConsumable =
        comp.ConsumableOrVMI ||
        (comp.Location &&
          (comp.Location.toUpperCase().includes('CONSUMABLE') ||
            comp.Location.toUpperCase().includes('VMI'))) ||
        (comp.LeadHandComments &&
          (comp.LeadHandComments.toUpperCase().includes('CONSUMABLE') ||
            comp.LeadHandComments.toUpperCase().includes('VMI')));

      let totalQty = 0;

      const isWire =
        comp.Description && comp.Description.toUpperCase().includes('WIRE');

      if (isConsumable) {
        // For consumables, show the per-fixture quantity (no multiplication)
        totalQty = comp.QuantityPerFixture || 0;
      } else if (comp.TDGPN.includes('LABEL')) {
        totalQty = 0;
      } else if (isWire) {
        totalQty = 0;
      } else {
        // For regular items, multiply by fixture quantity
        totalQty =
          (comp.QuantityPerFixture || 0) * excelFixtureDetail.tempQuantity;
      }

      const available = comp.QuantityAvailable || 0;

      // Highlight if quantity exceeds available and not consumable
      if (totalQty > available && !isConsumable) {
        row.getCell(6).fill = {
          type: 'pattern',
          pattern: 'lightTrellis',
          fgColor: { argb: 'FFFF0000' }, // red color for over quantity
        };
      }

      // Location-based coloring logic
      const loc = (comp.Location || '').toUpperCase();
      const locationNumberCheck = comp.Location
        ? isValidNumberLocationFormat(comp.Location)
        : false;

      const isGray =
        !comp.Location || // Location is missing or empty
        isConsumable || // Item is consumable
        loc.includes('INHOUSE') || // Location contains "INHOUSE"
        loc.includes('VMI') || // Location contains "VMI"
        (loc.includes('V') && !loc.includes('HV')) || // Location has "V" but not "HV"
        totalQty === 0 || // No quantity
        locationNumberCheck; // Location in wrong number format

      // here check location exist but not number- number -number format
      // not consumable
      // not VMI
      // not V (but not HV)
      // this all condition is false then set row to white
      let shouldBeWhite;
      if (comp.Location) {
        // Check for conditions that should set row to white
        shouldBeWhite =
          comp.Location && // Location exists
          !isConsumable && // Not consumable
          !loc.includes('VMI') && // Not VMI
          !(loc.includes('V') && !loc.includes('HV')) && // Not V (but not HV)
          !locationNumberCheck; // Location not in number-number-number format
      }

      // Apply gray fill for conditions (based on consumable, missing location, etc.)
      if (isGray) {
        for (let c = 1; c <= 9; c++) {
          row.getCell(c).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD3D3D3' }, // light gray color
          };
        }
      }

      if (shouldBeWhite) {
        for (let c = 1; c <= 9; c++) {
          row.getCell(c).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFFFF' }, // white color
          };
        }
      }
    });
    //-------------------------------------------------------------------------------------------------------------------

    // last row set gray color
    const finalRow = worksheet.getRow(sheetlistData.length + 8);
    for (let c = 1; c <= 9; c++) {
      finalRow.getCell(c).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFA9A9A9' },
      };
    }

    // Heading font + alignment
    worksheet.getCell('C1').font = { bold: true, size: 18 };
    worksheet.getCell('C1').alignment = {
      horizontal: 'center',
      vertical: 'top',
    };
    worksheet.getCell('C3').font = { size: 18 };
    ['A', 'B'].forEach((col) =>
      [1, 2, 3, 4, 5].forEach(
        (row) =>
          (worksheet.getCell(`${col}${row}`).alignment = {
            horizontal: col === 'A' ? 'left' : 'right',
          }),
      ),
    );
    ['G', 'H'].forEach((col) =>
      [1, 2, 3, 4, 5].forEach(
        (row) =>
          (worksheet.getCell(`${col}${row}`).alignment = {
            horizontal: 'center',
          }),
      ),
    );
    ['C1', 'C2', 'C3', 'C4', 'C5'].forEach(
      (cell) =>
        (worksheet.getCell(cell).alignment = {
          horizontal: 'center',
          vertical: 'top',
        }),
    );

    worksheet.getCell('C3').alignment = {
      horizontal: 'center',
      vertical: 'top',
      wrapText: true,
    };

    worksheet.views = [{ showGridLines: false }];

    const buffer = await workbookCreate.xlsx.writeBuffer();

    // Format: "-(fixture) Jul-24.xlsx"
    const now = new Date();
    const month = now.toLocaleString('en-US', { month: 'short' });
    const day = String(now.getDate()).padStart(2, '0');
    const dateStr = `${month}-${day}`;
    const safeFixture = excelFixtureDetail.fixture.replace(
      /[\\/:*?"<>|]/g,
      '-',
    ); // Avoid filename issues
    const fileName =
      excelFixtureDetail.sopNum + ' (' + safeFixture + ') ' + dateStr + '.xlsx';

    res.setHeader(
      'Content-Disposition',
      `attachment; filename=\"${fileName}\"`,
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );

    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');
    res.send(buffer);
  } catch (error) {
    console.log('error', error);
    res.failureResponse({ message: error.message });
  }
};

exports.SOPSerchService = async (req, res) => {
  try {
    if (!req.query.SOPNumber) {
      return res.badRequest({
        message: 'SOP Number is required',
      });
    }

    const { SOPNumber } = req.query;

    const sqlQuery = `
      SELECT *
      FROM SOPs
      WHERE SOPNum = @SOPNumber
    `;

    const params = { SOPNumber };

    let sopNumberResult = await query('sop', sqlQuery, params);

    if (sopNumberResult.length === 0) {
      return res.badRequest({ message: 'SOP not found' });
    }

    sopNumberResult = sopNumberResult[0];

    const LeadHandEntryResult = await query(
      'sop',
      `
        SELECT *
        FROM SOPLeadHandEntries
        WHERE SOPId = @SOPId
      `,
      { SOPId: sopNumberResult.SOPId },
    );

    const LeadHandEntryResults = LeadHandEntryResult;
    let backorderEntryResult = [];

    if (LeadHandEntryResults.length) {
      backorderEntryResult = await Promise.all(
        LeadHandEntryResults.map(async (e) => {
          const backorderResults = await query(
            'sop',
            `
            SELECT *
            FROM SOPBackorderEntries
            WHERE SOPLeadHandEntryId = @SOPLeadHandEntryId
          `,
            { SOPLeadHandEntryId: e.SOPLeadHandEntryId },
          );
          return { ...e, backorderEntry: backorderResults || [] };
        }),
      );
    }

    const customerResult = await query(
      'sop',
      `
        SELECT *
        FROM SOPCustomers
        WHERE SOPCustomerId = @SOPCustomerId
      `,
      { SOPCustomerId: sopNumberResult.SOPCustomerId },
    );

    const programResult = await query(
      'sop',
      `
        SELECT *
        FROM SOPPrograms
        WHERE SOPProgramId = @SOPProgramId
      `,
      { SOPProgramId: sopNumberResult.SOPProgramId },
    );

    const locationResult = await query(
      'sop',
      `
        SELECT *
        FROM SOPLocations
        WHERE SOPLocationId = @SOPLocationId
      `,
      { SOPLocationId: sopNumberResult.SOPLocationId },
    );

    const sopProductionManagerResult = await query(
      'sop',
      `
        SELECT *
        FROM SOPProductionManagers
        WHERE SOPProductionManagerId = @SOPProductionManagerId
      `,
      { SOPProductionManagerId: sopNumberResult.SOPProductionManagerId },
    );

    const productionEntryResult = await query(
      'sop',
      `
        SELECT *
        FROM SOPProductionEntries 
        WHERE SOPProductionEntryId = @SOPProductionEntryId
      `,
      { SOPProductionEntryId: sopNumberResult.SOPProductionEntryId },
    );

    const sopLeadHandId = productionEntryResult[0].SOPLeadHandId;

    const leadHandResult = await query(
      'sop',
      `
        SELECT *
        FROM SOPLeadHands 
        WHERE SOPLeadHandId = @SOPLeadHandId
      `,
      { SOPLeadHandId: sopLeadHandId },
    );

    const qaEntryResult = await query(
      'sop',
      `
        SELECT *
        FROM SOPQAEntries
        WHERE SOPQAEntryId = @SOPQAEntryId
      `,
      { SOPQAEntryId: sopNumberResult.SOPQAEntryId },
    );

    const shippingEntryResult = await query(
      'sop',
      `
        SELECT *
        FROM SOPShippingEntries
        WHERE SOPShippingEntryId = @SOPShippingEntryId
      `,
      { SOPShippingEntryId: sopNumberResult.SOPShippingEntryId },
    );

    const fixtureResults = await query(
      'sop',
      `
        SELECT *
        FROM SOPLeadHandEntries
        WHERE SOPId = @SOPId
      `,
      { SOPId: sopNumberResult.SOPId },
    );

    const db = await connectDB();
    const collection = db.BOMs.collection('Fixture');

    // Process all fixtures
    const fixturesDetailed = await Promise.all(
      fixtureResults.map(async (fixture) => {
        // Get assembler for this fixture
        let assemblerIdResult = [];
        if (fixture.SOPAssemblerId) {
          assemblerIdResult = await query(
            'sop',
            `
              SELECT *
              FROM SOPAssemblers
              WHERE SOPAssemblerId = @SOPAssemblerId
            `,
            { SOPAssemblerId: fixture.SOPAssemblerId },
          );
        }

        const fixtureName = fixFixtureName(fixture.FixtureNumber);

        // Get MongoDB data for this fixture
        const fixtureMongoData = await collection
          .find({ Name: fixtureName })
          .limit(10)
          .toArray();

        return {
          ...fixture,
          assembler: assemblerIdResult,
          fixtureMongoData,
        };
      }),
    );

    const responseData = {
      ...sopNumberResult,
      customer: customerResult,
      program: programResult,
      location: locationResult,
      sopProductionManager: sopProductionManagerResult,
      productionEntry: productionEntryResult,
      leadHandEntry: backorderEntryResult, // leadhand entry and backorder entry both in one array
      leadHand: leadHandResult,
      qaEntry: qaEntryResult,
      shippingEntry: shippingEntryResult,
      fixtures: fixturesDetailed, // <-- now an array of detailed fixture info
    };

    return res.ok({
      message: 'Successfully fetched SOP data',
      data: responseData,
    });
  } catch (err) {
    console.log('error:', err);
    return res.failureResponse({ message: err.message });
  }
};

exports.fixtureDetails = async (req, res) => {
  try {
    const { fixtureNumber } = req.query;

    const fixedFixtureNumber = fixFixtureName(fixtureNumber);

    const db = await connectDB();
    const collection = db.BOMs.collection('Fixture');
    // Get MongoDB data for this fixture
    const fixtureMongoData = await collection
      .find({ Name: fixedFixtureNumber })
      .limit(10)
      .toArray();

    if (fixtureMongoData.length === 0) {
      return res.badRequest({ message: 'Fixture not found' });
    }

    const openPickLists = await getOpenPickLists(fixedFixtureNumber);

    return res.ok({
      message: 'Successfully fetched fixture details',
      data: openPickLists,
    });
  } catch (err) {
    console.log('error:', err);
    return res.failureResponse({ message: err.message });
  }
};

exports.downloadPickList = async (req, res) => {
  try {
    const { fixture, lhrEntryId, user } = req.query;

    if (lhrEntryId) {
      // Call for existing pick list using LeadHandEntryId
      await generatePickLists(
        { LHREntries: [parseInt(lhrEntryId)] },
        user,
        null,
        res,
      );
    } else if (fixture) {
      const fixedFixture = fixFixtureName(fixture);
      // Call for blank pick list using fixture number
      await generatePickLists(null, user, fixedFixture, res);
    } else {
      return res.status(400).json({ message: 'Missing required parameters' });
    }
  } catch (err) {
    console.log('error:', err);
    return res.failureResponse({ message: err.message });
  }
};

exports.getSheetsBomData = async (req, res) => {
  try {
    const { lhrEntryId, user, fixture } = req.query;

    if (!user) {
      return res.badRequest({ message: 'User is required' });
    }

    let sheetBOMsData = [];

    if (lhrEntryId) {
      sheetBOMsData = await getPickListData(
        { LHREntries: [parseInt(lhrEntryId)] },
        user,
      );
    } else {
      const fixedFixture = fixFixtureName(fixture);
      // Call for blank pick list using fixture number
      sheetBOMsData = await getPickListData(null, user, fixedFixture);
    }

    return res.ok({
      message: 'Successfully fetched sheet BOMs data',
      data: sheetBOMsData || [],
    });
  } catch (err) {
    console.log('error:', err);
    return res.failureResponse({ message: err.message });
  }
};

exports.downloadupdatedDataSheets = async (req, res) => {
  try {
    const { sheetData, excelFixtureDetail } = req.body;

    if (!sheetData || sheetData.length === 0) {
      return res.badRequest({ message: 'Sheet Data is required' });
    }

    if (!excelFixtureDetail || Object.keys(excelFixtureDetail).length === 0) {
      return res.badRequest({ message: 'Excel Fixture Detail is required' });
    }

    await updatedSheetDownload(excelFixtureDetail, sheetData, res);
  } catch (err) {
    console.log('error:', err);
    return res.failureResponse({ message: err.message });
  }
};
