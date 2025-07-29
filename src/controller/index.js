const sql = require("mssql");
const ExcelJS = require("exceljs");
const path = require("path");
const getDbPool = require("../db/mssqlPool");
const connectDB = require("../db/conn");
const axios = require("axios");
const os = require("os");
const fs = require("fs");
const fsp = require("fs/promises");

exports.testing = async (req, res) => {
  try {
    const pool = await getDbPool("Purchasing");
    const result = await pool
      .request()
      .query(`SELECT * FROM [dbo].[PurchasingOrders]`);
    res.json({ connected: true, result: result.recordset });

    const db = await connectDB("BOMs");
    const collection = db.collection("ExcelFixture");

    // Demo: Get all documents
    const data = await collection.find({}).limit(10).toArray();

    res.json({ success: true, data });
  } catch (err) {
    console.log("err", err);
    res.status(500).json({ connected: false, error: err.message });
  }
};

const getMasterList = async () => {
  try {
    const pool = await getDbPool("Design");

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
    console.error("❌ Error fetching MasterList:", error);
  }
};

const getLeadHandEntry = async (SOPLeadHandEntryId) => {
  try {
    const pool = await getDbPool("SOP");

    const result = await pool
      .request()
      .input("SOPLeadHandEntryId", sql.Int, SOPLeadHandEntryId).query(`
        SELECT TOP 1 *
        FROM [SOP].[dbo].[SOPLeadHandEntries]
        WHERE SOPLeadHandEntryId = @SOPLeadHandEntryId
      `);

    if (result.recordset.length === 0) {
      return null;
    }

    return result.recordset[0];
  } catch (error) {
    console.error("❌ Error fetching LeadHandEntry:", error);
    return null;
  }
};

const getSOPIdtoSopData = async (SOPId) => {
  try {
    const pool = await getDbPool("SOP");

    const result = await pool.request().input("SOPId", sql.Int, SOPId).query(`
        SELECT TOP 1 *
        FROM [SOP].[dbo].[SOPs]
        WHERE SOPId = @SOPId
      `);

    if (result.recordset.length === 0) {
      return null;
    }

    return result.recordset[0];
  } catch (error) {
    console.error("❌ Error fetching SOPIdtoSopData:", error);
    return null;
  }
};

const getProgramNameByProgramId = async (SOPProgramId) => {
  try {
    const pool = await getDbPool("SOP");

    const result = await pool
      .request()
      .input("SOPProgramId", sql.Int, SOPProgramId).query(`
        SELECT TOP 1 *
        FROM [SOP].[dbo].[SOPPrograms]
        WHERE SOPProgramId = @SOPProgramId
      `);

    if (result.recordset.length === 0) {
      return null;
    }

    return result.recordset[0];
  } catch (error) {
    console.error("❌ Error fetching SOPIdtoSopData:", error);
    return null;
  }
};

const fixFixtureName = (fixture) => {
  if (!fixture) {
    return "";
  }

  let fixtureString = fixture.toUpperCase();
  fixtureString = fixtureString.replace(/-?WAR/g, "");
  fixtureString = fixtureString.replace(/-?RPR/g, "");
  fixtureString = fixtureString.replace(/-?EVAL/g, "");

  return fixtureString;
};

const getExplodedBOM = async (fixtureName, db) => {
  const Fixtures = db.collection("Fixture");
  const PDMSubAssemblies = db.collection("PDMSubAssembly");

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
  const parentSplit = parent.Level.split(".");
  const parentSplitCount = parentSplit.length;

  for (const item of componentPool) {
    const split = item.Level.split(".");
    const splitCount = split.length;

    if (splitCount > 1 && parentSplitCount < splitCount) {
      if (item.Level.startsWith(parent.Level + ".")) {
        children.push(item);
      }
    }
  }

  return children;
};

const getForceMakeBool = async (TDGPN) => {
  if (!TDGPN) return false;

  try {
    const pool = await getDbPool("Purchasing");

    const result = await pool
      .request()
      .input("TDGPN", sql.VarChar, TDGPN.toLowerCase()).query(`
        SELECT 1 
        FROM [Purchasing].[dbo].[MakePartNumbers]
        WHERE LOWER(TDGPN) = @TDGPN
      `);

    return result.recordset.length > 0;
  } catch (error) {
    console.error("❌ Error in getForceMakeBool:", error);
    return false;
  }
};

const getForceBuyBool = async (TDGPN) => {
  if (!TDGPN) return false;

  try {
    const pool = await getDbPool("Purchasing");

    const result = await pool
      .request()
      .input("TDGPN", sql.VarChar, TDGPN.toLowerCase()).query(`
        SELECT 1
        FROM [Purchasing].[dbo].[BuyPartNumbers]
        WHERE LOWER(TDGPN) = @TDGPN
      `);

    return result.recordset.length > 0;
  } catch (error) {
    console.error("❌ Error in getForceBuyBool:", error);
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

    if (component.Type === "P") {
      const group = component.Group;

      if (component.Quantity !== 0) {
        const potentialFamily = item.Components.filter(
          (x) =>
            x.Level.includes(component.Level) &&
            x.Level.split(".").length >= component.Level.split(".").length
        );

        const potentialChildren = getChildren(component, item.Components);
        const tempComponent = component;

        if (potentialFamily.length > 1 && potentialChildren.length >= 1) {
          if (group === "MetalPart" || group === "PCB") {
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
          } else if (group === "PlasticPart") {
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
    } else if (component.Type === "S" && component.Group !== "PCB") {
      const buy = getBuyBool(component.PathName, fixturePool);
      const levelString = component.Level + ".";
      const potentialChildren = getChildren(component, item.Components);

      if (buy) {
        const potentialFamily = item.Components.filter(
          (x) =>
            x.Level.includes(component.Level) &&
            x.Level.split(".").length >= component.Level.split(".").length
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

  const collection = db.collection("Fixture"); // Adjust name if needed

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
  if (!TDGPN || TDGPN.trim() === "") {
    return [];
  }

  try {
    const response = await axios.get(
      `http://192.168.2.175:62625/api/inventory/getlocations`,
      {
        params: { tdgpn: TDGPN },
      }
    );
    return response.data;
  } catch (err) {
    console.error("Error fetching inventory locations:", err.message);
    return [];
  }
};

/**
 * @param {string} TDGPN
 * @returns {Promise<InventoryEntry[]>}
 */
const GetINTLInventoryLocations = async (TDGPN) => {
  if (!TDGPN || TDGPN.trim() === "") {
    return [];
  }

  try {
    const response = await axios.get(
      `http://192.168.2.175:62625/api/inventory/getintllocations`,
      {
        params: { tdgpn: TDGPN },
      }
    );
    return response.data;
  } catch (err) {
    console.error("Error fetching INTL inventory locations:", err.message);
    return [];
  }
};

/**
 * @param {string} TDGPN
 * @param {boolean} INTL
 * @returns {Promise<{ location: string, type: boolean, quantity: number }>}
 */
const GetInventoryTuple = async (TDGPN, INTL = false) => {
  let returnString = "";
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

    if (returnString === "") {
      if (ConsumableType === "CONSUMABLE") {
        returnString += "CONSUMABLE" + newline + Location;
        applyConsumableOrVMI = true;
      } else if (ConsumableType === "INHOUSE") {
        returnString += "INHOUSE" + newline + Location;
        applyConsumableOrVMI = true;
      } else if (ConsumableType === "VMI") {
        returnString += Location;
        applyConsumableOrVMI = true;
      } else {
        returnQuantity += Quantity;
        returnString += `${Location} (${Math.floor(Quantity)})`;
      }
    } else {
      if (Location && Location !== "") {
        if (ConsumableType === "VMI") {
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
    const pool = await getDbPool("OVERVIEW");

    const result = await pool
      .request()
      .input("roleName", sql.NVarChar, roleName).query(`
      SELECT u.*
    FROM [OVERVIEW].[dbo].[AspNetUsers] AS u
    INNER JOIN [OVERVIEW].[dbo].[AspNetUserRoles] AS ur ON u.Id = ur.UserId
    INNER JOIN [OVERVIEW].[dbo].[AspNetRoles] AS r ON ur.RoleId = r.Id
    WHERE r.[Name] = @roleName
    `);

    return result.recordset;
  } catch (error) {
    console.error("❌ Error in getUsersInRole:", error);
    return false;
  }
};

const getUserByUsername = async (username) => {
  try {
    const pool = await getDbPool("OVERVIEW");
    const result = await pool
      .request()
      .input("username", sql.NVarChar, username).query(`
        SELECT * FROM [OVERVIEW].[dbo].[AspNetUsers]
        WHERE [UserName] = @username
      `);

    return result.recordset[0] || null; // return user or null if not found
  } catch (err) {
    console.error("Error fetching user by username:", err);
    throw err;
  }
};

// async function addSOP(
//   components,
//   fixtureDescription,
//   SOP,
//   workbook,
//   Project,
//   Fixture,
//   Quantity,
//   RequiredDate
// ) {
//   const worksheet = workbook.addWorksheet(SOP, {
//     pageSetup: {
//       orientation: "landscape",
//       fitToPage: true,
//       fitToWidth: 1,
//       fitToHeight: 0,
//     },
//   });

//   // Format the current date to "Month Day, Year"
//   const formattedPrintedDate = new Date().toLocaleDateString("en-US", {
//     year: "numeric",
//     month: "long",
//     day: "numeric",
//   });

//   // Format the required date to "D-MMM-YY" (e.g., 1-Jan-01)
//   const formattedRequiredDate = RequiredDate
//     ? new Date(RequiredDate)
//         .toLocaleDateString("en-GB", {
//           day: "numeric",
//           month: "short",
//           year: "2-digit",
//         })
//         .replace(/ /g, "-")
//     : "";

//   // Insert header rows (Row 1-5)
//   worksheet.insertRows(1, [
//     [
//       "SOP #",
//       SOP,
//       `PICK LIST #${SOP}`,
//       "",
//       "",
//       "",
//       "PICK LIST PRINTED ON",
//       "",
//       formattedPrintedDate,
//     ],
//     ["PROJECT", Project, "", "", "", "", "PICK LIST LOG NUMBER", "", ""],
//     ["FIXTURE", Fixture, fixtureDescription, "", "", "", "DATE PICKED", "", ""],
//     ["QUANTITY", Quantity, "", "", "", "", "LEAD HAND SIGN OFF", "", ""],
//     ["REQUIRED ON", formattedRequiredDate, "", "", "", "", "", "", ""],
//   ]);

//   // Merge cells (ExcelJS uses "A1:B1" format)
//   ["C1:F1", "G1:H1", "G2:H2", "C3:F5", "G3:H3", "G4:H5", "I4:I5"].forEach(range => worksheet.mergeCells(range));

//   // Column widths
//   const colWidths = [22.28, 69, 22.42, 27, 11.57, 13.28, 20.42, 21.57, 31.57];
//   colWidths.forEach((w, i) => (worksheet.getColumn(i + 1).width = w));

//   // Table header (Row 7)
//   worksheet.getRow(7).values = [
//     "TDG PART NO",
//     "DESCRIPTION",
//     "VENDOR",
//     "VENDOR P/N",
//     "PER FIX QTY.",
//     "TOTAL QTY NEEDED",
//     "ACTUAL QTY PICKED",
//     "LOCATION/ PURCHASING COMMENTS",
//     "LEAD HAND COMMENTS",
//   ];
//   // ✅ FIX: Apply bold font and background cell-by-cell
//   worksheet.getRow(7).eachCell((cell) => {
//     cell.font = { bold: true, size: 18, name: "Calibri" };
//     cell.alignment = {
//       horizontal: "center",
//       vertical: "middle",
//       wrapText: true,
//     };
//     cell.fill = {
//       type: "pattern",
//       pattern: "solid",
//       fgColor: { argb: "D9E1F2" },
//     };
//   });

//   // Set row height to auto
//   worksheet.getRow(7).height = undefined; // Let Excel auto-adjust

//   // Apply bold borders to header rows (Rows 1 to 5)
//   for (let row = 1; row <= 5; row++) {
//     for (let col = 1; col <= 9; col++) {
//       const cell = worksheet.getRow(row).getCell(col);
//       cell.border = {
//         top: { style: "thin" },
//         bottom: { style: "thin" },
//         left: { style: "thin" },
//         right: { style: "thin" },
//       };
//     }
//   }

//   // Apply bold borders to second table (starting from Row 7)
//   for (let row = 7; row <= worksheet.rowCount; row++) {
//     for (let col = 1; col <= 9; col++) {
//       const cell = worksheet.getRow(row).getCell(col);
//       cell.border = {
//         top: { style: "thin" },
//         bottom: { style: "thin" },
//         left: { style: "thin" },
//         right: { style: "thin" },
//       };
//     }
//   }
//   // ✅ FIX: Set background color for specific cells and center alignment
//   [
//     "A1",
//     "A2",
//     "A3",
//     "A4",
//     "A5",
//     "B1",
//     "B2",
//     "B3",
//     "B4",
//     "B5",
//     "G1",
//     "G2",
//     "G3",
//     "G4",
//   ].forEach((cell) => {
//     const currentCell = worksheet.getCell(cell);
//     currentCell.fill = {
//       type: "pattern",
//       pattern: "solid",
//       fgColor: { argb: "D9E1F2" }, // Light blue background color
//     };
//     currentCell.alignment = { horizontal: "center", vertical: "middle" };
//     currentCell.font = { bold: true, size: 18, name: "Calibri" };
//   });

//   ["I1", "I2", "I3", "I4", "I5"].forEach((cell) => {
//     worksheet.getCell(cell).font = { bold: true, size: 18, name: "Calibri" };
//   });

//   // Set the rest of the cells in column I (from I1 to I5) to have center alignment
//   for (let row = 1; row <= 5; row++) {
//     const cell = worksheet.getRow(row).getCell(9); // Column I is the 9th column (index 9)
//     cell.alignment = { horizontal: "center", vertical: "middle" }; // Center alignment
//   }

//   // Adjust the width of column I if necessary
//   worksheet.getColumn(9).width = 20; // Adjusting column I width to 20 if needed

//   // Data rows (starting from Row 8)
//   components.forEach((comp, idx) => {
//     const rowIdx = idx + 8;
//     const descParts = (comp.Description || "").split("<line>");
//     const goesInto = descParts[0] ? `GOES INTO ${descParts[0]}` : "";
//     const restDesc = descParts[1] || "";
//     const fullDesc = goesInto ? `${goesInto}\n${restDesc}` : restDesc;

//     worksheet.getRow(rowIdx).values = [
//       comp.TDGPN,
//       fullDesc,
//       comp.Vendor,
//       comp.VendorPN,
//       comp.QuantityPerFixture,
//       { formula: `\$B\$4*E${rowIdx}`, result: comp.QuantityNeeded || 0 },
//       "",
//       comp.Location,
//       comp.LeadHandComments,
//     ];

//     // Alignment & font styling
//     ["A", "C", "D", "E", "F"].forEach((col) => {
//       worksheet.getCell(`${col}${rowIdx}`).alignment = { horizontal: "center" };
//     });

//     worksheet.getCell(`B${rowIdx}`).alignment = { wrapText: true };
//     worksheet.getCell(`A${rowIdx}`).font = { bold: true };
//     ["E", "F", "G", "H"].forEach((col) => {
//       worksheet.getCell(`${col}${rowIdx}`).font = { bold: true };
//     });

//     // // Red fill if short
//     const quantity = comp.QuantityNeeded || 0;
//     const available = comp.QuantityAvailable || 0;

//     const isShort = quantity > available && !comp.ConsumableOrVMI;
//     if (isShort) {
//       worksheet.getCell(`F${rowIdx}`).fill = {
//         type: "pattern",
//         pattern: "solid",
//         fgColor: { argb: "FFFFC0CB" }, // Light pink
//       };
//     }

//     // Gray fill for INHOUSE/VMI/etc
//     const loc = (comp.Location || "").toUpperCase();
//     const isGray =
//       loc.includes("INHOUSE") ||
//       loc.includes("CONSUMABLE") ||
//       (loc.includes("V") && !loc.includes("HV")) ||
//       quantity === 0;
//     if (isGray) {
//       for (let col = 1; col <= 9; col++) {
//         worksheet.getRow(rowIdx).getCell(col).fill = {
//           type: "pattern",
//           pattern: "solid",
//           fgColor: { argb: "FFD3D3D3" },
//         };
//       }
//     }
//   });

//   const startRow = 8;
//   const endRow = startRow + components.length - 1;

//   // Table borders
//   for (let r = startRow; r <= endRow; r++) {
//     for (let c = 1; c <= 9; c++) {
//       const cell = worksheet.getRow(r).getCell(c);
//       cell.border = {
//         top: { style: "thin" },
//         bottom: { style: "thin" },
//         left: { style: "thin" },
//         right: { style: "thin" },
//       };
//       cell.alignment = { wrapText: true, vertical: "middle" };
//       cell.font = { size: 18 };
//     }
//   }

//   // Final gray row
//   const finalRowIdx = endRow + 1;
//   for (let col = 1; col <= 9; col++) {
//     worksheet.getRow(finalRowIdx).getCell(col).fill = {
//       type: "pattern",
//       pattern: "solid",
//       fgColor: { argb: "FFA9A9A9" },
//     };
//   }

//   // Heading styles
//   worksheet.getCell("C1").font = { bold: true, size: 18 };
//   worksheet.getCell("C3").font = { size: 18 };
//   worksheet.getCell("C1").alignment = { horizontal: "center", vertical: "top" };

//   // Header alignment
//   ["A1", "A2", "A3", "A4", "A5"].forEach((cell) => {
//     worksheet.getCell(cell).alignment = { horizontal: "left" };
//   });
//   ["B1", "B2", "B3", "B4", "B5"].forEach((cell) => {
//     worksheet.getCell(cell).alignment = { horizontal: "right" };
//   });
//   [
//     "G1",
//     "G2",
//     "G3",
//     "G4",
//     "G5",
//     "H1",
//     "H2",
//     "H3",
//     "H4",
//     "H5",
//     "I4",
//     "I5",
//   ].forEach((cell) => {
//     worksheet.getCell(cell).alignment = { horizontal: "center" };
//   });
//   ["C1", "C2", "C3", "C4", "C5"].forEach((cell) => {
//     worksheet.getCell(cell).alignment = {
//       horizontal: "center",
//       vertical: "top",
//     };
//   });

//   // Hide grid lines
//   worksheet.views = [{ showGridLines: false }];
// }

async function addSOP(
  components,
  fixtureDescription,
  SOP,
  workbook,
  Project,
  Fixture,
  Quantity,
  RequiredDate
) {
  const worksheet = workbook.addWorksheet(SOP, {
    pageSetup: {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    },
  });

  // Format the current date to "Month Day, Year"
  const formattedPrintedDate = new Date().toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });

  // Format the required date to "D-MMM-YY" (e.g., 1-Jan-01)
  const formattedRequiredDate = RequiredDate
    ? new Date(RequiredDate)
        .toLocaleDateString("en-GB", {
          day: "numeric",
          month: "short",
          year: "2-digit",
        })
        .replace(/ /g, "-")
    : "";

  // Insert header rows (Row 1-5)
  worksheet.insertRows(1, [
    [
      "SOP #",
      SOP,
      `PICK LIST #${SOP}`,
      "",
      "",
      "",
      "PICK LIST PRINTED ON",
      "",
      formattedPrintedDate,
    ],
    ["PROJECT", Project, "", "", "", "", "PICK LIST LOG NUMBER", "", ""],
    ["FIXTURE", Fixture, fixtureDescription, "", "", "", "DATE PICKED", "", ""],
    ["QUANTITY", Quantity, "", "", "", "", "LEAD HAND SIGN OFF", "", ""],
    ["REQUIRED ON", formattedRequiredDate, "", "", "", "", "", "", ""],
  ]);

  // Merges
  ["C1:F1", "G1:H1", "G2:H2", "C3:F5", "G3:H3", "G4:H5", "I4:I5"].forEach(
    (range) => worksheet.mergeCells(range)
  );

  // Column widths
  [22.28, 69, 22.42, 27, 11.57, 13.28, 20.42, 21.57, 31.57].forEach(
    (w, i) => (worksheet.getColumn(i + 1).width = w)
  );

  // Row 7 headers
  const headers = [
    "TDG PART NO",
    "DESCRIPTION",
    "VENDOR",
    "VENDOR P/N",
    "PER FIX QTY.",
    "TOTAL QTY NEEDED",
    "ACTUAL QTY PICKED",
    "LOCATION/ PURCHASING COMMENTS",
    "LEAD HAND COMMENTS",
  ];
  worksheet.getRow(7).values = headers;
  worksheet.getRow(7).height = undefined;
  worksheet.getRow(7).eachCell((cell) => {
    cell.font = { bold: true, size: 18, name: "Calibri" };
    cell.alignment = {
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
    };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "D9E1F2" },
    };
  });

  // Apply borders and fonts to header rows
  for (let r = 1; r <= 5; r++) {
    for (let c = 1; c <= 9; c++) {
      const cell = worksheet.getRow(r).getCell(c);
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
    }
  }

  // Apply bold borders to second table (starting from Row 7)
  for (let row = 7; row <= worksheet.rowCount; row++) {
    for (let col = 1; col <= 9; col++) {
      const cell = worksheet.getRow(row).getCell(col);
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
    }
  }

  // Font + fill for header info cells
  const headerFontCells = ["A", "B", "G"].flatMap((col) =>
    [1, 2, 3, 4].map((row) => `${col}${row}`)
  );
  headerFontCells
    .concat(["A5", "B5", "I1", "I2", "I3", "I4", "I5"])
    .forEach((cell) => {
      const c = worksheet.getCell(cell);
      c.font = { bold: true, size: 18, name: "Calibri" };
      c.alignment = { horizontal: "center", vertical: "middle" };
      if (
        cell.startsWith("A") ||
        cell.startsWith("B") ||
        cell.startsWith("G")
      ) {
        c.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "D9E1F2" },
        };
      }
    });

  for (let r = 1; r <= 5; r++) {
    worksheet.getCell(`I${r}`).alignment = {
      horizontal: "center",
      vertical: "middle",
    };
  }
  worksheet.getColumn(9).width = 20;

  // Data rows
  components.forEach((comp, i) => {
    const row = worksheet.getRow(i + 8);
    const descParts = (comp.Description || "").split("<line>");
    const parent = descParts[0];
    const isTopLevel = parent && parent.trim() !== "" && parent === comp.TDGPN;

    // Create rich text for description with bold "GOES INTO" part
    let richText = [];
    if (parent) {
      richText.push({
        text: `GOES INTO ${parent}\n`,
        font: { bold: true, size: 18, name: "Calibri" },
      });
    }
    if (descParts[1]) {
      richText.push({
        text: descParts[1],
        font: { bold: false, size: 18, name: "Calibri" },
      });
    }

    // Determine if this is a consumable item
    const isConsumable =
      comp.ConsumableOrVMI ||
      (comp.Location && comp.Location.toUpperCase().includes("CONSUMABLE")) ||
      (comp.LeadHandComments &&
        comp.LeadHandComments.toUpperCase().includes("CONSUMABLE"));

    // Calculate total quantity based on type
    let totalQty = 0;
    if (isConsumable) {
      // For consumables, show the per-fixture quantity (no multiplication)
      totalQty = comp.QuantityPerFixture || 0;
    } else {
      // For regular items, multiply by fixture quantity
      totalQty = (comp.QuantityPerFixture || 0) * Quantity;
    }

    row.values = [
      comp.TDGPN,
      richText.length > 0 ? { richText } : descParts[1] || "",
      comp.Vendor,
      comp.VendorPN,
      comp.QuantityPerFixture,
      totalQty,
      "",
      comp.Location,
      comp.LeadHandComments,
    ];

    // create table border for description.....
    for (let c = 1; c <= 9; c++) {
      const cell = row.getCell(c);
      // Apply normal font for B, C, D; bold for others
      const isNormalFont = c === 2 || c === 3 || c === 4;
      cell.font = { size: 18, name: "Calibri", bold: !isNormalFont };
      cell.alignment = {
        wrapText: true,
        vertical: "middle",
        ...(c <= 6 && c !== 2 ? { horizontal: "center" } : {}),
      };
      cell.border = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
    }

    const qty = totalQty;
    const available = comp.QuantityAvailable || 0;
    if (qty > available && !isConsumable) {
      row.getCell(6).fill = {
        type: "pattern",
        pattern: "lightTrellis",
        fgColor: { argb: "FFFF0000" },
      };
    }

    const loc = (comp.Location || "").toUpperCase();
    const isGray =
      isConsumable ||
      loc.includes("INHOUSE") ||
      (loc.includes("V") && !loc.includes("HV")) ||
      qty === 0;
    if (isGray) {
      for (let c = 1; c <= 9; c++) {
        row.getCell(c).fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFD3D3D3" },
        };
      }
    }
  });

  // last row set gray color
  const finalRow = worksheet.getRow(components.length + 8);
  for (let c = 1; c <= 9; c++) {
    finalRow.getCell(c).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFA9A9A9" },
    };
  }

  // Heading font + alignment
  worksheet.getCell("C1").font = { bold: true, size: 18 };
  worksheet.getCell("C1").alignment = { horizontal: "center", vertical: "top" };
  worksheet.getCell("C3").font = { size: 18 };
  ["A", "B"].forEach((col) =>
    [1, 2, 3, 4, 5].forEach(
      (row) =>
        (worksheet.getCell(`${col}${row}`).alignment = {
          horizontal: col === "A" ? "left" : "right",
        })
    )
  );
  ["G", "H"].forEach((col) =>
    [1, 2, 3, 4, 5].forEach(
      (row) =>
        (worksheet.getCell(`${col}${row}`).alignment = { horizontal: "center" })
    )
  );
  ["C1", "C2", "C3", "C4", "C5"].forEach(
    (cell) =>
      (worksheet.getCell(cell).alignment = {
        horizontal: "center",
        vertical: "top",
      })
  );

  worksheet.views = [{ showGridLines: false }];
}

// async function addSOP(components, fixtureDescription, SOP, workbook, Project, Fixture, Quantity, RequiredDate) {
//   const worksheet = workbook.addWorksheet(SOP, {
//     pageSetup: {
//       orientation: 'landscape',
//       fitToPage: true,
//       fitToWidth: 1,
//       fitToHeight: 0,
//     },
//     properties: { defaultColWidth: 20 },
//     views: [{ showGridLines: false }]
//   });

//   // Insert header rows
//   worksheet.insertRow(1, []);
//   worksheet.insertRow(2, []);
//   worksheet.insertRow(3, []);
//   worksheet.insertRow(4, []);
//   worksheet.insertRow(5, []);
//   worksheet.insertRow(6, []);

//   worksheet.getCell('A1').value = 'SOP #';
//   worksheet.getCell('A1').font = { size: 18, bold: true };
//   worksheet.getCell('B1').value = { formula: 'MID(CELL("filename",A1),FIND("]",CELL("filename",A1))+1,255)' };
//   worksheet.getCell('C1').value = { formula: '="PICK LIST #" & B1' };
//   worksheet.getCell('C1').font = { size: 18, bold: true };
//   worksheet.mergeCells('C1:F1');
//   worksheet.getCell('G1').value = 'PICK LIST PRINTED ON';
//   worksheet.getCell('I1').value = new Date();
//   worksheet.mergeCells('G1:H1');

//   worksheet.getCell('A2').value = 'PROJECT';
//   worksheet.getCell('B2').value = Project;
//   worksheet.getCell('G2').value = 'PICK LIST LOG NUMBER';
//   worksheet.mergeCells('G2:H2');

//   worksheet.getCell('A3').value = 'FIXTURE';
//   worksheet.getCell('B3').value = Fixture;
//   worksheet.getCell('C3').value = fixtureDescription;
//   worksheet.getCell('C3').font = { size: 18 };
//   worksheet.mergeCells('C3:F5');
//   worksheet.getCell('G3').value = 'DATE PICKED';
//   worksheet.mergeCells('G3:H3');

//   worksheet.getCell('A4').value = 'QUANTITY';
//   worksheet.getCell('B4').value = Quantity;
//   worksheet.getCell('G4').value = 'LEAD HAND SIGN OFF';
//   worksheet.mergeCells('G4:H5');

//   worksheet.getCell('A5').value = 'REQUIRED ON';
//   worksheet.getCell('B5').value = RequiredDate;
//   worksheet.mergeCells('I4:I5');

//   // Column headers
//   const headerRow = worksheet.addRow([
//     'TDG PART NO', 'DESCRIPTION', 'VENDOR', 'VENDOR P/N',
//     'PER FIX QTY.', 'TOTAL QTY NEEDED', 'ACTUAL QTY PICKED', 'LOCATION/ PURCHASING COMMENTS', 'LEAD HAND COMMENTS'
//   ]);

//   headerRow.eachCell((cell) => {
//     cell.font = { bold: true };
//     cell.alignment = { horizontal: 'center' };
//     cell.border = {
//       top: { style: 'thin' },
//       left: { style: 'thin' },
//       bottom: { style: 'thin' },
//       right: { style: 'thin' },
//     };
//   });
//   worksheet.getCell('C1').alignment = { horizontal: 'center' };
// worksheet.getCell('C3').alignment = { horizontal: 'center' };

//   // Start writing component data from row 8
//   components.forEach((comp, index) => {
//     const rowIndex = index + 8;
//     const row = worksheet.getRow(rowIndex);
//     row.height = 25;

//     // Description handling with rich text
//     const descParts = (comp.Description || '').split('<line>');
//     const richText = [];

//     if (descParts[0]) {
//       richText.push({ text: `GOES INTO ${descParts[0]}\n`, font: { bold: true } });
//     }
//     if (descParts[1]) {
//       richText.push({ text: descParts[1], font: { bold: false } });
//     }

//     worksheet.getCell(`A${rowIndex}`).value = comp.TDGPN;
//     worksheet.getCell(`B${rowIndex}`).value = { richText };
//     worksheet.getCell(`B${rowIndex}`).alignment = { wrapText: true };

//     worksheet.getCell(`C${rowIndex}`).value = comp.Vendor;
//     worksheet.getCell(`D${rowIndex}`).value = comp.VendorPN;
//     worksheet.getCell(`E${rowIndex}`).value = comp.QuantityPerFixture || 0;

//     // Formula in column F: = $B$4 * E{rowIndex}
//     worksheet.getCell(`F${rowIndex}`).value = {
//       formula: `=$B$4*E${rowIndex}`,
//       result: null
//     };

//     worksheet.getCell(`E${rowIndex}`).alignment = { horizontal: 'right' };
// worksheet.getCell(`F${rowIndex}`).alignment = { horizontal: 'right' };

//     worksheet.getCell(`G${rowIndex}`).value = '';
//     worksheet.getCell(`H${rowIndex}`).value = comp.Location;
//     worksheet.getCell(`I${rowIndex}`).value = comp.LeadHandComments;

//     // Styling
//     worksheet.getCell(`A${rowIndex}`).font = { bold: true };
//     worksheet.getCell(`A${rowIndex}`).alignment = { horizontal: 'center' };
//     worksheet.getCell(`C${rowIndex}`).alignment = { horizontal: 'center' };
//     worksheet.getCell(`D${rowIndex}`).alignment = { horizontal: 'center' };
//     worksheet.getCell(`E${rowIndex}`).alignment = { horizontal: 'center', wrapText: true };
//     worksheet.getCell(`F${rowIndex}`).alignment = { horizontal: 'center' };
//     worksheet.getCell(`G${rowIndex}`).alignment = { horizontal: 'center' };

//     // Red fill if shortage
//     if ((comp.QuantityNeeded > comp.QuantityAvailable) && !comp.ConsumableOrVMI) {
//       worksheet.getCell(`F${rowIndex}`).fill = {
//         type: 'pattern',
//         pattern: 'lightTrellis',
//         fgColor: { argb: 'FFFF0000' }
//       };
//     }

//     // Gray fill for specific locations or 0 quantity
//     if (
//       comp.Location?.includes('INHOUSE') ||
//       comp.Location?.includes('CONSUMABLE') ||
//       (comp.Location?.includes('V') && !comp.Location.includes('HV')) ||
//       comp.QuantityNeeded === 0
//     ) {
//       for (let col = 1; col <= 9; col++) {
//         worksheet.getCell(rowIndex, col).fill = {
//           type: 'pattern',
//           pattern: 'solid',
//           fgColor: { argb: 'FFD3D3D3' }
//         };
//       }
//     }
//   });

//   // Column widths
//   worksheet.columns = [
//     { width: 22.28 },
//     { width: 69 },
//     { width: 22.42 },
//     { width: 27 },
//     { width: 11.57 },
//     { width: 13.28 },
//     { width: 20.42 },
//     { width: 21.57 },
//     { width: 31.57 },
//   ];
// }

const generatePickLists = async (vmParam, userParam, fixtureParam, res) => {
  try {
    const vm = vmParam || { LHREntries: [0] };
    const user = userParam || null;
    let fixture = fixtureParam || null;

    let sopNum = "-";
    const ml = await getMasterList();

    const workbook = new ExcelJS.Workbook();

    // Fetch user role info only once
    const intlUsers = await getUsersInRole("INTL");
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
          tempSOP.SOPProgramId
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
          Program: { Name: "" },
          ODD: "0001-01-01T00:00:00.000Z", // DateTime.MinValue equivalent
        };
      }

      const db = await connectDB("BOMs");
      const tempFixture = await getExplodedBOM(fixture, db);
      const refFixture = await getStoredFixture(fixture, db);

      // Process components in parallel
      const tempComponents = await Promise.all(
        tempFixture.Components.map(async (comp) => {
          const split = comp.Level.split(".");
          const parentLevel = split.slice(0, -1).join(".");
          const tempParent = tempFixture.Components.find(
            (x) => x.Level === parentLevel
          );
          const parent = tempParent ? tempParent.TDGPN : "";

          let quantityPerFixture = Math.round(comp.Quantity);
          const refComp = refFixture.Components.find(
            (x) => x.Level === comp.Level
          );
          if (refComp) {
            quantityPerFixture = Math.round(refComp.Quantity);
          }

          const tempComp = {
            Description: `${parent}<line>${comp.Description}`,
            TDGPN: comp.TDGPN,
            QuantityPerFixture: quantityPerFixture,
            QuantityNeeded: quantityPerFixture * tempQuantity,
            Vendor: comp.Vendor,
            VendorPN: comp.VendorPN,
          };

          const isDieGroup = ml.find(
            (x) => x.TDGPN === comp.TDGPN && x.GroupingName === "Die"
          );
          if (isDieGroup) {
            tempComp.QuantityNeeded = 0;
            tempComp.QuantityPerFixture = 0;
          }

          const inventory = await GetInventoryTuple(comp.TDGPN, isIntlUser);
          tempComp.Location = inventory.location;
          tempComp.QuantityAvailable = inventory.quantity;
          tempComp.ConsumableOrVMI = inventory.type;

          return tempComp;
        })
      );

      await addSOP(
        tempComponents,
        tempFixture.Description,
        sopNum,
        workbook,
        tempSOP.Program.Name,
        fixture,
        tempQuantity,
        tempSOP.ODD
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
    const month = now.toLocaleString("en-US", { month: "short" });
    const day = String(now.getDate()).padStart(2, "0");
    const dateStr = `${month}-${day}`;
    const safeFixture = fixture.replace(/[\\/:*?"<>|]/g, "-"); // Avoid filename issues
    const fileName = sopNum + " (" + safeFixture + ") " + dateStr + ".xlsx";

    res.setHeader(
      "Content-Disposition",
      `attachment; filename=\"${fileName}\"`
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.send(buffer);
    // Do not send a JSON response after sending the file buffer
  } catch (err) {
    return res.failureResponse({ message: err.message });
  }
};

const getOpenPickLists = async (fixtureNumber) => {
  try {
    let fixtureLists = [];

    const pool = await getDbPool("SOP");

    const retriveAllFixturesResult = await pool
      .request()
      .input("FixtureNumber", sql.NVarChar, fixtureNumber).query(`
        SELECT *
        FROM [SOP].[dbo].[SOPLeadHandEntries]
        WHERE FixtureNumber = @FixtureNumber
      `);

    if (retriveAllFixturesResult.recordset.length === 0) {
      return fixtureLists;
    }

    const result = await pool
      .request()
      .input("fixture", sql.VarChar, fixtureNumber.toUpperCase()).query(`
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
  `);

    return result.recordset;
  } catch (err) {
    console.log("error:", err);
    return res.failureResponse({ message: err.message });
  }
};

exports.SOPSerchService = async (req, res) => {
  try {
    if (!req.query.SOPNumber) {
      return res.badRequest({ message: "SOP ID is required" });
    }

    const { SOPNumber } = req.query;

    const pool = await getDbPool("SOP");

    let sopNumberResult = await pool
      .request()
      .input("SOPNumber", sql.NVarChar, SOPNumber).query(`
        SELECT *
        FROM [SOP].[dbo].[SOPs]
        WHERE SOPNum = @SOPNumber
      `);

    if (sopNumberResult.recordset.length === 0) {
      return res.badRequest({ message: "SOP not found" });
    }

    sopNumberResult = sopNumberResult.recordset[0];

    const LeadHandEntryResult = await pool
      .request()
      .input("SOPId", sql.Int, sopNumberResult.SOPId).query(`
        SELECT *
        FROM [SOP].[dbo].[SOPLeadHandEntries]
        WHERE SOPId = @SOPId
      `);

    const customerResult = await pool
      .request()
      .input("SOPCustomerId", sql.Int, sopNumberResult.SOPCustomerId).query(`
        SELECT *
        FROM [SOP].[dbo].[SOPCustomers]
        WHERE SOPCustomerId = @SOPCustomerId
      `);

    const programResult = await pool
      .request()
      .input("SOPProgramId", sql.Int, sopNumberResult.SOPProgramId).query(`
        SELECT *
        FROM [SOP].[dbo].[SOPPrograms]
        WHERE SOPProgramId = @SOPProgramId
      `);

    const locationResult = await pool
      .request()
      .input("SOPLocationId", sql.Int, sopNumberResult.SOPLocationId).query(`
        SELECT *
        FROM [SOP].[dbo].[SOPLocations]
        WHERE SOPLocationId = @SOPLocationId
      `);

    const sopProductionManagerResult = await pool
      .request()
      .input(
        "SOPProductionManagerId",
        sql.Int,
        sopNumberResult.SOPProductionManagerId
      ).query(`
        SELECT *
        FROM [SOP].[dbo].[SOPProductionManagers]
        WHERE SOPProductionManagerId = @SOPProductionManagerId
      `);

    const productionEntryResult = await pool
      .request()
      .input(
        "SOPProductionEntryId",
        sql.Int,
        sopNumberResult.SOPProductionEntryId
      ).query(`
        SELECT *
        FROM [SOP].[dbo].[SOPProductionEntries] 
        WHERE SOPProductionEntryId = @SOPProductionEntryId
      `);

    const sopLeadHandId = productionEntryResult.recordset[0].SOPLeadHandId;

    const leadHandResult = await pool
      .request()
      .input("SOPLeadHandId", sql.Int, sopLeadHandId).query(`
        SELECT *
        FROM [SOP].[dbo].[SOPLeadHands] 
        WHERE SOPLeadHandId = @SOPLeadHandId
      `);

    const qaEntryResult = await pool
      .request()
      .input("SOPQAEntryId", sql.Int, sopNumberResult.SOPQAEntryId).query(`
        SELECT *
        FROM [SOP].[dbo].[SOPQAEntries]
        WHERE SOPQAEntryId = @SOPQAEntryId
      `);

    const shippingEntryResult = await pool
      .request()
      .input("SOPShippingEntryId", sql.Int, sopNumberResult.SOPShippingEntryId)
      .query(`
        SELECT *
        FROM [SOP].[dbo].[SOPShippingEntries]
        WHERE SOPShippingEntryId = @SOPShippingEntryId
      `);

    const fixtureResults = await pool
      .request()
      .input("SOPId", sql.Int, sopNumberResult.SOPId).query(`
        SELECT *
        FROM [SOP].[dbo].[SOPLeadHandEntries]
        WHERE SOPId = @SOPId
      `);

    const db = await connectDB("BOMs");
    const collection = db.collection("Fixture");

    // Process all fixtures
    const fixturesDetailed = await Promise.all(
      fixtureResults.recordset.map(async (fixture) => {
        // Get assembler for this fixture
        let assemblerIdResult = { recordset: [] };
        if (fixture.SOPAssemblerId) {
          assemblerIdResult = await pool
            .request()
            .input("SOPAssemblerId", sql.Int, fixture.SOPAssemblerId).query(`
              SELECT *
              FROM [SOP].[dbo].[SOPAssemblers]
              WHERE SOPAssemblerId = @SOPAssemblerId
            `);
        }

        // Get MongoDB data for this fixture
        const fixtureMongoData = await collection
          .find({ Name: fixture.FixtureNumber })
          .limit(10)
          .toArray();

        return {
          ...fixture,
          assembler: assemblerIdResult.recordset,
          fixtureMongoData,
        };
      })
    );

    const responseData = {
      ...sopNumberResult,
      customer: customerResult.recordset,
      program: programResult.recordset,
      location: locationResult.recordset,
      sopProductionManager: sopProductionManagerResult.recordset,
      productionEntry: productionEntryResult.recordset,
      leadHandEntry: LeadHandEntryResult.recordset,
      leadHand: leadHandResult.recordset,
      qaEntry: qaEntryResult.recordset,
      shippingEntry: shippingEntryResult.recordset,
      fixtures: fixturesDetailed, // <-- now an array of detailed fixture info
    };

    return res.ok({
      message: "Successfully fetched SOP data",
      data: responseData,
    });
  } catch (err) {
    console.log("error:", err);
    return res.failureResponse({ message: err.message });
  }
};

exports.fixtureDetails = async (req, res) => {
  try {
    const { fixtureNumber } = req.query;

    const fixedFixtureNumber = fixFixtureName(fixtureNumber);

    const db = await connectDB("BOMs");
    const collection = db.collection("Fixture");
    // Get MongoDB data for this fixture
    const fixtureMongoData = await collection
      .find({ Name: fixedFixtureNumber })
      .limit(10)
      .toArray();

    if (fixtureMongoData.length === 0) {
      return res.badRequest({ message: "Fixture not found" });
    }

    const openPickLists = await getOpenPickLists(fixtureNumber);

    return res.ok({
      message: "Successfully fetched fixture details",
      data: openPickLists,
    });
  } catch (err) {
    console.log("error:", err);
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
        res
      );
    } else if (fixture) {
      // Call for blank pick list using fixture number
      await generatePickLists(null, user, fixture, res);
    } else {
      return res.status(400).json({ message: "Missing required parameters" });
    }
  } catch (err) {
    console.log("error:", err);
    return res.failureResponse({ message: err.message });
  }
};
