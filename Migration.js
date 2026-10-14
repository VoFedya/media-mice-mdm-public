/*************************************************
 * FIND EXISTING VENDOR
 *************************************************/
function findExistingVendor(queueRow) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.MASTER_VENDORS);

  const data = sheet.getDataRange().getValues();

  const vendorId =
    String(queueRow[4] || "").trim();

  const email =
    String(queueRow[8] || "")
      .trim()
      .toLowerCase();

  const website =
    String(queueRow[9] || "")
      .trim()
      .toLowerCase()
      .replace(/^https?:\/\//, "")
      .replace(/^www\./, "")
      .replace(/\/$/, "");

  const matches = {

    vendorId: null,
    website: null,
    email: null

  };

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const masterVendorId =
      String(row[0] || "").trim();

    const masterEmail =
      String(row[2] || "")
        .trim()
        .toLowerCase();

    const masterWebsite =
      String(row[3] || "")
        .trim()
        .toLowerCase()
        .replace(/^https?:\/\//, "")
        .replace(/^www\./, "")
        .replace(/\/$/, "");

    if (
      !matches.vendorId &&
      vendorId &&
      vendorId === masterVendorId
    ) {

      matches.vendorId = {
        vendorId: masterVendorId,
        row: i + 1
      };

    }

    if (
      !matches.website &&
      website &&
      website === masterWebsite
    ) {

      matches.website = {
        vendorId: masterVendorId,
        row: i + 1
      };

    }

    if (
      !matches.email &&
      email &&
      email === masterEmail
    ) {

      matches.email = {
        vendorId: masterVendorId,
        row: i + 1
      };

    }

    if (
      matches.vendorId &&
      matches.website &&
      matches.email
    ) {
      break;
    }

  }

  return matches;

}

/*************************************************
 * CHECK VENDOR REQUEST
 *************************************************/
function checkVendorRequest() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.VENDOR_UPDATE_QUEUE);

  const row = sheet.getActiveCell().getRow();

  const queueRow = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  const exact =
    findExistingVendor(queueRow);

  const fuzzy =
    findFuzzyVendorMatches(queueRow);

  let differences = {};

  const vendorIds = new Set();

  if (exact.vendorId) {
    vendorIds.add(exact.vendorId.vendorId);
  }

  if (exact.website) {
    vendorIds.add(exact.website.vendorId);
  }

  if (exact.email) {
    vendorIds.add(exact.email.vendorId);
  }

  if (vendorIds.size === 1) {

    const vendorId = [...vendorIds][0];

    const master =
      getMasterVendorRow(vendorId);

    differences =
      compareVendorData(queueRow, master.values);

  }

  return {

    row,

    queue: {
      vendorId: queueRow[4],
      vendorName: queueRow[5],
      vendorType: queueRow[6],
      contactPerson: queueRow[7],
      email: queueRow[8],
      website: queueRow[9],
      portfolio: queueRow[10]
    },

    exact,

    fuzzy,

    differences

  };

}

/*************************************************
 * OPEN VENDOR REVIEW
 *************************************************/
function openVendorReview() {

  const html = HtmlService
    .createHtmlOutputFromFile("VendorCheck")
    .setTitle("Vendor Review");

  SpreadsheetApp.getUi().showSidebar(html);

}

/*************************************************
 * GET MASTER VENDOR ROW
 *************************************************/
function getMasterVendorRow(vendorId) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.MASTER_VENDORS);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    if (String(data[i][0]).trim() === String(vendorId).trim()) {

      return {
        row: i + 1,
        values: data[i]
      };

    }

  }

  throw new Error(
    `Vendor ${vendorId} not found in Master_Vendors.`
  );

}

/*************************************************
 * COMPARE VENDOR DATA
 *************************************************/
function compareVendorData(queueRow, masterRow) {

  const differences = {};

  const fields = [
    {
      name: "vendor_name",
      queue: 5,
      master: 1
    },
    {
      name: "email",
      queue: 8,
      master: 2
    },
    {
      name: "website",
      queue: 9,
      master: 3
    },
    {
      name: "contact_person",
      queue: 7,
      master: 4
    },
    {
      name: "portfolio",
      queue: 10,
      master: 5
    }
  ];

  fields.forEach(field => {

    let submitted =
      String(queueRow[field.queue] || "").trim();

    let master =
      String(masterRow[field.master] || "").trim();

    if (field.name === "vendor_name") {

      submitted = submitted
        .toLowerCase()
        .replace(/\s+/g, " ");

      master = master
        .toLowerCase()
        .replace(/\s+/g, " ");

    }

    if (submitted !== master) {

      differences[field.name] = {

        master,
        submitted

      };

    }

  });

  return differences;

}

/*************************************************
 * APPROVE VENDOR
 *************************************************/
function approveVendor(row, selections) {

  const queueSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.VENDOR_UPDATE_QUEUE);

  const queueRow = queueSheet
    .getRange(row, 1, 1, queueSheet.getLastColumn())
    .getValues()[0];

  const matches =
    findExistingVendor(queueRow);

  const vendorIds = new Set();

  if (matches.vendorId) {
    vendorIds.add(matches.vendorId.vendorId);
  }

  if (matches.website) {
    vendorIds.add(matches.website.vendorId);
  }

  if (matches.email) {
    vendorIds.add(matches.email.vendorId);
  }

  if (vendorIds.size === 0) {
    throw new Error("No exact match found.");
  }

  if (vendorIds.size > 1) {
    throw new Error("Conflicting exact matches found.");
  }

  const vendorId = [...vendorIds][0];

  const category = queueRow[6];

  addVendorCategory(vendorId, category);

  const eventId = getEventIdBySheetId(queueRow[1]);

  addVendorEvent(
    eventId,
    vendorId,
    category,
    queueRow[5],
    queueRow[11],
    queueRow[12],
    queueRow[14],
    queueRow[15]
  );

  const master =
    getMasterVendorRow(vendorId);
  
  const masterSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.MASTER_VENDORS);

  Logger.log(master);

  Logger.log(
    JSON.stringify(selections, null, 2)
  );

  if (selections.email === "submitted") {

    masterSheet
      .getRange(master.row, 3)
      .setValue(queueRow[8]);

  }

  if (selections.website === "submitted") {

    masterSheet
      .getRange(master.row, 4)
      .setValue(queueRow[9]);

  }

  if (selections.contact_person === "submitted") {

    masterSheet
      .getRange(master.row, 5)
      .setValue(queueRow[7]);

  }

  if (selections.portfolio === "submitted") {

    masterSheet
      .getRange(master.row, 6)
      .setValue(queueRow[10]);

  }

  queueSheet
    .getRange(row, QUEUE.VENDOR_ID)
    .setValue(vendorId);

  queueSheet
    .getRange(row, QUEUE.STATUS)
    .setValue("Approved");

  queueSheet
    .getRange(row, QUEUE.PROCESSED_AT)
    .setValue(new Date());

  // Update the original Event Workbook
  const eventSS = SpreadsheetApp
    .openById(queueRow[0]);

  const eventSheet = eventSS
    .getSheetByName(queueRow[2]);

  const eventRow = queueRow[3];

  // Column A is reserved for vendor_id.
  eventSheet
    .getRange(eventRow, 1)
    .setValue(vendorId);

}

/*************************************************
 * GENERATE VENDOR ID
 *************************************************/
function generateVendorId() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.MASTER_VENDORS);

  const lastRow = sheet.getLastRow();

  const lastId = sheet
    .getRange(lastRow, 1)
    .getValue();

  const number = Number(
    String(lastId).replace("V", "")
  );

  const vendorId =
    "V" + String(number + 1).padStart(4, "0");

  Logger.log(vendorId);

  return vendorId;

}

/*************************************************
 * CREATE NEW VENDOR
 *************************************************/
function createNewVendor(row) {

  const queueSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.VENDOR_UPDATE_QUEUE);

  const queueRow = queueSheet
    .getRange(row, 1, 1, queueSheet.getLastColumn())
    .getValues()[0];

  const vendorId =
    generateVendorId();

  const masterSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.MASTER_VENDORS);

  masterSheet.appendRow([
    vendorId,
    queueRow[5],   // vendor_name
    queueRow[8],   // email
    queueRow[9],   // website
    queueRow[7],   // contact_person
    queueRow[10]   // portfolio
  ]);

  const category = queueRow[6];

  addVendorCategory(
    vendorId,
    category
  );

  const eventId =
    getEventIdBySheetId(queueRow[1]);

  addVendorEvent(
    eventId,
    vendorId,
    category,
    queueRow[5],
    queueRow[11],
    queueRow[12],
    queueRow[14],
    queueRow[15]
  );

  queueSheet
    .getRange(row, QUEUE.VENDOR_ID)
    .setValue(vendorId);

  queueSheet
    .getRange(row, QUEUE.STATUS)
    .setValue("Approved");

  queueSheet
    .getRange(row, QUEUE.PROCESSED_AT)
    .setValue(new Date());

  const eventSS = SpreadsheetApp
    .openById(queueRow[0]);

  const eventSheet = eventSS
    .getSheetByName(queueRow[2]);

  const eventRow = queueRow[3];

  // Column A is reserved for vendor_id.
  eventSheet
    .getRange(eventRow, 1)
    .setValue(vendorId);

}

/*************************************************
 * GET EVENT ID
 *************************************************/
function getEventIdBySheetId(sheetId) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.EVENTS);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    if (String(data[i][1]).trim() === String(sheetId).trim()) {

      const eventId = String(data[i][2] || "").trim();

      if (!eventId) {
        throw new Error(
          `Event for sheet ${sheetId} does not have an event_id yet.`
        );
      }

      return eventId;
    }
  }

  throw new Error(
    `Event with sheet_id ${sheetId} not found in Events.`
  );
}

/*************************************************
 * ADD VENDOR CATEGORY
 *************************************************/
function addVendorCategory(vendorId, category) {

  if (!vendorId || !category) {
    return;
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.VENDOR_CATEGORIES);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const existingVendorId =
      String(data[i][0] || "").trim();

    const existingCategory =
      String(data[i][1] || "").trim();

    if (
      existingVendorId === String(vendorId).trim() &&
      existingCategory.toLowerCase() ===
        String(category).trim().toLowerCase()
    ) {
      return;
    }

  }

  sheet.appendRow([
    vendorId,
    category
  ]);

}

function addVendorEvent(eventId, vendorId, category, vendorName, quotationNotes, equipmentList, decision, rating) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName("Vendor_Event");

  const data = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 8)
    .getValues();

  const exists = data.some(row =>
    row[0] === eventId &&
    row[1] === vendorId &&
    row[2] === category
  );

  if (exists) {
    return;
  }

  sheet.appendRow([
    eventId,
    vendorId,
    category,
    vendorName,
    decision, 
    rating,
    quotationNotes,
    equipmentList
  ]);

}