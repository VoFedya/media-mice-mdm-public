/*************************************************
 * OPEN SIDEBAR
 *************************************************/
function openVendorAdd() {

  const html = HtmlService
    .createHtmlOutputFromFile("VendorAdd")
    .setTitle("Vendor Add");

  SpreadsheetApp.getUi().showSidebar(html);

}

/*************************************************
 * LOAD VENDOR SEARCH
 *************************************************/
function getVendorSearch() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.VENDOR_SEARCH);

  const data = sheet.getDataRange().getValues();

  const vendors = [];

  for (let i = 1; i < data.length; i++) {

    vendors.push({

      id: data[i][0],
      name: data[i][1],
      type: data[i][2],
      country: data[i][3],
      city: data[i][4],
      evaluated: data[i][5],
      selected: data[i][6],
      avgRating: data[i][7],
      email: data[i][8],
      website: data[i][9],
      contact: data[i][10],
      portfolio: data[i][11]

    });

  }

  return vendors;

}

/*************************************************
 * CREATE VENDOR
 *************************************************/
function createVendor(vendor) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.MASTER_VENDORS);

  const lastId = sheet
    .getRange(sheet.getLastRow(), CONFIG.COL.MASTER.VENDOR_ID)
    .getValue();

  const nextNumber =
    Number(String(lastId).replace("V", "")) + 1;

  const vendorId =
    "V" + String(nextNumber).padStart(4, "0");

  sheet.appendRow([

    vendorId,
    vendor.name,
    vendor.type,
    vendor.email,
    vendor.website,
    vendor.contact,
    vendor.portfolio

  ]);

  return vendorId;

}

/*************************************************
 * CHECK VENDOR
 *************************************************/
function checkVendor(vendor) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.MASTER_VENDORS);

  const data = sheet.getDataRange().getValues();

  const matches = [];

  const newName =
    vendor.name.trim().toLowerCase();

  const newWebsite =
    normalizeWebsite(vendor.website);

  const newEmail =
    vendor.email.trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {

    const existingName =
      String(data[i][CONFIG.COL.MASTER.VENDOR_NAME - 1]).trim();

    const existingWebsite =
      String(data[i][CONFIG.COL.MASTER.WEBSITE - 1]);

    const existingEmail =
      String(data[i][CONFIG.COL.MASTER.EMAIL - 1]);

    let reason = "";

    if (existingName.toLowerCase() === newName) {

      reason = "Exact name";

    } else if (
      newWebsite &&
      normalizeWebsite(existingWebsite) === newWebsite
    ) {

      reason = "Same website";

    } else if (
      newEmail &&
      existingEmail.trim().toLowerCase() === newEmail
    ) {

      reason = "Same email";

    }

    if (reason) {

      matches.push({

        id: data[i][CONFIG.COL.MASTER.VENDOR_ID - 1],
        name: existingName,
        type: data[i][CONFIG.COL.MASTER.VENDOR_TYPE - 1],
        email: existingEmail,
        website: existingWebsite,
        contact: data[i][CONFIG.COL.MASTER.CONTACT_PERSON - 1],
        portfolio: data[i][CONFIG.COL.MASTER.PORTFOLIO - 1],
        reason: reason

      });

    }

  }

  return matches;

}

function normalizeWebsite(url){

  return String(url)
    .toLowerCase()
    .replace(/^https?:\/\//,"")
    .replace(/^www\./,"")
    .replace(/\/$/,"")
    .trim();

}