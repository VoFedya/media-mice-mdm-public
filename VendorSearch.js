/*************************************************
 * REBUILD VENDOR SEARCH
 *************************************************/
function rebuildVendorSearch() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const masterSheet = ss.getSheetByName(
    CONFIG.SHEET.MASTER_VENDORS
  );

  const vendorEventSheet = ss.getSheetByName(
    "Vendor_Event"
  );

  const eventsSheet = ss.getSheetByName(
    CONFIG.SHEET.EVENTS
  );

  const searchSheet = ss.getSheetByName(
    CONFIG.SHEET.VENDOR_SEARCH
  );

  const master =
    masterSheet.getDataRange().getValues();

  const vendorEvent =
    vendorEventSheet.getDataRange().getValues();

  const events =
    eventsSheet.getDataRange().getValues();

  /*************************************************
   * MASTER VENDOR LOOKUP
   *************************************************/

  const masterMap = new Map();

  for (let i = 1; i < master.length; i++) {

    const vendorId =
      master[i][CONFIG.COL.MASTER.VENDOR_ID - 1];

    if (!vendorId) continue;

    masterMap.set(String(vendorId).trim(), {

      name:
        master[i][CONFIG.COL.MASTER.VENDOR_NAME - 1],

      email:
        master[i][CONFIG.COL.MASTER.EMAIL - 1],

      website:
        master[i][CONFIG.COL.MASTER.WEBSITE - 1],

      contact:
        master[i][CONFIG.COL.MASTER.CONTACT_PERSON - 1],

      portfolio:
        master[i][CONFIG.COL.MASTER.PORTFOLIO - 1]

    });

  }

  /*************************************************
   * EVENT LOOKUP
   *************************************************/

  const eventMap = new Map();

  for (let i = 1; i < events.length; i++) {

    const eventId =
      events[i][CONFIG.COL.EVENTS.EVENT_ID - 1];

    if (!eventId) continue;

    eventMap.set(String(eventId).trim(), {

      country:
        events[i][CONFIG.COL.EVENTS.COUNTRY - 1],

      city:
        events[i][CONFIG.COL.EVENTS.CITY - 1]

    });

  }

  /*************************************************
   * AGGREGATE VENDOR EVENT DATA
   *************************************************/

  const agg = new Map();

  for (let i = 1; i < vendorEvent.length; i++) {

    const eventId =
      String(vendorEvent[i][0] || "").trim();

    const vendorId =
      String(vendorEvent[i][1] || "").trim();

    const category =
      String(vendorEvent[i][2] || "").trim();

    const selected =
      vendorEvent[i][4];

    const rating =
      vendorEvent[i][5];

    if (!eventId || !vendorId || !category) {
      continue;
    }

    const event =
      eventMap.get(eventId);

    if (!event) {
      continue;
    }

    const key = [
      vendorId,
      category,
      event.country,
      event.city
    ].join("|");

    if (!agg.has(key)) {

      agg.set(key, {

        vendorId,
        category,
        country: event.country,
        city: event.city,

        evaluated: 0,
        selected: 0,

        ratingSum: 0,
        ratingCount: 0

      });

    }

    const row =
      agg.get(key);

    const decision =
      String(selected || "").trim().toLowerCase();

    if (
      decision === "evaluate" ||
      decision === "selected"
    ) {
      row.evaluated++;
    }

    if (decision === "selected") {
      row.selected++;
    }

    if (rating !== "" && rating != null) {

      row.ratingSum += Number(rating);
      row.ratingCount++;

    }
  }  
  /*************************************************
   * BUILD OUTPUT
   *************************************************/

  const output = [];

  agg.forEach(row => {

    const masterInfo =
      masterMap.get(row.vendorId);

    if (!masterInfo) {
      return;
    }

    output.push([

      row.vendorId,

      masterInfo.name,

      row.category,

      row.country,

      row.city,

      row.evaluated,

      row.selected,

      row.ratingCount
        ? row.ratingSum / row.ratingCount
        : "",

      masterInfo.email,

      masterInfo.website,

      masterInfo.contact,

      masterInfo.portfolio

    ]);

  });

  /*************************************************
   * SORT
   *************************************************/

  output.sort((a, b) =>
    String(a[0]).localeCompare(String(b[0]))
  );

  /*************************************************
   * WRITE
   *************************************************/

  searchSheet.clearContents();

  searchSheet.appendRow([
    "vendor_id",
    "vendor_name",
    "vendor_type",
    "country",
    "city",
    "evaluated",
    "selected",
    "avg_rating",
    "email",
    "website",
    "contact_person",
    "portfolio"
  ]);

  if (output.length > 0) {

    searchSheet
      .getRange(
        2,
        1,
        output.length,
        output[0].length
      )
      .setValues(output);

  }

  ss.toast(
    `${output.length} vendor search records rebuilt.`,
    "MDM",
    3
  );

}