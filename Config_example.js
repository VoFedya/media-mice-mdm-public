const CONFIG = {
  EVENT_WORKBOOK_ID:
    "...",

  SHEET: {
    MASTER_VENDORS: "Master_Vendors",
    HISTORICAL_VENDORS: "Historical_Vendors",
    EVENTS: "Events",
    VENDOR_CATEGORIES: "Vendor_Categories",
    VENDOR_COVERAGE: "Vendor_Coverage",
    VENDOR_METRICS: "Vendor_Metrics",
    VENDOR_DIRECTORY: "Vendor_Directory",
    VENDOR_SEARCH: "Vendor_Search",
    VENDOR_UPDATE_QUEUE: "VENDOR_UPDATE_QUEUE"
  },

  COL: {

    MASTER: {
      VENDOR_ID: 1,
      VENDOR_NAME: 2,
      EMAIL: 3,
      WEBSITE: 4,
      CONTACT_PERSON: 5,
      PORTFOLIO: 6
    },

    HISTORICAL: {
      EVENT_ID: 3,
      VENDOR_ID: 7
    },

    EVENTS: {
      EVENT_ID: 3,
      COUNTRY: 6,
      CITY: 7
    },

    COVERAGE: {
      VENDOR_ID: 1,
      COUNTRY: 2,
      CITY: 3,
      EVENTS_COUNT: 4
    }

  },

  EVENT_WORKBOOK: {

    VENDOR_ID: 1,
    VENDOR_NAME: 2,
    STATUS: 3,
    VENDOR_TYPE: 4,
    CONTACT: 5,
    EMAIL: 6,
    WEBSITE: 7,

    SEND_EMAIL: 8,
    RFQ_SENT: 9,
    LAST_ACTIVITY: 10,
    NEW_ACTIVITY: 11,

    SAMPLE: 12,
    QUOTATION: 13,
    FULL_DAY: 14,
    HALF_DAY: 15,
    EQUIPMENT: 16,
    COMMENTS: 17,

    RFQ_STATUS: 18,
    REVIEW: 19,
    PM_NOTES: 20

  }

};

const QUEUE = {
  VENDOR_ID: 5,
  Rating: 16,
  STATUS: 17,
  PROCESSED_AT: 19
};