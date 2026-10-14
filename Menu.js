function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu("MDM Tools")
    .addItem("Vendor Review", "openVendorReview")
    .addItem("Rebuild Vendor Search", "rebuildVendorSearch")
    .addToUi();

}