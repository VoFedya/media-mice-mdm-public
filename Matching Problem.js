/*************************************************
 * FIND FUZZY VENDOR MATCHES
 *************************************************/
function findFuzzyVendorMatches(queueRow) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET.MASTER_VENDORS);

  const data = sheet.getDataRange().getValues();

  const vendorName =
    String(queueRow[5] || "").trim().toLowerCase();

  const matches = [];

  for (let i = 1; i < data.length; i++) {

    const masterName =
      String(data[i][1] || "").trim().toLowerCase();

    if (!masterName) continue;

    const distance =
      levenshtein(vendorName, masterName);

    const maxLength =
      Math.max(vendorName.length, masterName.length);

    const score =
      Math.round((1 - distance / maxLength) * 100);

    if (score >= 80) {

      matches.push({

        vendorId: data[i][0],
        vendorName: data[i][1],
        row: i + 1,
        score: score

      });

    }

  }

  matches.sort((a, b) => b.score - a.score);

  return matches.slice(0, 3);

}

/*************************************************
 * LEVENSHTEIN DISTANCE
 *************************************************/
function levenshtein(a, b) {

  const matrix = [];

  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }

  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= b.length; i++) {

    for (let j = 1; j <= a.length; j++) {

      if (b.charAt(i - 1) === a.charAt(j - 1)) {

        matrix[i][j] = matrix[i - 1][j - 1];

      } else {

        matrix[i][j] = Math.min(

          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1

        );

      }

    }

  }

  return matrix[b.length][a.length];

}

function testFuzzy() {

  const sheet = SpreadsheetApp.getActiveSheet();

  const row = sheet.getActiveCell().getRow();

  const queueRow =
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  Logger.log(
    JSON.stringify(
      findFuzzyVendorMatches(queueRow),
      null,
      2
    )
  );

}