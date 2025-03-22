function doGet() {
  return HtmlService.createHtmlOutputFromFile('dashboard');
}

function getTicketData() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
    if (!sheet) {
      throw new Error("Sheet 'Sheet1' not found.");
    }

    var data = {};
    data.totalSold = sheet.getRange('M2').getValue();
    data.totalRevenue = sheet.getRange('N2').getValue();
    data.ticketsRemaining = sheet.getRange('O2').getValue();
    data.percentageSold = sheet.getRange('P2').getValue();

    data.gradeDistribution = [
      ['9th Grade', sheet.getRange('R2').getValue()],
      ['10th Grade', sheet.getRange('R3').getValue()],
      ['11th Grade', sheet.getRange('R4').getValue()],
      ['12th Grade', sheet.getRange('R5').getValue()],
    ];

    data.groupDistribution = [
      ['RPBHS', sheet.getRange('R12').getValue()],
      ['District Guest', sheet.getRange('R13').getValue()],
      ['External Guest', sheet.getRange('R14').getValue()],
    ];

    Logger.log("Fetched Data: " + JSON.stringify(data));
    Logger.log("Grade Data: " + JSON.stringify(data.gradeDistribution));
    Logger.log("Group Data: " + JSON.stringify(data.groupDistribution));

    return data;
  } catch (e) {
    Logger.log("Error in getTicketData: " + e.toString());
    return null;
  } // This closing brace was missing
} // This closing brace was also missing
