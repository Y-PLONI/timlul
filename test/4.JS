function sendSimpleEmail() {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("גיליון1");
    var destEmail = ss.getRange("L2").getValue();
    GmailApp.sendEmail(destEmail, "שלום", "שלום");
  }
  sendSimpleEmail();
  