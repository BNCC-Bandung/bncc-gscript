function sendZoomInvitationEmail() {
  // Change with selected form
  var file = SpreadsheetApp.openById(
    "1DLbDTCEFxyyYl_frt8ltJnX6UWIjcu8P9U4k1LjVKNk"
  );
  SpreadsheetApp.setActiveSpreadsheet(file);

  // Webinar Data
  const link = "https://binus.zoom.us/";
  const title = "BNCC Techbinar: AI in Industry 4.0";
  const date = "Kamis, 1 July 2020";
  const time = "13.00 - 14.00 WIB";
  const passcode = "3fdas21";
  const meetingId = "846 9220 1721";

  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;
  var rows = 7;
  var dataRange = sheet.getRange(startRow, 1, rows, 7); //startrow, start col, row len, col len
  var data = dataRange.getValues();

  const sent_col = 3;
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; // First column
    var name = row[1];
    var status = row[2];
    var sent = row[sent_col];

    var content = {
      name,
      title,
      email: emailAddress,
      link,
      date,
      time,
      passcode,
      meetingId,
    };

    var subject = `${title} Confirmation`;
    var htmlBody = generateZoomTemplate(content);
    if (status == "INVT" && sent != "SENT") {
      MailApp.sendEmail({
        name: "BNCC Bandung",
        noReply: true,
        to: emailAddress,
        subject: subject,
        htmlBody: htmlBody,
      });
      sheet.getRange(startRow + i, sent_col + 1).setValue("SENT");
      SpreadsheetApp.flush();
    }
  }
}

function generateZoomTemplate(data) {
  var template = HtmlService.createTemplateFromFile("zoominvitation");
  template.data = data;
  var htmlBody = template.evaluate().getContent();
  return htmlBody;
}
