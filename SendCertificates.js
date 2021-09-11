function sendCertificateEmail() {
  var file = SpreadsheetApp.openById(
    "1Ube_pqbgJqoXdTcP-BKo-nzCVGD3uaOqRRJNl4ynM2c"
  );
  SpreadsheetApp.setActiveSpreadsheet(file);
  var certSrcFolder = DriveApp.getFolderById(
    "1vzbrF68-yDaoUUJXHR_UJAAOOd4jE1N6"
  );

  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;
  var rows = 7; //Change this to maximum row
  var dataRange = sheet.getRange(startRow, 1, rows, 7); //startrow, start col, row len, col len
  var data = dataRange.getValues();

  const senderName = "BNCC Bandung";
  const title = "BNCC Techbinar: AI in Industry 4.0";
  const sent_col = 4;

  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; // First column

    var name = row[1];
    var certLink = row[2];
    var status = row[3];
    var sent = row[sent_col];
    var content = {
      name,
      title,
      link: certLink,
      email: emailAddress,
    };

    var subject = `${title} Certificate`;
    var htmlBody = generateCertTemplate(content);
    if (status == "INVT" && sent != "SENT" && sent != "FAILED") {
      if (certLink != "") {
        // send link instead
        MailApp.sendEmail({
          name: senderName,
          noReply: true,
          to: emailAddress,
          subject: subject,
          htmlBody: htmlBody,
        });
        Logger.log("link sent");
      } else {
        // send attachment
        const certiFile = getCertificateFile(certSrcFolder, name);
        if (!certiFile) {
          sheet.getRange(startRow + i, sent_col + 1).setValue("FAILED");
          SpreadsheetApp.flush();
          continue;
        } else {
          MailApp.sendEmail({
            name: senderName,
            noReply: true,
            to: emailAddress,
            subject: subject,
            htmlBody: htmlBody,
            attachments: [certiFile],
          });
        }
        Logger.log("attachment sent");
      }

      sheet.getRange(startRow + i, sent_col + 1).setValue("SENT");
      SpreadsheetApp.flush();
    }
  }
}

// https://developers.google.com/apps-script/reference/base/mime-type
function getCertificateFile(folder, filename) {
  const mime = MimeType.PDF;
  var files = folder.searchFiles(
    'mimeType = "' + mime + '" and title contains "' + filename + '"'
  );
  while (files.hasNext()) {
    return files.next().getAs(mime);
  }
  return null;
}

function generateCertTemplate(data) {
  var template = HtmlService.createTemplateFromFile("webinarcertificate");
  template.data = data;
  var htmlBody = template.evaluate().getContent();
  return htmlBody;
}
