function doGet(e) {
  var action = (e.parameter.action || "").toLowerCase();
  var timestamp = e.parameter.timestamp;
  var missionaryEmail = e.parameter.missionary;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var timestampCol = headers.indexOf("Timestamp");
  var statusCol = headers.indexOf("Status");

  var statusText = action === "approved" ? "Approved" : "Denied";
  var actionText = action === "approved" ? "approved ✅" : "denied ❌";

  // Update sheet
  for (var i = 1; i < data.length; i++) {
    if (data[i][timestampCol] == timestamp) {
      sheet.getRange(i + 1, statusCol + 1).setValue(statusText);
      break;
    }
  }

  // Send email to missionary
  var subject = "Temple Visit Request Update";

  var message = `
  <div style="font-family:Arial, sans-serif; background-color:#f4f6f8; padding:30px;">
    <div style="max-width:600px; margin:auto; background:white; padding:25px; border-radius:10px; border:1px solid #ddd; text-align:center;">
      
      <h2 style="color:${action === 'approved' ? '#2E7D32' : '#C62828'};">
        Temple Visit Request Update
      </h2>

      <p>Hello,</p>

      <p>
        Your temple visit request submitted on 
        <strong>${timestamp}</strong> has been
      </p>

      <h3 style="color:${action === 'approved' ? '#2E7D32' : '#C62828'};">
        ${actionText}
      </h3>

      <p style="margin-top:20px;">
        Thank you for your preparation and dedication.
      </p>

    </div>
  </div>
  `;

  MailApp.sendEmail({
    to: missionaryEmail,
    subject: subject,
    htmlBody: message
  });

  // Styled confirmation page
  return HtmlService.createHtmlOutput(`
  <div style="font-family:Arial, sans-serif; background-color:#f4f6f8; height:100vh; display:flex; align-items:center; justify-content:center;">
    <div style="background:white; padding:30px; border-radius:10px; border:1px solid #ddd; text-align:center; max-width:400px;">
      
      <h2 style="color:${action === 'approved' ? '#2E7D32' : '#C62828'};">
        ${action === 'approved' ? ' Approved' : ' Denied'}
      </h2>

      <p style="margin-top:10px;">
        The submission has been successfully <strong>${actionText}</strong>.
      </p>

    </div>
  </div>
  `);
}