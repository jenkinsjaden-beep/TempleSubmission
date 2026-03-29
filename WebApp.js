/**
 * Temple Visit Request System
 * ---------------------------------------
 * BY ELDER JENKINS
 * Created: March 2026
 * Description: 
 *   This Google Apps Script handles temple visit requests submitted via Google Forms.
 *   - Sends formatted emails to the mission president for approval/denial
 *   - Updates the form responses sheet with Approved/Denied status
 *   - Sends personalized confirmation emails to missionaries
 *   - Web App provides clickable Approve/Deny buttons and a confirmation page
 *
 * Notes:
 *   - Make sure column names in the response sheet match exactly with the script
 *   - Missionary names are pulled from the "Names of Missionaries" form question -- Still working on this as of March 29th 2026
 *   - Timestamps are used as unique identifiers for each request, I'm debating or not to keep em or not
 *   - Style elements (colors, fonts, and buttons) are customizable
 */

function doGet(e) {
  var action = (e.parameter.action || "").toLowerCase();
  var timestamp = e.parameter.timestamp;
  var missionaryEmail = e.parameter.missionary;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var timestampCol = headers.indexOf("Timestamp");
  var statusCol = headers.indexOf("Status");

  var MISSION_NAME = "Philippines Manila Mission";
var PRIMARY_COLOR = "#1B5E20"; // deep green

   var formattedDate = Utilities.formatDate(
  new Date(timestamp),
  Session.getScriptTimeZone(),
  "MMMM d, yyyy 'at' h:mm a"
);

var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
var data = sheet.getDataRange().getValues();
var headers = data[0];

var timestampCol = headers.indexOf("Timestamp");
var statusCol = headers.indexOf("Status");
var nameCol = headers.indexOf("Names of missionaries: "); 

var missionaryName = "Missionary"; // default

for (var i = 1; i < data.length; i++) {
  if (data[i][timestampCol] == timestamp) {
    missionaryName = data[i][nameCol] || "Missionary";
    sheet.getRange(i + 1, statusCol + 1).setValue(statusText);
    break;
  }
}
  var isApproved = action === "approve";

  var statusText = isApproved ? "Approved" : "Denied";
  var actionText = isApproved ? "approved ✅" : "denied ❌";

// Update sheet (with duplicate protection)
for (var i = 1; i < data.length; i++) {
  if (data[i][timestampCol] == timestamp) {

    var currentStatus = data[i][statusCol];

    // 🚫 Prevent double processing
    if (currentStatus === "Approved" || currentStatus === "Denied") {
      return HtmlService.createHtmlOutput(`
        <div style="font-family:Arial; text-align:center; padding:40px;">
          <h2>⚠️ Already Processed</h2>
          <p>This request has already been <strong>${currentStatus}</strong>.</p>
        </div>
      `);
    }

    //  Update status
    sheet.getRange(i + 1, statusCol + 1).setValue(statusText);
    break;
  }
}

////        SEND EMAIL TO MISSIONARY        \\\\
  var subject = "Temple Visit Request Update";

  var message = `
  <div style="font-family:Arial, sans-serif; background-color:#f4f6f8; padding:30px;">
    <div style="max-width:600px; margin:auto; background:#ffffff; padding:30px; border-radius:10px; border:1px solid #ddd; text-align:center;">
      
      <h2 style="color:${isApproved ? '#2E7D32' : '#C62828'}; margin-top:0;">
        Temple Visit Request Update
      </h2>

      <p style="font-size:15px; color:#333;">
        Hello Missionaries,
      </p>

      <p style="font-size:15px; color:#333;">
        Your temple visit request submitted on<br>
        <strong>${formattedDate}</strong>
      </p>

      <p style="font-size:16px; margin:20px 0;">
        has been
      </p>

      <h3 style="color:${isApproved ? '#2E7D32' : '#C62828'}; margin:10px 0;">
        ${isApproved ? 'Approved ✅' : 'Denied ❌'}
      </h3>

      <p style="margin-top:20px; font-size:14px; color:#555;">
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

  //  confirmation page
  return HtmlService.createHtmlOutput(`
  <!DOCTYPE html>
  <html>
  <head>
    <title>Temple Request</title>
  </head>
  <body style="margin:0; font-family:Arial, sans-serif; background:#f4f6f8;">

    <div style="display:flex; align-items:center; justify-content:center; height:100vh;">
      
      <div style="background:#fff; padding:40px; border-radius:12px; 
                  box-shadow:0 4px 12px rgba(0,0,0,0.1); text-align:center; max-width:420px;">
        
        <div style="font-size:50px;">
          ${isApproved ? '✅' : '❌'}
        </div>

        <h2 style="color:${isApproved ? '#2E7D32' : '#C62828'}; margin-top:10px;">
          ${isApproved ? 'Request Approved' : 'Request Denied'}
        </h2>

        <p style="font-size:15px; color:#444;">
          The temple visit request has been successfully
          <strong>${isApproved ? 'approved' : 'denied'}</strong>.
        </p>

        <p style="margin-top:20px; font-size:13px; color:#888;">
          Thank you President, you may now close this window.
        </p>

      </div>

    </div>

  </body>
  </html>
  `);
}