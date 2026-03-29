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


////        DO GET        \\\\
function sendEmailOnFormSubmit(e) {
  var responses = e.namedValues;

  var missionaryEmail = responses["Email Address"][0];
  var presidentEmail = "500464809@missionary.org";
  var timestamp = responses["Timestamp"][0];

  var formattedDate = Utilities.formatDate(
  new Date(timestamp),
  Session.getScriptTimeZone(),
  "MMMM d, yyyy 'at' h:mm a"
);

var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
var data = sheet.getDataRange().getValues();
var headers = data[0];

var missionaryName = "Missionary"; 

for (var i = 1; i < data.length; i++) {
  if (data[i][timestampCol] == timestamp) {
    missionaryName = data[i][nameCol] || "Missionary"; // grab the name from that row
    sheet.getRange(i + 1, statusCol + 1).setValue(statusText);
    break;
  }
}

var timestampCol = headers.indexOf("Timestamp");
var statusCol = headers.indexOf("Status");
var nameCol = headers.indexOf("Names of missionaries: ");  




  var MISSION_NAME = "Philippines Manila Mission";
  var PRIMARY_COLOR = "#1B5E20"; // deep green

  var subject = "Temple Visit Request";

  var formUrl = "https://docs.google.com/forms/d/e/1FAIpQLSfjiRnQNweAHFHlRJ9pklTdcLKPwSAoA_YtEVU2PrEAkfmv6g/viewform?usp=header";

  var webAppUrl = "https://script.google.com/a/macros/missionary.org/s/AKfycbwwiPSCi-8yN3ITgVC0RzxD1Qt_fqAkzA5CgJDcp5JPKZ3yfVV9hpqVJADAKTHJ5Eqm/exec";

  var message = `
  <div style="font-family:Arial, sans-serif; background-color:#f4f6f8; padding:30px;">
  <div style="max-width:650px; margin:auto; background:white; padding:25px; border-radius:10px; border:1px solid #ddd;">
    
    <!-- TOP COLOR BAR -->
    <div style="height:6px; background:${PRIMARY_COLOR}; border-radius:10px 10px 0 0; margin:-25px -25px 20px -25px;"></div>
    
    <!-- BRANDING -->
    <div style="text-align:center; margin-bottom:20px;">
      <div style="font-size:20px; font-weight:bold; color:${PRIMARY_COLOR};">
         ${MISSION_NAME}
      </div>
      <div style="font-size:14px; color:#555; margin-top:4px;">
        Temple Visit Request System
      </div>
    </div>

    <h2 style="color:${PRIMARY_COLOR}; margin-top:0; text-align:center;">
      New Request Submitted
    </h2>
      
      <p>President Fernandez,</p>

      <p>
        A new temple visit request has been submitted and is awaiting your review.
        Please review the details below and select an action.
      </p>

      <table style="width:100%; border-collapse:collapse; margin-top:20px; font-size:14px;">
        <tr style="background-color:#1B5E20; color:white;">
          <th style="text-align:left; padding:10px;">Question</th>
          <th style="text-align:left; padding:10px;">Response</th>
        </tr>
  `;

var rowIndex = 0;

for (var question in responses) {
  var answer = responses[question].toString().trim();
  question = question.replace(/:+$/, '').trim();

  // Alternate row colors
  var rowColor = rowIndex % 2 === 0 ? "#f9f9f9" : "#ffffff";

  message += `
    <tr style="background:${rowColor}; border-radius:8px;">
      <td style="padding:12px; font-weight:bold; border-radius:8px 0 0 8px;">
        ${question}
      </td>
      <td style="padding:12px; border-radius:0 8px 8px 0; background:#eef6ff;">
        ${answer}
      </td>
    </tr>
  `;
  rowIndex++;
}

  message += `</table>`;

  message += `
    <div style="margin-top:25px; text-align:center;">
      <a href="${webAppUrl}?action=approve&timestamp=${encodeURIComponent(timestamp)}&missionary=${encodeURIComponent(missionaryEmail)}"
         style="display:inline-block; background:#2E7D32; color:white; padding:12px 20px; text-decoration:none; border-radius:6px; font-weight:bold; margin-right:10px;">
         ✅ Approve
      </a>

      <a href="${webAppUrl}?action=deny&timestamp=${encodeURIComponent(timestamp)}&missionary=${encodeURIComponent(missionaryEmail)}"
         style="display:inline-block; background:#C62828; color:white; padding:12px 20px; text-decoration:none; border-radius:6px; font-weight:bold;">
         ❌ Deny
      </a>
    </div>

    <div style="margin-top:20px; text-align:center;">
      <a href="${formUrl}" 
         style="display:inline-block; background:#455A64; color:white; padding:10px 16px; text-decoration:none; border-radius:6px;">
        View Form
      </a>
    </div>

    <p style="margin-top:25px; font-size:12px; color:#777; text-align:center;">
      Submitted on ${timestamp}
    </p>

    </div>
  </div>
  `;

  MailApp.sendEmail({
    to: presidentEmail,
    subject: subject,
    htmlBody: message
  });
}