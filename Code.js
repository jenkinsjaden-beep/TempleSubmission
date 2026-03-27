function sendEmailOnFormSubmit(e) {
  var responses = e.namedValues;

  var missionaryEmail = responses["Email Address"][0];
  var presidentEmail = "500464809@missionary.org";
  var timestamp = responses["Timestamp"][0];

  var subject = "Temple Visit Request";

  var formUrl = "https://docs.google.com/forms/d/e/1FAIpQLSfjiRnQNweAHFHlRJ9pklTdcLKPwSAoA_YtEVU2PrEAkfmv6g/viewform?usp=header";

  var webAppUrl = "https://script.google.com/a/macros/missionary.org/s/AKfycbwwiPSCi-8yN3ITgVC0RzxD1Qt_fqAkzA5CgJDcp5JPKZ3yfVV9hpqVJADAKTHJ5Eqm/exec";

  var message = `
  <div style="font-family:Arial, sans-serif; background-color:#f4f6f8; padding:30px;">
    <div style="max-width:650px; margin:auto; background:white; padding:25px; border-radius:10px; border:1px solid #ddd;">
      
      <h2 style="color:#1B5E20; margin-top:0;">Temple Visit Request</h2>
      
      <p>Hello President,</p>

      <p>
        A new temple visit request has been submitted and is awaiting your review.
        Please review the details below and choose an action.
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
    var rowColor = rowIndex % 2 === 0 ? "#f9f9f9" : "#ffffff";

    message += `
      <tr style="background-color:${rowColor};">
        <td style="padding:10px; border:1px solid #ddd;"><strong>${question}</strong></td>
        <td style="padding:10px; border:1px solid #ddd;">${answer}</td>
      </tr>
    `;
    rowIndex++;
  }

  message += `</table>`;

  message += `
    <div style="margin-top:25px; text-align:center;">
      <a href="${webAppUrl}?action=approve&timestamp=${encodeURIComponent(timestamp)}&missionary=${encodeURIComponent(missionaryEmail)}"
         style="background-color:#2E7D32; color:white; padding:12px 18px; text-decoration:none; border-radius:6px; font-weight:bold; margin-right:10px;">
         Approve
      </a>

      <a href="${webAppUrl}?action=deny&timestamp=${encodeURIComponent(timestamp)}&missionary=${encodeURIComponent(missionaryEmail)}"
         style="background-color:#C62828; color:white; padding:12px 18px; text-decoration:none; border-radius:6px; font-weight:bold;">
         Deny
      </a>
    </div>

    <div style="margin-top:20px; text-align:center;">
      <a href="${formUrl}" 
         style="background-color:#455A64; color:white; padding:10px 16px; text-decoration:none; border-radius:6px;">
        View All Responses
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