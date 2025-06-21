// MAIN SCRIPT SEND BATCH EMAILS 
function sendBatchEmails() { 
  Logger.log("⏱️ Trigger started at " + new Date());
  
  const today = new Date(); 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("recofauto");
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");

  const quota = MailApp.getRemainingDailyQuota();
  Logger.log("📧 Quota remaining: " + quota);

  if (quota < 10) {
    Logger.log("🚫 Stopping: Not enough quota left (" + quota + ")");
    settingsSheet.getRange("A15").setValue("⛔ Email batch stopped");
    settingsSheet.getRange("B15").setValue("Quota too low: " + quota + " at " + new Date());
    return;
  }

  let sentCount = parseInt(settingsSheet.getRange("B3").getValue(), 10) || 0;

  const todayKey = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const lastRunKey = settingsSheet.getRange("B2").getValue();

  if (todayKey !== lastRunKey) {
    sentCount = 0;
    settingsSheet.getRange("B3").setValue(0);
    settingsSheet.getRange("B2").setValue(todayKey);
    Logger.log("🔄 New day – daily counter reset.");
  }

  const lastRow = sheet.getLastRow();
  const currentHour = new Date().getHours();
  const validHours = Array.from({ length: 16 }, (_, i) => i + 6);
  if (!validHours.includes(currentHour)) return;

  const docIds = {
    "Template 1": "1SCLFX17i_MmI41XCqqVWhbm6vF3l8ztxxsEC8BVgcq8",
    "Template 2": "1ls7ZF9ySF-FAhWaWwVXyS6p9bvjnrSYD85OSkJebnl8",
    "Template 3": "1yWCS8pd07Hv03liengFkbNfi0ssXZ0apXSHrL-d9leY"
  };

  const headerImages = {
    "Template 1": "https://i.imgur.com/L9ebqgP.png",
    "Template 2": "https://i.imgur.com/3snPbcw.png",
    "Template 3": "https://i.imgur.com/WoIG3Qo.png"
  };

  const footerUrl = "https://i.imgur.com/ztrqm9a.png";

  const subjectLines = {
    "Template 1": "Your Real Estate Career Starts Here 🏡",
    "Template 2": "Need Help Starting in Real Estate? We’ve Got You",
    "Template 3": "Take the First Step—Become a Florida Real Estate Agent"
  };

  const trackingBase = "https://script.google.com/macros/s/AKfycbzdJz7ZNZq_sQW0K8tYcw8mXXK4ZQNTc1Gb0XapO8w8BM2L5BzOu03AOm_1nv-KMJki/exec";
  const batchSize = 50;
  let totalSent = 0;
  const templateBodies = {};
  for (const key in docIds) {
    const html = UrlFetchApp.fetch(`https://docs.google.com/document/d/${docIds[key]}/export?format=html`).getContentText();
    templateBodies[key] = html;
  }

  let currentBatch = parseInt(settingsSheet.getRange("B1").getValue(), 10) || 1;
  sentCount = parseInt(settingsSheet.getRange("B3").getValue(), 10) || 0;
  const templateOptions = ["Template 1", "Template 2", "Template 3"];
  const rawTemplate = settingsSheet.getRange("A4").getValue().toString().trim();
  Logger.log(`🔍 Template pulled from A4: '${rawTemplate}'`);
  const matchedTemplate = templateOptions.find(t => t.toLowerCase() === rawTemplate.toLowerCase());
  const template = matchedTemplate || "Template 1";
  Logger.log(`✅ Final template used: ${template}`);

  const batchColumn = sheet.getRange(2, 5, lastRow - 1).getValues().flat();
  const statusColumn = sheet.getRange(2, 6, lastRow - 1).getValues().flat();
  const dateColumn = sheet.getRange(2, 4, lastRow - 1).getValues().flat();
  const rows = [];

  for (let i = 0; i < batchColumn.length; i++) {
    const sentDate = dateColumn[i];
    const sentToday = sentDate instanceof Date && sentDate.toDateString() === today.toDateString();
    if (batchColumn[i] == currentBatch && !sentToday && statusColumn[i] !== "undeliverable") {
      rows.push(i + 2);
    }
  }

  if (rows.length === 0) {
    const nextBatch = currentBatch >= 11 ? 1 : currentBatch + 1;
    if (nextBatch === 1) {
      const templateIndex = templateOptions.findIndex(t => t === template);
      const nextTemplateIndex = (templateIndex + 1) % templateOptions.length;
      settingsSheet.getRange("A4").setValue(templateOptions[nextTemplateIndex]);
      sheet.getRange(2, 4, lastRow - 1).clearContent();
      sheet.getRange(2, 6, lastRow - 1).clearContent();
      sheet.getRange(2, 7, lastRow - 1).clearContent();
      sheet.getRange(2, 12, lastRow - 1).clearContent();
      sheet.getRange(2, 17, lastRow - 1).clearContent();
    }
    settingsSheet.getRange("B1").setValue(nextBatch);
    settingsSheet.getRange("B3").setValue(0);
    return;
  }

  const startTime = new Date().getTime();
  const maxExecutionTime = 330000;

  if (sentCount >= 2000) {
    const nextBatch = currentBatch >= 11 ? 1 : currentBatch + 1;
    settingsSheet.getRange("B1").setValue(nextBatch);
    settingsSheet.getRange("B3").setValue(0);
    return;
  }

  for (let j = 0; j < Math.min(batchSize, rows.length); j++) {
    if (new Date().getTime() - startTime > maxExecutionTime) break;
    const row = rows[j];
    const email = sheet.getRange(row, 3).getValue();
    if (!isValidEmail(email)) continue;
    const encodedEmail = encodeURIComponent(email);
    const openTrack = `<img src="${trackingBase}?action=open&email=${encodedEmail}" width="1" height="1" style="display:none;" />`;
    const unsubscribeLink = `${trackingBase}?email=${encodedEmail}`;
    const name = sheet.getRange(row, 1).getValue();
    const firstName = (name && typeof name === 'string' && name.trim() !== '' && !name.includes('[') && name !== '-') ? name.trim() : '';
    let bodyHtml = templateBodies[template];
    bodyHtml = bodyHtml.replace(/{{\s*First Name\s*}}/gi, firstName || '');
    const headerUrl = headerImages[template];

    const htmlBody = `
      <html>
        <head>
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <style>
            body {
              margin: 0;
              padding: 0;
              width: 100% !important;
              font-family: Arial, sans-serif;
              background: #ffffff;
            }
            .button {
              background-color: #11056d;
              color: #ffffff !important;
              text-decoration: none;
              padding: 15px 25px;
              font-size: 16px;
              display: inline-block;
              border-radius: 6px;
              margin-top: 20px;
            }
          </style>
        </head>
        <body style="margin:0;padding:0;width:100%;">
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="${headerUrl}" width="100%" style="display:block;" alt="Header" /></td>
            </tr>
            <tr>
              <td style="padding: 20px; font-size: 16px; line-height: 1.6; color: #333333;">
                ${bodyHtml}
                <div style="text-align:center;">
                  <a href="https://www.myrealestatecampus.com/pages/home" class="button">Explore Courses</a>
                </div>
              </td>
            </tr>
            <tr>
              <td><img src="${footerUrl}" width="100%" style="display:block;" alt="Footer" /></td>
            </tr>
            <tr>
              <td style="font-size:12px; text-align:center; color:#999999; padding: 15px;">
                Real Estate Campus of Florida<br>
                295 NW Peacock Blvd #881013<br>
                Port St. Lucie, FL 34986 <br><br>
                If you no longer wish to receive emails from us, 
                <a href="${unsubscribeLink}" style="color:#999999;">unsubscribe here</a>.
              </td>
            </tr>
          </table>
          ${openTrack}
        </body>
      </html>`;

    MailApp.sendEmail({
      to: email,
      subject: subjectLines[template],
      htmlBody: htmlBody,
      from: "info@myrecampus.com",
      replyTo: "info@myrecampus.com",
      name: "Real Estate Campus of Florida"
    });

    sheet.getRange(row, 6).setValue("✅ Sent").setFontColor("#0000FF");
    sheet.getRange(row, 4).setValue(new Date());
    sheet.getRange(row, 7).setValue(template);
    sheet.getRange(row, 12).setValue("");
    sheet.getRange(row, 17).setValue(subjectLines[template]);
    SpreadsheetApp.flush();

    totalSent++;
    sentCount++;
    settingsSheet.getRange("B3").setValue(sentCount);
    if (totalSent % 25 === 0) Logger.log("📨 Sent " + totalSent + " emails so far...");
    if (j % 10 === 0) Utilities.sleep(1000);
  }

  if (sentCount >= 2000 || rows.length <= batchSize) {
    const nextBatch = currentBatch >= 11 ? 1 : currentBatch + 1;
    settingsSheet.getRange('B1').setValue(nextBatch);
    settingsSheet.getRange('B3').setValue(0);
    Logger.log(`🔁 Batch advanced to ${nextBatch}`);
  }
}

function isValidEmail(email) {
  return typeof email === 'string' && email.includes('@') && email.includes('.') && !email.includes('..');
}


function highlightRecofRowAndMove(email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("recofauto");
  const targetSheet = ss.getSheetByName("clickedfollowupauto");

  if (!sourceSheet || !targetSheet) return;

  const lastRow = sourceSheet.getLastRow();
  const emailColumn = sourceSheet.getRange(2, 3, lastRow - 1).getValues().flat(); // Column C

  for (let i = 0; i < emailColumn.length; i++) {
    if ((emailColumn[i] || "").toLowerCase() === email.toLowerCase()) {
      const rowIndex = i + 2;
      const rowData = sourceSheet.getRange(rowIndex, 1, 1, sourceSheet.getLastColumn()).getValues()[0];

      // Highlight in recofauto
      sourceSheet.getRange(rowIndex, 1, 1, sourceSheet.getLastColumn()).setBackground("#fce5cd");

      // Mark as clicked in column N (column 14)
      sourceSheet.getRange(rowIndex, 14).setValue("Yes");

      // Avoid adding duplicate in clickedfollowupauto
      const clickedEmails = targetSheet.getRange(2, 3, targetSheet.getLastRow() - 1).getValues().flat(); // Col C = email
      if (!clickedEmails.includes(email)) {
        targetSheet.appendRow(rowData);
        Logger.log(`✅ Row ${rowIndex} copied to clickedfollowupauto.`);
      } else {
        Logger.log(`⚠️ Email ${email} already exists in clickedfollowupauto.`);
      }

      break;
    }
  }
}

