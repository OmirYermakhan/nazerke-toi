/**
 * Қонақтардың жауабын Google кестесіне жинау.
 *
 * Орнату:
 * 1. sheets.new ашып, жаңа кесте жасаңыз.
 * 2. Кестеде: Extensions → Apps Script. Ашылған терезедегі кодты өшіріп, осы файлды түгел қойыңыз, сақтаңыз.
 * 3. Deploy → New deployment → түрі: Web app.
 *    Execute as: Me.  Who has access: Anyone.  Deploy басып, рұқсат беріңіз.
 * 4. Шыққан Web app URL сілтемесін (…/exec деп бітеді) index.html ішіндегі TOI.sheetUrl жолына қойыңыз.
 *
 * Кодты өзгертсеңіз: Deploy → Manage deployments → Edit → Version: New version.
 */
const HEAD = ["Уақыты", "Аты-жөні", "Жауабы", "Ересек", "Бала", "Шақырылған есім", "id"];

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getRange(1, 1).getValue() !== HEAD[0]) {
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight("bold");
    sh.getRange("I1:N1").setValues([["Ересек:", "=SUM(D2:D)", "Бала:", "=SUM(E2:E)", "Барлығы:", "=J1+L1"]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const sh = sheet_();
    const yes = d.come === "yes";
    const row = [new Date(), String(d.name || "").slice(0, 100), yes ? "Келеді" : "Келмейді",
      yes ? Math.min(20, Number(d.adults) || 1) : 0, yes ? Math.min(20, Number(d.kids) || 0) : 0,
      String(d.invited || "").slice(0, 100), String(d.id || "").slice(0, 40)];
    // Қонақ жауабын өзгертсе, жаңа жол қосылмайды — ескісі жаңарады
    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, HEAD.length, last - 1, 1).getValues().map(r => r[0]) : [];
    const at = row[6] ? ids.indexOf(row[6]) : -1;
    if (at >= 0) sh.getRange(at + 2, 1, 1, row.length).setValues([row]);
    else sh.getRange(Math.max(last, 1) + 1, 1, 1, row.length).setValues([row]);
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  sheet_();
  return ContentService.createTextOutput("Жұмыс істеп тұр");
}
