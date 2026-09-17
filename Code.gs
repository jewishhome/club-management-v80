/**
 * 👑 ניהול קורסים ומועדון V80 - Data Driven & Personalized Drive
 */

const scriptProperties = PropertiesService.getScriptProperties();

function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) throw new Error("לשונית Settings לא נמצאה!");
  
  const data = sheet.getDataRange().getValues();
  const settings = {};
  
  for (let i = 1; i < data.length; i++) {
    const [pName, folderId, pulseemId, yemotId, type] = data[i];
    if (pName) {
      settings[String(pName).trim()] = {
        folderId: String(folderId).trim(),
        pulseemId: String(pulseemId).trim(),
        yemotId: String(yemotId).trim(),
        type: String(type).trim()
      };
    }
  }
  return settings;
}

function processUserRow(rowIdx) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("מנויים");
  const settings = getSettings();
  
  const rowData = sheet.getRange(rowIdx, 1, 1, 10).getValues()[0];
  const status = String(rowData[0]).trim(); // עמודה A: סטטוס
  const name = String(rowData[2]).trim();   // עמודה C: שם
  const phone = String(rowData[3]).trim();  // עמודה D: טלפון
  const email = String(rowData[4]).trim().toLowerCase(); // עמודה E: מייל
  const purchasedItems = String(rowData[5] || "").split(",").map(s => s.trim()); // עמודה F: מוצרים לרכישה
  let personalFolderId = String(rowData[6] || "").trim(); // עמודה G: תיקייה אישית

  if (!email || !isValidEmail(email)) return;

  if (status === "פעיל") {
    // 1. יצירת/איתור תיקייה אישית
    let userFolder;
    if (!personalFolderId) {
      userFolder = DriveApp.createFolder(`התכנים של ${name}`);
      userFolder.addViewer(email);
      personalFolderId = userFolder.getId();
      sheet.getRange(rowIdx, 7).setValue(personalFolderId);
    } else {
      userFolder = DriveApp.getFolderById(personalFolderId);
    }

    // 2. עדכון קיצורי דרך בתיקייה האישית
    purchasedItems.forEach(item => {
      if (settings[item] && settings[item].folderId) {
        const targetFolder = DriveApp.getFolderById(settings[item].folderId);
        createShortcutIfNotExists(userFolder, targetFolder, item);
      }
    });

    // 3. עדכון יומן
    sheet.getRange(rowIdx, 2).setValue("✅ סונכרן " + Utilities.formatDate(new Date(), "GMT+3", "dd/MM/yy HH:mm"));
  } else {
    // במידה ולא פעיל - הסרת הרשאות מהתיקייה האישית
    if (personalFolderId) {
      try {
        const userFolder = DriveApp.getFolderById(personalFolderId);
        userFolder.removeViewer(email);
      } catch(e) { console.warn(e); }
    }
    sheet.getRange(rowIdx, 2).setValue("⛔ הוסר " + Utilities.formatDate(new Date(), "GMT+3", "dd/MM/yy HH:mm"));
  }
}

function createShortcutIfNotExists(parentFolder, targetFolder, name) {
  const existing = parentFolder.getFiles();
  while (existing.hasNext()) {
    const file = existing.next();
    if (file.getName() === name) return;
  }
  parentFolder.createShortcut(targetFolder.getId()).setName(name);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
}