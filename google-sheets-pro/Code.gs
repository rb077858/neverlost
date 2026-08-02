/**
 * NeverLost — Pro users feed
 *
 * מגדיר את המנויים ב-Pro (כולל "אינסוף") דרך גיליון Google Sheets, במקום קובץ pro.txt.
 * ראה SETUP.md לפריסה מלאה.
 *
 * מבנה הגיליון (שורה ראשונה = כותרות, מתעלמים ממנה):
 *   A: Email                     — כתובת המייל של המשתמש (חייבת להתאים למייל שמחוברים איתו ל-Google)
 *   B: Limit                     — מספר (כמה תגים מותר), או "+" לכמות בלתי מוגבלת
 *   C: Expiry (אופציונלי)        — תאריך תפוגה בפורמט DD/MM/YYYY. תא ריק = בלי תפוגה בכלל (לעולם לא פג).
 *                                   אפשר גם להקליד מספר בלבד (למשל 30) — הוא יתחלף אוטומטית,
 *                                   כעבור כ-5 שניות, בתאריך שהוא כ-30 יום מהיום (ראו onEdit למטה).
 *   D: Days left (אופציונלי)     — עמודת נוסחה שמתעדכנת לבד, ראו הסבר ב-SETUP.md.
 *
 * אפשר להוסיף כמה שורות לאותו מייל — הכמויות (Limit) מצטברות, ואם אחת מהשורות היא "+" המשתמש יקבל אינסוף.
 */

const EXPIRY_COLUMN = 3; // עמודה C

/**
 * הופך ערך גולמי מהתא לאובייקט Date, או null אם אין תאריך בכלל (=לעולם לא פג).
 * לא מסתמכים על new Date(string) הרגיל של JS כי הוא מפרש "01/02/1990" כ-MM/DD
 * (פורמט אמריקאי) — וזה הפוך מהפורמט DD/MM/YYYY שאנחנו רוצים.
 */
function parseExpiryDate(raw) {
  if (raw === "" || raw === null || raw === undefined) return null;
  if (raw instanceof Date) return raw;

  const str = String(raw).trim();
  if (!str) return null;

  const match = str.match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})$/);
  if (match) {
    const [, d, m, y] = match;
    return new Date(Number(y), Number(m) - 1, Number(d));
  }

  return null; // ערך לא מזוהה כתאריך — מתייחסים אליו כ"בלי תפוגה" ולא חוסמים בטעות
}

function doGet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const rows = sheet.getDataRange().getValues();

  const result = {};
  const today = new Date();

  for (let i = 1; i < rows.length; i++) { // מדלגים על שורת הכותרות
    const email = String(rows[i][0] || "").trim().toLowerCase();
    const limitRaw = String(rows[i][1] || "").trim();
    const expDate = parseExpiryDate(rows[i][EXPIRY_COLUMN - 1]);

    if (!email) continue;
    if (expDate && expDate < today) continue; // מנוי שפג תוקפו — מתעלמים משורה זו. בלי תאריך = לעולם לא פג.

    if (!result[email]) result[email] = { unlimited: false, limit: 0 };

    if (limitRaw === "+") {
      result[email].unlimited = true;
    } else {
      result[email].limit += parseInt(limitRaw) || 0;
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Trigger פשוט שרץ אוטומטית כשעורכים תא בגיליון.
 * אם עורכים את עמודת Expiry (C) ומזינים מספר בלבד (למשל 30, בלי תאריך) —
 * ממתין 5 שניות ואז מחליף את התא בתאריך אמיתי שהוא היום + אותו מספר ימים,
 * מפורמט כ-DD/MM/YYYY.
 */
function onEdit(e) {
  const range = e.range;
  if (range.getColumn() !== EXPIRY_COLUMN || range.getRow() === 1) return;

  const value = e.value;
  if (value === undefined) return; // תא נמחק — משאירים ריק (=בלי תפוגה)

  const isPlainNumber = /^\d+$/.test(String(value).trim());
  if (!isPlainNumber) return; // כבר תאריך, או טקסט אחר — לא נוגעים

  const days = parseInt(value, 10);

  Utilities.sleep(5000);

  // אם המשתמש הספיק לשנות את התא שוב תוך כדי ההמתנה, לא דורסים את מה שהוא הזין
  if (String(range.getValue()).trim() !== String(value).trim()) return;

  const newDate = new Date();
  newDate.setDate(newDate.getDate() + days);

  range.setValue(newDate);
  range.setNumberFormat("dd/mm/yyyy");
}
