// ===============================
// BotBackend.gs
// הוסף קובץ זה לפרויקט GAS הקיים
// ואז מחק / הפוך ל-comment את doPost הישן ב-Code.gs
// ===============================

const BOT_CONFIG = {
  SCRIPTS_SHEET: 'סקריפטים',
  TECH_PASSWORD: '06536368',
  EMAIL: 'sos@1979.co.il'
};

// ── SETUP (הרץ פעם אחת) ─────────────────────────
function setupBotSheets() {
  const ss = SpreadsheetApp.openById(getSheetId());
  setupScriptsSheet(ss);
  Logger.log('✅ גיליון סקריפטים הוגדר!');
}

function setupScriptsSheet(ss) {
  let sheet = ss.getSheetByName(BOT_CONFIG.SCRIPTS_SHEET);
  if (!sheet) sheet = ss.insertSheet(BOT_CONFIG.SCRIPTS_SHEET);
  sheet.clear();

  const headers = ['ID','שם','תיאור','תוכן','קטגוריה','פומבי','מילות_מפתח','תאריך'];
  sheet.getRange(1,1,1,headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');

  // דוגמאות
  const samples = [
    ['1','תיקון תור הדפסה','מנקה תור הדפסה תקוע',
     'Stop-Service Spooler\nRemove-Item "C:\\Windows\\System32\\spool\\PRINTERS\\*" -Force\nStart-Service Spooler',
     'מדפסות','TRUE','מדפסת,תור,spooler,הדפסה',new Date()],
    ['2','איפוס DNS ורשת','מאפס הגדרות DNS ומחדש IP',
     'ipconfig /flushdns\nipconfig /release\nipconfig /renew\nnetsh winsock reset',
     'רשת','TRUE','dns,אינטרנט,רשת,חיבור,ip',new Date()],
    ['3','הסרת תוכנה בכוח','מסיר תוכנה שלא נמחקת',
     '$app = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*שם_תוכנה*" }\n$app.Uninstall()',
     'התקנות','FALSE','הסרה,תוכנה,uninstall,מחיקה',new Date()]
  ];
  sheet.getRange(2,1,samples.length,8).setValues(samples);
}

// ── CRUD סקריפטים ────────────────────────────────
function getScripts(userType) {
  const ss = SpreadsheetApp.openById(getSheetId());
  const sheet = ss.getSheetByName(BOT_CONFIG.SCRIPTS_SHEET);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // skip header
  const result = [];

  for (const r of rows) {
    if (!r[0]) continue; // skip empty rows
    const isPublic = r[5] === true || String(r[5]).toUpperCase() === 'TRUE';
    if (userType !== 'technician' && !isPublic) continue;
    result.push({
      id: String(r[0]),
      name: r[1],
      description: r[2],
      content: r[3],
      category: r[4],
      isPublic: isPublic,
      keywords: String(r[6]).split(',').map(k => k.trim())
    });
  }
  return result;
}

function addScript(data) {
  const ss = SpreadsheetApp.openById(getSheetId());
  let sheet = ss.getSheetByName(BOT_CONFIG.SCRIPTS_SHEET);
  if (!sheet) { setupScriptsSheet(ss); sheet = ss.getSheetByName(BOT_CONFIG.SCRIPTS_SHEET); }

  const id = String(Date.now());
  sheet.appendRow([
    id, data.name, data.description, data.content,
    data.category || 'כללי',
    data.isPublic ? 'TRUE' : 'FALSE',
    data.keywords || '',
    new Date()
  ]);
  return id;
}

function togglePublicScript(scriptId, isPublic) {
  const ss = SpreadsheetApp.openById(getSheetId());
  const sheet = ss.getSheetByName(BOT_CONFIG.SCRIPTS_SHEET);
  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(scriptId)) {
      sheet.getRange(i+1, 6).setValue(isPublic ? 'TRUE' : 'FALSE');
      return true;
    }
  }
  return false;
}

function deleteScript(scriptId) {
  const ss = SpreadsheetApp.openById(getSheetId());
  const sheet = ss.getSheetByName(BOT_CONFIG.SCRIPTS_SHEET);
  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(scriptId)) {
      sheet.deleteRow(i+1);
      return true;
    }
  }
  return false;
}

// ── חיפוש סקריפט לפי בעיה ───────────────────────
function searchScripts(problem, userType) {
  const scripts = getScripts(userType);
  const q = problem.toLowerCase();
  let best = null, bestScore = 0;

  for (const s of scripts) {
    let score = 0;
    s.keywords.forEach(k => {
      if (k && (q.includes(k.toLowerCase()) || k.toLowerCase().includes(q.split(' ')[0]))) score += 2;
    });
    const words = q.split(/\s+/);
    words.forEach(w => {
      if (w.length > 2 && (s.name.includes(w) || (s.description||'').includes(w))) score++;
    });
    if (score > bestScore) { bestScore = score; best = s; }
  }

  return bestScore >= 1 ? best : null;
}

// ── מייל ─────────────────────────────────────────
function sendEmailSummary(conversation, userType) {
  const now = new Date().toLocaleString('he-IL');
  let body = `סיכום שיחת IT - ${now}\n`;
  body += `סוג משתמש: ${userType === 'technician' ? 'טכנאי' : 'לקוח'}\n`;
  body += '='.repeat(40) + '\n\n';

  (conversation || []).forEach(m => {
    if (m.role === 'user') body += `👤 שאלה: ${m.txt}\n\n`;
    else {
      body += `🤖 פתרון: ${m.txt}\n`;
      if (m.script) body += `\n📋 סקריפט:\n${m.script}\n`;
      body += '\n';
    }
  });

  GmailApp.sendEmail(BOT_CONFIG.EMAIL,
    `סיכום שיחת IT - ${now}`, body);
}

// ── היסטוריה ─────────────────────────────────────
function getHistory() {
  const ss = SpreadsheetApp.openById(getSheetId());
  const sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  return data.slice(1).reverse().slice(0, 50).map(r => ({
    date: r[0], time: r[1], problem: r[2], source: r[3]
  }));
}

// ── ROUTER ראשי ──────────────────────────────────
function handleBotRequest(data) {
  const action = data.action || '';
  const userType = data.userType || 'client';
  const isTech = data.password === BOT_CONFIG.TECH_PASSWORD || userType === 'technician';

  try {
    if (action === 'chat') {
      // חפש קודם בסקריפטים המקומיים
      const localScript = searchScripts(data.problem || '', userType);
      if (localScript) {
        logQuery(data.problem, localScript.name, 'מאגר סקריפטים');
        return {
          success: true, source: 'מאגר פנימי',
          solution: `פתרון: ${localScript.name}\n${localScript.description}`,
          script: localScript.content
        };
      }
      // אחרת - Claude
      const result = handleSupportRequest(data.problem);
      return result;
    }

    if (action === 'getScripts') {
      return { success: true, scripts: getScripts(isTech ? 'technician' : 'client') };
    }

    if (action === 'addScript') {
      if (!isTech) return { success: false, error: 'אין הרשאה' };
      const id = addScript(data);
      return { success: true, id };
    }

    if (action === 'togglePublic') {
      if (!isTech) return { success: false, error: 'אין הרשאה' };
      togglePublicScript(data.scriptId, data.isPublic);
      return { success: true };
    }

    if (action === 'deleteScript') {
      if (!isTech) return { success: false, error: 'אין הרשאה' };
      deleteScript(data.scriptId);
      return { success: true };
    }

    if (action === 'sendEmail') {
      sendEmailSummary(data.conversation, userType);
      return { success: true };
    }

    if (action === 'getHistory') {
      if (!isTech) return { success: false, error: 'אין הרשאה' };
      return { success: true, history: getHistory() };
    }

    return { success: false, error: 'פעולה לא מוכרת: ' + action };

  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ── doPost מעודכן (החלף את הישן ב-Code.gs) ───────
function doPost(e) {
  const headers = { 'Access-Control-Allow-Origin': '*' };
  try {
    const data = JSON.parse(e.postData.contents);
    let result;

    // בקשות בוט חדשות
    if (data.action) {
      result = handleBotRequest(data);
    } else {
      // תאימות לאחור - בקשות ישנות
      const problem = data.problem || data.description || '';
      if (!problem) {
        result = { success: false, error: 'חסר תיאור בעיה' };
      } else {
        result = handleSupportRequest(problem);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
