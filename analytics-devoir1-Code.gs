/**
 * MOUKRIM MATHS ANALYTICS — Devoir N°1 — 3APIC (37 questions)
 * À COLLER dans Extensions > Apps Script du classeur Google Sheets concerné.
 * Ne pas coller dans le projet Apps Script de l'ancien quiz.
 * Le script associé au classeur utilise SpreadsheetApp.getActiveSpreadsheet().
 * Publier : Deploy > New deployment > Web app.
 * Exécuter en tant que : Moi. Accès : Tout le monde (nécessaire pour un quiz public).
 * IMPORTANT : Les requêtes publiques peuvent être envoyées par des tiers.
 * Ne jamais stocker mots de passe ni données sensibles des élèves.
 */
const CONFIG = Object.freeze({
  quiz: 'devoir-1-3apic-v1',
  total: 37,
  keys: 'ABCABCABCABCABCABCABCABCABCABCABCABCA',
  resultsSheet: 'Résultats',
  visitsSheet: 'Visites'
});

function doGet() {
  return ContentService.createTextOutput('Moukrim Maths Analytics — service actif');
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw || raw.length > 12000) throw Error('Invalid payload');
    const data = JSON.parse(raw);
    if (data.quiz !== CONFIG.quiz || !['visit', 'finish'].includes(data.event)) throw Error('Invalid event');
    const sid = String(data.sid || '').trim().slice(0, 100);
    if (!/^[a-zA-Z0-9_-]{12,100}$/.test(sid)) throw Error('Invalid session');
    const name = safeCell_(String(data.name || '').trim().slice(0, 60) || 'Anonyme');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw Error('No linked spreadsheet');
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      if (data.event === 'visit') {
        const sheet = requiredSheet_(ss, CONFIG.visitsSheet);
        const ids = sheet.getLastRow() > 1
          ? sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat()
          : [];
        if (ids.includes(sid)) return response_('already-recorded');
        sheet.appendRow([new Date(), sid, name]);
      } else {
        const answers = data.answers;
        if (!Array.isArray(answers) || answers.length !== CONFIG.total) throw Error('Invalid answers');
        const normalized = answers.map(a => {
          if (typeof a === 'number' && Number.isInteger(a) && a >= 0 && a <= 2) return 'ABC'[a];
          if (typeof a === 'string' && /^[ABC]$/.test(a)) return a;
          throw Error('Invalid answer choice');
        });
        const score = normalized.filter((a, i) => a === CONFIG.keys[i]).length;
        const sheet = requiredSheet_(ss, CONFIG.resultsSheet);
        const ids = sheet.getLastRow() > 1
          ? sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat()
          : [];
        if (ids.includes(sid)) return response_('already-recorded');
        sheet.appendRow([new Date(), sid, name, score, CONFIG.total, score / CONFIG.total, ...normalized]);
      }
    } finally {
      lock.releaseLock();
    }
    return response_('ok');
  } catch (err) {
    console.error(err);
    return response_('error');
  }
}

function requiredSheet_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) throw Error('Missing sheet: ' + name);
  return sh;
}

function safeCell_(value) {
  return /^[=+@\-\t\r]/.test(value) ? "'" + value : value;
}

function response_(value) {
  return ContentService.createTextOutput(value);
}
