/**
 * Hochzeit Kim & Pere: Backend für hochzeit.pere-dalic.ch
 * Tabs: Zusagen (Antworten der Gäste), Inhalte (Texte der Seite), Einstellungen
 * Fotos landen im Drive-Ordner "Hochzeit Kim & Pere Fotos".
 * Einmalig ausführen: setup()
 */
const SHEET_RSVP = 'Zusagen';
const SHEET_TEXT = 'Inhalte';
const SHEET_SET = 'Einstellungen';
const SHEET_GUESTS = 'Gäste';
const FOLDER_NAME = 'Hochzeit Kim & Pere Fotos';
const SHEET_ID = '1eDePisfs74wirlVFGT2e-g4F7Z92GXjxmaEejEyz0Us';
function book() { return SpreadsheetApp.openById(SHEET_ID); }

function setup() {
  const ss = book();
  const bold = sh => sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');

  const r = ss.getSheetByName(SHEET_RSVP) || ss.insertSheet(SHEET_RSVP);
  if (r.getLastRow() === 0) {
    r.appendRow(['Zeit', 'Name', 'Kommt', 'Personen', 'Erwachsene', 'Kinder', 'Essenswünsche', 'Song', 'Nachricht', 'Sprache']);
    r.setFrozenRows(1); bold(r);
  }

  const t = ss.getSheetByName(SHEET_TEXT) || ss.insertSheet(SHEET_TEXT);
  if (t.getLastRow() === 0) {
    t.appendRow(['Schlüssel', 'Deutsch', 'Kroatisch', 'Tschechisch']);
    t.setFrozenRows(1); bold(t);
    t.getRange('B:D').setNumberFormat('@').setWrap(true);
    t.setColumnWidth(1, 140); t.setColumnWidths(2, 3, 380);
  }

  const s = ss.getSheetByName(SHEET_SET) || ss.insertSheet(SHEET_SET);
  if (s.getLastRow() === 0) {
    s.getRange('B:B').setNumberFormat('@');
    s.getRange(1, 1, 4, 3).setValues([
      ['Einstellung', 'Wert', 'Hinweis'],
      ['album_url', '', 'Link zu einem geteilten Fotoalbum (optional)'],
      ['whatsapp', '41793897169', 'Nummer für WhatsApp-Links, international ohne +'],
      ['upload', 'ja', 'ja = Gäste dürfen Fotos hochladen, nein = Upload gesperrt']
    ]);
    s.setFrozenRows(1); bold(s);
    s.setColumnWidth(1, 120); s.setColumnWidth(2, 260); s.setColumnWidth(3, 380);
  }

  const first = ss.getSheets()[0];
  if (![SHEET_RSVP, SHEET_TEXT, SHEET_SET].includes(first.getName()) && first.getLastRow() === 0) ss.deleteSheet(first);

  const props = PropertiesService.getScriptProperties();
  let fid = props.getProperty('FOLDER_ID');
  if (!fid) {
    const f = DriveApp.createFolder(FOLDER_NAME);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fid = f.getId();
    props.setProperty('FOLDER_ID', fid);
  }
  Logger.log('Fotoordner: https://drive.google.com/drive/folders/' + fid);
}

function doGet(e) {
  const a = (e && e.parameter && e.parameter.action) || '';
  try {
    if (a === 'photos') return json({ ok: true, photos: listPhotos() });
    if (a === 'content') return json({ ok: true, texts: readTexts(), settings: readSettings() });
  } catch (err) {
    return json({ ok: false, error: 'server' });
  }
  return json({ ok: false, error: 'unknown' });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'bad_json' }); }
  try {
    if (d.action === 'rsvp') return json(saveRsvp(d));
    if (d.action === 'upload') return json(saveUpload(d));
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: 'server', detail: String((err && err.message) || err).slice(0, 300) });
  }
  return json({ ok: false, error: 'unknown' });
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function clip(v, n) { return String(v == null ? '' : v).slice(0, n).trim(); }
// Verhindert, dass Eingaben als Formel ausgeführt werden
function cell(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }

function saveRsvp(d) {
  const name = clip(d.name, 600);
  if (!name) return { ok: false, error: 'name' };
  const yes = d.yes === true;
  const n = yes ? Math.max(1, Math.min(20, parseInt(d.n, 10) || 1)) : 0;
  const a = yes ? Math.max(0, Math.min(20, parseInt(d.a, 10) || 0)) : 0;
  const k = yes ? Math.max(0, Math.min(20, parseInt(d.k, 10) || 0)) : 0;
  const now = new Date();
  const guests = (Array.isArray(d.guests) ? d.guests : []).slice(0, 20).map(g => [
    now, cell(clip(g && g.f, 80)), cell(clip(g && g.l, 80)), g && g.kid ? 'Kind' : 'Erwachsen',
    yes ? 'Ja' : 'Nein', cell(clip(g && g.diet, 300))
  ]);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = book();
    const r = ss.getSheetByName(SHEET_RSVP);
    rsvpColumns(r);
    r.appendRow([
      now, cell(name), yes ? 'Ja' : 'Nein', n, a, k,
      cell(clip(d.food, 1500)), cell(clip(d.song, 200)), cell(clip(d.msg, 1000)), clip(d.lang, 5)
    ]);
    if (guests.length) {
      const g = guestSheet(ss);
      g.getRange(g.getLastRow() + 1, 1, guests.length, 6).setValues(guests);
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

// Ergänzt die Spalten Erwachsene und Kinder, falls sie noch fehlen
function rsvpColumns(sh) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (head.indexOf('Erwachsene') >= 0) return;
  sh.insertColumnsAfter(4, 2);
  sh.getRange(1, 5, 1, 2).setValues([['Erwachsene', 'Kinder']]).setFontWeight('bold');
}

// Eine Zeile pro Gast, praktisch für Sitzplan und Küche
function guestSheet(ss) {
  let g = ss.getSheetByName(SHEET_GUESTS);
  if (g) return g;
  g = ss.insertSheet(SHEET_GUESTS, 1);
  g.getRange(1, 1, 1, 6).setValues([['Zeit', 'Vorname', 'Nachname', 'Erwachsen/Kind', 'Kommt', 'Essenswunsch']]).setFontWeight('bold');
  g.setFrozenRows(1);
  g.setColumnWidths(2, 2, 140); g.setColumnWidth(6, 260);
  return g;
}

function saveUpload(d) {
  if (String(readSettings().upload || 'ja').toLowerCase() === 'nein') return { ok: false, error: 'closed' };
  const mime = String(d.mime || '');
  if (!/^image\/(jpeg|png|webp|gif)$/.test(mime)) return { ok: false, error: 'type' };
  const b64 = String(d.data || '');
  if (!b64 || b64.length > 20000000) return { ok: false, error: 'size' };
  const ext = mime === 'image/jpeg' ? 'jpg' : mime.split('/')[1];
  const stamp = Utilities.formatDate(new Date(), 'Europe/Zurich', 'yyyyMMdd-HHmmss');
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, 'foto-' + stamp + '-' + Math.random().toString(36).slice(2, 6) + '.' + ext);
  const f = folder().createFile(blob);
  // Freigabe kommt vom Ordner (Jeder mit Link), einzelne Dateien muessen nicht extra freigegeben werden
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  CacheService.getScriptCache().remove('photos');
  return { ok: true, id: f.getId() };
}

function folder() {
  return DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
}

function listPhotos() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('photos');
  if (hit) return JSON.parse(hit);
  const out = [];
  const it = folder().getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.isTrashed() || !/^image\//.test(f.getMimeType())) continue;
    out.push({ id: f.getId(), t: f.getDateCreated().getTime() });
  }
  out.sort((a, b) => b.t - a.t);
  const s = JSON.stringify(out);
  if (s.length < 90000) cache.put('photos', s, 60);
  return out;
}

function readTexts() {
  const sh = book().getSheetByName(SHEET_TEXT);
  const out = { de: {}, hr: {}, cs: {} };
  if (!sh || sh.getLastRow() < 2) return out;
  const langs = ['de', 'hr', 'cs'];
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues().forEach(row => {
    const k = String(row[0]).trim();
    if (!k) return;
    langs.forEach((l, i) => { if (row[i + 1] !== '') out[l][k] = row[i + 1]; });
  });
  return out;
}

function readSettings() {
  const sh = book().getSheetByName(SHEET_SET);
  const o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(row => {
    const k = String(row[0]).trim();
    if (k) o[k] = String(row[1]).trim();
  });
  return o;
}
