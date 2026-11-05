/**
 * Christ Connect Berlin · Anmeldungen von der Link-Seite
 * Schreibt jede Anmeldung von christconnectberlin.github.io als neue Zeile in dieses Sheet.
 *
 * Einrichtung (einmalig):
 * 1. Im Sheet: Erweiterungen → Apps Script → diesen Code einfügen → Speichern.
 * 2. Oben die Funktion "setup" auswählen → Ausführen → Zugriff erlauben.
 * 3. Bereitstellen → Neue Bereitstellung → Typ: Web-App
 *    - Ausführen als: Ich
 *    - Zugriff: Jeder
 *    → Bereitstellen → die Web-App-URL (endet auf /exec) kopieren und an Claude schicken.
 */

const SHEET_NAME = ''; // leer = erstes Tabellenblatt
const HEADERS = ['Datum', 'Name', 'Kontakt', 'Bezirk', 'Interessen', 'Quelle', 'Zuständig', 'Status', 'Notizen'];
const STATUS = ['neu', 'kontaktiert', 'getroffen', 'in Bibelgruppe', 'kein Interesse'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
}

/** Einmal ausführen: Kopfzeile, Status-Auswahl, Formatierung. */
function setup() {
  const sh = sheet_();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUS, true).setAllowInvalid(false).build();
  sh.getRange(2, 8, sh.getMaxRows() - 1, 1).setDataValidation(rule);
  sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setNumberFormat('dd.MM.yyyy HH:mm');
  sh.setColumnWidths(1, HEADERS.length, 150);
}

function clean_(v, max) {
  v = String(v == null ? '' : v).trim().slice(0, max);
  // Formel-Injektion verhindern
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return out_({ ok: true });                  // Spam-Falle: Bots füllen das versteckte Feld
    if (!d.name || !d.contact || !d.district || !d.consent) return out_({ ok: false, error: 'missing' });
    lock.waitLock(10000);
    sheet_().appendRow([
      new Date(),
      clean_(d.name, 80),
      clean_(d.contact, 120),
      clean_(d.district, 60),
      clean_(d.interests, 200),
      clean_(d.source || 'Linkpage', 40),
      '',
      'neu',
      ''
    ]);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function doGet() {
  return out_({ ok: true, service: 'Christ Connect Anmeldungen' });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
