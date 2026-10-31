/**
 * Omiyage: raccoglie in un Foglio Google le registrazioni del sito.
 *
 * Ogni registrazione diventa una riga del foglio "Registrazioni".
 * Il foglio "Riepilogo" si aggiorna da solo: a sinistra quante persone
 * hanno scelto ogni regalo, a destra la lista di chi vuole cosa.
 * Se una persona si registra più volte con lo stesso nome, nel riepilogo
 * conta solo la sua ultima scelta.
 *
 * Installazione: Foglio Google > Estensioni > Apps Script, incolla questo
 * codice al posto di quello presente, salva, poi Esegui il deployment >
 * Nuovo deployment > App web (Esegui come: Me, Chi ha accesso: Chiunque).
 */

var REGISTRAZIONI = 'Registrazioni';
var RIEPILOGO = 'Riepilogo';
var INTESTAZIONI = ['Data', 'Nome', 'Regalo', 'Gruppo', 'Note'];

// Riceve una registrazione dal sito.
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var nome = testo_(d.nome);
    var regalo = testo_(d.regalo);
    // _honey lo compilano solo i bot
    if (!d._honey && nome && regalo) {
      foglio_(REGISTRAZIONI).appendRow([new Date(), sicuro_(nome), sicuro_(regalo), sicuro_(testo_(d.gruppo)), sicuro_(testo_(d.note))]);
      aggiornaRiepilogo();
    }
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Aprendo l'indirizzo dell'app web nel browser si vede se lo script è attivo.
function doGet() {
  return json_({ ok: true, messaggio: 'Lo script Omiyage è attivo.' });
}

// Menu "Omiyage" nel foglio, utile dopo aver aggiunto o corretto righe a mano.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Omiyage').addItem('Aggiorna riepilogo', 'aggiornaRiepilogo').addToUi();
}

function aggiornaRiepilogo() {
  var righe = foglio_(REGISTRAZIONI).getDataRange().getValues().slice(1).filter(function (r) {
    return String(r[1]).trim() && String(r[2]).trim();
  });

  // Vale l'ultima scelta di ogni persona
  var perNome = {};
  righe.forEach(function (r) { perNome[String(r[1]).trim().toLowerCase()] = r; });
  var scelte = Object.keys(perNome).map(function (k) { return perNome[k]; });

  var conteggi = {};
  scelte.forEach(function (r) { conteggi[r[2]] = (conteggi[r[2]] || 0) + 1; });
  var regali = Object.keys(conteggi).sort(function (a, b) {
    return conteggi[b] - conteggi[a] || a.localeCompare(b);
  });
  scelte.sort(function (a, b) {
    return regali.indexOf(a[2]) - regali.indexOf(b[2]) || String(a[1]).localeCompare(String(b[1]));
  });

  var f = foglio_(RIEPILOGO);
  f.clear();
  var sinistra = [['Regalo', 'Persone']].concat(regali.map(function (g) { return [sicuro_(g), conteggi[g]]; }));
  sinistra.push(['', ''], ['Totale', scelte.length]);
  f.getRange(1, 1, sinistra.length, 2).setValues(sinistra);
  var destra = [['Regalo', 'Nome', 'Note']].concat(scelte.map(function (r) {
    return [sicuro_(r[2]), sicuro_(r[1]), sicuro_(r[4])];
  }));
  f.getRange(1, 4, destra.length, 3).setValues(destra);
  f.getRange('A1:F1').setFontWeight('bold');
  f.getRange(sinistra.length, 1, 1, 2).setFontWeight('bold');
  f.setFrozenRows(1);
  f.autoResizeColumns(1, 6);
}

function foglio_(nome) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = ss.getSheetByName(nome) || ss.insertSheet(nome);
  if (nome === REGISTRAZIONI && f.getLastRow() === 0) {
    f.appendRow(INTESTAZIONI);
    f.getRange(1, 1, 1, INTESTAZIONI.length).setFontWeight('bold');
    f.setFrozenRows(1);
  }
  return f;
}

function testo_(v) {
  return String(v == null ? '' : v).trim().slice(0, 500);
}

// Evita che un testo che inizia con = + - @ venga letto come formula
function sicuro_(v) {
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
