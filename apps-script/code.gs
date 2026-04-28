/**
 * API Google Apps Script — Planilha Escala + Disponibilidades
 * Implante como “Aplicativo da Web” e use a URL /exec no index.html (SCRIPT_URL).
 */

var SHEET_ID  = '1ZUdYCsNxyt4z8B5KgdtQNpBKHJxuMQAI0EW4fT68akc';
var DATE_COLS = ['06/mai', '09/mai', '20/mai', '24/mai', '30/mai'];

/* ════════════════════════════════════════
   ROTEADOR
   ════════════════════════════════════════ */
function doGet(e) {
  var action   = (e.parameter && e.parameter.action)   || '';
  var callback = (e.parameter && e.parameter.callback) || '';
  var result;
  try {
    if      (action === 'getAll')        result = getAllAvailability();
    else if (action === 'getSchedule')   result = getScheduleData();
    else if (action === 'setSchedule')   result = setScheduleData(e.parameter.sched);
    else if (action === 'addPerson')     result = addPerson(e.parameter.date, e.parameter.name);
    else if (action === 'removePerson')  result = removePerson(e.parameter.date, e.parameter.name);
    else if (action === 'clearSchedule') result = clearScheduleData();
    else if (e.parameter && e.parameter.data) {
      registrarEscala(JSON.parse(e.parameter.data));
      result = { status: 'ok' };
    } else {
      result = { status: 'ok' };
    }
  } catch (err) {
    result = { error: String(err.message || err) };
  }
  var json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var dados = e.parameter && e.parameter.data
      ? JSON.parse(e.parameter.data)
      : (e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : null);
    if (dados) registrarEscala(dados);
  } catch (err) {
    console.error('Erro no doPost:', err.message);
  }
  return ContentService.createTextOutput('OK');
}

/* ════════════════════════════════════════
   DISPONIBILIDADES — matriz
   ════════════════════════════════════════ */
function getDispSheet() {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var aba = ss.getSheetByName('Disponibilidades') || ss.getSheets()[0];

  if (aba.getLastRow() === 0 || String(aba.getRange(1, 1).getValue()).trim() === '') {
    aba.getRange(1, 2, 1, DATE_COLS.length).setNumberFormat('@STRING@');
    var header = ['Nome'].concat(DATE_COLS);
    aba.getRange(1, 1, 1, header.length).setValues([header]);
  }
  return aba;
}

function getAllAvailability() {
  var aba     = getDispSheet();
  var rows    = aba.getDataRange().getValues();
  var headers = rows[0];
  var result  = [];
  var FALSY   = ['', '-', '—', 'não', 'nao', 'no', 'false'];

  for (var i = 1; i < rows.length; i++) {
    var nome = String(rows[i][0] || '').trim();
    if (!nome) continue;

    var tempResults = [];
    for (var j = 1; j < headers.length; j++) {
      var hdr  = headers[j];
      var date = (hdr instanceof Date)
        ? normalizeDate(hdr)
        : normalizeDate(String(hdr || '').trim());
      if (!date) continue;
      var val = String(rows[i][j] || '').trim().toLowerCase();
      if (FALSY.indexOf(val) !== -1) continue;
      tempResults.push([nome, date, 'Disponivel']);
    }

    if (tempResults.length > 0) {
      tempResults.forEach(function(r) { result.push(r); });
    } else {
      result.push([nome, 'INDISPONIVEL_TOTAL', 'Indisponivel']);
    }
  }
  return result;
}

function registrarEscala(selecionados) {
  if (!selecionados || selecionados.length === 0) return;

  var aba  = getDispSheet();
  var nome = String(selecionados[0][0] || '').trim().toUpperCase();

  var headers = aba.getRange(1, 1, 1, DATE_COLS.length + 1).getValues()[0];
  var colMap  = {};
  for (var h = 1; h < headers.length; h++) {
    var hdr = headers[h];
    var iso = (hdr instanceof Date)
      ? normalizeDate(hdr)
      : normalizeDate(String(hdr || '').trim());
    if (iso) colMap[iso] = h + 1;
  }

  var lastRow   = aba.getLastRow();
  var targetRow = lastRow + 1;
  if (lastRow > 1) {
    var names = aba.getRange(2, 1, lastRow, 1).getValues();
    for (var n = 0; n < names.length; n++) {
      if (String(names[n][0] || '').trim().toUpperCase() === nome) {
        targetRow = n + 2;
        break;
      }
    }
  }

  aba.getRange(targetRow, 1).setValue(nome);

  selecionados.forEach(function(item) {
    var iso = normalizeDate(String(item[1] || '').trim());
    var col = colMap[iso];
    if (!col) return;
    aba.getRange(targetRow, col).setValue(item[2] === 'Disponivel' ? '✓' : '');
  });
}

/* ════════════════════════════════════════
   UTIL — datas ISO YYYY-MM-DD
   ════════════════════════════════════════ */
function normalizeDate(rawDate) {
  if (rawDate instanceof Date) {
    var yr = rawDate.getFullYear();
    var mo = rawDate.getMonth() + 1;
    var dy = rawDate.getDate();
    return yr + '-' + (mo < 10 ? '0' + mo : mo) + '-' + (dy < 10 ? '0' + dy : dy);
  }
  var s = String(rawDate || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var m = s.match(/^(\d{1,2})\/mai$/i);
  if (m) {
    var d = parseInt(m[1], 10);
    return '2026-05-' + (d < 10 ? '0' + d : String(d));
  }
  return s;
}

/* ════════════════════════════════════════
   ESCALA — unicidade (data + nome)
   ════════════════════════════════════════ */
function getEscalaSheet() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Escala');
  if (!sheet) {
    sheet = ss.insertSheet('Escala');
    sheet.appendRow(['Data', 'Nome']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  return sheet;
}

function getScheduleData() {
  var sheet  = getEscalaSheet();
  var data   = sheet.getDataRange().getValues();
  var result = {};
  var seen   = {};

  for (var i = 1; i < data.length; i++) {
    var date = normalizeDate(data[i][0]);
    var name = String(data[i][1] || '').trim();
    if (!date || !name) continue;

    var pairKey = date + '\t' + name.toUpperCase();
    if (seen[pairKey]) continue;
    seen[pairKey] = true;

    if (!result[date]) result[date] = [];
    result[date].push(name);
  }
  return result;
}

function setScheduleData(schedJson) {
  if (!schedJson) return { status: 'error', msg: 'missing sched' };

  var sched   = JSON.parse(schedJson);
  var sheet   = getEscalaSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  var rows = [];
  var seen = {};

  Object.keys(sched).forEach(function(rawDate) {
    var date = normalizeDate(rawDate);
    if (!date) return;

    var names = sched[rawDate];
    if (!Array.isArray(names)) return;

    for (var n = 0; n < names.length; n++) {
      var name = String(names[n] || '').trim();
      if (!name) continue;

      var pairKey = date + '\t' + name.toUpperCase();
      if (seen[pairKey]) continue;
      seen[pairKey] = true;

      rows.push([date, name]);
    }
  });

  rows.sort(function(a, b) {
    if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1;
    return a[1].toUpperCase() < b[1].toUpperCase() ? -1 : 1;
  });

  for (var r = 0; r < rows.length; r++) {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1).setNumberFormat('@STRING@');
    sheet.getRange(newRow, 1, 1, 2).setValues([rows[r]]);
  }

  return { status: 'ok' };
}

function addPerson(date, name) {
  if (!date || !name) return { status: 'error', msg: 'missing params' };
  date = normalizeDate(date);
  name = String(name).trim();
  if (!date || !name) return { status: 'error', msg: 'invalid params' };

  var sheet = getEscalaSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeDate(data[i][0]) === date &&
        String(data[i][1] || '').trim().toUpperCase() === name.toUpperCase()) {
      return { status: 'exists' };
    }
  }
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1).setNumberFormat('@STRING@');
  sheet.getRange(newRow, 1, 1, 2).setValues([[date, name]]);
  return { status: 'ok' };
}

function removePerson(date, name) {
  if (!date || !name) return { status: 'error', msg: 'missing params' };
  date = normalizeDate(date);
  name = String(name).trim();
  var sheet = getEscalaSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (normalizeDate(data[i][0]) === date &&
        String(data[i][1] || '').trim().toUpperCase() === name.toUpperCase()) {
      sheet.deleteRow(i + 1);
      return { status: 'ok' };
    }
  }
  return { status: 'not_found' };
}

function clearScheduleData() {
  var sheet   = getEscalaSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  return { status: 'ok' };
}
