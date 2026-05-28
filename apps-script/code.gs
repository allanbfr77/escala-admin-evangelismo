/**
 * API Google Apps Script — Planilha Escala + Disponibilidades
 * Implante como “Aplicativo da Web” e use a URL /exec no index.html (SCRIPT_URL).
 */

var SHEET_ID = '1ZUdYCsNxyt4z8B5KgdtQNpBKHJxuMQAI0EW4fT68akc';
var DATE_COLS = ['03/jun', '06/jun', '20/jun', '24/jun', '28/jun'];
var RECEIVED_AT_COL = 9; // Coluna I

/** Ordem cronológica das datas da escala (igual ao GROUP_DATES no front). */
var SCHEDULE_DATE_KEYS = ['2026-06-03', '2026-06-06', '2026-06-20', '2026-06-24', '2026-06-28'];

/* ════════════════════════════════════════
   ROTEADOR
   ════════════════════════════════════════ */
function doGet(e) {
  var action = (e.parameter && e.parameter.action) || '';
  var callback = (e.parameter && e.parameter.callback) || '';
  var result;
  try {
    if (action === 'getAll') result = getAllAvailability();
    else if (action === 'getUsedNames') result = getUsedNames();
    else if (action === 'nomesOcupados') result = getOccupiedNames(e.parameter.periodo);
    else if (action === 'getSchedule') result = getScheduleData();
    else if (action === 'setSchedule') result = setScheduleData(e.parameter.sched, e.parameter.leaders);
    else if (action === 'addPerson') result = addPerson(e.parameter.date, e.parameter.name);
    else if (action === 'removePerson') result = removePerson(e.parameter.date, e.parameter.name);
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
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var aba = ss.getSheetByName('Disponibilidades') || ss.getSheets()[0];

  var primeiraLinha = aba.getLastRow() === 0 ? '' : String(aba.getRange(1, 1).getValue()).trim();
  if (!primeiraLinha) {
    aba.getRange(1, 2, 1, DATE_COLS.length).setNumberFormat('@STRING@');
    var header = ['Nome'].concat(DATE_COLS);
    aba.getRange(1, 1, 1, header.length).setValues([header]);
    aba.getRange(1, RECEIVED_AT_COL).setValue('Recebido em');
  }
  return aba;
}

function garantirFormatoRecebidoEm(aba, targetRow) {
  aba.getRange(targetRow, RECEIVED_AT_COL).setNumberFormat('dd/MM/yyyy HH:mm:ss');
}

function getAllAvailability() {
  var aba = getDispSheet();
  var rows = aba.getDataRange().getValues();
  var headers = rows[0];
  var result = [];
  var FALSY = ['', '-', '—', 'não', 'nao', 'no', 'false'];

  for (var i = 1; i < rows.length; i++) {
    var nome = String(rows[i][0] || '').trim();
    if (!nome) continue;

    var tempResults = [];
    for (var j = 1; j < headers.length; j++) {
      var hdr = headers[j];
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

function getOccupiedNames(periodo) {
  var registros = getAllAvailability();
  var namesMap = {};
  var yearMonth = periodFromText(periodo);
  for (var i = 0; i < registros.length; i++) {
    var nome = normalizePersonName(registros[i][0]);
    var dataIso = String(registros[i][1] || '').trim();
    if (!nome) continue;

    if (dataIso === 'INDISPONIVEL_TOTAL') {
      namesMap[nome] = true;
      continue;
    }
    if (!yearMonth || dataIso.indexOf(yearMonth + '-') === 0) {
      namesMap[nome] = true;
    }
  }
  return { ocupados: Object.keys(namesMap).sort() };
}

function getUsedNames() {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var aba = ss.getSheetByName('Disponibilidades') || ss.getSheets()[0];
  var lastRow = aba.getLastRow();
  if (lastRow <= 1) return { ocupados: [] };

  var values = aba.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  var namesMap = {};
  for (var i = 0; i < values.length; i++) {
    var nome = normalizePersonName(values[i][0]);
    if (!nome) continue;
    namesMap[nome] = true;
  }
  return { ocupados: Object.keys(namesMap).sort() };
}

function invalidarCacheNomes() { /* cache removido — leitura sempre ao vivo */ }

function registrarEscala(selecionados) {
  if (!selecionados || selecionados.length === 0) return;

  var aba = getDispSheet();
  var nome = String(selecionados[0][0] || '').trim();
  var nomeNorm = normalizePersonName(nome);

  var headers = aba.getRange(1, 1, 1, DATE_COLS.length + 1).getValues()[0];
  var colMap = {};
  for (var h = 1; h < headers.length; h++) {
    var hdr = headers[h];
    var iso = (hdr instanceof Date)
      ? normalizeDate(hdr)
      : normalizeDate(String(hdr || '').trim());
    if (iso) colMap[iso] = h + 1;
  }

  var lastRow = aba.getLastRow();
  var targetRow = -1;
  if (lastRow > 1) {
    var names = aba.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    for (var n = 0; n < names.length; n++) {
      var existingNameNorm = normalizePersonName(names[n][0]);
      if (existingNameNorm && existingNameNorm === nomeNorm) {
        targetRow = n + 2;
        break;
      }
    }

    if (targetRow === -1) {
      for (var e = 0; e < names.length; e++) {
        if (!String(names[e][0] || '').trim()) {
          targetRow = e + 2;
          break;
        }
      }
    }
  }

  if (targetRow === -1) targetRow = Math.max(2, lastRow + 1);

  aba.getRange(targetRow, 1).setValue(nome.toUpperCase());
  garantirFormatoRecebidoEm(aba, targetRow);
  aba.getRange(targetRow, RECEIVED_AT_COL).setValue(new Date());
  invalidarCacheNomes();

  selecionados.forEach(function(item) {
    var iso = normalizeDate(String(item[1] || '').trim());
    var col = colMap[iso];
    if (!col) return;
    aba.getRange(targetRow, col).setValue(item[2] === 'Disponivel' ? '✓' : '');
  });
}

function normalizePersonName(rawName) {
  return String(rawName || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
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
  var m = s.match(/^(\d{1,2})\/jun$/i);
  if (m) {
    var d = parseInt(m[1], 10);
    return '2026-06-' + (d < 10 ? '0' + d : String(d));
  }
  return s;
}

function periodFromText(periodo) {
  var s = String(periodo || '').trim().toLowerCase();
  if (!s) return '';
  if (s.indexOf('jun') !== -1 && s.indexOf('2026') !== -1) return '2026-06';
  return '';
}

/* ════════════════════════════════════════
   ESCALA — unicidade (data + nome)
   ════════════════════════════════════════ */
function getEscalaSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Escala');
  if (!sheet) {
    sheet = ss.insertSheet('Escala');
    sheet.appendRow(['Data', 'Nome', 'Lider']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  } else if (String(sheet.getRange(1, 3).getValue() || '').trim() !== 'Lider') {
    sheet.getRange(1, 3).setValue('Lider').setFontWeight('bold');
  }
  return sheet;
}

function isLeaderCell(val) {
  var s = String(val || '').trim().toLowerCase();
  return s === '✓' || s === '1' || s === 'sim' || s === 'true' || s === 'lider' || s === 'líder';
}

function readEscalaRows(sheet) {
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var date = normalizeDate(data[i][0]);
    var name = String(data[i][1] || '').trim();
    if (!date || !name) continue;
    rows.push([date, name]);
  }
  return rows;
}

function getScheduleData() {
  var sheet = getEscalaSheet();
  var rows = sheet.getDataRange().getValues();
  var result = {};
  var leaders = {};
  var seen = {};

  for (var i = 1; i < rows.length; i++) {
    var date = normalizeDate(rows[i][0]);
    var name = String(rows[i][1] || '').trim();
    if (!date || !name) continue;

    var pairKey = date + '\t' + name.toUpperCase();
    if (seen[pairKey]) continue;
    seen[pairKey] = true;

    if (!result[date]) result[date] = [];
    result[date].push(name);
    if (isLeaderCell(rows[i][2]) && !leaders[date]) leaders[date] = name;
  }
  return { schedule: result, leaders: leaders };
}

function setScheduleData(schedJson, leadersJson) {
  if (!schedJson) return { status: 'error', msg: 'missing sched' };

  var sched = JSON.parse(schedJson);
  var leaders = {};
  if (leadersJson) {
    try { leaders = JSON.parse(leadersJson) || {}; } catch (e) { leaders = {}; }
  }

  var sheet = getEscalaSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  var rows = [];
  var seen = {};

  function appendRowsForDateKey(rawDateKey) {
    if (!sched.hasOwnProperty(rawDateKey)) return;
    var date = normalizeDate(rawDateKey);
    if (!date) return;

    var names = sched[rawDateKey];
    if (!Array.isArray(names)) return;

    var leaderName = String(leaders[rawDateKey] || leaders[date] || '').trim();

    for (var n = 0; n < names.length; n++) {
      var name = String(names[n] || '').trim();
      if (!name) continue;

      var pairKey = date + '\t' + name.toUpperCase();
      if (seen[pairKey]) continue;
      seen[pairKey] = true;

      var isLeader = leaderName && name.toUpperCase() === leaderName.toUpperCase();
      rows.push([date, name, isLeader ? '✓' : '']);
    }
  }

  for (var di = 0; di < SCHEDULE_DATE_KEYS.length; di++) {
    appendRowsForDateKey(SCHEDULE_DATE_KEYS[di]);
  }

  Object.keys(sched).forEach(function(k) {
    if (SCHEDULE_DATE_KEYS.indexOf(k) !== -1) return;
    appendRowsForDateKey(k);
  });

  for (var r = 0; r < rows.length; r++) {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1).setNumberFormat('@STRING@');
    sheet.getRange(newRow, 1, 1, 3).setValues([rows[r]]);
  }

  if (sheet.getLastRow() > 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
  }

  return { status: 'ok' };
}

function addPerson(date, name) {
  if (!date || !name) return { status: 'error', msg: 'missing params' };
  date = normalizeDate(date);
  name = String(name).trim();
  if (!date || !name) return { status: 'error', msg: 'invalid params' };

  var sheet = getEscalaSheet();
  var data = readEscalaRows(sheet);
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === date && data[i][1].toUpperCase() === name.toUpperCase()) {
      return { status: 'exists' };
    }
  }
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1).setNumberFormat('@STRING@');
  sheet.getRange(newRow, 1, 1, 3).setValues([[date, name, '']]);
  if (sheet.getLastRow() > 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
  }
  return { status: 'ok' };
}

function removePerson(date, name) {
  if (!date || !name) return { status: 'error', msg: 'missing params' };
  date = normalizeDate(date);
  name = String(name).trim();
  var sheet = getEscalaSheet();
  var data = readEscalaRows(sheet);
  var targetRow = -1;
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === date && data[i][1].toUpperCase() === name.toUpperCase()) {
      targetRow = i + 2;
      break;
    }
  }
  if (targetRow === -1) return { status: 'not_found' };
  sheet.deleteRow(targetRow);
  if (sheet.getLastRow() > 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
  }
  return { status: 'ok' };
}

function clearScheduleData() {
  var sheet = getEscalaSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  return { status: 'ok' };
}
