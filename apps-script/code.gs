/**
 * API Google Apps Script — Planilha Escala + Disponibilidades
 * Implante como “Aplicativo da Web” e use a URL /exec no index.html (SCRIPT_URL).
 */

var SHEET_ID = '1ZUdYCsNxyt4z8B5KgdtQNpBKHJxuMQAI0EW4fT68akc';
var DATE_COLS = ['04/jul', '08/jul', '18/jul', '22/jul', '26/jul', '29/jul'];
var RECEIVED_AT_COL = 9; // Coluna I
var MINISTRY_NAME = 'Evangelismo & Integração';
var ROLE_VOLUNTEER = 'Voluntário';
var ROLE_LEADER = 'Líder';
var ESCALA_HEADERS = ['Data', 'Nome', 'Função', 'Ministério', 'Lider'];

/** Ordem cronológica das datas da escala (igual ao GROUP_DATES no front). */
var SCHEDULE_DATE_KEYS = ['2026-07-04', '2026-07-08', '2026-07-18', '2026-07-22', '2026-07-26', '2026-07-29'];
var ACTIVE_PERIOD = '2026-07';

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
    else if (action === 'getDates') result = getScheduleDates();
    else if (action === 'setDates') result = setScheduleDates(e.parameter.dates);
    else if (action === 'addDate') result = addScheduleDate(e.parameter.date);
    else if (action === 'removeDate') result = removeScheduleDate(e.parameter.key);
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
    rebuildDisponibilidadesColumns(getDefaultScheduleDates());
  }
  return aba;
}

function garantirFormatoRecebidoEm(aba, targetRow, colNum) {
  var col = colNum || RECEIVED_AT_COL;
  aba.getRange(targetRow, col).setNumberFormat('dd/MM/yyyy HH:mm:ss');
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
      var date = headerCellToIso(hdr);
      if (!date || date.indexOf(ACTIVE_PERIOD + '-') !== 0) continue;
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

  var lastCol = Math.max(aba.getLastColumn(), 1);
  var headers = aba.getRange(1, 1, 1, lastCol).getValues()[0];
  var recvCol = findReceivedAtColumn(headers);
  var colMap = {};
  for (var h = 1; h < headers.length; h++) {
    var colNum = h + 1;
    if (colNum === recvCol) continue;
    var iso = headerCellToIso(headers[h]);
    if (iso) colMap[iso] = colNum;
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
  aba.getRange(targetRow, recvCol).setValue(new Date());
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
var MONTH_ABBR_MAP = {
  jan: 1, fev: 2, mar: 3, abr: 4, mai: 5, jun: 6,
  jul: 7, ago: 8, set: 9, out: 10, nov: 11, dez: 12
};

function cleanDispHeaderText(raw) {
  return String(raw || '')
    .trim()
    .replace(/\s+/g, '')
    .replace(/\.+$/g, '')
    .replace(/\/+/g, '/');
}

function normalizeDate(rawDate) {
  if (rawDate instanceof Date) {
    var yr = rawDate.getFullYear();
    var mo = rawDate.getMonth() + 1;
    var dy = rawDate.getDate();
    return yr + '-' + (mo < 10 ? '0' + mo : mo) + '-' + (dy < 10 ? '0' + dy : dy);
  }
  var s = cleanDispHeaderText(rawDate);
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return s.slice(0, 10);

  var m = s.match(/^(\d{1,2})\/(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)$/i);
  if (m) {
    var day = parseInt(m[1], 10);
    var moNum = MONTH_ABBR_MAP[m[2].toLowerCase()];
    if (!day || !moNum) return '';
    var year = 2026;
    return year + '-'
      + (moNum < 10 ? '0' + moNum : String(moNum)) + '-'
      + (day < 10 ? '0' + day : String(day));
  }
  return '';
}

function normalizeDispHeaderLabel(label, isoKey) {
  var iso = normalizeDate(isoKey || label);
  if (iso) return isoToShortLabel(iso);
  var cleaned = cleanDispHeaderText(label);
  return cleaned || '';
}

function headerCellToIso(hdr) {
  if (hdr instanceof Date) return normalizeDate(hdr);
  return normalizeDate(hdr);
}

function periodFromText(periodo) {
  var s = String(periodo || '').trim().toLowerCase();
  if (!s) return '';
  if (s.indexOf('jul') !== -1 && s.indexOf('2026') !== -1) return '2026-07';
  return '';
}

/* ════════════════════════════════════════
   ESCALA — unicidade (data + nome)
   ════════════════════════════════════════ */
function ensureEscalaHeaders(sheet) {
  var current = sheet.getRange(1, 1, 1, Math.max(5, sheet.getLastColumn())).getValues()[0];
  var h3 = String(current[2] || '').trim();
  if (h3 === 'Lider' || h3 === 'Líder') {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var oldData = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      sheet.deleteRows(2, lastRow - 1);
      var migrated = [];
      for (var i = 0; i < oldData.length; i++) {
        var date = oldData[i][0];
        var name = oldData[i][1];
        var leader = oldData[i][2];
        if (!date || !name) continue;
        migrated.push([
          date,
          name,
          isLeaderCell(leader) ? ROLE_LEADER : ROLE_VOLUNTEER,
          MINISTRY_NAME,
          isLeaderCell(leader) ? '✓' : ''
        ]);
      }
      sheet.getRange(1, 1, 1, ESCALA_HEADERS.length).setValues([ESCALA_HEADERS]);
      sheet.getRange(1, 1, 1, ESCALA_HEADERS.length).setFontWeight('bold');
      for (var r = 0; r < migrated.length; r++) {
        var newRow = sheet.getLastRow() + 1;
        sheet.getRange(newRow, 1).setNumberFormat('@STRING@');
        sheet.getRange(newRow, 1, 1, ESCALA_HEADERS.length).setValues([migrated[r]]);
      }
    } else {
      sheet.getRange(1, 1, 1, ESCALA_HEADERS.length).setValues([ESCALA_HEADERS]);
      sheet.getRange(1, 1, 1, ESCALA_HEADERS.length).setFontWeight('bold');
    }
    return;
  }
  var needsHeader = String(current[0] || '').trim() !== 'Data'
    || String(current[2] || '').trim() !== 'Função'
    || String(current[3] || '').trim() !== 'Ministério';
  if (needsHeader) {
    sheet.getRange(1, 1, 1, ESCALA_HEADERS.length).setValues([ESCALA_HEADERS]);
    sheet.getRange(1, 1, 1, ESCALA_HEADERS.length).setFontWeight('bold');
  }
}

function scheduleRowValues(date, name, isLeader) {
  return [
    date,
    name,
    isLeader ? ROLE_LEADER : ROLE_VOLUNTEER,
    MINISTRY_NAME,
    isLeader ? '✓' : ''
  ];
}

function getEscalaSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Escala');
  if (!sheet) {
    sheet = ss.insertSheet('Escala');
    sheet.getRange(1, 1, 1, ESCALA_HEADERS.length).setValues([ESCALA_HEADERS]);
    sheet.getRange(1, 1, 1, ESCALA_HEADERS.length).setFontWeight('bold');
  } else {
    ensureEscalaHeaders(sheet);
  }
  return sheet;
}

function getLeaderColumnIndex() {
  return 5;
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
  var leaderCol = getLeaderColumnIndex() - 1;

  for (var i = 1; i < rows.length; i++) {
    var date = normalizeDate(rows[i][0]);
    var name = String(rows[i][1] || '').trim();
    if (!date || !name || date.indexOf(ACTIVE_PERIOD + '-') !== 0) continue;

    var pairKey = date + '\t' + name.toUpperCase();
    if (seen[pairKey]) continue;
    seen[pairKey] = true;

    if (!result[date]) result[date] = [];
    result[date].push(name);
    var leaderVal = rows[i][leaderCol];
    if (leaderVal === undefined || leaderVal === '') {
      leaderVal = rows[i][2];
    }
    if (isLeaderCell(leaderVal) && !leaders[date]) leaders[date] = name;
    if (!leaders[date] && String(rows[i][2] || '').trim() === ROLE_LEADER) leaders[date] = name;
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
      rows.push(scheduleRowValues(date, name, isLeader));
    }
  }

  var orderedKeys = getScheduleDateKeysOrdered();
  var used = {};
  for (var di = 0; di < orderedKeys.length; di++) {
    used[orderedKeys[di]] = true;
    appendRowsForDateKey(orderedKeys[di]);
  }

  Object.keys(sched).forEach(function(k) {
    if (used[k]) return;
    var date = normalizeDate(k);
    if (!date || date.indexOf(ACTIVE_PERIOD + '-') !== 0) return;
    appendRowsForDateKey(k);
  });

  var colCount = ESCALA_HEADERS.length;
  for (var r = 0; r < rows.length; r++) {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1).setNumberFormat('@STRING@');
    sheet.getRange(newRow, 1, 1, colCount).setValues([rows[r]]);
  }

  if (sheet.getLastRow() > 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, colCount).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
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
  sheet.getRange(newRow, 1, 1, ESCALA_HEADERS.length).setValues([scheduleRowValues(date, name, false)]);
  if (sheet.getLastRow() > 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, ESCALA_HEADERS.length).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
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
    sheet.getRange(2, 1, sheet.getLastRow() - 1, ESCALA_HEADERS.length).sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
  }
  return { status: 'ok' };
}

function clearScheduleData() {
  var sheet = getEscalaSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  return { status: 'ok' };
}

/* ════════════════════════════════════════
   DATAS DA ESCALA (configuração dinâmica)
   ════════════════════════════════════════ */
function getDefaultScheduleDates() {
  return [
    { key: '2026-07-04', label: '04/jul', full: '04 de julho', day: 'Sábado',  hora: '10h às 11h30' },
    { key: '2026-07-08', label: '08/jul', full: '08 de julho', day: 'Quarta',  hora: '18h às 19h30' },
    { key: '2026-07-18', label: '18/jul', full: '18 de julho', day: 'Sábado',  hora: '09h30 às 11h' },
    { key: '2026-07-22', label: '22/jul', full: '22 de julho', day: 'Quarta',  hora: '18h às 19h30' },
    { key: '2026-07-26', label: '26/jul', full: '26 de julho', day: 'Domingo', hora: '17h às 18h30' },
    { key: '2026-07-29', label: '29/jul', full: '29 de julho', day: 'Quarta',  hora: '18h às 19h30' }
  ];
}

function isoToShortLabel(iso) {
  var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  var mo = parseInt(m[2], 10);
  var dy = parseInt(m[3], 10);
  var abbr = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
  return dy + '/' + (abbr[mo - 1] || 'jul');
}

function getDatesSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('DatasEscala');
  if (!sheet) {
    sheet = ss.insertSheet('DatasEscala');
    sheet.getRange(1, 1, 1, 5).setValues([['key', 'label', 'full', 'day', 'hora']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  return sheet;
}

function isPollutedDateText(val) {
  var s = String(val || '').trim();
  if (!s) return false;
  if (s.length > 36) return true;
  return /GMT|Horário|Horario|00:00:00|\b(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/i.test(s);
}

function cellToCleanText(val) {
  if (val instanceof Date) return '';
  var s = String(val || '').trim();
  if (!s || isPollutedDateText(s)) return '';
  return s;
}

function buildDateMetaFromIso(iso, horaOpt) {
  var parts = String(iso || '').split('-');
  if (parts.length < 3) return null;
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  var d = parseInt(parts[2], 10);
  if (!y || !m || !d) return null;
  var dt = new Date(y, m - 1, d);
  var days = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
  var months = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
  var abbr = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  return {
    label: d + '/' + abbr[m - 1],
    full: d + ' de ' + months[m - 1],
    day: days[dt.getDay()],
    hora: String(horaOpt || '').trim() || '18h às 19h30'
  };
}

function sanitizeDateRecord(d) {
  if (!d) return null;
  var key = normalizeDate(d.key);
  if (!key) return null;
  var meta = buildDateMetaFromIso(key, d.hora);
  if (!meta) return null;
  return {
    key: key,
    label: cellToCleanText(d.label) || meta.label,
    full: cellToCleanText(d.full) || meta.full,
    day: cellToCleanText(d.day) || meta.day,
    hora: cellToCleanText(d.hora) || meta.hora
  };
}

function getScheduleDates() {
  var sheet = getDatesSheet();
  var rows = sheet.getDataRange().getValues();
  var dates = [];
  for (var i = 1; i < rows.length; i++) {
    var clean = sanitizeDateRecord({
      key: rows[i][0],
      label: rows[i][1],
      full: rows[i][2],
      day: rows[i][3],
      hora: rows[i][4]
    });
    if (!clean) continue;
    dates.push(clean);
  }
  dates = dates.filter(function(d) { return d.key.indexOf(ACTIVE_PERIOD + '-') === 0; });
  if (dates.length === 0) dates = getDefaultScheduleDates();
  dates.sort(function(a, b) {
    return a.key < b.key ? -1 : a.key > b.key ? 1 : 0;
  });
  return { dates: dates };
}

function getScheduleDateKeysOrdered() {
  return getScheduleDates().dates.map(function(d) { return d.key; });
}

function findReceivedAtColumn(headers) {
  for (var h = 0; h < headers.length; h++) {
    if (String(headers[h] || '').trim().toLowerCase() === 'recebido em') return h + 1;
  }
  return headers.length + 1;
}

/**
 * Reescreve a aba Disponibilidades com colunas de data em ordem cronológica,
 * preservando os checks (✓) de cada voluntário na coluna correta.
 */
function rebuildDisponibilidadesColumns(scheduleDates) {
  var aba = getDispSheet();
  var lastRow = Math.max(aba.getLastRow(), 1);
  var lastCol = Math.max(aba.getLastColumn(), 1);
  var allData = aba.getRange(1, 1, lastRow, lastCol).getValues();
  var oldHeaders = allData[0] || ['Nome'];
  var recvColIdx = -1;
  var oldDateCols = [];

  for (var h = 1; h < oldHeaders.length; h++) {
    var hdrLower = String(oldHeaders[h] || '').trim().toLowerCase();
    if (hdrLower === 'recebido em') {
      recvColIdx = h;
      continue;
    }
    var iso = headerCellToIso(oldHeaders[h]);
    if (!iso) continue;
    oldDateCols.push({
      iso: iso,
      label: normalizeDispHeaderLabel(oldHeaders[h], iso),
      colIdx: h
    });
  }

  var dateMap = {};
  (scheduleDates || []).forEach(function(d) {
    var clean = sanitizeDateRecord(d);
    if (!clean) return;
    dateMap[clean.key] = normalizeDispHeaderLabel(clean.label, clean.key);
  });

  var sortedKeys = Object.keys(dateMap).sort(function(a, b) {
    return a < b ? -1 : a > b ? 1 : 0;
  });

  var newHeaders = ['Nome'];
  sortedKeys.forEach(function(k) { newHeaders.push(dateMap[k]); });
  newHeaders.push('Recebido em');
  var numCols = newHeaders.length;
  var recvColNum = numCols;

  var newRows = [newHeaders];
  for (var r = 1; r < allData.length; r++) {
    var row = allData[r];
    var newRow = [row[0] != null ? row[0] : ''];
    for (var ki = 0; ki < sortedKeys.length; ki++) {
      var iso = sortedKeys[ki];
      var srcIdx = -1;
      for (var o = 0; o < oldDateCols.length; o++) {
        if (oldDateCols[o].iso === iso) {
          srcIdx = oldDateCols[o].colIdx;
          break;
        }
      }
      newRow.push(srcIdx >= 0 && row[srcIdx] !== undefined && row[srcIdx] !== null ? row[srcIdx] : '');
    }
    var recvVal = recvColIdx >= 0 && row[recvColIdx] !== undefined ? row[recvColIdx] : '';
    newRow.push(recvVal);
    newRows.push(newRow);
  }

  if (lastRow > 0 && lastCol > 0) {
    aba.getRange(1, 1, lastRow, lastCol).clearContent();
  }
  if (newRows.length && numCols) {
    aba.getRange(1, 1, newRows.length, numCols).setValues(newRows);
    aba.getRange(1, 1, 1, numCols).setNumberFormat('@STRING@');
    if (sortedKeys.length) {
      aba.getRange(1, 2, newRows.length, 1 + sortedKeys.length).setNumberFormat('@STRING@');
    }
    aba.getRange(1, recvColNum).setValue('Recebido em').setNumberFormat('@STRING@');
    for (var rr = 2; rr <= newRows.length; rr++) {
      if (newRows[rr - 1][recvColNum - 1]) {
        aba.getRange(rr, recvColNum).setValue(newRows[rr - 1][recvColNum - 1]);
        garantirFormatoRecebidoEm(aba, rr, recvColNum);
      }
    }
  }
}

function setScheduleDates(datesJson) {
  if (!datesJson) return { status: 'error', msg: 'missing dates' };
  var dates = JSON.parse(datesJson);
  if (!Array.isArray(dates)) return { status: 'error', msg: 'invalid dates' };

  var sheet = getDatesSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  var cleanedDates = [];
  for (var i = 0; i < dates.length; i++) {
    var clean = sanitizeDateRecord(dates[i]);
    if (!clean) continue;
    cleanedDates.push(clean);
  }
  cleanedDates.sort(function(a, b) {
    return a.key < b.key ? -1 : a.key > b.key ? 1 : 0;
  });

  var rows = [];
  for (var j = 0; j < cleanedDates.length; j++) {
    var c = cleanedDates[j];
    rows.push([c.key, c.label, c.full, c.day, c.hora]);
  }

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
    sheet.getRange(2, 1, rows.length, 5).setNumberFormat('@STRING@');
  }

  if (cleanedDates.length) {
    rebuildDisponibilidadesColumns(cleanedDates);
  }
  return { status: 'ok', count: rows.length };
}

function addScheduleDate(dateJson) {
  if (!dateJson) return { status: 'error', msg: 'missing date' };
  var incoming = sanitizeDateRecord(JSON.parse(dateJson));
  if (!incoming || !incoming.key) return { status: 'error', msg: 'invalid date' };

  var existing = getScheduleDates().dates;
  var key = incoming.key;
  var found = false;
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].key === key) {
      existing[i] = incoming;
      found = true;
      break;
    }
  }
  if (!found) existing.push(incoming);
  existing.sort(function(a, b) {
    return a.key < b.key ? -1 : a.key > b.key ? 1 : 0;
  });
  return setScheduleDates(JSON.stringify(existing));
}


function removeEscalaRowsForDate(isoKey) {
  var sheet = getEscalaSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (normalizeDate(data[i][0]) === isoKey) {
      sheet.deleteRow(i + 2);
    }
  }
}

function removeScheduleDate(rawKey) {
  var isoKey = normalizeDate(rawKey);
  if (!isoKey) return { status: 'error', msg: 'invalid key' };

  var dates = getScheduleDates().dates.filter(function(d) {
    return d.key !== isoKey;
  });
  if (dates.length === 0) {
    return { status: 'error', msg: 'cannot remove last date' };
  }

  var sheet = getDatesSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  var rows = [];
  for (var i = 0; i < dates.length; i++) {
    rows.push([dates[i].key, dates[i].label, dates[i].full, dates[i].day, dates[i].hora]);
  }
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
    sheet.getRange(2, 1, rows.length, 5).setNumberFormat('@STRING@');
  }

  removeEscalaRowsForDate(isoKey);
  rebuildDisponibilidadesColumns(dates);

  return { status: 'ok', key: isoKey };
}
