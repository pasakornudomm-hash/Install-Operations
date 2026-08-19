const APP = {
  rawSheet: 'RawData',
  historySheet: 'ImportHistory',
  configSheet: 'Config',
  defaultFolderName: 'MyCenter Install Dashboard Uploads',
  terminationText: 'termination',
  addressColumnIndex: 18, // Column R in the imported source file.
  dashboardSampleLimit: 300,
  searchLimit: 1000,
  historyLimit: 100,
  metaHeaders: ['_Import File', '_Import At', '_Source File ID', '_Source URL', '_Row Hash'],
  cacheKey: 'raw_data_cache_v2',
  cacheDuration: 300 // 5 minutes
};

// ── In-memory cache (survives within same execution) ──
let _memoryCache = null;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MyCenter Dashboard')
    .addItem('Setup / Refresh sheets', 'setupWorkbook')
    .addToUi();
}

function doGet() {
  setupWorkbook();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('MyCenter Install Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action || body.fn || '';
    const args = body.args || [];
    let result;
    switch (action) {
      case 'getBootstrapData':
        result = getBootstrapData();
        break;
      case 'searchRecords':
        result = searchRecords(args[0] || {});
        break;
      case 'uploadFile':
        result = uploadFile(args[0] || {});
        break;
      case 'getImportHistory':
        result = getImportHistory();
        break;
      case 'getDashboardData':
        result = getBootstrapData();
        break;
      default:
        throw new Error('Unknown action: ' + action);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, data: result })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function setupWorkbook() {
  const ss = SpreadsheetApp.getActive();
  const raw = getOrCreateSheet_(ss, APP.rawSheet);
  const history = getOrCreateSheet_(ss, APP.historySheet);
  const config = getOrCreateSheet_(ss, APP.configSheet);

  if (history.getLastRow() === 0) {
    history.appendRow([
      'Import At', 'File Name', 'Total Rows', 'Imported Rows',
      'Skipped Termination', 'Duplicate Rows', 'Status', 'Source URL'
    ]);
  }
  if (config.getLastRow() === 0) {
    config.appendRow(['Key', 'Value']);
  }
  const folder = ensureUploadFolder_();
  ensureConfigValue_(config, 'Termination Filter', 'Termination');
  raw.setFrozenRows(Math.min(raw.getLastRow(), 1));
  history.setFrozenRows(1);
  return { ok: true, spreadsheetUrl: ss.getUrl(), folderUrl: folder.getUrl() };
}

function getBootstrapData() {
  setupWorkbook();
  const rawData = readRawData_();
  return buildInitialDataPayload_(rawData);
}

function buildInitialDataPayload_(rawData) {
  rawData = rawData || { headers: [], columns: [], records: [] };
  const defaultMonth = getDefaultMonth_(rawData.records);
  const filters = defaultMonth ? getMonthFilters_(defaultMonth) : {};
  const records = defaultMonth
    ? filterRecords_(rawData.records, rawData.headers, filters)
    : rawData.records;
  return {
    dashboard: buildDashboardData_(records, rawData.columns),
    columns: rawData.columns,
    history: getImportHistory(),
    defaultMonth,
    allTeams: getTeamNames_(rawData.records)
  };
}

function uploadFile(payload) {
  setupWorkbook();
  if (!payload || !payload.name || !payload.data) {
    throw new Error('ไม่พบไฟล์ที่อัปโหลด');
  }

  const bytes = Utilities.base64Decode(payload.data);
  const blob = Utilities.newBlob(bytes, payload.mimeType || MimeType.MICROSOFT_EXCEL, payload.name);
  const folder = ensureUploadFolder_();
  const rawFile = folder.createFile(blob);

  try {
    const result = importDriveFile_(rawFile.getId(), rawFile.getName(), rawFile.getUrl());
    // Invalidate cache after upload
    invalidateCache_();
    const rawData = readRawData_();
    return Object.assign({
      ok: true,
      fileName: rawFile.getName(),
      result
    }, buildInitialDataPayload_(rawData));
  } catch (err) {
    appendHistory_(rawFile.getName(), 0, 0, 0, 0, 'ERROR: ' + err.message, rawFile.getUrl());
    throw err;
  }
}

function importExistingDriveFile(fileId) {
  setupWorkbook();
  const file = DriveApp.getFileById(fileId);
  try {
    const result = importDriveFile_(fileId, file.getName(), file.getUrl());
    invalidateCache_();
    const rawData = readRawData_();
    return Object.assign({ ok: true, fileName: file.getName(), result }, buildInitialDataPayload_(rawData));
  } catch (err) {
    appendHistory_(file.getName(), 0, 0, 0, 0, 'ERROR: ' + err.message, file.getUrl());
    throw err;
  }
}

// ── Optimized searchRecords with caching ──
function searchRecords(filters) {
  const rawData = readRawData_();
  const records = filterRecords_(rawData.records, rawData.headers, filters || {});
  return {
    rows: records.slice(0, APP.searchLimit),
    total: records.length,
    columns: rawData.columns,
    allTeams: getTeamNames_(rawData.records),
    dashboard: buildDashboardData_(records, rawData.columns)
  };
}

function getImportHistory() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(APP.historySheet);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastRow = sheet.getLastRow();
  const startRow = Math.max(2, lastRow - APP.historyLimit + 1);
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 8).getDisplayValues();
  return values.reverse().map(row => ({
    importAt: row[0],
    fileName: row[1],
    totalRows: row[2],
    importedRows: row[3],
    skippedTermination: row[4],
    duplicateRows: row[5],
    status: row[6],
    sourceUrl: row[7]
  }));
}

function getFilteredRecords_(filters) {
  const rawData = readRawData_();
  return filterRecords_(rawData.records, rawData.headers, filters || {});
}

// ── Optimized readRawData_ with dual cache (memory + CacheService) ──
function readRawData_() {
  // 1. Memory cache (same execution)
  if (_memoryCache) return _memoryCache;

  // 2. CacheService cache (cross-execution, ~5 min)
  try {
    const cached = CacheService.getScriptCache().get(APP.cacheKey);
    if (cached) {
      _memoryCache = JSON.parse(cached);
      return _memoryCache;
    }
  } catch (e) { /* ignore cache errors */ }

  // 3. Read from sheet (slow path)
  const raw = SpreadsheetApp.getActive().getSheetByName(APP.rawSheet);
  if (!raw || raw.getLastRow() < 1 || raw.getLastColumn() < 1) {
    return { headers: [], columns: [], records: [] };
  }
  const values = raw.getRange(1, 1, raw.getLastRow(), raw.getLastColumn()).getDisplayValues();
  const headers = values[0].map(String);
  const columns = headers.filter(String);
  const records = values.slice(1)
    .filter(row => row.some(v => String(v).trim() !== ''))
    .map(row => rowToObject_(headers, row));
  const result = { headers, columns, records };

  // Save to both caches
  _memoryCache = result;
  try {
    const serialized = JSON.stringify(result);
    if (serialized.length < 100000) { // CacheService limit
      CacheService.getScriptCache().put(APP.cacheKey, serialized, APP.cacheDuration);
    }
  } catch (e) { /* ignore */ }

  return result;
}

function invalidateCache_() {
  _memoryCache = null;
  try { CacheService.getScriptCache().remove(APP.cacheKey); } catch (e) {}
}

function filterRecords_(records, headers, filters) {
  records = records || [];
  headers = (headers || []).filter(String);
  filters = filters || {};
  const query = String(filters.query || '').trim().toLowerCase();
  const fromDate = String(filters.fromDate || '').trim();
  const toDate = String(filters.toDate || '').trim();
  const columnFilters = filters.columnFilters || {};

  return records.filter(record => {
    const appointmentDate = normalizeDate_(record['Appointment Date']);
    if (query && !headers.some(h => String(record[h] || '').toLowerCase().includes(query))) return false;
    if ((fromDate || toDate) && !isNormalizedDate_(appointmentDate)) return false;
    if (fromDate && appointmentDate < fromDate) return false;
    if (toDate && appointmentDate > toDate) return false;
    return Object.keys(columnFilters).every(col => {
      const val = String(columnFilters[col] || '').trim().toLowerCase();
      return !val || String(record[col] || '').toLowerCase().includes(val);
    });
  });
}

// ── Helper functions ──

function buildDashboardData_(records, columns) {
  records = records || [];
  const byHandler = groupBy_(records, 'Handler');
  const byFlag = groupBy_(records, 'Install Flag');
  const byProvince = groupBy_(records, 'Province');
  const byDistrict = groupBy_(records, 'District');
  const bySubDistrict = groupBy_(records, 'Sub-District');
  const byArea = groupBy_(records, 'Area');
  const byTeam = buildTeamData_(records);
  const daily = buildDailyData_(records);
  const calendarDays = buildCalendarDays_(records);
  const handlerStats = buildHandlerStats_(records);
  const handlerLocations = buildHandlerLocations_(records);
  const detailedInsights = buildDetailedInsights_(records, byHandler, byFlag, byArea, byDistrict, bySubDistrict);
  const sampleRows = records.slice(0, APP.dashboardSampleLimit).map(r => r);
  const monthlyLocations = buildMonthlyLocationData_(records, columns);

  return {
    total: records.length,
    byHandler, byFlag, byProvince, byDistrict, bySubDistrict, byArea, byTeam,
    daily, calendarDays, handlerStats, handlerLocations,
    detailedInsights, sampleRows, monthlyLocations, columns
  };
}

function buildTeamData_(records) {
  const teamMap = {};
  records.forEach(r => {
    const team = r['Team'] || '(blank)';
    if (!teamMap[team]) teamMap[team] = { team, total: 0, installation: 0, change: 0, handlerCount: 0, topHandlers: '', topAreas: '' };
    teamMap[team].total++;
    if (String(r['Install Flag']).toLowerCase() === 'installation') teamMap[team].installation++;
    if (String(r['Install Flag']).toLowerCase() === 'change') teamMap[team].change++;
  });
  Object.values(teamMap).forEach(t => {
    const teamRecords = records.filter(r => (r['Team'] || '(blank)') === t.team);
    t.handlerCount = new Set(teamRecords.map(r => r['Handler'])).size;
    t.topHandlers = groupBy_(teamRecords, 'Handler').slice(0, 5).map(x => x.name + ' (' + x.count + ')').join(', ');
    t.topAreas = groupBy_(teamRecords, 'Area').slice(0, 5).map(x => x.name + ' (' + x.count + ')').join(', ');
  });
  return Object.values(teamMap).sort((a, b) => b.total - a.total || a.team.localeCompare(b.team));
}

function buildDailyData_(records) {
  const dayMap = {};
  records.forEach(r => {
    const date = normalizeDate_(r['Appointment Date']);
    if (!date) return;
    if (!dayMap[date]) dayMap[date] = { date, total: 0, installation: 0, change: 0 };
    dayMap[date].total++;
    if (String(r['Install Flag']).toLowerCase() === 'installation') dayMap[date].installation++;
    if (String(r['Install Flag']).toLowerCase() === 'change') dayMap[date].change++;
  });
  return Object.values(dayMap).sort((a, b) => a.date.localeCompare(b.date));
}

function buildCalendarDays_(records) {
  const dayMap = {};
  records.forEach(r => {
    const date = normalizeDate_(r['Appointment Date']);
    if (!date) return;
    if (!dayMap[date]) {
      dayMap[date] = {
        date, total: 0, installation: 0, change: 0,
        handlerCount: 0, teamCount: 0,
        topHandlers: '', topTeams: '', topAreas: '',
        subDistricts: [], districts: [], provinces: [], teams: [], handlers: [],
        rows: [], areaText: ''
      };
    }
    const d = dayMap[date];
    d.total++;
    const flag = String(r['Install Flag']).toLowerCase();
    if (flag === 'installation') d.installation++;
    if (flag === 'change') d.change++;
    d.rows.push({
      team: r['Team'], handler: r['Handler'], flag: r['Install Flag'],
      subDistrict: r['Sub-District'], district: r['District'],
      province: r['Province'], workOrder: r['Work Order No.']
    });
  });
  Object.values(dayMap).forEach(d => {
    d.handlers = groupBy_(d.rows, 'handler');
    d.teams = groupBy_(d.rows, 'team');
    d.subDistricts = groupBy_(d.rows, 'subDistrict');
    d.districts = groupBy_(d.rows, 'district');
    d.provinces = groupBy_(d.rows, 'province');
    const areas = groupBy_(d.rows, 'team');
    d.handlerCount = d.handlers.length;
    d.teamCount = d.teams.length;
    d.topHandlers = d.handlers.slice(0, 4).map(x => x.name + ' ' + x.count).join(', ');
    d.topTeams = d.teams.slice(0, 4).map(x => x.name + ' ' + x.count).join(', ');
    d.topAreas = d.subDistricts.slice(0, 4).map(x => x.name + ' ' + x.count).join(', ');
    d.areaText = d.subDistricts.slice(0, 3).map(x => x.name).join(', ') + ' / ' + d.provinces.length + ' จังหวัด / ' + d.districts.length + ' อำเภอ / ' + d.subDistricts.length + ' ตำบล';
  });
  return Object.values(dayMap).sort((a, b) => a.date.localeCompare(b.date));
}

function buildHandlerStats_(records) {
  const handlerMap = {};
  records.forEach(r => {
    const handler = r['Handler'] || '(blank)';
    if (!handlerMap[handler]) {
      handlerMap[handler] = {
        handler, total: 0, installation: 0, change: 0,
        teamText: '', primaryTeam: '', subDistrictCount: 0, areaCount: 0,
        topAreas: '', topSubDistricts: '', dates: ''
      };
    }
    handlerMap[handler].total++;
    const flag = String(r['Install Flag']).toLowerCase();
    if (flag === 'installation') handlerMap[handler].installation++;
    if (flag === 'change') handlerMap[handler].change++;
  });
  Object.values(handlerMap).forEach(h => {
    const hrs = records.filter(r => (r['Handler'] || '(blank)') === h.handler);
    const teams = groupBy_(hrs, 'Team');
    h.teamText = teams.slice(0, 8).map(x => x.name + ' (' + x.count + ')').join(', ');
    h.primaryTeam = teams[0] ? teams[0].name : '';
    h.subDistrictCount = new Set(hrs.map(r => r['Sub-District'])).size;
    h.areaCount = new Set(hrs.map(r => r['Area'])).size;
    const subDists = groupBy_(hrs, 'Sub-District');
    const areas = groupBy_(hrs, 'Area');
    h.topSubDistricts = subDists.slice(0, 4).map(x => x.name + ' (' + x.count + ')').join(', ');
    h.topAreas = areas.slice(0, 4).map(x => x.name + ' (' + x.count + ')').join(', ');
    h.dates = [...new Set(hrs.map(r => normalizeDate_(r['Appointment Date'])).filter(Boolean))].join(', ');
  });
  return Object.values(handlerMap).sort((a, b) => b.total - a.total);
}

function buildHandlerLocations_(records) {
  const handlerMap = {};
  records.forEach(r => {
    const handler = r['Handler'] || '(blank)';
    if (!handlerMap[handler]) handlerMap[handler] = [];
    handlerMap[handler].push(r);
  });
  return Object.entries(handlerMap).map(([handler, hrs]) => {
    const areaGroups = groupBy_(hrs, 'Area');
    return {
      handler,
      locations: areaGroups.map(x => {
        const sample = hrs.find(r => r.Area === x.name) || {};
        return {
          location: x.name + ' | ' + (sample['Sub-District'] || '') + ' / ' + (sample['District'] || '') + ' / ' + (sample['Province'] || ''),
          count: x.count,
          dates: [...new Set(hrs.map(r => normalizeDate_(r['Appointment Date'])).filter(Boolean))].join(', ')
        };
      })
    };
  });
}

function buildDetailedInsights_(records, byHandler, byFlag, byArea, byDistrict, bySubDistrict) {
  return {
    total: records.length,
    installation: (byFlag.find(x => x.name === 'Installation') || {}).count || 0,
    change: (byFlag.find(x => x.name === 'Change') || {}).count || 0,
    handlerCount: byHandler.length,
    areaCount: byArea.length,
    subDistrictCount: bySubDistrict.length,
    topHandler: byHandler[0] || null,
    topArea: byArea[0] || null,
    topDistrict: byDistrict[0] || null,
    topSubDistrict: bySubDistrict[0] || null
  };
}

function buildMonthlyLocationData_(records, columns) {
  const monthMap = {};
  records.forEach(r => {
    const date = normalizeDate_(r['Appointment Date']);
    if (!date) return;
    const monthKey = date.substring(0, 7);
    if (!monthMap[monthKey]) monthMap[monthKey] = [];
    monthMap[monthKey].push(r);
  });
  return Object.entries(monthMap).map(([month, monthRecords]) => {
    const colR = columns[APP.addressColumnIndex] || 'Column R';
    const subDists = groupBy_(monthRecords, 'Sub-District');
    const dists = groupBy_(monthRecords, 'District');
    const provs = groupBy_(monthRecords, 'Province');
    const installationRecords = monthRecords.filter(r => String(r['Install Flag']).toLowerCase() === 'installation');
    const changeRecords = monthRecords.filter(r => String(r['Install Flag']).toLowerCase() === 'change');
    return {
      month,
      total: monthRecords.length,
      withAddress: monthRecords.filter(r => r[colR]).length,
      subDistrictTotal: subDists.reduce((s, x) => s + x.count, 0),
      districtTotal: dists.reduce((s, x) => s + x.count, 0),
      provinceTotal: provs.reduce((s, x) => s + x.count, 0),
      subDistrictUnique: subDists.length,
      districtUnique: dists.length,
      provinceUnique: provs.length,
      subDistricts: subDists.slice(0, 20),
      districts: dists.slice(0, 20),
      provinces: provs.slice(0, 20),
      subDistrictBreakdown: subDists.map(x => {
        const recs = monthRecords.filter(r => r['Sub-District'] === x.name);
        return { name: x.name, installation: recs.filter(r => String(r['Install Flag']).toLowerCase() === 'installation').length, change: recs.filter(r => String(r['Install Flag']).toLowerCase() === 'change').length, total: x.count };
      }),
      districtBreakdown: dists.map(x => {
        const recs = monthRecords.filter(r => r['District'] === x.name);
        return { name: x.name, installation: recs.filter(r => String(r['Install Flag']).toLowerCase() === 'installation').length, change: recs.filter(r => String(r['Install Flag']).toLowerCase() === 'change').length, total: x.count };
      }),
      provinceBreakdown: provs.map(x => {
        const recs = monthRecords.filter(r => r['Province'] === x.name);
        return { name: x.name, installation: recs.filter(r => String(r['Install Flag']).toLowerCase() === 'installation').length, change: recs.filter(r => String(r['Install Flag']).toLowerCase() === 'change').length, total: x.count };
      }),
      needReview: 0,
      reviewRows: [],
      columnRHeader: colR,
      byType: {
        installation: { total: installationRecords.length },
        change: { total: changeRecords.length }
      }
    };
  });
}

function groupBy_(records, key) {
  const map = {};
  records.forEach(r => {
    const name = r[key] || '(blank)';
    if (!map[name]) map[name] = { name, count: 0 };
    map[name].count++;
  });
  return Object.values(map).sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
}

function getTeamNames_(records) {
  return [...new Set(records.map(r => r['Team']).filter(Boolean))].sort();
}

function getDefaultMonth_(records) {
  const dates = records.map(r => normalizeDate_(r['Appointment Date'])).filter(Boolean).sort();
  return dates.length ? dates[dates.length - 1].substring(0, 7) : '';
}

function getMonthFilters_(monthKey) {
  const [year, month] = monthKey.split('-').map(Number);
  const start = year + '-' + pad_(month) + '-01';
  const end = year + '-' + pad_(month) + '-' + pad_(new Date(year, month, 0).getDate());
  return { fromDate: start, toDate: end };
}

function normalizeDate_(value) {
  if (!value) return '';
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(text)) {
    const [d, m, y] = text.split('/');
    return y + '-' + pad_(+m) + '-' + pad_(+d);
  }
  const d = new Date(text);
  if (isNaN(d.getTime())) return '';
  return d.getFullYear() + '-' + pad_(d.getMonth() + 1) + '-' + pad_(d.getDate());
}

function isNormalizedDate_(dateStr) {
  return /^\d{4}-\d{2}-\d{2}$/.test(dateStr);
}

function pad_(n) { return String(n).padStart(2, '0'); }

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => { if (h) obj[h] = row[i]; });
  return obj;
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureUploadFolder_() {
  const folders = DriveApp.getFoldersByName(APP.defaultFolderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(APP.defaultFolderName);
}

function ensureConfigValue_(config, key, defaultValue) {
  const values = config.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === key) return values[i][1];
  }
  config.appendRow([key, defaultValue]);
  return defaultValue;
}

function appendHistory_(fileName, totalRows, importedRows, skipped, duplicates, status, sourceUrl) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(APP.historySheet);
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([timestamp, fileName, totalRows, importedRows, skipped, duplicates, status, sourceUrl || '']);
}

function importDriveFile_(fileId, fileName, fileUrl) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const contentType = blob.getContentType() || '';
  let rows;

  if (contentType.includes('csv') || fileName.toLowerCase().endsWith('.csv')) {
    rows = parseCSV_(blob.getDataAsString());
  } else {
    const ss = SpreadsheetApp.open(file);
    const sheet = ss.getSheets()[0];
    rows = sheet.getDataRange().getValues();
  }

  if (!rows || rows.length < 2) {
    throw new Error('ไฟล์ไม่มีข้อมูล (น้อยกว่า 2 แถว)');
  }

  const raw = SpreadsheetApp.getActive().getSheetByName(APP.rawSheet);
  const existingHeaders = raw.getLastRow() > 0 ? raw.getRange(1, 1, 1, raw.getLastColumn()).getValues()[0] : [];
  const newHeaders = rows[0].map(String);
  const dataRows = rows.slice(1);

  // Add metadata columns
  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const metaStartCol = newHeaders.length + 1;

  if (existingHeaders.length === 0) {
    const allHeaders = newHeaders.concat(APP.metaHeaders);
    raw.appendRow(allHeaders);
  }

  const finalHeaders = raw.getRange(1, 1, 1, raw.getLastColumn()).getValues()[0];
  const headerMap = {};
  finalHeaders.forEach((h, i) => { if (h) headerMap[h] = i; });

  let imported = 0, skipped = 0, duplicates = 0;
  const existingHashes = new Set();

  // Get existing hashes for duplicate detection
  if (raw.getLastRow() > 1) {
    const hashCol = headerMap['_Row Hash'];
    if (hashCol !== undefined) {
      const hashes = raw.getRange(2, hashCol + 1, raw.getLastRow() - 1, 1).getValues();
      hashes.forEach(h => { if (h[0]) existingHashes.add(String(h[0])); });
    }
  }

  const terminationCol = headerMap['Operation Status'];
  const newRows = [];

  dataRows.forEach(row => {
    const rowHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, row.join('|')).map(b => (b + 256) % 256).map(b => b.toString(16).padStart(2, '0')).join('');

    // Check termination
    if (terminationCol !== undefined) {
      const status = String(row[terminationCol] || '').toLowerCase();
      if (status.includes(APP.terminationText)) { skipped++; return; }
    }

    // Check duplicate
    if (existingHashes.has(rowHash)) { duplicates++; return; }
    existingHashes.add(rowHash);

    const outRow = finalHeaders.map((h, i) => {
      if (APP.metaHeaders.includes(h)) {
        if (h === '_Import File') return fileName;
        if (h === '_Import At') return timestamp;
        if (h === '_Source File ID') return fileId;
        if (h === '_Source URL') return fileUrl;
        if (h === '_Row Hash') return rowHash;
        return '';
      }
      const srcIdx = newHeaders.indexOf(h);
      return srcIdx >= 0 ? row[srcIdx] : '';
    });
    newRows.push(outRow);
    imported++;
  });

  // Batch append rows
  if (newRows.length > 0) {
    const startRow = raw.getLastRow() + 1;
    raw.getRange(startRow, 1, newRows.length, finalHeaders.length).setValues(newRows);
  }

  appendHistory_(fileName, dataRows.length, imported, skipped, duplicates, 'OK', fileUrl);
  return { importedRows: imported, duplicateRows: duplicates, skippedTermination: skipped, totalRows: dataRows.length };
}

function parseCSV_(text) {
  const lines = text.split(/\r?\n/);
  const result = [];
  lines.forEach(line => {
    if (!line.trim()) return;
    const row = [];
    let inQuote = false, cell = '';
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') { inQuote = !inQuote; continue; }
      if (ch === ',' && !inQuote) { row.push(cell); cell = ''; continue; }
      cell += ch;
    }
    row.push(cell);
    result.push(row);
  });
  return result;
}
