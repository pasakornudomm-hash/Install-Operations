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
  metaHeaders: ['_Import File', '_Import At', '_Source File ID', '_Source URL', '_Row Hash']
};

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
  const result = importDriveFile_(fileId, file.getName(), file.getUrl());
  const rawData = readRawData_();
  return Object.assign({
    ok: true,
    fileName: file.getName(),
    result
  }, buildInitialDataPayload_(rawData));
}

function importDriveFile_(fileId, fileName, sourceUrl) {
  const ext = String(fileName).split('.').pop().toLowerCase();
  let values;
  let tempSheetId = '';

  if (ext === 'csv') {
    values = Utilities.parseCsv(DriveApp.getFileById(fileId).getBlob().getDataAsString('UTF-8'));
  } else {
    const converted = Drive.Files.copy(
      {
        title: 'TMP_IMPORT_' + new Date().getTime() + '_' + fileName,
        mimeType: MimeType.GOOGLE_SHEETS,
        parents: [{ id: ensureUploadFolder_().getId() }]
      },
      fileId
    );
    tempSheetId = converted.id;
    const temp = SpreadsheetApp.openById(tempSheetId);
    values = temp.getSheets()[0].getDataRange().getDisplayValues();
  }

  try {
    const result = appendCleanRows_(values, fileName, fileId, sourceUrl);
    appendHistory_(
      fileName,
      result.totalRows,
      result.importedRows,
      result.skippedTermination,
      result.duplicateRows,
      'OK',
      sourceUrl
    );
    return result;
  } finally {
    if (tempSheetId) {
      DriveApp.getFileById(tempSheetId).setTrashed(true);
    }
  }
}

function appendCleanRows_(values, fileName, fileId, sourceUrl) {
  if (!values || values.length < 2) throw new Error('ไม่พบข้อมูลในไฟล์');

  const headerRowIndex = findAdaptiveHeaderRow_(values);
  const originalHeaders = uniquifyHeaders_(values[headerRowIndex].map(String));
  const sourceHeaders = uniquifyHeaders_(originalHeaders.map(canonicalHeader_));
  const dataRows = values.slice(headerRowIndex + 1).filter(row => row.some(v => String(v).trim() !== ''));
  const installFlagIndex = findColumnIndex_(sourceHeaders, ['Install Flag']);
  const operationStatusIndex = findColumnIndex_(sourceHeaders, ['Operation Status']);
  const workOrderIndex = findColumnIndex_(sourceHeaders, ['Work Order No.']);
  const accessIndex = findColumnIndex_(sourceHeaders, ['Access Number', 'Service Access No.']);
  if (installFlagIndex < 0 && operationStatusIndex < 0) {
    throw new Error('ไม่พบคอลัมน์ Install Flag หรือ Operation Status');
  }
  if (workOrderIndex < 0 && accessIndex < 0) {
    throw new Error('ไม่พบเลขอ้างอิงงาน: Work Order No., Access Number หรือ Service Access No.');
  }

  const ss = SpreadsheetApp.getActive();
  const raw = getOrCreateSheet_(ss, APP.rawSheet);
  const existingHeaders = getColumns_();
  const desiredHeaders = mergeHeaders_(existingHeaders, sourceHeaders.concat(APP.metaHeaders));

  if (raw.getLastRow() === 0) {
    raw.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
    styleHeader_(raw, desiredHeaders.length);
  } else if (desiredHeaders.length > existingHeaders.length) {
    raw.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
    styleHeader_(raw, desiredHeaders.length);
  }

  const finalHeaders = getColumns_();
  const hashIndex = finalHeaders.indexOf('_Row Hash');
  const existingHashes = getExistingHashes_(raw, hashIndex);
  const existingIds = getExistingWorkIds_(raw, finalHeaders);
  const importAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  let skippedTermination = 0;
  let duplicateRows = 0;
  let invalidRows = 0;
  const warnings = [];
  const rowsToAppend = [];

  dataRows.forEach((row, rowOffset) => {
    const record = {};
    sourceHeaders.forEach((h, i) => record[h] = row[i] === undefined ? '' : String(row[i]).trim());
    normalizeImportedRecord_(record);

    const statusText = [record['Install Flag'], record['Operation Status'], record['Return Reason'], record['Return Reason(Thai)']].join(' ').toLowerCase();
    if (/termination|terminate|ยกเลิก|ถอดถอน/.test(statusText)) {
      skippedTermination++;
      return;
    }

    const workId = normalizeImportKey_(record['Work Order No.'] || record['Access Number'] || record['Service Access No.']);
    if (!workId) {
      invalidRows++;
      if (warnings.length < 10) warnings.push('แถว ' + (headerRowIndex + rowOffset + 2) + ': ไม่มีเลขอ้างอิงงาน');
      return;
    }

    record['_Import File'] = fileName;
    record['_Import At'] = importAt;
    record['_Source File ID'] = fileId;
    record['_Source URL'] = sourceUrl;
    record['_Row Hash'] = makeRowHash_(record);

    if (existingIds.has(workId) || existingHashes.has(record['_Row Hash'])) {
      duplicateRows++;
      return;
    }
    existingIds.add(workId);
    existingHashes.add(record['_Row Hash']);
    rowsToAppend.push(finalHeaders.map(h => record[h] === undefined ? '' : record[h]));
  });

  if (rowsToAppend.length) {
    const startRow = raw.getLastRow() + 1;
    raw.getRange(startRow, 1, rowsToAppend.length, finalHeaders.length).setValues(rowsToAppend);
    if (startRow === 2) raw.autoResizeColumns(1, Math.min(finalHeaders.length, 20));
  }

  return {
    totalRows: dataRows.length,
    importedRows: rowsToAppend.length,
    skippedTermination,
    duplicateRows,
    invalidRows,
    headerRow: headerRowIndex + 1,
    warnings
  };
}

function findAdaptiveHeaderRow_(values) {
  let bestRow = -1;
  let bestScore = -1;
  const limit = Math.min(values.length, 60);
  for (let r = 0; r < limit; r++) {
    const headers = values[r].map(v => canonicalHeader_(String(v)));
    const anchors = ['Work Order No.', 'Access Number', 'Service Access No.', 'Install Flag', 'Operation Status', 'Appointment Date'];
    let score = 0;
    anchors.forEach(h => { if (headers.indexOf(h) >= 0) score += 10; });
    headers.forEach(h => { if (h && h !== 'Column') score++; });
    if (score > bestScore) { bestScore = score; bestRow = r; }
  }
  if (bestRow < 0 || bestScore < 22) throw new Error('หาหัวตารางไม่พบใน 60 แถวแรก');
  return bestRow;
}

function canonicalHeader_(value) {
  const raw = String(value || '').replace(/^\uFEFF/, '').trim();
  const key = raw.toLowerCase().replace(/[\s_\-().\/\\:]+/g, '');
  const aliases = {
    workorderno: 'Work Order No.', workorder: 'Work Order No.', wono: 'Work Order No.', wo: 'Work Order No.',
    'เลขที่ใบงาน': 'Work Order No.', 'เลขที่งาน': 'Work Order No.',
    accessnumber: 'Access Number', accessno: 'Access Number', 'เลขหมาย': 'Access Number',
    serviceaccessno: 'Service Access No.', serviceaccessnumber: 'Service Access No.',
    handler: 'Handler', 'ผู้รับผิดชอบ': 'Handler', 'ช่างผู้รับผิดชอบ': 'Handler',
    appointmentdate: 'Appointment Date', appointmentdatetime: 'Appointment Date', 'วันนัดหมาย': 'Appointment Date', 'วันที่นัดหมาย': 'Appointment Date',
    installflag: 'Install Flag', installationflag: 'Install Flag', 'ประเภทงาน': 'Install Flag', worktype: 'Install Flag',
    operationstatus: 'Operation Status', status: 'Operation Status', 'สถานะงาน': 'Operation Status',
    area: 'Area', 'พื้นที่': 'Area', areathai: 'Area', 'areaภาษาไทย': 'Area',
    address: 'Address', 'ที่อยู่': 'Address', customeraddress: 'Address', installaddress: 'Address',
    province: 'Province', 'จังหวัด': 'Province', changwat: 'Province',
    district: 'District', 'อำเภอ': 'District', amphoe: 'District',
    subdistrict: 'Sub-District', subdistrictname: 'Sub-District', 'ตำบล': 'Sub-District', tambon: 'Sub-District',
    team: 'Team', 'ทีม': 'Team', 'ทีมช่าง': 'Team', technicianteam: 'Team'
  };
  return aliases[key] || raw || 'Column';
}

function normalizeImportedRecord_(record) {
  const flag = String(record['Install Flag'] || '');
  if (/termination|terminate|ยกเลิก|ถอดถอน/i.test(flag)) record['Install Flag'] = 'Termination';
  else if (/change|ย้าย|เปลี่ยน/i.test(flag)) record['Install Flag'] = 'Change';
  else if (/install|ติดตั้ง/i.test(flag)) record['Install Flag'] = 'Installation';
  if (record['Appointment Date']) record['Appointment Date'] = normalizeDate_(record['Appointment Date']);

  if ((!record['Province'] || !record['District'] || !record['Sub-District']) && record['Address']) {
    const location = parseThaiAddress_(String(record['Address']));
    if (!record['Province']) record['Province'] = location.province || '';
    if (!record['District']) record['District'] = location.district || '';
    if (!record['Sub-District']) record['Sub-District'] = location.subDistrict || '';
  }
  if (!record['Area']) record['Area'] = [record['District'], record['Sub-District']].filter(String).join(' - ');
}

function normalizeImportKey_(value) {
  return String(value || '').toLowerCase().replace(/[\s_\-().\/\\:]+/g, '');
}

function getExistingWorkIds_(raw, headers) {
  const ids = new Set();
  if (raw.getLastRow() < 2) return ids;
  const indexes = ['Work Order No.', 'Access Number', 'Service Access No.'].map(h => headers.indexOf(h));
  raw.getRange(2, 1, raw.getLastRow() - 1, headers.length).getDisplayValues().forEach(row => {
    let value = '';
    indexes.some(i => { if (i >= 0 && row[i]) { value = row[i]; return true; } return false; });
    const key = normalizeImportKey_(value);
    if (key) ids.add(key);
  });
  return ids;
}

function getDashboardData(filters) {
  const rawData = readRawData_();
  const records = filterRecords_(rawData.records, rawData.headers, filters || {});
  return buildDashboardData_(records, rawData.columns);
}

function buildDashboardData_(records, columns) {
  records = records || [];
  columns = columns || [];
  const columnRHeader = getColumnRHeader_(columns);
  const total = records.length;
  const byFlag = groupCount_(records, 'Install Flag');
  const byHandler = groupCount_(records, 'Handler').slice(0, 30);
  const byProvince = groupCount_(records, 'Province');
  const byDistrict = groupCount_(records, 'District').slice(0, 40);
  const bySubDistrict = groupCount_(records, 'Sub-District').slice(0, 80);
  const byArea = groupCount_(records, 'Area').slice(0, 80);
  const byTeam = buildTeamStats_(records);
  const daily = groupCount_(records, 'Appointment Date').sort((a, b) => String(a.name).localeCompare(String(b.name)));
  const handlerLocations = buildHandlerLocations_(records);
  const handlerStats = buildHandlerStats_(records);
  const calendarDays = buildCalendarDays_(records, columnRHeader);
  const monthlyLocations = buildMonthlyLocationSummary_(records, columnRHeader);
  const detailedInsights = buildDetailedInsights_(records);

  return {
    total,
    byFlag,
    byHandler,
    byProvince,
    byDistrict,
    bySubDistrict,
    byArea,
    byTeam,
    daily,
    handlerLocations,
    handlerStats,
    calendarDays,
    monthlyLocations,
    columnRHeader,
    detailedInsights,
    sampleRows: records.slice(0, APP.dashboardSampleLimit),
    columns
  };
}

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

function readRawData_() {
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
  return { headers, columns, records };
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

function getDefaultMonth_(records) {
  const months = {};
  (records || []).forEach(record => {
    const date = normalizeDate_(record['Appointment Date']);
    if (/^\d{4}-\d{2}-\d{2}$/.test(date)) months[date.slice(0, 7)] = true;
  });
  const keys = Object.keys(months).sort();
  return keys.length ? keys[keys.length - 1] : '';
}

function getMonthFilters_(month) {
  const parts = String(month || '').split('-').map(Number);
  if (parts.length !== 2 || !parts[0] || !parts[1]) {
    return { query: '', fromDate: '', toDate: '', columnFilters: {} };
  }
  const end = new Date(parts[0], parts[1], 0);
  return {
    query: '',
    fromDate: parts[0] + '-' + ('0' + parts[1]).slice(-2) + '-01',
    toDate: end.getFullYear() + '-' + ('0' + (end.getMonth() + 1)).slice(-2) + '-' + ('0' + end.getDate()).slice(-2),
    columnFilters: {}
  };
}

function getTeamNames_(records) {
  return groupCount_(records || [], 'Team')
    .map(row => row.name)
    .filter(name => name && name !== '(blank)');
}

function buildHandlerLocations_(records) {
  const map = {};
  records.forEach(r => {
    const handler = valueOrBlank_(r['Handler']);
    const areaName = getAreaName_(r);
    const detail = [
      valueOrBlank_(r['Sub-District']),
      valueOrBlank_(r['District']),
      valueOrBlank_(r['Province'])
    ].join(' / ');
    const key = areaName + ' | ' + detail;
    const date = normalizeDate_(r['Appointment Date']) || valueOrBlank_(r['Appointment Date']);
    if (!map[handler]) map[handler] = {};
    if (!map[handler][key]) map[handler][key] = { count: 0, dates: {} };
    map[handler][key].count++;
    if (date) map[handler][key].dates[date] = true;
  });

  return Object.keys(map).sort().map(handler => ({
    handler,
    locations: Object.keys(map[handler]).sort().map(location => ({
      location,
      count: map[handler][location].count,
      dates: Object.keys(map[handler][location].dates).sort().join(', ')
    }))
  }));
}

function buildHandlerStats_(records) {
  const map = {};
  records.forEach(r => {
    const handler = valueOrBlank_(r['Handler']);
    if (!map[handler]) {
      map[handler] = {
        handler,
        total: 0,
        installation: 0,
        change: 0,
        flags: {},
        teams: {},
        areas: {},
        provinces: {},
        districts: {},
        subDistricts: {},
        dates: {},
        workOrders: {}
      };
    }
    const item = map[handler];
    const flag = valueOrBlank_(r['Install Flag']);
    item.total++;
    if (flag.toLowerCase() === 'installation') item.installation++;
    if (flag.toLowerCase() === 'change') item.change++;
    item.flags[flag] = (item.flags[flag] || 0) + 1;
    item.teams[valueOrBlank_(r['Team'])] = (item.teams[valueOrBlank_(r['Team'])] || 0) + 1;
    item.areas[getAreaName_(r)] = (item.areas[getAreaName_(r)] || 0) + 1;
    item.provinces[valueOrBlank_(r['Province'])] = true;
    item.districts[valueOrBlank_(r['District'])] = true;
    item.subDistricts[valueOrBlank_(r['Sub-District'])] = (item.subDistricts[valueOrBlank_(r['Sub-District'])] || 0) + 1;
    item.dates[normalizeDate_(r['Appointment Date']) || valueOrBlank_(r['Appointment Date'])] = true;
    item.workOrders[valueOrBlank_(r['Work Order No.'])] = true;
  });

  return Object.keys(map).map(handler => {
    const item = map[handler];
    return {
      handler,
      total: item.total,
      installation: item.installation,
      change: item.change,
      flagText: Object.keys(item.flags).sort().map(k => k + ': ' + item.flags[k]).join(', '),
      teamText: objectTop_(item.teams, 8).join(', '),
      primaryTeam: objectTop_(item.teams, 1).join(''),
      provinceCount: Object.keys(item.provinces).length,
      districtCount: Object.keys(item.districts).length,
      subDistrictCount: Object.keys(item.subDistricts).length,
      areaCount: Object.keys(item.areas).length,
      dates: Object.keys(item.dates).sort().join(', '),
      topAreas: objectTop_(item.areas, 8).join(', '),
      topSubDistricts: objectTop_(item.subDistricts, 8).join(', ')
    };
  }).sort((a, b) => b.total - a.total || String(a.handler).localeCompare(String(b.handler)));
}

function buildCalendarDays_(records, columnRHeader) {
  const map = {};
  records.forEach(r => {
    const date = normalizeDate_(r['Appointment Date']);
    if (!date) return;
    if (!map[date]) {
      map[date] = {
        date,
        total: 0,
        installation: 0,
        change: 0,
        handlers: {},
        teams: {},
        areas: {},
        provinces: {},
        districts: {},
        subDistricts: {},
        rows: []
      };
    }
    const item = map[date];
    const flag = valueOrBlank_(r['Install Flag']).toLowerCase();
    const handler = valueOrBlank_(r['Handler']);
    const team = valueOrBlank_(r['Team']);
    const location = parseColumnRLocation_(r, columnRHeader);
    item.total++;
    if (flag === 'installation') item.installation++;
    if (flag === 'change') item.change++;
    item.handlers[handler] = (item.handlers[handler] || 0) + 1;
    item.teams[team] = (item.teams[team] || 0) + 1;
    item.areas[getAreaName_(r)] = (item.areas[getAreaName_(r)] || 0) + 1;
    if (location.province) addCount_(item.provinces, location.province);
    if (location.district) addCount_(item.districts, location.district);
    if (location.subDistrict) addCount_(item.subDistricts, location.subDistrict);
    if (item.rows.length < 200) {
      item.rows.push({
        team,
        handler,
        flag: valueOrBlank_(r['Install Flag']),
        subDistrict: location.subDistrict || valueOrBlank_(r['Sub-District']),
        district: location.district || valueOrBlank_(r['District']),
        province: location.province || valueOrBlank_(r['Province']),
        workOrder: valueOrBlank_(r['Work Order No.']),
        address: getColumnRValue_(r, columnRHeader)
      });
    }
  });

  return Object.keys(map).sort().map(date => {
    const item = map[date];
    return {
      date,
      total: item.total,
      installation: item.installation,
      change: item.change,
      handlerCount: Object.keys(item.handlers).length,
      teamCount: Object.keys(item.teams).length,
      topHandlers: objectTop_(item.handlers, 5).join(', '),
      topTeams: objectTop_(item.teams, 5).join(', '),
      topAreas: objectTop_(item.areas, 5).join(', '),
      subDistricts: objectCountRows_(item.subDistricts),
      districts: objectCountRows_(item.districts),
      provinces: objectCountRows_(item.provinces),
      teams: objectCountRows_(item.teams),
      handlers: objectCountRows_(item.handlers),
      rows: item.rows,
      areaText: [
        objectTop_(item.areas, 3).join(', '),
        Object.keys(item.provinces).length + ' จังหวัด',
        Object.keys(item.districts).length + ' อำเภอ',
        Object.keys(item.subDistricts).length + ' ตำบล'
      ].filter(String).join(' / ')
    };
  });
}

function buildTeamStats_(records) {
  const map = {};
  records.forEach(r => {
    const team = valueOrBlank_(r['Team']);
    if (!map[team]) {
      map[team] = {
        team,
        total: 0,
        installation: 0,
        change: 0,
        handlers: {},
        areas: {},
        dates: {}
      };
    }
    const item = map[team];
    const flag = valueOrBlank_(r['Install Flag']).toLowerCase();
    item.total++;
    if (flag === 'installation') item.installation++;
    if (flag === 'change') item.change++;
    item.handlers[valueOrBlank_(r['Handler'])] = (item.handlers[valueOrBlank_(r['Handler'])] || 0) + 1;
    item.areas[getAreaName_(r)] = (item.areas[getAreaName_(r)] || 0) + 1;
    item.dates[normalizeDate_(r['Appointment Date']) || valueOrBlank_(r['Appointment Date'])] = true;
  });

  return Object.keys(map).map(team => {
    const item = map[team];
    return {
      team,
      total: item.total,
      installation: item.installation,
      change: item.change,
      handlerCount: Object.keys(item.handlers).length,
      topHandlers: objectTop_(item.handlers, 6).join(', '),
      topAreas: objectTop_(item.areas, 6).join(', '),
      dates: Object.keys(item.dates).sort().join(', ')
    };
  }).sort((a, b) => b.total - a.total || String(a.team).localeCompare(String(b.team)));
}

function buildDetailedInsights_(records) {
  const total = records.length;
  const change = records.filter(r => valueOrBlank_(r['Install Flag']).toLowerCase() === 'change').length;
  const installation = records.filter(r => valueOrBlank_(r['Install Flag']).toLowerCase() === 'installation').length;
  const handlers = groupCount_(records, 'Handler');
  const provinces = groupCount_(records, 'Province');
  const districts = groupCount_(records, 'District');
  const subDistricts = groupCount_(records, 'Sub-District');
  const areas = groupCount_(records, 'Area');
  const dates = groupCount_(records, 'Appointment Date');

  return {
    total,
    installation,
    change,
    handlerCount: handlers.length,
    provinceCount: provinces.length,
    districtCount: districts.length,
    subDistrictCount: subDistricts.length,
    areaCount: areas.length,
    dayCount: dates.length,
    topHandler: handlers[0] || { name: '-', count: 0 },
    topArea: areas[0] || { name: '-', count: 0 },
    topProvince: provinces[0] || { name: '-', count: 0 },
    topDistrict: districts[0] || { name: '-', count: 0 },
    topSubDistrict: subDistricts[0] || { name: '-', count: 0 }
  };
}

function buildMonthlyLocationSummary_(records, columnRHeader) {
  const map = {};
  records.forEach(r => {
    const date = normalizeDate_(r['Appointment Date']);
    if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return;

    const month = date.slice(0, 7);
    if (!map[month]) map[month] = createMonthlyLocationBucket_(month);

    const item = map[month];
    const rawAddress = getColumnRValue_(r, columnRHeader);
    const location = parseThaiAddress_(rawAddress);
    const typeKey = getWorkTypeKey_(r['Install Flag']);

    addLocationStats_(item, location);
    if (typeKey && item.byType[typeKey]) {
      addLocationStats_(item.byType[typeKey], location);
    }

    if (!location.complete) {
      if (item.reviewRows.length < 20) {
        item.reviewRows.push({
          date,
          flag: valueOrBlank_(r['Install Flag']),
          address: rawAddress || '(blank)',
          found: [
            location.subDistrict ? 'ต.' + location.subDistrict : '',
            location.district ? 'อ.' + location.district : '',
            location.province ? 'จ.' + location.province : ''
          ].filter(String).join(' / ') || '-'
        });
      }
    }
  });

  return Object.keys(map).sort().map(month => {
    const item = map[month];
    const installation = item.byType.installation;
    const change = item.byType.change;
    const out = serializeLocationStats_(item);
    out.month = month;
    out.byType = {
      installation: serializeLocationStats_(installation),
      change: serializeLocationStats_(change)
    };
    out.subDistrictBreakdown = buildLocationBreakdownRows_(installation.subDistricts, change.subDistricts);
    out.districtBreakdown = buildLocationBreakdownRows_(installation.districts, change.districts);
    out.provinceBreakdown = buildLocationBreakdownRows_(installation.provinces, change.provinces);
    return out;
  });
}

function createMonthlyLocationBucket_(month) {
  const item = createLocationStatsBucket_();
  item.month = month;
  item.byType = {
    installation: createLocationStatsBucket_(),
    change: createLocationStatsBucket_()
  };
  return item;
}

function createLocationStatsBucket_() {
  return {
    total: 0,
    withAddress: 0,
    subDistrictTotal: 0,
    districtTotal: 0,
    provinceTotal: 0,
    completeTotal: 0,
    needReview: 0,
    subDistricts: {},
    districts: {},
    provinces: {},
    reviewRows: []
  };
}

function addLocationStats_(item, location) {
  item.total++;
  if (location.raw) item.withAddress++;
  if (location.subDistrict) {
    item.subDistrictTotal++;
    addCount_(item.subDistricts, location.subDistrict);
  }
  if (location.district) {
    item.districtTotal++;
    addCount_(item.districts, location.district);
  }
  if (location.province) {
    item.provinceTotal++;
    addCount_(item.provinces, location.province);
  }
  if (location.complete) {
    item.completeTotal++;
  } else {
    item.needReview++;
  }
}

function serializeLocationStats_(item) {
  return {
    total: item.total,
    withAddress: item.withAddress,
    subDistrictTotal: item.subDistrictTotal,
    districtTotal: item.districtTotal,
    provinceTotal: item.provinceTotal,
    completeTotal: item.completeTotal,
    needReview: item.needReview,
    subDistrictUnique: Object.keys(item.subDistricts).length,
    districtUnique: Object.keys(item.districts).length,
    provinceUnique: Object.keys(item.provinces).length,
    subDistricts: objectCountRows_(item.subDistricts),
    districts: objectCountRows_(item.districts),
    provinces: objectCountRows_(item.provinces),
    reviewRows: item.reviewRows || []
  };
}

function buildLocationBreakdownRows_(installationMap, changeMap) {
  const keys = {};
  Object.keys(installationMap || {}).forEach(name => keys[name] = true);
  Object.keys(changeMap || {}).forEach(name => keys[name] = true);
  return Object.keys(keys)
    .map(name => {
      const installation = Number((installationMap || {})[name] || 0);
      const change = Number((changeMap || {})[name] || 0);
      return { name, installation, change, total: installation + change };
    })
    .sort((a, b) => b.total - a.total || b.installation - a.installation || String(a.name).localeCompare(String(b.name)));
}

function getWorkTypeKey_(value) {
  const flag = String(value || '').trim().toLowerCase();
  if (flag === 'installation') return 'installation';
  if (flag === 'change') return 'change';
  return '';
}

function parseColumnRLocation_(record, columnRHeader) {
  return parseThaiAddress_(getColumnRValue_(record, columnRHeader));
}

function getColumnRHeader_(columns) {
  return (columns || [])[APP.addressColumnIndex - 1] || '';
}

function getColumnRValue_(record, columnRHeader) {
  if (!columnRHeader) return '';
  return record[columnRHeader] === undefined ? '' : record[columnRHeader];
}

function parseThaiAddress_(value) {
  const raw = String(value || '').trim();
  const text = normalizeThaiAddressText_(raw);
  if (!text) {
    return { raw, subDistrict: '', district: '', province: '', complete: false };
  }

  const subDistrict = cleanAddressPart_(extractThaiAddressPart_(
    text,
    'ตำบล|ต\\.|แขวง',
    'อำเภอ|อ\\.|เขต|จังหวัด|จ\\.|กรุงเทพฯ|กรุงเทพมหานคร|กทม\\.?|\\d{5}|$'
  ));
  const district = cleanAddressPart_(extractThaiAddressPart_(
    text,
    'อำเภอ|อ\\.|เขต',
    'ตำบล|ต\\.|แขวง|จังหวัด|จ\\.|กรุงเทพฯ|กรุงเทพมหานคร|กทม\\.?|\\d{5}|$'
  ));
  let province = cleanAddressPart_(extractThaiAddressPart_(
    text,
    'จังหวัด|จ\\.',
    'รหัส|ไปรษณีย์|ประเทศไทย|\\d{5}|$'
  ));

  if (!province && /(กรุงเทพฯ|กรุงเทพมหานคร|กทม\.?)/i.test(text)) {
    province = 'กรุงเทพมหานคร';
  }

  if (!subDistrict || !district || !province) {
    const fallback = parseUnlabeledThaiAddress_(text);
    return {
      raw,
      subDistrict: subDistrict || fallback.subDistrict,
      district: district || fallback.district,
      province: province || fallback.province,
      complete: Boolean((subDistrict || fallback.subDistrict) && (district || fallback.district) && (province || fallback.province))
    };
  }

  return {
    raw,
    subDistrict,
    district,
    province,
    complete: Boolean(subDistrict && district && province)
  };
}

function normalizeThaiAddressText_(value) {
  return String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/[，、]/g, ',')
    .replace(/[|]/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/(^|[\s,;/()])ต\s+/g, '$1ต.')
    .replace(/(^|[\s,;/()])อ\s+/g, '$1อ.')
    .replace(/(^|[\s,;/()])จ\s+/g, '$1จ.')
    .trim();
}

function parseUnlabeledThaiAddress_(text) {
  const parts = String(text || '')
    .replace(/\d{5}\b/g, ' ')
    .split(/[\s,;/()]+/)
    .map(cleanAddressPart_)
    .filter(part => part && !isIgnoredAddressToken_(part));

  if (parts.length < 3) {
    return { subDistrict: '', district: '', province: '' };
  }

  const province = parts[parts.length - 1];
  const district = parts[parts.length - 2];
  const subDistrict = parts[parts.length - 3];
  return { subDistrict, district, province };
}

function isIgnoredAddressToken_(value) {
  const text = String(value || '').trim();
  if (!text) return true;
  if (/^\d+(?:\/\d+)*(?:-\d+)?$/.test(text)) return true;
  if (/^(บ้านเลขที่|เลขที่)$/i.test(text)) return true;
  if (/^(หมู่|หมู่ที่|ม\.?)\s*\d*$/i.test(text)) return true;
  if (/^หมู่\d+$/i.test(text)) return true;
  if (/^(ถนน|ถ\.|ซอย|ซ\.|อาคาร|ชั้น|ห้อง)$/i.test(text)) return true;
  return false;
}

function extractThaiAddressPart_(text, labelPattern, stopPattern) {
  const re = new RegExp('(?:^|[\\s,;/()])(?:' + labelPattern + ')\\s*([^,;/()]+?)(?=\\s*(?:' + stopPattern + ')\\s*|\\s*$)', 'i');
  const match = String(text || '').match(re);
  return match ? match[1] : '';
}

function cleanAddressPart_(value) {
  return String(value || '')
    .replace(/^(ตำบล|ต\.|อำเภอ|อ\.|จังหวัด|จ\.|แขวง|เขต)\s*/i, '')
    .replace(/\s*(บ้านเลขที่|เลขที่|หมู่บ้าน|หมู่ที่|หมู่|ม\.|ถนน|ถ\.|ซอย|ซ\.|อาคาร|ชั้น|ห้อง|โทร|เบอร์|รหัส|ไปรษณีย์).*$/i, '')
    .replace(/\d{5}.*$/, '')
    .replace(/^[\s.,:;/-]+|[\s.,:;/-]+$/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function addCount_(obj, key) {
  if (!key) return;
  obj[key] = (obj[key] || 0) + 1;
}

function objectCountRows_(obj) {
  return Object.keys(obj)
    .map(name => ({ name, count: obj[name] }))
    .sort((a, b) => b.count - a.count || String(a.name).localeCompare(String(b.name)));
}

function objectTop_(obj, limit) {
  return Object.keys(obj)
    .map(name => ({ name, count: obj[name] === true ? 1 : obj[name] }))
    .sort((a, b) => b.count - a.count || String(a.name).localeCompare(String(b.name)))
    .slice(0, limit)
    .map(r => r.count === 1 ? r.name : r.name + ' (' + r.count + ')');
}

function getAreaName_(record) {
  const area = String(record['Area'] || '').trim();
  if (area) return area;
  return [
    valueOrBlank_(record['District']),
    valueOrBlank_(record['Sub-District'])
  ].filter(v => v && v !== '(blank)').join(' - ') || '(blank)';
}

function groupCount_(records, column) {
  const map = {};
  records.forEach(r => {
    const key = valueOrBlank_(r[column]);
    map[key] = (map[key] || 0) + 1;
  });
  return Object.keys(map)
    .map(name => ({ name, count: map[name] }))
    .sort((a, b) => b.count - a.count || String(a.name).localeCompare(String(b.name)));
}

function findHeaderRow_(values) {
  for (let i = 0; i < Math.min(values.length, 20); i++) {
    const row = values[i].map(v => String(v).trim().toLowerCase());
    if (row.includes('handler') && row.includes('install flag')) return i;
  }
  return 0;
}

function findColumnIndex_(headers, candidates) {
  const lower = headers.map(h => String(h).trim().toLowerCase());
  for (const candidate of candidates) {
    const idx = lower.indexOf(String(candidate).trim().toLowerCase());
    if (idx >= 0) return idx;
  }
  return -1;
}

function uniquifyHeaders_(headers) {
  const seen = {};
  return headers.map((h, i) => {
    const base = String(h || ('Column ' + (i + 1))).trim();
    seen[base] = (seen[base] || 0) + 1;
    return seen[base] === 1 ? base : base + ' ' + seen[base];
  });
}

function mergeHeaders_(a, b) {
  const out = [];
  a.concat(b).forEach(h => {
    if (h && out.indexOf(h) < 0) out.push(h);
  });
  return out;
}

function getColumns_() {
  const raw = SpreadsheetApp.getActive().getSheetByName(APP.rawSheet);
  if (!raw || raw.getLastRow() === 0) return [];
  return raw.getRange(1, 1, 1, raw.getLastColumn()).getDisplayValues()[0].filter(String);
}

function getExistingHashes_(sheet, hashIndex) {
  const set = new Set();
  if (sheet.getLastRow() < 2 || hashIndex < 0) return set;
  sheet.getRange(2, hashIndex + 1, sheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .forEach(r => { if (r[0]) set.add(r[0]); });
  return set;
}

function makeRowHash_(record) {
  const keys = ['Access Number', 'Service Access No.', 'Work Order No.', 'Appointment Date', 'Handler', 'Install Flag'];
  const raw = keys.map(k => record[k] || '').join('|');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i] || '');
  return obj;
}

function normalizeDate_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const m = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return formatDateParts_(normalizeYear_(m[1]), Number(m[2]), Number(m[3])) || text;

  const dmy = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})(?:\s|$)/);
  if (dmy) {
    let day = Number(dmy[1]);
    let month = Number(dmy[2]);
    const year = normalizeYear_(dmy[3]);
    if (day <= 12 && month > 12) {
      day = Number(dmy[2]);
      month = Number(dmy[1]);
    }
    return formatDateParts_(year, month, day) || text;
  }

  const thaiDate = text.match(/^(\d{1,2})\s*([ก-๙.]+)\s*(\d{2,4})/);
  if (thaiDate) {
    const month = getThaiMonthNumber_(thaiDate[2]);
    const year = normalizeYear_(thaiDate[3]);
    const formatted = formatDateParts_(year, month, Number(thaiDate[1]));
    if (formatted) return formatted;
  }

  const date = new Date(text);
  return isNaN(date.getTime()) ? text : Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function isNormalizedDate_(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || ''));
}

function normalizeYear_(year) {
  let y = Number(year);
  if (y < 100) y += 2000;
  if (y > 2400) y -= 543;
  return y;
}

function formatDateParts_(year, month, day) {
  if (!year || !month || !day) return '';
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return '';
  return [
    String(year),
    ('0' + month).slice(-2),
    ('0' + day).slice(-2)
  ].join('-');
}

function getThaiMonthNumber_(value) {
  const key = String(value || '').replace(/\./g, '').toLowerCase();
  const months = {
    'มกราคม': 1, 'มค': 1,
    'กุมภาพันธ์': 2, 'กพ': 2,
    'มีนาคม': 3, 'มีค': 3,
    'เมษายน': 4, 'เมย': 4,
    'พฤษภาคม': 5, 'พค': 5,
    'มิถุนายน': 6, 'มิย': 6,
    'กรกฎาคม': 7, 'กค': 7,
    'สิงหาคม': 8, 'สค': 8,
    'กันยายน': 9, 'กย': 9,
    'ตุลาคม': 10, 'ตค': 10,
    'พฤศจิกายน': 11, 'พย': 11,
    'ธันวาคม': 12, 'ธค': 12
  };
  return months[key] || 0;
}

function valueOrBlank_(value) {
  const text = String(value || '').trim();
  return text || '(blank)';
}

function appendHistory_(fileName, totalRows, importedRows, skippedTermination, duplicateRows, status, sourceUrl) {
  const sheet = getOrCreateSheet_(SpreadsheetApp.getActive(), APP.historySheet);
  sheet.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    fileName,
    totalRows,
    importedRows,
    skippedTermination,
    duplicateRows,
    status,
    sourceUrl
  ]);
}

function ensureConfigValue_(sheet, key, value) {
  if (!sheet) return;
  if (sheet.getLastRow() >= 2) {
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
    if (values.some(row => row[0] === key)) return;
  }
  sheet.appendRow([key, value]);
}

function ensureUploadFolder_() {
  const ss = SpreadsheetApp.getActive();
  const config = ss.getSheetByName(APP.configSheet);
  if (config && config.getLastRow() >= 2) {
    const values = config.getRange(2, 1, config.getLastRow() - 1, 2).getDisplayValues();
    const row = values.find(r => r[0] === 'Upload Folder ID' && r[1]);
    if (row) {
      try {
        return DriveApp.getFolderById(row[1]);
      } catch (err) {
        // Fall through and create a new folder.
      }
    }
  }
  const folder = DriveApp.createFolder(APP.defaultFolderName + ' - ' + ss.getName());
  if (config) config.appendRow(['Upload Folder ID', folder.getId()]);
  return folder;
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function styleHeader_(sheet, width) {
  sheet.getRange(1, 1, 1, width)
    .setBackground('#E60012')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}


/** REST API for the GitHub Pages frontend. */
function doPost(e) {
  try {
    const request = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const action = String(request.action || "");
    const args = Array.isArray(request.args) ? request.args : [];
    let data;
    switch (action) {
      case "getBootstrapData": data = getBootstrapData(); break;
      case "getDashboardData": data = getDashboardData(args[0] || {}); break;
      case "searchRecords": data = searchRecords(args[0] || {}); break;
      case "uploadFile": data = uploadFile(args[0] || {}); break;
      case "getImportHistory": data = getImportHistory(); break;
      case "importExistingDriveFile": data = importExistingDriveFile(args[0]); break;
      default: throw new Error("ไม่รองรับคำสั่ง API: " + action);
    }
    return jsonResponse_({ ok: true, data: data });
  } catch (error) {
    return jsonResponse_({ ok: false, error: error && error.message ? error.message : String(error) });
  }
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
