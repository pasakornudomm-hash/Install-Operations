const fs = require('fs');
let c = fs.readFileSync('index.html', 'utf8');
c = c.replace(/\r\n/g, '\n');

// 1. Add debounce + cache helpers after state declaration
const stateEnd = c.indexOf('chartsReady: false\n    };');
if (stateEnd === -1) { console.error('Cannot find state'); process.exit(1); }

const helpers = [
  '',
  '    /* debounce + filter cache */',
  '    let _filterTimer = null;',
  '    function debounce(fn, ms) { return function() { clearTimeout(_filterTimer); _filterTimer = setTimeout(() => fn.apply(this, arguments), ms); }; }',
  '    function filterCacheKey(f) { return JSON.stringify({ m: f.fromDate, t: f.toDate, c: f.columnFilters, q: f.query }); }',
  '    function getFilterCache(key, maxAge) {',
  '      try { const c = JSON.parse(sessionStorage.getItem("fc_" + key) || "null"); return c && Date.now() - c.at < maxAge ? c.data : null; } catch(_) { return null; }',
  '    }',
  '    function setFilterCache(key, data) {',
  '      try { sessionStorage.setItem("fc_" + key, JSON.stringify({ at: Date.now(), data: data })); } catch(_) {}',
  '    }',
  ''
].join('\n');

c = c.slice(0, stateEnd) + helpers + c.slice(stateEnd);

// 2. Replace applyGlobalFilters
const oldStart = '    function applyGlobalFilters(showStatus) {';
const oldEnd = '        .searchRecords(filters);\n    }';
const oldIdx = c.indexOf(oldStart);
const oldEndIdx = c.indexOf(oldEnd, oldIdx);
if (oldIdx === -1 || oldEndIdx === -1) { console.error('Cannot find applyGlobalFilters'); process.exit(1); }

const lines = [
  '    function applyGlobalFilters(showStatus) {',
  '      if (state.applyingFilters) return;',
  '      state.applyingFilters = true;',
  '      const filters = getGlobalFilters();',
  '      updateGlobalFilterNote();',
  '',
  '      /* instant: show local data immediately */',
  '      const localRows = filterRowsLocally(state.allRows || [], filters);',
  '      if (localRows.length > 0 || !hasAppsScript()) {',
  '        state.dashboard = summarizeMock(localRows);',
  '        state.currentMonth = selectedMonthDate();',
  '        drawDashboard(); renderCalendar(); renderHandlers(); renderSearchPreview();',
  '        const lt = [...new Set(localRows.map(r => r["Team"]).filter(Boolean))].sort();',
  '        if (lt.length && !filters.columnFilters.Team) { state.allTeams = lt; renderGlobalTeamOptions(); }',
  '      }',
  '',
  '      /* check client cache */',
  '      const ck = filterCacheKey(filters);',
  '      const cached = getFilterCache(ck, 30 * 60 * 1000);',
  '      if (cached) {',
  '        state.dashboard = cached.dashboard || state.dashboard;',
  '        state.currentMonth = selectedMonthDate();',
  '        if (cached.allTeams) { state.allTeams = cached.allTeams; renderGlobalTeamOptions(); }',
  '        drawDashboard(); renderCalendar(); renderHandlers();',
  '        if (cached.rows) renderSearchResult({ rows: cached.rows, total: cached.total, columns: cached.columns });',
  '        setStatus("\\u0e41\\u0e2a\\u0e14\\u0e07\\u0e02\\u0e49\\u0e21\\u0e39\\u0e15\\u0e32\\u0e25\\u0e31\\u0e07 (\\u0e41\\u0e04\\u0e27)");',
  '        state.applyingFilters = false;',
  '        return;',
  '      }',
  '',
  '      /* background API call */',
  '      if (showStatus) setStatus("\\u0e01\\u0e33\\u0e25\\u0e31\\u0e07\\u0e01\\u0e32\\u0e23\\u0e23\\u0e49\\u0e32\\u0e19\\u0e40\\u0e14\\u0e34\\u0e48\\u0e21\\u0e44\\u0e21\\u0e48\\u0e1e\\u0e37\\u0e48\\u0e2d\\u0e40\\u0e17\\u0e35\\u0e48\\u0e22\\u0e27...");',
  '      google.script.run',
  '        .withSuccessHandler(res => {',
  '          state.dashboard = res.dashboard;',
  '          state.currentMonth = selectedMonthDate();',
  '          if ((res.allTeams || []).length && !filters.columnFilters.Team) {',
  '            state.allTeams = res.allTeams;',
  '            renderGlobalTeamOptions();',
  '          }',
  '          drawDashboard(); renderCalendar(); renderHandlers(); renderSearchResult(res);',
  '          setFilterCache(ck, { dashboard: res.dashboard, allTeams: res.allTeams, rows: (res.rows || []).slice(0, 500), total: res.total, columns: res.columns });',
  '          setStatus("");',
  '          state.applyingFilters = false;',
  '        })',
  '        .withFailureHandler(err => {',
  '          state.applyingFilters = false;',
  '          setStatus("\\u0e41\\u0e2a\\u0e14\\u0e07\\u0e02\\u0e49\\u0e21\\u0e39 \\u2014 \\u0e01\\u0e23\\u0e49\\u0e2d Refresh");',
  '        })',
  '        .searchRecords(filters);',
  '    }'
].join('\n');

c = c.substring(0, oldIdx) + lines + c.substring(oldEndIdx + oldEnd.length);

// 3. Add debounce to change events
c = c.replace(
  "document.getElementById('globalMonth').addEventListener('change', () => applyGlobalFilters(true));\n    document.getElementById('globalTeam').addEventListener('change', () => applyGlobalFilters(true));",
  "document.getElementById('globalMonth').addEventListener('change', debounce(() => applyGlobalFilters(true), 400));\n    document.getElementById('globalTeam').addEventListener('change', debounce(() => applyGlobalFilters(true), 400));"
);

// 4. Limit renderTable rows
c = c.replace(
  "    function renderTable(id, headers, rows, allowHtml) {\n      const el = document.getElementById(id);\n      const head = '<thead><tr>' + headers.map(h => '<th>' + escapeHtml(h) + '</th>').join('') + '</tr></thead>';\n      const body = '<tbody>' + rows.map(row => '<tr>' + row.map(v => '<td>' + (allowHtml ? (v || '') : escapeHtml(v)) + '</td>').join('') + '</tr>').join('') + '</tbody>';\n      el.innerHTML = head + body;\n    }",
  '    function renderTable(id, headers, rows, allowHtml) {\n      const el = document.getElementById(id);\n      if (!el) return;\n      const MAX = 400;\n      const shown = rows.length > MAX ? rows.slice(0, MAX) : rows;\n      const head = \'<thead><tr>\' + headers.map(h => \'<th>\' + escapeHtml(h) + \'</th>\').join(\'\') + \'</tr></thead>\';\n      const body = \'<tbody>\' + shown.map(row => \'<tr>\' + row.map(v => \'<td>\' + (allowHtml ? (v || \'\') : escapeHtml(v)) + \'</td>\').join(\'\') + \'</tr>\').join(\'\') + (rows.length > MAX ? \'<tr><td colspan="\' + headers.length + \'" style="text-align:center;color:#94a3b8;font-weight:700">แสดง \' + fmt(MAX) + \' จาก \' + fmt(rows.length) + \' แถว</td></tr>\' : \'\') + \'</tbody>\';\n      el.innerHTML = head + body;\n    }'
);

fs.writeFileSync('index.html', c, 'utf8');
console.log('OK: index.html optimized');
