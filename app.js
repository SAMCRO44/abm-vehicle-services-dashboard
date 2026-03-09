/* ABM Rental Operations Dashboard - Excel import & charts */
(function () {
  'use strict';

  var XLSX = window.XLSX;
  var Chart = window.Chart;

  function $(id) { return document.getElementById(id); }
  function setText(id, v) { var e = $(id); if (e) e.textContent = v; }

  var state = {
    rawMoves: [],
    taskTypes: [],
    startLocations: [],
    endLocations: [],
    workers: [],
    subMinPivot: [],
    scans: [],
    namesIds: {},
    sourceFiles: []
  };

  var charts = {};
  var PALETTE = ['#22d3ee', '#4ade80', '#fb923c', '#a78bfa', '#f87171', '#e879f9', '#2dd4bf', '#fbbf24'];

  function trimKeys(row) {
    var out = {};
    for (var k in row) {
      var key = (k && typeof k === 'string') ? k.trim() : k;
      out[key] = row[k];
    }
    return out;
  }

  function col(row, names) {
    for (var i = 0; i < names.length; i++) {
      var v = row[names[i]];
      if (v !== undefined && v !== null && v !== '') return v;
    }
    return null;
  }

  function parseNum(v) {
    if (v === null || v === undefined || v === '') return null;
    var n = Number(v);
    return isNaN(n) ? null : n;
  }

  function parseTimestamp(ts) {
    if (!ts) return null;
    if (typeof ts === 'number' && ts > 100000) return new Date((ts - 25569) * 86400 * 1000);
    var d = new Date(ts);
    return isNaN(d.getTime()) ? null : d;
  }

  function addFieldOpsRows(sheet, firstRowIsHeader) {
    var data = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
    var headers = (data[0] || []).map(function (h) { return (h || '').trim(); });
    var start = firstRowIsHeader ? 1 : 0;
    for (var r = start; r < data.length; r++) {
      var row = data[r];
      if (!row || !row.length) continue;
      var obj = {};
      for (var c = 0; c < headers.length; c++) if (headers[c]) obj[headers[c]] = row[c];
      obj = trimKeys(obj);
      var taskType = col(obj, ['Task Type ', 'Task Type']);
      var name = col(obj, ['Name']);
      var id = col(obj, ['Id', 'ID']);
      var startLoc = col(obj, ['Start Location ', 'Start Location']);
      var endLoc = col(obj, ['End Location ', 'End Location']);
      var dur = parseNum(col(obj, ['Duration Taken Seconds']));
      var startTs = col(obj, ['Start Timestamp']);
      var blocked = col(obj, ['Is Blocked By Foundry']);
      if (blocked === true || blocked === 'true' || blocked === 'TRUE' || blocked === 1) blocked = true; else blocked = false;
      if (!taskType && !name && !startLoc) continue;
      state.rawMoves.push({
        taskType: taskType || '',
        name: (name || '').trim(),
        id: (id || '').toString().trim(),
        startLocation: (startLoc || '').trim(),
        endLocation: (endLoc || '').trim(),
        durationSecs: dur,
        startTimestamp: startTs,
        startDate: parseTimestamp(startTs),
        blocked: !!blocked
      });
    }
  }

  function addPivotTable(sheet, headers) {
    var data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    var h = (data[0] || []).map(function (x) { return (x || '').trim(); });
    var taskCol = h.indexOf('Task Type Description');
    var startCol = h.indexOf('Start Location Title');
    var endCol = h.indexOf('End Location Title');
    var countCol = h.indexOf('Count');
    var avgMins = h.indexOf('Average time taken (mins)');
    var avgSecs = h.indexOf('Average time taken (secs)');
    var fastestSecs = h.indexOf('Fastest time (secs');
    if (fastestSecs === -1) fastestSecs = h.indexOf('Fastest time (secs)');
    var slowestSecs = h.indexOf('Slowest time (secs)');
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (!row || !row.length) continue;
      var label = (row[taskCol] ?? row[startCol] ?? row[endCol] ?? '').toString().trim();
      var count = parseNum(row[countCol]);
      if (!label) continue;
      if (taskCol >= 0 && row[taskCol]) {
        state.taskTypes.push({
          type: label,
          count: count,
          avgMins: parseNum(row[avgMins]),
          avgSecs: parseNum(row[avgSecs]),
          fastestSecs: parseNum(row[fastestSecs]),
          slowestSecs: parseNum(row[slowestSecs])
        });
      } else if (startCol >= 0 && row[startCol]) {
        state.startLocations.push({ location: label, count: count });
      } else if (endCol >= 0 && row[endCol]) {
        state.endLocations.push({ location: label, count: count });
      }
    }
  }

  function addSummary(sheet) {
    var data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    var h = (data[0] || []).map(function (x) { return (x || '').trim(); });
    var taskCol = h.indexOf('Task Type Description');
    var countCol = h.indexOf('Count');
    var avgMins = h.indexOf('Average time taken (mins)');
    var avgSecs = h.indexOf('Average time taken (secs)');
    var fastestSecs = h.indexOf('Fastest time (secs');
    if (fastestSecs === -1) fastestSecs = h.indexOf('Fastest time (secs)');
    var slowestSecs = h.indexOf('Slowest time (secs)');
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (!row || !row.length) continue;
      var label = (row[taskCol] || '').toString().trim();
      if (!label) continue;
      state.taskTypes.push({
        type: label,
        count: parseNum(row[countCol]),
        avgMins: parseNum(row[avgMins]),
        avgSecs: parseNum(row[avgSecs]),
        fastestSecs: parseNum(row[fastestSecs]),
        slowestSecs: parseNum(row[slowestSecs])
      });
    }
  }

  function addNamesIds(sheet) {
    var data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    var h = (data[0] || []).map(function (x) { return (x || '').trim(); });
    var idCol = h.indexOf('ID');
    var nameCol = h.indexOf('Name');
    if (idCol < 0 || nameCol < 0) return;
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var id = (row[idCol] ?? '').toString().trim();
      var name = (row[nameCol] ?? '').toString().trim();
      if (id) state.namesIds[id] = name;
    }
  }

  function addSubMinPivot(sheet) {
    var data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    var h = (data[0] || []).map(function (x) { return (x || '').trim(); });
    var labelCol = h.indexOf('Row Labels');
    var countCol = h.indexOf('Count of End Location ');
    if (labelCol < 0) labelCol = 0;
    if (countCol < 0) countCol = 1;
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var label = (row[labelCol] ?? '').toString().trim();
      var count = parseNum(row[countCol]);
      if (label && /^[\d,]|^[A-Za-z]/.test(label)) state.subMinPivot.push({ name: label, count: count || 0 });
    }
  }

  function addScans(sheet) {
    var data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    var h = (data[0] || []).map(function (x) { return (x || '').trim(); });
    var nameCol = h.indexOf('Completed by Name');
    var countCol = h.indexOf('COUNT');
    var avgCol = h.indexOf('Average Time taken (secs)');
    if (nameCol < 0 || countCol < 0) return;
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var name = (row[nameCol] ?? '').toString().trim();
      if (!name) continue;
      state.scans.push({
        name: name,
        count: parseNum(row[countCol]) || 0,
        avgSecs: parseNum(row[avgCol])
      });
    }
  }

  function processWorkbook(wb, fileName) {
    state.sourceFiles.push(fileName);
    for (var i = 0; i < wb.SheetNames.length; i++) {
      var name = wb.SheetNames[i];
      var sheet = wb.Sheets[name];
      var data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      var firstRow = data[0] || [];
      var headers = firstRow.map(function (x) { return (x || '').trim(); });

      if (name === 'Field Ops Task') addFieldOpsRows(sheet, true);
      if (name === '<1 Minute DATA') {
        var h = headers;
        var obj = {};
        for (var c = 0; c < h.length; c++) if (h[c]) obj[h[c]] = (data[1] || [])[c];
        obj = trimKeys(obj);
        var taskType = col(obj, ['Task Type ', 'Task Type']);
        var nm = col(obj, ['Name']);
        var id = col(obj, ['Id', 'ID']);
        var startLoc = col(obj, ['Start Location ', 'Start Location']);
        var endLoc = col(obj, ['End Location ', 'End Location']);
        var dur = parseNum(col(obj, ['Duration Taken Seconds']));
        var startTs = col(obj, ['Start Timestamp']);
        var blocked = col(obj, ['Is Blocked By Foundry']);
        if (blocked === true || blocked === 'true' || blocked === 'TRUE' || blocked === 1) blocked = true; else blocked = false;
        for (var r = 1; r < data.length; r++) {
          var row = data[r];
          if (!row || !row.length) continue;
          var o = {};
          for (var c = 0; c < headers.length; c++) if (headers[c]) o[headers[c]] = row[c];
          o = trimKeys(o);
          var tt = col(o, ['Task Type ', 'Task Type']);
          var n = (col(o, ['Name']) || '').trim();
          var startL = (col(o, ['Start Location ', 'Start Location']) || '').trim();
          var endL = (col(o, ['End Location ', 'End Location']) || '').trim();
          var d = parseNum(col(o, ['Duration Taken Seconds']));
          var st = col(o, ['Start Timestamp']);
          var bl = col(o, ['Is Blocked By Foundry']);
          bl = bl === true || bl === 'true' || bl === 'TRUE' || bl === 1;
          if (!tt && !n && !startL) continue;
          state.rawMoves.push({
            taskType: tt || '',
            name: n,
            id: (col(o, ['Id', 'ID']) || '').toString().trim(),
            startLocation: startL,
            endLocation: endL,
            durationSecs: d,
            startTimestamp: st,
            startDate: parseTimestamp(st),
            blocked: !!bl
          });
        }
      }
      if (name === 'Pivot Table') addPivotTable(sheet, headers);
      if (name === 'Summary') addSummary(sheet);
      if (name === 'Names-IDs') addNamesIds(sheet);
      if (name === '<1 Minute Pivot') addSubMinPivot(sheet);
      if (name.indexOf('Scans') === 0 && name.indexOf('3.5') !== -1) addScans(sheet);
    }
  }

  function aggregateFromRaw() {
    var moves = state.rawMoves;
    if (!moves.length) return;

    var byWorker = {};
    var byTask = {};
    var byStart = {};
    var byEnd = {};
    var hourly = {};
    var totalDur = 0;
    var countDur = 0;
    var blockedCount = 0;
    var subMinCount = 0;

    for (var i = 0; i < moves.length; i++) {
      var m = moves[i];
      var name = m.name || state.namesIds[m.id] || m.id || 'Unknown';
      if (!byWorker[name]) byWorker[name] = { name: name, count: 0, totalSecs: 0, fastest: null, slowest: null, subMin: 0 };
      byWorker[name].count++;
      if (m.durationSecs != null) {
        byWorker[name].totalSecs += m.durationSecs;
        if (byWorker[name].fastest == null || m.durationSecs < byWorker[name].fastest) byWorker[name].fastest = m.durationSecs;
        if (byWorker[name].slowest == null || m.durationSecs > byWorker[name].slowest) byWorker[name].slowest = m.durationSecs;
        if (m.durationSecs < 60) byWorker[name].subMin++;
      }
      totalDur += m.durationSecs || 0;
      if (m.durationSecs != null) countDur++;
      if (m.blocked) blockedCount++;
      if (m.durationSecs != null && m.durationSecs < 60) subMinCount++;

      var tt = m.taskType || 'Other';
      if (!byTask[tt]) byTask[tt] = { type: tt, count: 0, totalSecs: 0, fastest: null, slowest: null };
      byTask[tt].count++;
      if (m.durationSecs != null) {
        byTask[tt].totalSecs += m.durationSecs;
        if (byTask[tt].fastest == null || m.durationSecs < byTask[tt].fastest) byTask[tt].fastest = m.durationSecs;
        if (byTask[tt].slowest == null || m.durationSecs > byTask[tt].slowest) byTask[tt].slowest = m.durationSecs;
      }

      var sl = m.startLocation || 'Unknown';
      byStart[sl] = (byStart[sl] || 0) + 1;
      var el = m.endLocation || 'Unknown';
      byEnd[el] = (byEnd[el] || 0) + 1;

      if (m.startDate) {
        var h = m.startDate.getHours();
        hourly[h] = (hourly[h] || 0) + 1;
      }
    }

    if (!state.taskTypes.length) {
      for (var t in byTask) state.taskTypes.push({
        type: byTask[t].type,
        count: byTask[t].count,
        avgSecs: byTask[t].totalSecs / byTask[t].count,
        avgMins: (byTask[t].totalSecs / byTask[t].count) / 60,
        fastestSecs: byTask[t].fastest,
        slowestSecs: byTask[t].slowest
      });
    }
    if (!state.startLocations.length) {
      for (var s in byStart) state.startLocations.push({ location: s, count: byStart[s] });
      state.startLocations.sort(function (a, b) { return b.count - a.count; });
    }
    if (!state.endLocations.length) {
      for (var e in byEnd) state.endLocations.push({ location: e, count: byEnd[e] });
      state.endLocations.sort(function (a, b) { return b.count - a.count; });
    }

    state.workers = [];
    for (var w in byWorker) {
      var x = byWorker[w];
      state.workers.push({
        name: x.name,
        count: x.count,
        avgSecs: x.count ? x.totalSecs / x.count : null,
        fastestSecs: x.fastest,
        slowestSecs: x.slowest,
        subMinCount: x.subMin || 0
      });
    }
    state.workers.sort(function (a, b) { return b.count - a.count; });

    state._aggregate = {
      totalMoves: moves.length,
      uniqueWorkers: state.workers.length,
      avgDurationSecs: countDur ? totalDur / countDur : null,
      blockedCount: blockedCount,
      subMinCount: subMinCount,
      hourly: hourly
    };
  }

  function mergeTaskTypes() {
    var byType = {};
    state.taskTypes.forEach(function (t) {
      var key = (t.type || '').trim();
      if (!byType[key]) byType[key] = { type: key, count: 0, totalSecs: 0, fastest: null, slowest: null, n: 0 };
      byType[key].count += t.count || 0;
      if (t.avgSecs != null) { byType[key].totalSecs += t.avgSecs * (t.count || 0); byType[key].n += t.count || 0; }
      if (t.fastestSecs != null && (byType[key].fastest == null || t.fastestSecs < byType[key].fastest)) byType[key].fastest = t.fastestSecs;
      if (t.slowestSecs != null && (byType[key].slowest == null || t.slowestSecs > byType[key].slowest)) byType[key].slowest = t.slowestSecs;
    });
    state.taskTypes = [];
    for (var k in byType) {
      var x = byType[k];
      state.taskTypes.push({
        type: x.type,
        count: x.count,
        avgSecs: x.n ? x.totalSecs / x.n : null,
        avgMins: x.n ? (x.totalSecs / x.n) / 60 : null,
        fastestSecs: x.fastest,
        slowestSecs: x.slowest
      });
    }
    state.taskTypes.sort(function (a, b) { return (b.count || 0) - (a.count || 0); });
  }

  function mergeLocations(arr, key) {
    var byKey = {};
    arr.forEach(function (x) {
      var k = (x[key] || x.location || '').trim();
      if (!k) return;
      byKey[k] = (byKey[k] || 0) + (x.count || 0);
    });
    var out = [];
    for (var k in byKey) out.push(key === 'location' ? { location: k, count: byKey[k] } : { [key]: k, count: byKey[k] });
    out.sort(function (a, b) { return b.count - a.count; });
    return out;
  }

  function fmtSecs(s) {
    if (s == null || isNaN(s)) return '—';
    if (s < 60) return Math.round(s) + 's';
    var m = Math.floor(s / 60), sc = Math.round(s % 60);
    return m + ':' + String(sc).padStart(2, '0');
  }

  function toast(msg, type) {
    type = type || 'ok';
    var container = document.getElementById('toast-container');
    if (!container) return;
    var el = document.createElement('div');
    el.className = 'toast ' + type;
    el.textContent = msg;
    container.appendChild(el);
    setTimeout(function () { el.remove(); }, 4000);
  }

  function render() {
    var agg = state._aggregate || {};
    setText('kpi-moves', agg.totalMoves != null ? agg.totalMoves.toLocaleString() : '—');
    setText('kpi-workers', agg.uniqueWorkers != null ? agg.uniqueWorkers : '—');
    setText('kpi-avg-time', agg.avgDurationSecs != null ? fmtSecs(agg.avgDurationSecs) : '—');
    setText('kpi-submin', agg.subMinCount != null ? agg.subMinCount.toLocaleString() : '—');
    setText('kpi-blocked', agg.blockedCount != null ? agg.blockedCount : '—');
    setText('header-meta', state.rawMoves.length ? state.rawMoves.length + ' moves from ' + state.sourceFiles.length + ' file(s)' : 'Import Excel to begin');
    setText('perf-sub', agg.subMinCount != null ? agg.subMinCount.toLocaleString() : '—');
    setText('perf-blocked', agg.blockedCount != null ? agg.blockedCount : '—');

    var chips = document.getElementById('source-chips');
    if (chips && state.sourceFiles.length) {
      chips.innerHTML = state.sourceFiles.map(function (f) { return '<span class="source-chip">' + f + '</span>'; }).join('');
    }

    var hourly = agg.hourly || {};
    var hours = [];
    for (var h = 0; h < 24; h++) hours.push({ hour: h, count: hourly[h] || 0 });
    renderChart('chart-hourly', 'bar', hours.map(function (x) { return x.hour + ':00'; }), hours.map(function (x) { return x.count; }), 'Moves');

    var taskLabels = state.taskTypes.slice(0, 8).map(function (t) { return t.type; });
    var taskCounts = state.taskTypes.slice(0, 8).map(function (t) { return t.count || 0; });
    renderChart('chart-task-mix', 'doughnut', taskLabels, taskCounts);

    var topWorkers = state.workers.slice(0, 10);
    renderChart('chart-top-workers', 'bar', topWorkers.map(function (w) { return w.name; }), topWorkers.map(function (w) { return w.count; }), 'Moves');

    var topDest = state.endLocations.slice(0, 10);
    renderChart('chart-destinations', 'bar', topDest.map(function (d) { return d.location; }), topDest.map(function (d) { return d.count; }), 'Count');

    renderChart('chart-tasks-bar', 'bar', state.taskTypes.map(function (t) { return t.type; }), state.taskTypes.map(function (t) { return t.count || 0; }), 'Count');

    var tbody = document.getElementById('table-tasks');
    if (tbody) {
      tbody.innerHTML = state.taskTypes.map(function (t) {
        return '<tr><td class="cell-name">' + (t.type || '—') + '</td><td>' + (t.count != null ? t.count.toLocaleString() : '—') + '</td><td>' + (t.avgMins != null ? t.avgMins.toFixed(2) : '—') + '</td><td>' + (t.avgSecs != null ? Math.round(t.avgSecs) : '—') + '</td><td>' + fmtSecs(t.fastestSecs) + '</td><td>' + fmtSecs(t.slowestSecs) + '</td></tr>';
      }).join('');
    }

    var startBody = document.getElementById('table-start');
    if (startBody) {
      startBody.innerHTML = state.startLocations.map(function (x) { return '<tr><td class="cell-name">' + (x.location || '—') + '</td><td>' + (x.count != null ? x.count.toLocaleString() : '—') + '</td></tr>'; }).join('');
    }
    var endBody = document.getElementById('table-end');
    if (endBody) {
      endBody.innerHTML = state.endLocations.map(function (x) { return '<tr><td class="cell-name">' + (x.location || '—') + '</td><td>' + (x.count != null ? x.count.toLocaleString() : '—') + '</td></tr>'; }).join('');
    }

    var workersBody = document.getElementById('table-workers');
    if (workersBody) {
      workersBody.innerHTML = state.workers.map(function (w, i) {
        return '<tr><td>' + (i + 1) + '</td><td class="cell-name">' + (w.name || '—') + '</td><td>' + (w.count != null ? w.count : '—') + '</td><td>' + fmtSecs(w.avgSecs) + '</td><td>' + fmtSecs(w.fastestSecs) + '</td><td>' + fmtSecs(w.slowestSecs) + '</td><td>' + (w.subMinCount != null ? w.subMinCount : '—') + '</td></tr>';
      }).join('');
    }

    var subMinBody = document.getElementById('table-submin');
    if (subMinBody) {
      var subMin = state.subMinPivot.length ? state.subMinPivot : state.workers.filter(function (w) { return (w.subMinCount || 0) > 0; }).map(function (w) { return { name: w.name, count: w.subMinCount }; }).sort(function (a, b) { return (b.count || 0) - (a.count || 0); });
      subMinBody.innerHTML = subMin.slice(0, 50).map(function (x) { return '<tr><td class="cell-name">' + (x.name || '—') + '</td><td>' + (x.count != null ? x.count : '—') + '</td></tr>'; }).join('');
    }
  }

  function renderChart(canvasId, type, labels, data, label) {
    var canvas = document.getElementById(canvasId);
    if (!canvas) return;
    if (charts[canvasId]) { charts[canvasId].destroy(); charts[canvasId] = null; }
    var ctx = canvas.getContext('2d');
    var bgColors = type === 'bar' ? data.map(function (_, i) { return PALETTE[i % PALETTE.length]; }) : PALETTE;
    charts[canvasId] = new Chart(ctx, {
      type: type,
      data: {
        labels: labels,
        datasets: [{ label: label || 'Count', data: data, backgroundColor: bgColors, borderColor: 'transparent', borderWidth: 0 }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: type === 'doughnut', labels: { color: '#a1a1aa' } } },
        scales: type !== 'doughnut' ? {
          x: { ticks: { color: '#a1a1aa', maxRotation: 45 }, grid: { color: 'rgba(255,255,255,.06)' } },
          y: { ticks: { color: '#a1a1aa' }, grid: { color: 'rgba(255,255,255,.06)' } }
        } : {}
      }
    });
  }

  function resetState() {
    state.rawMoves = [];
    state.taskTypes = [];
    state.startLocations = [];
    state.endLocations = [];
    state.workers = [];
    state.subMinPivot = [];
    state.scans = [];
    state.sourceFiles = [];
    state._aggregate = null;
  }

  function onFiles(files) {
    if (!files || !files.length) return;
    resetState();
    var read = 0;
    var total = files.length;
    function done() {
      read++;
      if (read >= total) {
        mergeTaskTypes();
        state.startLocations = mergeLocations(state.startLocations, 'location');
        state.endLocations = mergeLocations(state.endLocations, 'location');
        aggregateFromRaw();
        render();
        toast('Imported ' + total + ' file(s). ' + state.rawMoves.length + ' moves loaded.', 'ok');
      }
    }
    for (var i = 0; i < files.length; i++) {
      (function (file) {
        var r = new FileReader();
        r.onload = function (e) {
          try {
            var wb = XLSX.read(e.target.result, { type: 'array' });
            processWorkbook(wb, file.name);
          } catch (err) {
            toast('Error reading ' + file.name + ': ' + err.message, 'err');
          }
          done();
        };
        r.readAsArrayBuffer(file);
      })(files[i]);
    }
  }

  document.getElementById('file-input').addEventListener('change', function () {
    var files = this.files;
    if (files && files.length) onFiles(Array.prototype.slice.call(files));
    this.value = '';
  });

  document.getElementById('tabs').addEventListener('click', function (e) {
    var t = e.target.closest('.tab');
    if (!t || !t.dataset.tab) return;
    document.querySelectorAll('.tab').forEach(function (el) { el.classList.remove('active'); });
    document.querySelectorAll('.panel').forEach(function (el) { el.classList.remove('active'); });
    t.classList.add('active');
    var panel = document.getElementById('panel-' + t.dataset.tab);
    if (panel) panel.classList.add('active');
  });

  if (typeof window.toast === 'undefined') window.toast = toast;
})();
