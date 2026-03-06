/* ABM Rental Operations Dashboard - Rebuilt with robust event handling */
(function () {
  'use strict';

  var XLSX = window.XLSX;
  var Chart = window.Chart;

  var SEED = {
    summary: { date: "2026-02-07", total_moves: 1650, total_workers: 63, avg_time_secs: 251.3, sub_minute_total: 459, sub_minute_pct: 27.8, blocked_moves: 35 },
    workers: [
      { name: "Cardoza, Israel", count: 113, avg_secs: 242.77, fastest_secs: 8.34, slowest_secs: 2124.01, sub_minute_count: 41 },
      { name: "Barrow, Ryvan", count: 90, avg_secs: 215.98, fastest_secs: 6.47, slowest_secs: 596.06, sub_minute_count: 20 },
      { name: "Hughes, Joe", count: 85, avg_secs: 205.93, fastest_secs: 71.05, slowest_secs: 499.93, sub_minute_count: 0 }
    ],
    task_types: [
      { type: "Move Car to Production", count: 677, avg_secs: 208.34, avg_mins: 3.47, fastest_secs: 4.78, slowest_secs: 2533.32 },
      { type: "Move Car to Readyline", count: 616, avg_secs: 326.63, avg_mins: 5.44, fastest_secs: 9.88, slowest_secs: 2149.29 }
    ],
    start_locations: [{ location: "Return Area", count: 824 }, { location: "A Side Production", count: 771 }],
    end_locations: [{ location: "A Side Production", count: 705 }, { location: "Readyline", count: 600 }],
    hourly_activity: Array.from({ length: 24 }, function (_, i) { return { hour: i, count: i >= 6 && i <= 20 ? 50 + Math.floor(Math.random() * 80) : Math.floor(Math.random() * 20) }; }),
    payroll: { reg: 918.82, lunch: 34.08, overtime: 235.1, pto: 0, employees: 94, total_paid: 1153.92 },
    sub_minute_pivot: [{ name: "Cardoza, Israel", sub_minute_count: 41, avg_secs: 26.95 }, { name: "Shepelev, Lev", sub_minute_count: 35, avg_secs: 17.2 }],
    worker_pivot: [],
    hours_overview: [],
    gap_stats: [{ name: "Mukengeshay, Dedy", avg_gap_secs: 11, samples: 29 }, { name: "Cardoza, Israel", avg_gap_secs: 242.3, samples: 113 }],
    names_ids: {}
  };

  var PAL = ['#0071e3', '#30d158', '#ff9f0a', '#ff453a', '#bf5af2', '#5ac8fa', '#ff2d55', '#ffcc00', '#5856d6', '#30b0c7'];
  var charts = {};
  var DB = JSON.parse(JSON.stringify(SEED));
  DB.raw_moves = DB.raw_moves || [];
  var gapMap = {};
  (DB.gap_stats || []).forEach(function (g) { gapMap[g.name.trim()] = g; });

  var FILTER = { start: null, end: null };
  var wSort = 'count';
  var cmpMode = 'weekvweek';

  function $(id) { return document.getElementById(id); }
  function set(id, v) { var e = $(id); if (e) e.textContent = v; }
  function fmtS(s) {
    if (s == null || isNaN(s)) return '—';
    if (s < 60) return Math.round(s) + 's';
    var m = Math.floor(s / 60), sc = Math.round(s % 60);
    return m + ':' + String(sc).padStart(2, '0');
  }
  function fmtH(h) { return h < 12 ? (h || 12) + ' AM' : h === 12 ? '12 PM' : (h - 12) + ' PM'; }
  function fmtDate(str) {
    if (!str) return '—';
    var d = new Date(str + 'T12:00:00');
    return isNaN(d) ? str : d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }
  function cx(id) {
    if (charts[id]) { charts[id].destroy(); delete charts[id]; }
    var el = $(id);
    return el ? el.getContext('2d') : null;
  }
  var axBase = { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { color: '#86868b', font: { size: 11 } } };
  var boBase = { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } };

  function uOv() {
    if (!Chart) return;
    var s = DB.summary;
    var moves = (DB.workers || []).reduce(function (a, b) { return a + b.count; }, 0) || s.total_moves || 0;
    var wavg = moves > 0 ? (DB.workers || []).reduce(function (a, b) { return a + b.avg_secs * b.count; }, 0) / moves : 0;
    set('o-moves', moves.toLocaleString());
    set('o-date', fmtDate(s.date));
    set('o-wkrs', (DB.workers || []).length || s.total_workers);
    set('o-avg', fmtS(wavg));
    set('o-sub', s.sub_minute_total || '—');
    set('o-subpct', (s.sub_minute_pct || 0) + '% of total');
    set('o-blk', s.blocked_moves || 35);
    if (DB.payroll) set('o-phrs', Math.round(DB.payroll.total_paid || 0).toLocaleString());
    if (s.date) set('hdr-date', fmtDate(s.date));

    var hly = DB.hourly_activity || [];
    var data = hly.map(function (h) { return h.count; });
    var pk = Math.max.apply(null, data.concat([1]));
    if ($('c-hly')) {
      charts['c-hly'] = new Chart(cx('c-hly'), { type: 'bar', data: { labels: hly.map(function (h) { return fmtH(h.hour); }), datasets: [{ data: data, backgroundColor: data.map(function (v) { return v === pk ? '#0071e3' : 'rgba(0,113,227,0.2)'; }), borderRadius: 6, borderWidth: 0 }] }, options: Object.assign({}, boBase, { scales: { x: Object.assign({}, axBase, { grid: { display: false } }), y: axBase } }) });
    }
    var ph = hly.reduce(function (a, b) { return b.count > (a.count || 0) ? b : a; }, {});
    set('o-pk', ph.hour != null ? 'Peak ' + fmtH(ph.hour) + ' · ' + (ph.count || 0) + ' moves' : '—');

    var top = (DB.task_types || []).slice(0, 7);
    if (top.length && $('c-tdn')) {
      charts['c-tdn'] = new Chart(cx('c-tdn'), { type: 'doughnut', data: { labels: top.map(function (t) { return t.type.replace(/Move Car to /i, ''); }), datasets: [{ data: top.map(function (t) { return t.count; }), backgroundColor: PAL.slice(0, 7), borderWidth: 2, borderColor: '#fff' }] }, options: Object.assign({}, boBase, { cutout: '60%', plugins: { legend: { display: true, position: 'right' } } }) });
    }
    set('o-ttl', (DB.task_types || []).length + ' types');

    var topW = (DB.workers || []).slice(0, 10);
    if (topW.length && $('c-topw')) {
      charts['c-topw'] = new Chart(cx('c-topw'), { type: 'bar', data: { labels: topW.map(function (w) { return w.name.split(',')[0].trim(); }), datasets: [{ data: topW.map(function (w) { return w.count; }), backgroundColor: '#0071e3', borderRadius: 6, borderWidth: 0 }] }, options: Object.assign({}, boBase, { indexAxis: 'y', scales: { x: axBase, y: Object.assign({}, axBase, { grid: { display: false } }) } }) });
    }
    var topE = (DB.end_locations || []).slice(0, 10);
    if (topE.length && $('c-dest')) {
      charts['c-dest'] = new Chart(cx('c-dest'), { type: 'bar', data: { labels: topE.map(function (l) { return l.location.replace(/A Side /i, ''); }), datasets: [{ data: topE.map(function (l) { return l.count; }), backgroundColor: PAL, borderRadius: 6, borderWidth: 0 }] }, options: Object.assign({}, boBase, { indexAxis: 'y', scales: { x: axBase, y: Object.assign({}, axBase, { grid: { display: false } }) } }) });
    }
  }

  function wList() {
    var q = ($('worker-search') && $('worker-search').value || '').toLowerCase();
    return (DB.workers || []).filter(function (w) { return w.name.toLowerCase().includes(q); });
  }
  function renderWt(wks) {
    var tb = $('wtb');
    if (!tb) return;
    if (!wks.length) { tb.innerHTML = '<tr><td colspan="9" class="empty">No results</td></tr>'; return; }
    var maxC = Math.max.apply(null, wks.map(function (w) { return w.count; }));
    var nameToId = {};
    if (DB.names_ids) Object.keys(DB.names_ids).forEach(function (id) { nameToId[DB.names_ids[id].trim()] = id; });
    tb.innerHTML = wks.map(function (w, i) {
      var pct = Math.round(w.count / maxC * 100);
      var gap = gapMap[w.name && w.name.trim()];
      var wid = w.id || (nameToId[w.name && w.name.trim()] || '—');
      return '<tr><td>' + (i + 1) + '</td><td class="cell-name">' + w.name + '</td><td class="cell-mono">' + wid + '</td><td><strong>' + w.count + '</strong></td><td>' + fmtS(w.avg_secs) + '</td><td class="positive">' + fmtS(w.fastest_secs) + '</td><td class="negative">' + fmtS(w.slowest_secs) + '</td><td>' + ((w.sub_minute_count || 0) > 0 ? '<span class="badge badge-purple">' + w.sub_minute_count + '</span>' : '') + '</td><td>' + (gap ? fmtS(gap.avg_gap_secs) : '—') + '</td><td><div class="progress"><div class="progress-fill" style="width:' + pct + '%;background:#0071e3"></div></div></td></tr>';
    }).join('');
  }

  function uWk() {
    var wks = DB.workers || [];
    if (!wks.length) return;
    var tot = wks.reduce(function (a, b) { return a + b.count; }, 0);
    var top = wks[0];
    var fast = wks.slice().sort(function (a, b) { return (a.avg_secs || 0) - (b.avg_secs || 0); })[0];
    var fastGap = (DB.gap_stats || [])[0];
    set('w-tot', wks.length);
    set('w-top', top && top.name && top.name.split(',')[0].trim() || '—');
    set('w-top2', (top && top.count || 0) + ' moves');
    set('w-fast', fmtS(fast && fast.avg_secs));
    set('w-fast2', fast && fast.name ? fast.name.split(',')[0] : '');
    set('w-avgm', Math.round(tot / wks.length));
    set('w-gap', fastGap ? fmtS(fastGap.avg_gap_secs) : '—');
    set('w-gap2', fastGap ? (fastGap.name && fastGap.name.split(',')[0]) + ' (fastest)' : 'between moves');
    var list = wList();
    if (wSort === 'avg_secs') list.sort(function (a, b) { return (a.avg_secs || 0) - (b.avg_secs || 0); });
    else if (wSort === 'sub') list.sort(function (a, b) { return (b.sub_minute_count || 0) - (a.sub_minute_count || 0); });
    else if (wSort === 'gap') list.sort(function (a, b) { return (gapMap[a.name && a.name.trim()] && gapMap[a.name.trim()].avg_gap_secs || 9999) - (gapMap[b.name && b.name.trim()] && gapMap[b.name.trim()].avg_gap_secs || 9999); });
    renderWt(list);
  }

  function uTk() {
    if (!Chart) return;
    var tks = DB.task_types || [];
    if (!tks.length) return;
    var total = tks.reduce(function (a, b) { return a + b.count; }, 0);
    var maxC = Math.max.apply(null, tks.map(function (t) { return t.count; }));
    if ($('c-tbar')) {
      charts['c-tbar'] = new Chart(cx('c-tbar'), { type: 'bar', data: { labels: tks.map(function (t) { return t.type.replace(/Move Car to /i, ''); }), datasets: [{ data: tks.map(function (t) { return t.avg_mins || t.avg_secs / 60; }), backgroundColor: PAL, borderRadius: 6, borderWidth: 0 }] }, options: Object.assign({}, boBase, { scales: { x: Object.assign({}, axBase, { grid: { display: false } }), y: axBase } }) });
    }
    set('tk-lbl', tks.length + ' types · ' + total.toLocaleString());
    var ttb = $('ttb');
    if (ttb) ttb.innerHTML = tks.map(function (t, i) { return '<tr><td class="cell-name">' + t.type + '</td><td><strong>' + t.count + '</strong></td><td>' + (t.avg_mins || (t.avg_secs / 60).toFixed(2)) + '</td><td>' + (t.avg_secs || 0).toFixed(1) + 's</td><td class="positive">' + fmtS(t.fastest_secs) + '</td><td class="negative">' + fmtS(t.slowest_secs) + '</td><td>' + (total ? Math.round(t.count / total * 100) : 0) + '%</td><td><div class="progress"><div class="progress-fill" style="width:' + (maxC ? Math.round(t.count / maxC * 100) : 0) + '%;background:' + PAL[i % PAL.length] + '"></div></div></td></tr>'; }).join('');
  }

  function uLc() {
    var lr = function (l, mx, c) { return '<div class="loc-row"><div class="loc-name">' + l.location + '</div><div class="loc-bar"><div class="progress"><div class="progress-fill" style="width:' + (mx ? Math.round(l.count / mx * 100) : 0) + '%;background:' + c + '"></div></div></div><div class="loc-count">' + l.count + '</div></div>'; };
    var sl = DB.start_locations || [];
    if (sl.length) {
      var ms = Math.max.apply(null, sl.map(function (l) { return l.count; }));
      set('sl-tot', sl.reduce(function (a, b) { return a + b.count; }, 0).toLocaleString() + ' departures');
      var slg = $('slg');
      if (slg) slg.innerHTML = sl.map(function (l) { return lr(l, ms, '#0071e3'); }).join('');
    }
    var el = DB.end_locations || [];
    if (el.length) {
      var me = Math.max.apply(null, el.map(function (l) { return l.count; }));
      set('el-tot', el.reduce(function (a, b) { return a + b.count; }, 0).toLocaleString() + ' arrivals');
      var elg = $('elg');
      if (elg) elg.innerHTML = el.map(function (l) { return lr(l, me, '#30d158'); }).join('');
    }
    var wp = DB.worker_pivot || [];
    if (wp.length) {
      var allDests = [].concat.apply([], wp.map(function (w) { return Object.keys(w.destinations || {}); }));
      allDests = allDests.filter(function (v, i, a) { return a.indexOf(v) === i; }).sort();
      set('piv-lbl', wp.length + ' × ' + allDests.length + ' dest');
      var pivHdr = $('piv-hdr'), pivBody = $('piv-body');
      if (pivHdr && pivBody) {
        pivHdr.innerHTML = '<tr><th>Worker</th>' + allDests.map(function (d) { return '<th style="font-size:10px">' + d.substring(0, 10) + '</th>'; }).join('') + '<th>Total</th></tr>';
        var maxPiv = Math.max.apply(null, wp.map(function (w) { return w.total || 0; }));
        pivBody.innerHTML = wp.map(function (w) { return '<tr><td class="cell-name">' + w.name + '</td>' + allDests.map(function (d) { var v = (w.destinations || {})[d] || 0; return '<td>' + (v ? '<span style="background:rgba(0,113,227,' + (0.08 + v / maxPiv * 0.6) + ');padding:2px 6px;border-radius:4px">' + v + '</span>' : '·') + '</td>'; }).join('') + '<td><strong>' + (w.total || 0) + '</strong></td></tr>'; }).join('');
      }
    } else {
      var pivBody = $('piv-body');
      if (pivBody) pivBody.innerHTML = '<tr><td colspan="10" class="empty">Import worker pivot data</td></tr>';
    }
  }

  function uPy() {
    if (!Chart) return;
    var p = DB.payroll;
    if (!p) return;
    var tot = (p.reg || 0) + (p.overtime || 0);
    set('p-reg', (p.reg || 0).toLocaleString());
    set('p-ot', (p.overtime || 0).toLocaleString());
    set('p-otpct', tot ? Math.round((p.overtime || 0) / tot * 100) + '% of paid' : '');
    set('p-lu', (p.lunch || 0).toFixed(1));
    set('p-emp', p.employees || 0);
    set('p-tot', Math.round((p.total_paid || tot) || 0).toLocaleString());
    var rows = [{ l: 'Regular', v: p.reg || 0, c: '#0071e3' }, { l: 'Overtime', v: p.overtime || 0, c: '#ff453a' }, { l: 'Lunch', v: p.lunch || 0, c: '#ff9f0a' }, { l: 'PTO', v: p.pto || 0, c: '#30d158' }];
    var psum = $('p-sum');
    if (psum) psum.innerHTML = rows.map(function (r) { return '<div class="pay-row"><div class="pay-dot" style="background:' + r.c + '"></div><div class="pay-label">' + r.l + '</div><div class="pay-bar"><div class="progress"><div class="progress-fill" style="width:' + (tot ? Math.round(r.v / tot * 100) : 0) + '%;background:' + r.c + '"></div></div></div><div class="pay-value">' + r.v.toFixed(1) + '</div></div>'; }).join('');
    if ($('c-pdn')) {
      charts['c-pdn'] = new Chart(cx('c-pdn'), { type: 'doughnut', data: { labels: rows.map(function (r) { return r.l; }), datasets: [{ data: rows.map(function (r) { return r.v; }), backgroundColor: rows.map(function (r) { return r.c; }), borderWidth: 2, borderColor: '#fff' }] }, options: Object.assign({}, boBase, { cutout: '62%', plugins: { legend: { display: true, position: 'bottom' } } }) });
    }
  }

  function uPf() {
    if (!Chart) return;
    var s = DB.summary;
    set('pf-sub', s.sub_minute_total || 459);
    set('pf-wkrs', (DB.sub_minute_pivot || []).length);
    var fast = (DB.sub_minute_pivot || []).sort(function (a, b) { return (a.avg_secs || 0) - (b.avg_secs || 0); })[0];
    if (fast) { set('pf-fast', Math.round(fast.avg_secs) + 's'); set('pf-fast2', fast.name && fast.name.split(',')[0] || ''); }
    set('pf-blk', s.blocked_moves || 35);
    set('pf-pct', (s.sub_minute_pct || 27.8) + '%');
    set('sub-lbl', (DB.sub_minute_pivot || []).length + ' workers');
    var sub = (DB.sub_minute_pivot || []).slice(0, 20);
    if (sub.length && $('c-sub')) {
      charts['c-sub'] = new Chart(cx('c-sub'), { type: 'bar', data: { labels: sub.map(function (w) { return w.name && w.name.split(',')[0].trim(); }), datasets: [{ data: sub.map(function (w) { return w.sub_minute_count; }), backgroundColor: '#bf5af2', borderRadius: 6, borderWidth: 0 }] }, options: Object.assign({}, boBase, { indexAxis: 'y', scales: { x: axBase, y: Object.assign({}, axBase, { grid: { display: false } }) } }) });
    }
    var busy = (DB.workers || []).slice(0, 20);
    var gapTb = $('gap-tb');
    if (gapTb) {
      var maxG = Math.max.apply(null, busy.map(function (x) { var g = gapMap[x.name && x.name.trim()]; return g ? g.avg_gap_secs : 0; }).concat([1]));
      gapTb.innerHTML = busy.map(function (w) {
        var g = gapMap[w.name && w.name.trim()];
        if (!g) return '<tr><td class="cell-name">' + w.name + '</td><td>—</td><td>' + w.count + '</td><td></td></tr>';
        var pct = Math.round(g.avg_gap_secs / maxG * 100);
        return '<tr><td class="cell-name">' + w.name + '</td><td style="font-weight:600">' + fmtS(g.avg_gap_secs) + '</td><td>' + g.samples + '</td><td><div class="progress"><div class="progress-fill" style="width:' + pct + '%;background:#30d158"></div></div></td></tr>';
      }).join('');
    }
    var hmapDiv = $('hmap-div');
    if (hmapDiv) {
      var sph = DB.scans_per_hour || [];
      if (!sph.length) { hmapDiv.innerHTML = '<div class="empty">Import Excel with Scans per hour sheet</div>'; set('hmap-lbl', '—'); }
      else {
        var dateKeys = [];
        sph.forEach(function (r) { Object.keys(r).forEach(function (k) { if (k !== 'name' && dateKeys.indexOf(k) < 0) dateKeys.push(k); }); });
        dateKeys.sort();
        var maxVal = 0; sph.forEach(function (r) { dateKeys.forEach(function (d) { var v = parseFloat(r[d]); if (!isNaN(v) && v > maxVal) maxVal = v; }); });
        maxVal = maxVal || 1;
        var ths = '<tr><th class="hl-name">Worker</th>' + dateKeys.slice(0, 31).map(function (d) { return '<th class="hl-dh">' + d + '</th>'; }).join('') + '<th class="hl-avg">Avg</th></tr>';
        var rows = sph.map(function (r) {
          var cells = dateKeys.slice(0, 31).map(function (d) { var v = parseFloat(r[d]); var n = isNaN(v) ? 0 : v; var pct = Math.round(n / maxVal * 100); return '<td class="dc" style="background:rgba(0,113,227,' + (0.1 + pct / 100 * 0.7) + ')">' + (n > 0 ? n : '') + '</td>'; }).join('');
          var avg = 0, cnt = 0; dateKeys.forEach(function (d) { var v = parseFloat(r[d]); if (!isNaN(v)) { avg += v; cnt++; } }); avg = cnt ? Math.round(avg / cnt * 10) / 10 : 0;
          return '<tr><td class="hl-name">' + (r.name || '').substring(0, 20) + '</td>' + cells + '<td class="hl-avg">' + avg + '</td></tr>';
        }).join('');
        hmapDiv.innerHTML = '<table class="hmap"><thead>' + ths + '</thead><tbody>' + rows + '</tbody></table>';
        set('hmap-lbl', sph.length + ' workers × ' + dateKeys.length + ' days');
      }
    }
  }

  function filteredHO() {
    var ho = DB.hours_overview || [];
    if (!FILTER.start && !FILTER.end) return ho;
    return ho.filter(function (d) {
      if (FILTER.start && d.date < FILTER.start) return false;
      if (FILTER.end && d.date > FILTER.end) return false;
      return true;
    });
  }

  function uTr() {
    if (!Chart) return;
    var ho = filteredHO();
    if (!ho.length) return;
    var sorted = ho.slice().sort(function (a, b) { return a.date.localeCompare(b.date); });
    var withData = sorted.filter(function (d) { return d.actual_hours > 0; });
    var avg = withData.length ? Math.round(withData.reduce(function (a, b) { return a + b.actual_hours; }, 0) / withData.length * 10) / 10 : 0;
    var wN = withData.filter(function (d) { return d.needed_hours; });
    var avgN = wN.length ? Math.round(wN.reduce(function (a, b) { return a + (b.needed_hours || 0); }, 0) / wN.length * 10) / 10 : 0;
    var wV = withData.filter(function (d) { return d.variance_hours != null; });
    var avgV = wV.length ? Math.round(wV.reduce(function (a, b) { return a + (b.variance_hours || 0); }, 0) / wV.length * 10) / 10 : 0;
    var wF = withData.filter(function (d) { return d.fte_variance != null; });
    var avgFTE = wF.length ? Math.round(wF.reduce(function (a, b) { return a + (b.fte_variance || 0); }, 0) / wF.length * 10) / 10 : 0;
    var monthArr = sorted.map(function (d) { return d.month; });
    var months = monthArr.filter(function (v, i, a) { return a.indexOf(v) === i; });
    set('tr-days', sorted.length);
    set('tr-months', months.length + ' months');
    set('tr-avg', avg);
    set('tr-need', avgN);
    set('tr-var', avgV > 0 ? '+' + avgV : avgV);
    set('tr-fte', avgFTE > 0 ? '+' + avgFTE : avgFTE);
    set('tr-lbl', (sorted[0] && sorted[0].date ? sorted[0].date.slice(0, 7) : '') + ' – ' + (sorted[sorted.length - 1] && sorted[sorted.length - 1].date ? sorted[sorted.length - 1].date.slice(0, 7) : ''));
    set('tr-fte-lbl', '+ over · − under');
    var lbl = sorted.map(function (d) { var dt = new Date(d.date + 'T12:00:00'); return dt.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }); });
    if ($('c-trend')) charts['c-trend'] = new Chart(cx('c-trend'), { type: 'line', data: { labels: lbl, datasets: [{ label: 'Actual', data: sorted.map(function (d) { return d.actual_hours; }), borderColor: '#0071e3', backgroundColor: 'rgba(0,113,227,0.08)', borderWidth: 2, fill: true, tension: 0.3, pointRadius: 1.5 }] }, options: Object.assign({}, boBase, { scales: { x: Object.assign({}, axBase, { grid: { display: false } }), y: axBase } }) });
    var wN2 = sorted.filter(function (d) { return d.needed_hours != null; });
    if (wN2.length && $('c-need')) charts['c-need'] = new Chart(cx('c-need'), { type: 'line', data: { labels: wN2.map(function (d) { var dt = new Date(d.date + 'T12:00:00'); return dt.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }); }), datasets: [{ label: 'Actual', data: wN2.map(function (d) { return d.actual_hours; }), borderColor: '#0071e3', borderWidth: 2, tension: 0.3 }, { label: 'Needed', data: wN2.map(function (d) { return d.needed_hours; }), borderColor: '#30d158', borderWidth: 2, borderDash: [5, 3], tension: 0.3 }] }, options: Object.assign({}, boBase, { plugins: { legend: { display: true, position: 'top' } }, scales: { x: Object.assign({}, axBase, { grid: { display: false } }), y: axBase } }) });
    var wFTE = sorted.filter(function (d) { return d.fte_variance != null; });
    if (wFTE.length && $('c-fte')) charts['c-fte'] = new Chart(cx('c-fte'), { type: 'bar', data: { labels: wFTE.map(function (d) { var dt = new Date(d.date + 'T12:00:00'); return dt.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }); }), datasets: [{ data: wFTE.map(function (d) { return d.fte_variance; }), backgroundColor: wFTE.map(function (d) { return (d.fte_variance || 0) > 0 ? 'rgba(48,209,88,0.6)' : 'rgba(255,69,58,0.6)'; }), borderRadius: 4, borderWidth: 0 }] }, options: Object.assign({}, boBase, { scales: { x: Object.assign({}, axBase, { grid: { display: false } }), y: axBase } }) });
    var byMonth = {};
    sorted.forEach(function (d) { if (!byMonth[d.month]) byMonth[d.month] = []; byMonth[d.month].push(d); });
    var trTb = $('tr-tb');
    if (trTb) trTb.innerHTML = Object.keys(byMonth).map(function (m) {
      var days = byMonth[m];
      var aAvg = Math.round(days.reduce(function (a, b) { return a + b.actual_hours; }, 0) / days.length * 10) / 10;
      var nd = days.filter(function (d) { return d.needed_hours; });
      var nAvg = nd.length ? Math.round(nd.reduce(function (a, b) { return a + (b.needed_hours || 0); }, 0) / nd.length * 10) / 10 : null;
      var vd = days.filter(function (d) { return d.variance_hours != null; });
      var vAvg = vd.length ? Math.round(vd.reduce(function (a, b) { return a + (b.variance_hours || 0); }, 0) / vd.length * 10) / 10 : null;
      var fd = days.filter(function (d) { return d.fte_variance != null; });
      var fAvg = fd.length ? Math.round(fd.reduce(function (a, b) { return a + (b.fte_variance || 0); }, 0) / fd.length * 10) / 10 : null;
      var pk = days.reduce(function (a, b) { return b.actual_hours > a.actual_hours ? b : a; });
      return '<tr><td class="cell-name">' + m + '</td><td>' + days.length + '</td><td>' + aAvg + '</td><td>' + (nAvg != null ? nAvg : '—') + '</td><td style="font-weight:600" class="' + (vAvg > 0 ? 'positive' : 'negative') + '">' + (vAvg != null ? (vAvg > 0 ? '+' : '') + vAvg : '—') + '</td><td style="font-weight:600" class="' + (fAvg > 0 ? 'positive' : 'negative') + '">' + (fAvg != null ? (fAvg > 0 ? '+' : '') + fAvg : '—') + '</td><td>' + Math.round(pk.actual_hours) + ' ' + fmtDate(pk.date) + '</td></tr>';
    }).join('');
  }

  function applyDateFilter() {
    var s = $('date-start') && $('date-start').value;
    var e = $('date-end') && $('date-end').value;
    FILTER.start = s || null;
    FILTER.end = e || null;
    document.querySelectorAll('.date-pill').forEach(function (p) { p.classList.remove('active'); });
    var hasFilt = FILTER.start || FILTER.end;
    var dc = $('date-clear');
    if (dc) dc.style.display = hasFilt ? '' : 'none';
    var di = $('date-info');
    if (di) di.textContent = hasFilt ? filteredHO().length + ' days' : '';
    var at = document.querySelector('.tab.active');
    if (at && at.dataset && at.dataset.tab) refreshTab(at.dataset.tab);
  }

  function setPreset(p) {
    document.querySelectorAll('.date-pill').forEach(function (x) { x.classList.remove('active'); });
    var active = document.querySelector('.date-pill[data-preset="' + p + '"]');
    if (active) active.classList.add('active');
    var now = new Date();
    var fmt = function (d) { return d.toISOString().split('T')[0]; };
    var monday = function (d) { var day = d.getDay(); var diff = d.getDate() - day + (day === 0 ? -6 : 1); return new Date(d.getFullYear(), d.getMonth(), diff); };
    var s, e;
    if (p === 'today') s = e = fmt(now);
    else if (p === 'week') { var m = monday(new Date()); s = fmt(m); e = fmt(new Date(m.getTime() + 6 * 86400000)); }
    else if (p === 'lastweek') { var m = monday(new Date()); var lm = new Date(m.getTime() - 7 * 86400000); s = fmt(lm); e = fmt(new Date(lm.getTime() + 6 * 86400000)); }
    else if (p === 'month') { s = fmt(new Date(now.getFullYear(), now.getMonth(), 1)); e = fmt(new Date(now.getFullYear(), now.getMonth() + 1, 0)); }
    else if (p === 'lastmonth') { s = fmt(new Date(now.getFullYear(), now.getMonth() - 1, 1)); e = fmt(new Date(now.getFullYear(), now.getMonth(), 0)); }
    else if (p === 'all') s = e = '';
    var ds = $('date-start'), de = $('date-end');
    if (ds) ds.value = s || '';
    if (de) de.value = e || '';
    applyDateFilter();
  }

  function clearFilter() {
    FILTER = { start: null, end: null };
    var ds = $('date-start'), de = $('date-end');
    if (ds) ds.value = '';
    if (de) de.value = '';
    document.querySelectorAll('.date-pill').forEach(function (p) { p.classList.remove('active'); });
    var dc = $('date-clear');
    if (dc) dc.style.display = 'none';
    var di = $('date-info');
    if (di) di.textContent = '';
    var at = document.querySelector('.tab.active');
    if (at && at.dataset && at.dataset.tab) refreshTab(at.dataset.tab);
  }

  var UP = { overview: uOv, workers: uWk, tasks: uTk, locations: uLc, payroll: uPy, performance: uPf, trends: uTr, compare: uCmp };
  function refreshTab(id) {
    var fn = UP[id];
    if (fn) { try { fn(); } catch (e) { console.error('Tab error:', e); } }
  }

  function switchTab(id, tabEl) {
    if (!id) return;
    document.querySelectorAll('.tab').forEach(function (x) { x.classList.remove('active'); });
    document.querySelectorAll('.panel').forEach(function (x) { x.classList.remove('active'); });
    if (tabEl) tabEl.classList.add('active');
    var panel = $('panel-' + id);
    if (panel) panel.classList.add('active');
    setTimeout(function () { refreshTab(id); }, 50);
  }

  function setCmpMode(mode) {
    cmpMode = mode;
    document.querySelectorAll('#cmp-mode .segmented-item').forEach(function (b) { b.classList.toggle('active', b.dataset.mode === mode); });
    var ww = $('cmp-weekvweek'), dow = $('cmp-dow');
    if (ww) ww.style.display = mode === 'weekvweek' ? '' : 'none';
    if (dow) dow.style.display = mode === 'dow' ? '' : 'none';
    if (mode === 'weekvweek') renderWeekVWeek();
    else renderDow();
  }

  function weekBounds(offsetWeeks) {
    var now = new Date();
    var day = now.getDay();
    var mon = new Date(now);
    mon.setDate(now.getDate() - ((day + 6) % 7) + offsetWeeks * 7);
    var sun = new Date(mon);
    sun.setDate(mon.getDate() + 6);
    return { start: mon.toISOString().split('T')[0], end: sun.toISOString().split('T')[0] };
  }
  function hoForRange(start, end) { return (DB.hours_overview || []).filter(function (d) { return d.date >= start && d.date <= end; }); }
  function weekStats(days) {
    if (!days.length) return null;
    var ah = days.reduce(function (a, b) { return a + b.actual_hours; }, 0);
    var nh = days.filter(function (d) { return d.needed_hours; }).reduce(function (a, b) { return a + (b.needed_hours || 0); }, 0);
    var vh = days.filter(function (d) { return d.variance_hours != null; }).reduce(function (a, b) { return a + (b.variance_hours || 0); }, 0);
    var fh = days.filter(function (d) { return d.fte_variance != null; }).reduce(function (a, b) { return a + (b.fte_variance || 0); }, 0);
    var rents = days.reduce(function (a, b) { return a + b.rent; }, 0);
    var returns = days.reduce(function (a, b) { return a + b.returns; }, 0);
    return { days: days.length, total_hours: Math.round(ah * 10) / 10, avg_hours: Math.round(ah / days.length * 10) / 10, needed: Math.round(nh * 10) / 10, variance: Math.round(vh * 10) / 10, fte: Math.round(fh * 10) / 10, rents: rents, returns: returns };
  }

  function renderWeekVWeek() {
    var tw = weekBounds(0), lw = weekBounds(-1);
    var twDays = hoForRange(tw.start, tw.end).sort(function (a, b) { return a.date.localeCompare(b.date); });
    var lwDays = hoForRange(lw.start, lw.end).sort(function (a, b) { return a.date.localeCompare(b.date); });
    var twS = weekStats(twDays), lwS = weekStats(lwDays);
    var fmt2 = function (d) { return d ? new Date(d + 'T12:00:00').toLocaleDateString('en-US', { month: 'short', day: 'numeric' }) : '—'; };
    set('ww-a-range', twDays.length ? fmt2(twDays[0].date) + ' – ' + fmt2(twDays[twDays.length - 1].date) : 'No data');
    set('ww-b-range', lwDays.length ? fmt2(lwDays[0].date) + ' – ' + fmt2(lwDays[lwDays.length - 1].date) : 'No data');
    var statRows = function (s, color) {
      if (!s) return '<div class="empty">No data</div>';
      var rows = [{ k: 'Total Actual Hrs', v: s.total_hours }, { k: 'Avg Daily', v: s.avg_hours }, { k: 'Needed', v: s.needed }, { k: 'Variance', v: s.variance > 0 ? '+' + s.variance : s.variance }, { k: 'FTE Var', v: s.fte > 0 ? '+' + s.fte : s.fte }, { k: 'Rents', v: s.rents }, { k: 'Returns', v: s.returns }];
      return rows.map(function (r) { return '<div class="cmp-stat"><span class="cmp-k">' + r.k + '</span><span class="cmp-v" style="color:' + color + '">' + r.v + '</span></div>'; }).join('');
    };
    var wwA = $('ww-a-stats'), wwB = $('ww-b-stats');
    if (wwA) wwA.innerHTML = statRows(twS, '#0071e3');
    if (wwB) wwB.innerHTML = statRows(lwS, '#ff9f0a');
    var wwDiff = $('ww-diff-rows');
    if (twS && lwS && wwDiff) {
      var metrics = [{ k: 'Total Actual', a: twS.total_hours, b: lwS.total_hours }, { k: 'Avg Daily', a: twS.avg_hours, b: lwS.avg_hours }, { k: 'Variance', a: twS.variance, b: lwS.variance }, { k: 'FTE', a: twS.fte, b: lwS.fte }];
      wwDiff.innerHTML = metrics.map(function (m) {
        var diff = Math.round((m.a - m.b) * 10) / 10;
        var cls = diff > 0 ? 'diff-pos' : diff < 0 ? 'diff-neg' : '';
        return '<div class="cmp-stat"><span class="cmp-k">' + m.k + '</span><span>This: ' + m.a + ' vs Last: ' + m.b + (diff !== 0 ? ' <span class="diff-badge ' + cls + '">' + (diff > 0 ? '↑' : '↓') + ' ' + Math.abs(diff) + '</span>' : '') + '</span></div>';
      }).join('');
    } else if (wwDiff) wwDiff.innerHTML = '<div class="empty">Import Hours Overview to compare</div>';
    var days7 = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
    var twByDay = days7.map(function (_, i) { var d = twDays[i]; return d ? d.actual_hours : null; });
    var lwByDay = days7.map(function (_, i) { var d = lwDays[i]; return d ? d.actual_hours : null; });
    if ($('c-ww') && Chart) {
      charts['c-ww'] = new Chart(cx('c-ww'), { type: 'bar', data: { labels: days7, datasets: [{ label: 'This Week', data: twByDay, backgroundColor: 'rgba(0,113,227,0.7)', borderRadius: 6 }, { label: 'Last Week', data: lwByDay, backgroundColor: 'rgba(255,159,10,0.6)', borderRadius: 6 }] }, options: Object.assign({}, boBase, { plugins: { legend: { display: true } }, scales: { x: Object.assign({}, axBase, { grid: { display: false } }), y: axBase } }) });
    }
  }

  function renderDow() {
    var ho = (DB.hours_overview || []).filter(function (d) { return d.actual_hours > 0; });
    var DAYS = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    var SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    var byDay = {};
    DAYS.forEach(function (d) { byDay[d] = []; });
    ho.forEach(function (d) { if (byDay[d.day]) byDay[d.day].push(d); });
    var stats = DAYS.map(function (d, i) {
      var rows = byDay[d];
      if (!rows.length) return { day: SHORT[i], count: 0, avg_hrs: 0, avg_needed: 0, avg_var: 0, avg_fte: 0, avg_rent: 0, avg_ret: 0 };
      var avg = function (arr) { return arr.length ? Math.round(arr.reduce(function (a, b) { return a + b; }, 0) / arr.length * 10) / 10 : 0; };
      return { day: SHORT[i], count: rows.length, avg_hrs: avg(rows.map(function (r) { return r.actual_hours; })), avg_needed: avg(rows.filter(function (r) { return r.needed_hours; }).map(function (r) { return r.needed_hours; })), avg_var: avg(rows.filter(function (r) { return r.variance_hours != null; }).map(function (r) { return r.variance_hours; })), avg_fte: avg(rows.filter(function (r) { return r.fte_variance != null; }).map(function (r) { return r.fte_variance; })), avg_rent: avg(rows.map(function (r) { return r.rent; })), avg_ret: avg(rows.map(function (r) { return r.returns; })) };
    });
    set('dow-range-lbl', ho.length + ' days');
    var dowCards = $('dow-cards');
    if (dowCards) {
      var maxHrs = Math.max.apply(null, stats.map(function (s) { return s.avg_hrs; })) || 1;
      dowCards.innerHTML = stats.map(function (s) { return '<div class="dow-card"><div class="dow-name">' + s.day + '</div><div class="dow-val">' + s.avg_hrs + '</div><div style="margin-top:8px;height:4px;background:rgba(0,0,0,.08);border-radius:2px;overflow:hidden"><div style="width:' + Math.round(s.avg_hrs / maxHrs * 100) + '%;height:100%;background:#0071e3;border-radius:2px"></div></div></div>'; }).join('');
    }
    if ($('c-dow') && Chart) {
      charts['c-dow'] = new Chart(cx('c-dow'), {
        type: 'radar',
        data: {
          labels: stats.map(function (s) { return s.day; }),
          datasets: [
            { label: 'Avg Actual', data: stats.map(function (s) { return s.avg_hrs; }), borderColor: '#0071e3', backgroundColor: 'rgba(0,113,227,0.12)', borderWidth: 2 },
            { label: 'Avg Needed', data: stats.map(function (s) { return s.avg_needed; }), borderColor: '#30d158', backgroundColor: 'rgba(48,209,88,0.08)', borderWidth: 2, borderDash: [5, 3] }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: true } },
          scales: { r: { grid: { color: 'rgba(0,0,0,0.06)' }, ticks: { color: '#86868b' } } }
        }
      });
    }
    var dowTable = $('dow-table');
    if (dowTable) dowTable.innerHTML = stats.map(function (s) { return '<tr><td class="cell-name">' + s.day + '</td><td>' + s.count + '</td><td style="font-weight:600;color:#0071e3">' + s.avg_hrs + '</td><td>' + (s.avg_needed || '—') + '</td><td class="' + (s.avg_var > 0 ? 'positive' : 'negative') + '" style="font-weight:600">' + (s.avg_var > 0 ? '+' : '') + s.avg_var + '</td><td class="' + (s.avg_fte > 0 ? 'positive' : 'negative') + '" style="font-weight:600">' + (s.avg_fte > 0 ? '+' : '') + s.avg_fte + '</td><td>' + Math.round(s.avg_rent) + '</td><td>' + Math.round(s.avg_ret) + '</td></tr>'; }).join('');
  }

  function uCmp() {
    if (cmpMode === 'weekvweek') renderWeekVWeek();
    else renderDow();
  }

  function findHdrRow(rows) {
    for (var i = 0; i < Math.min(rows.length, 20); i++) {
      var row = rows[i] || [];
      var str = row.map(function (x) { return String(x || '').toLowerCase(); }).join(' ');
      if (str.indexOf('actual driver hours') >= 0 && (str.indexOf('reporting period') >= 0 || str.indexOf('rent') >= 0)) return i;
    }
    for (var i = 0; i < Math.min(rows.length, 15); i++) { if ((rows[i] || []).filter(function (x) { return x != null && x !== ''; }).length >= 2) return i; }
    return 0;
  }
  function parseHoursSheet(sn, allRows) {
    var hc = -1, dc = 1, ri = -1, rti = -1, ds = -1;
    for (var i = 0; i < allRows.length; i++) {
      var row = allRows[i];
      for (var j = 0; j < row.length; j++) {
        var v = String(row[j] || '').toLowerCase();
        if (v.indexOf('actual driver hours') >= 0) hc = j;
        if (v.indexOf('reporting period') >= 0 || (v === 'date' && dc < 0)) dc = j;
        if ((v.indexOf('rent') >= 0 || v === 'actual rent') && v.indexOf('return') < 0 && v.indexOf('trans') < 0) ri = j;
        if ((v.indexOf('return') >= 0 || v === 'actual return') && v.indexOf('trans') < 0) rti = j;
      }
      if (hc >= 0) { ds = i + 1; break; }
    }
    if (ds < 0) return [];
    if (ri < 0) ri = 6;
    if (rti < 0) rti = 7;
    var dayAlias = { sun: 'sunday', mon: 'monday', tue: 'tuesday', wed: 'wednesday', thu: 'thursday', fri: 'friday', sat: 'saturday' };
    var days = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    var results = [];
    for (var i = ds; i < allRows.length; i++) {
      var r = allRows[i];
      var dn = String(r[0] || '').toLowerCase().trim().replace(/\s/g, '');
      dn = dayAlias[dn] || dn;
      if (days.indexOf(dn) < 0) continue;
      var dv = r[dc];
      var dateStr = null;
      if (dv instanceof Date) dateStr = dv.toISOString().split('T')[0];
      else if (typeof dv === 'number' && dv > 40000) dateStr = new Date((dv - 25569) * 86400 * 1000).toISOString().split('T')[0];
      else if (typeof dv === 'string') { var m = dv.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); if (m) { var y = parseInt(m[3], 10); if (y < 100) y += 2000; dateStr = y + '-' + (parseInt(m[1], 10) < 10 ? '0' : '') + parseInt(m[1], 10) + '-' + (parseInt(m[2], 10) < 10 ? '0' : '') + parseInt(m[2], 10); } else dateStr = dv.substring(0, 10); }
      var ah = parseFloat(String(r[hc] || '').replace(/,/g, ''));
      if (!dateStr || isNaN(ah)) continue;
      var rent = parseInt(String(r[ri] || '').replace(/,/g, ''), 10) || 0;
      var ret = parseInt(String(r[rti] || '').replace(/,/g, ''), 10) || 0;
      var trans = (rent + ret) / 2;
      var needed = trans / 2.2;
      results.push({ date: dateStr, day: dn, month: sn, actual_hours: Math.round(ah * 100) / 100, rent: rent, returns: ret, trans: Math.round(trans * 100) / 100, needed_hours: Math.round(needed * 100) / 100, variance_hours: Math.round((ah - needed) * 100) / 100, fte_variance: Math.round((ah - needed) / 7.5 * 100) / 100 });
    }
    return results;
  }

  function parseFieldOpsTask(allRows, hi) {
    var hdr = (allRows[hi] || []).map(function (x) { return String(x || '').toLowerCase().trim(); });
    var idx = { status: -1, name: -1, id: -1, taskType: -1, startLoc: -1, endLoc: -1, date: -1, duration: -1, centralTime: -1, blocked: -1, gapSec: -1, mva: -1, startTs: -1 };
    hdr.forEach(function (h, i) { if (h.indexOf('status') >= 0) idx.status = i; if (h.indexOf('name') >= 0 && h.indexOf('row') < 0) idx.name = i; if (h === 'id' || h === 'id ') idx.id = i; if (h.indexOf('task type') >= 0) idx.taskType = i; if (h.indexOf('start location') >= 0) idx.startLoc = i; if (h.indexOf('end location') >= 0) idx.endLoc = i; if ((h === 'date' || h.indexOf('reporting period') >= 0) && idx.date < 0) idx.date = i; if (h.indexOf('duration taken') >= 0 || (h.indexOf('duration') >= 0 && h.indexOf('avg') < 0)) idx.duration = i; if (h.indexOf('central time start') >= 0) idx.centralTime = i; if (h.indexOf('blocked') >= 0) idx.blocked = i; if (h.indexOf('gapsec') >= 0) idx.gapSec = i; if (h.indexOf('mva') >= 0 || h === 'mva ') idx.mva = i; if (h.indexOf('start timestamp') >= 0) idx.startTs = i; });
    if (idx.name < 0 || idx.duration < 0) return;
    var seenMva = {};
    for (var i = hi + 1; i < allRows.length; i++) {
      var r = allRows[i];
      var name = String(r[idx.name] || '').trim();
      if (!name || name.toLowerCase() === 'total' || name.toLowerCase() === 'grand total') continue;
      var status = String(r[idx.status] || '').toLowerCase();
      if (status !== 'completed') continue;
      var dur = parseFloat(r[idx.duration]);
      if (isNaN(dur)) continue;
      var mva = idx.mva >= 0 ? String(r[idx.mva] || '') : 'r' + i;
      if (mva && seenMva[mva]) continue;
      if (mva) seenMva[mva] = true;
      var hr = 0;
      if (idx.centralTime >= 0 && r[idx.centralTime]) { var ct = String(r[idx.centralTime]); var m = ct.match(/(\d+)\s*:\s*(\d+)\s*(AM|PM)/i); if (m) { hr = parseInt(m[1], 10); if (m[3].toUpperCase() === 'PM' && hr !== 12) hr += 12; else if (m[3].toUpperCase() === 'AM' && hr === 12) hr = 0; } }
      else if (idx.startTs >= 0 && r[idx.startTs]) { var ts = r[idx.startTs]; if (ts instanceof Date) hr = ts.getUTCHours(); else { var dt = new Date(ts); if (!isNaN(dt)) hr = dt.getUTCHours(); } }
      else if (idx.date >= 0 && r[idx.date]) { var d = r[idx.date]; if (d instanceof Date) hr = d.getHours(); }
      var taskType = String(r[idx.taskType] || '').trim() || 'Unknown';
      var startLoc = idx.startLoc >= 0 ? String(r[idx.startLoc] || '').trim() || 'Unknown' : 'Unknown';
      var endLoc = idx.endLoc >= 0 ? String(r[idx.endLoc] || '').trim() || 'Unknown' : 'Unknown';
      var blocked = idx.blocked >= 0 && (r[idx.blocked] === true || String(r[idx.blocked] || '').toUpperCase() === 'TRUE');
      var gapSec = idx.gapSec >= 0 && r[idx.gapSec] != null && r[idx.gapSec] !== '' ? parseFloat(r[idx.gapSec]) : null;
      var dateStr = '';
      if (idx.date >= 0 && r[idx.date]) { var dv = r[idx.date]; if (dv instanceof Date) dateStr = dv.toISOString().split('T')[0]; else { var ds = String(dv); var mm = ds.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); if (mm) { var y = parseInt(mm[3], 10); if (y < 100) y += 2000; dateStr = y + '-' + (parseInt(mm[1], 10) < 10 ? '0' : '') + parseInt(mm[1], 10) + '-' + (parseInt(mm[2], 10) < 10 ? '0' : '') + parseInt(mm[2], 10); } } }
      else if (idx.startTs >= 0 && r[idx.startTs]) { var ts = r[idx.startTs]; var dt = ts instanceof Date ? ts : new Date(ts); if (!isNaN(dt)) dateStr = dt.toISOString().split('T')[0]; }
      DB.raw_moves.push({ name: name, id: idx.id >= 0 ? String(r[idx.id] || '') : '', taskType: taskType, startLoc: startLoc, endLoc: endLoc, duration: dur, blocked: blocked, hour: hr, gapSec: gapSec, date: dateStr });
    }
    var workersMap = {}, taskMap = {}, startMap = {}, endMap = {}, hourMap = {}, gapByWorker = {}, blockedCount = 0;
    for (var i = 0; i < DB.raw_moves.length; i++) {
      var r = DB.raw_moves[i];
      var name = r.name;
      var dur = r.duration, taskType = r.taskType, startLoc = r.startLoc, endLoc = r.endLoc, blocked = r.blocked, hr = r.hour, gapSec = r.gapSec;
      if (blocked) blockedCount++;
      if (!workersMap[name]) workersMap[name] = { name: name, id: r.id || '', count: 0, totalSec: 0, minSec: 999999, maxSec: 0, sub_minute_count: 0 };
      workersMap[name].count++; workersMap[name].totalSec += dur; workersMap[name].minSec = Math.min(workersMap[name].minSec, dur); workersMap[name].maxSec = Math.max(workersMap[name].maxSec, dur);
      if (dur < 60) workersMap[name].sub_minute_count++;
      if (!taskMap[taskType]) taskMap[taskType] = { type: taskType, count: 0, totalSec: 0, minSec: 999999, maxSec: 0 };
      taskMap[taskType].count++; taskMap[taskType].totalSec += dur; taskMap[taskType].minSec = Math.min(taskMap[taskType].minSec, dur); taskMap[taskType].maxSec = Math.max(taskMap[taskType].maxSec, dur);
      startMap[startLoc] = (startMap[startLoc] || 0) + 1; endMap[endLoc] = (endMap[endLoc] || 0) + 1;
      hourMap[hr] = (hourMap[hr] || 0) + 1;
      if (gapSec != null && !isNaN(gapSec) && gapSec >= 0) { if (!gapByWorker[name]) gapByWorker[name] = { sum: 0, n: 0 }; gapByWorker[name].sum += gapSec; gapByWorker[name].n++; }
    }
    var workers = Object.keys(workersMap).map(function (k) { var w = workersMap[k]; return { name: w.name, id: w.id || '', count: w.count, avg_secs: Math.round(w.totalSec / w.count * 10) / 10, fastest_secs: w.minSec, slowest_secs: w.maxSec, sub_minute_count: w.sub_minute_count }; }).sort(function (a, b) { return b.count - a.count; });
    var task_types = Object.keys(taskMap).map(function (k) { var t = taskMap[k]; return { type: t.type, count: t.count, avg_secs: Math.round(t.totalSec / t.count * 10) / 10, avg_mins: Math.round(t.totalSec / t.count / 60 * 100) / 100, fastest_secs: t.minSec, slowest_secs: t.maxSec }; }).sort(function (a, b) { return b.count - a.count; });
    var start_locations = Object.keys(startMap).filter(function (k) { return k && k !== 'Unknown'; }).map(function (k) { return { location: k, count: startMap[k] }; }).sort(function (a, b) { return b.count - a.count; });
    var end_locations = Object.keys(endMap).filter(function (k) { return k && k !== 'Unknown'; }).map(function (k) { return { location: k, count: endMap[k] }; }).sort(function (a, b) { return b.count - a.count; });
    var hourly_activity = [];
    for (var h = 0; h < 24; h++) hourly_activity.push({ hour: h, count: hourMap[h] || 0 });
    var gap_stats = Object.keys(gapByWorker).map(function (k) { var g = gapByWorker[k]; return { name: k.trim(), avg_gap_secs: Math.round(g.sum / g.n * 10) / 10, samples: g.n }; }).sort(function (a, b) { return a.avg_gap_secs - b.avg_gap_secs; });
    var totalMoves = workers.reduce(function (a, b) { return a + b.count; }, 0);
    var subMinTotal = workers.reduce(function (a, b) { return a + (b.sub_minute_count || 0); }, 0);
    var wAvg = totalMoves ? workers.reduce(function (a, b) { return a + b.avg_secs * b.count; }, 0) / totalMoves : 0;
    var lastDate = '';
    for (var j = DB.raw_moves.length - 1; j >= 0; j--) if (DB.raw_moves[j].date) { lastDate = DB.raw_moves[j].date; break; }
    DB.workers = workers; DB.task_types = task_types; DB.start_locations = start_locations; DB.end_locations = end_locations; DB.hourly_activity = hourly_activity; DB.gap_stats = gap_stats;
    DB.summary = DB.summary || {}; DB.summary.total_moves = totalMoves; DB.summary.total_workers = workers.length; DB.summary.avg_time_secs = Math.round(wAvg * 10) / 10; DB.summary.sub_minute_total = subMinTotal; DB.summary.sub_minute_pct = totalMoves ? Math.round(subMinTotal / totalMoves * 1000) / 10 : 0; DB.summary.blocked_moves = blockedCount; if (lastDate) DB.summary.date = lastDate;
    workers.forEach(function (w) { var g = gap_stats.find(function (gs) { return gs.name.trim() === w.name.trim(); }); if (g) gapMap[w.name.trim()] = g; });
    return true;
  }

  function processWb(wb) {
    var found = [];
    for (var si = 0; si < wb.SheetNames.length; si++) {
      var sn = wb.SheetNames[si];
      var ws = wb.Sheets[sn];
      if (!ws) continue;
      var allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, cellDates: true });
      if (allRows.length < 2) continue;
      var hi = findHdrRow(allRows);
      var hdr = (allRows[hi] || []).map(function (x) { return String(x || '').toLowerCase().trim(); });

      if (sn.indexOf('Field Ops Task') >= 0 && hdr.some(function (x) { return x.indexOf('task type') >= 0; }) && hdr.some(function (x) { return x.indexOf('duration') >= 0 || x.indexOf('duration taken') >= 0; })) {
        if (parseFieldOpsTask(allRows, hi)) found.push({ t: 'fieldops', l: 'Field Ops Task (' + (DB.workers || []).length + ' workers)' });
      }
      if (hdr.indexOf('task type description') >= 0 && hdr.indexOf('count') >= 0 && hdr.some(function (x) { return x.indexOf('average time') >= 0; })) {
        var tti = hdr.findIndex(function (x) { return x.indexOf('task type description') >= 0; });
        var tci = hdr.findIndex(function (x) { return x === 'count'; });
        var tsi = hdr.findIndex(function (x) { return x.indexOf('average time taken (secs)') >= 0; });
        var tmi = hdr.findIndex(function (x) { return x.indexOf('average time taken (mins)') >= 0; });
        var tfs = hdr.findIndex(function (x) { return x.indexOf('fastest time') >= 0 || x.indexOf('fastest time (secs)') >= 0; });
        var tsl = hdr.findIndex(function (x) { return x.indexOf('slowest time') >= 0 || x.indexOf('slowest time (secs)') >= 0; });
        if (tti >= 0 && tci >= 0) {
          var tts = allRows.slice(hi + 1).filter(function (r) { return r[tti] && r[tci]; }).map(function (r) {
            var type = String(r[tti]).trim();
            var count = parseInt(r[tci], 10) || 0;
            var avgSec = tsi >= 0 ? parseFloat(r[tsi]) : 0;
            var avgMin = tmi >= 0 ? parseFloat(r[tmi]) : (avgSec / 60);
            var fast = tfs >= 0 ? parseFloat(r[tfs]) : 0;
            var slow = tsl >= 0 ? parseFloat(r[tsl]) : 0;
            return { type: type, count: count, avg_secs: avgSec, avg_mins: avgMin, fastest_secs: fast, slowest_secs: slow };
          }).filter(function (t) { return t.count > 0; });
          if (tts.length) { DB.task_types = tts.sort(function (a, b) { return b.count - a.count; }); found.push({ t: 'summary', l: 'Summary (' + tts.length + ' types)' }); }
        }
      }
      if (hdr.indexOf('completed by name') >= 0 && hdr.indexOf('count') >= 0 && hdr.some(function (x) { return x.indexOf('time taken') >= 0 || x.indexOf('average time') >= 0; })) {
        var ni = hdr.findIndex(function (x) { return x.indexOf('completed by name') >= 0; });
        var ci = hdr.findIndex(function (x) { return x === 'count'; });
        var ti = hdr.findIndex(function (x) { return x.indexOf('time taken') >= 0 || x.indexOf('average time') >= 0; });
        var wks = allRows.slice(hi + 1).filter(function (r) { return r[ni] && r[ci] && ['total', 'grand total'].indexOf(String(r[ni]).toLowerCase()) < 0; }).map(function (r) {
          var n = String(r[ni]).trim();
          var c = parseInt(r[ci], 10) || 0;
          var avg = ti >= 0 ? parseFloat(r[ti]) || 0 : 0;
          return { name: n, count: c, avg_secs: Math.round(avg * 10) / 10, fastest_secs: avg, slowest_secs: avg, sub_minute_count: 0 };
        }).filter(function (w) { return w.count > 0; });
        if (wks.length && !found.some(function (f) { return f.t === 'fieldops'; })) { DB.workers = wks.sort(function (a, b) { return b.count - a.count; }); found.push({ t: 'workers', l: 'Workers (' + wks.length + ')' }); }
        else if (wks.length && found.some(function (f) { return f.t === 'fieldops'; })) {
          var wMap = {}; (DB.workers || []).forEach(function (w) { wMap[w.name.trim()] = w; });
          wks.forEach(function (sw) { var ex = wMap[sw.name.trim()]; if (ex) { ex.avg_secs = Math.round((sw.avg_secs || ex.avg_secs) * 10) / 10; } });
        }
      }
      if (hdr.indexOf('row labels') >= 0 && hdr.some(function (x) { return x.indexOf('count of end location') >= 0; }) && sn.indexOf('1 Minute') >= 0) {
        var rli = hdr.findIndex(function (x) { return x.indexOf('row labels') >= 0; });
        var coi = hdr.findIndex(function (x) { return x.indexOf('count of end location') >= 0; });
        var subCnt = {};
        allRows.slice(hi + 1).forEach(function (r) { var n = String(r[rli] || '').trim(); var c = parseInt(r[coi], 10) || 0; if (n && c > 0) subCnt[n] = c; });
        Object.keys(subCnt).forEach(function (n) { var w = (DB.workers || []).find(function (x) { return x.name.trim() === n.trim(); }); if (w) w.sub_minute_count = subCnt[n]; });
        if (Object.keys(subCnt).length) {
          var st = (DB.workers || []).reduce(function (a, b) { return a + (b.sub_minute_count || 0); }, 0);
          var tm = (DB.workers || []).reduce(function (a, b) { return a + b.count; }, 0);
          if (DB.summary) { DB.summary.sub_minute_total = st; DB.summary.sub_minute_pct = tm ? Math.round(st / tm * 1000) / 10 : 0; }
          found.push({ t: 'submin', l: 'Sub-Minute Pivot' });
        }
      }
      if (sn.indexOf('<1 Minute AVG') >= 0 && hdr.indexOf('completed by name') >= 0 && hdr.indexOf('count') >= 0) {
        var avNi = hdr.findIndex(function (x) { return x.indexOf('completed by name') >= 0; });
        var avCi = hdr.findIndex(function (x) { return x === 'count'; });
        var avTi = hdr.findIndex(function (x) { return x.indexOf('average time') >= 0; });
        var subPiv = allRows.slice(hi + 1).filter(function (r) { return r[avNi] && r[avCi]; }).map(function (r) { return { name: String(r[avNi]).trim(), sub_minute_count: parseInt(r[avCi], 10) || 0, avg_secs: avTi >= 0 ? parseFloat(r[avTi]) || 0 : 0 }; }).filter(function (x) { return x.sub_minute_count > 0; });
        if (subPiv.length) { DB.sub_minute_pivot = subPiv.sort(function (a, b) { return b.sub_minute_count - a.sub_minute_count; }); found.push({ t: 'subminavg', l: 'Sub-Minute AVG' }); }
      }
      if (hdr.indexOf('row labels') >= 0 && hdr.some(function (x) { return x.indexOf('column labels') >= 0 || x.indexOf('count of end') >= 0; }) && sn === 'Pivot') {
        var pvRowLab = 0, pvColRow = 1;
        var dests = (allRows[pvColRow] || []).slice(1).filter(function (c) { return c && String(c).trim(); }).map(function (c) { return String(c).trim(); });
        var wp = [];
        for (var pi = 2; pi < allRows.length; pi++) {
          var prow = allRows[pi];
          var wname = String(prow[0] || '').trim();
          if (!wname) break;
          var destsObj = {}; var tot = 0;
          dests.forEach(function (d, di) { var v = parseInt(prow[di + 1], 10) || 0; if (v) { destsObj[d] = v; tot += v; } });
          if (tot > 0) wp.push({ name: wname, destinations: destsObj, total: tot });
        }
        if (wp.length) { DB.worker_pivot = wp; found.push({ t: 'pivot', l: 'Worker×Dest (' + wp.length + ')' }); }
      }
      if (hdr.indexOf('id') >= 0 && hdr.indexOf('name') >= 0 && sn.indexOf('Names') >= 0) {
        var idi = hdr.findIndex(function (x) { return x === 'id'; });
        var nami = hdr.findIndex(function (x) { return x === 'name'; });
        DB.names_ids = {}; allRows.slice(hi + 1).forEach(function (r) { var id = String(r[idi] || '').trim(); var nm = String(r[nami] || '').trim(); if (id && nm) DB.names_ids[id] = nm; });
        if (Object.keys(DB.names_ids || {}).length) found.push({ t: 'names', l: 'Names-IDs' });
      }
      if (hdr.some(function (x) { return x.indexOf('rate type') >= 0; }) && hdr.indexOf('hours') >= 0) {
        var ri = hdr.findIndex(function (x) { return x.indexOf('rate type') >= 0; });
        var hi2 = hdr.findIndex(function (x) { return x === 'hours'; });
        var p = { reg: 0, lunch: 0, overtime: 0, pto: 0 };
        for (var rj = hi + 1; rj < allRows.length; rj++) {
          var row = allRows[rj];
          if (!row[0]) continue;
          var rt = String(row[ri] || '').toLowerCase();
          var h = parseFloat(row[hi2]);
          if (isNaN(h)) continue;
          if (rt.indexOf('reg') >= 0) p.reg += h;
          else if (rt.indexOf('lunch') >= 0) p.lunch += h;
          else if (rt.indexOf('over') >= 0) p.overtime += h;
          else if (rt.indexOf('pto') >= 0) p.pto += h;
        }
        if (p.reg > 0 || p.overtime > 0) { DB.payroll = { reg: p.reg, lunch: p.lunch, overtime: p.overtime, pto: p.pto, employees: DB.payroll && DB.payroll.employees || 0, total_paid: p.reg + p.overtime }; found.push({ t: 'payroll', l: 'Payroll' }); }
      }
      if (/^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)/i.test(sn.trim()) || (hdr.some(function (x) { return x.indexOf('actual driver hours') >= 0; }) && hdr.some(function (x) { return x.indexOf('reporting period') >= 0 || x.indexOf('rent') >= 0; }))) {
        var parsed = parseHoursSheet(sn, allRows);
        if (parsed.length) { DB.hours_overview = (DB.hours_overview || []).filter(function (x) { return x.month !== sn; }); DB.hours_overview = DB.hours_overview.concat(parsed); DB.hours_overview.sort(function (a, b) { return a.date.localeCompare(b.date); }); if (!found.some(function (f) { return f.t === 'ho'; })) found.push({ t: 'ho', l: 'Hours (' + parsed.length + ' days)' }); }
      }
      if ((sn.indexOf('Scans per hour') >= 0 || sn.indexOf('scans per hour') >= 0) && hdr.some(function (x) { return x.indexOf('row labels') >= 0; })) {
        var sphRowLab = hdr.findIndex(function (x) { return x.indexOf('row labels') >= 0; });
        var sphHdr = allRows[hi] || [];
        var dateCols = []; sphHdr.forEach(function (c, j) { if (j > 0 && c && String(c).trim()) dateCols.push({ j: j, label: String(c).trim() }); });
        var sphData = []; for (var sph = hi + 1; sph < allRows.length; sph++) { var r = allRows[sph]; var name = String(r[sphRowLab] || '').trim(); if (!name) continue; var row = { name: name }; dateCols.forEach(function (dc) { var v = parseFloat(String(r[dc.j] || '').replace(/,/g, '')); row[dc.label] = isNaN(v) ? 0 : v; }); sphData.push(row); }
        if (sphData.length && dateCols.length) { DB.scans_per_hour = sphData; found.push({ t: 'sph', l: 'Scans/Hour' }); }
      }
    }
    return found;
  }

  function toast(msg, k) {
    var el = document.createElement('div');
    el.className = 'toast ' + (k || 'inf');
    el.innerHTML = '<span>' + msg + '</span>';
    var tc = $('toast-container');
    if (tc) tc.appendChild(el);
    setTimeout(function () { el.remove(); }, 4000);
  }

  function handleFiles(files) {
    if (!files || !files.length) return;
    DB.raw_moves = [];
    var i = 0;
    function next() {
      if (i >= files.length) {
        var at = document.querySelector('.tab.active');
        if (at && at.dataset && at.dataset.tab) refreshTab(at.dataset.tab);
        return;
      }
      var f = files[i++];
      toast('Reading ' + f.name + '…', 'inf');
      var reader = new FileReader();
      reader.onload = function (ev) {
        try {
          var buf = ev.target.result;
          var wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
          var found = processWb(wb);
          if (!found.length) { toast(f.name + ' — no recognized sheets', 'err'); next(); return; }
          toast('✓ ' + f.name + ' → ' + found.map(function (x) { return x.l; }).join(', '), 'ok');
          next();
        } catch (e) { toast(f.name + ' — ' + e.message, 'err'); next(); }
      };
      reader.readAsArrayBuffer(f);
    }
    next();
  }

  function init() {
    document.addEventListener('click', function (e) {
      var t = e.target;
      while (t && t !== document.body) {
        if (t.classList && t.classList.contains('tab') && t.getAttribute('data-tab')) {
          switchTab(t.getAttribute('data-tab'), t);
          e.preventDefault();
          return;
        }
        t = t.parentElement;
      }
    }, true);
    document.querySelectorAll('.date-pill').forEach(function (p) {
      p.addEventListener('click', function () { setPreset(this.dataset.preset); });
    });
    var dc = $('date-clear');
    if (dc) dc.addEventListener('click', clearFilter);
    var ds = $('date-start'), de = $('date-end');
    if (ds) ds.addEventListener('change', applyDateFilter);
    if (de) de.addEventListener('change', applyDateFilter);

    document.querySelectorAll('#worker-sort .segmented-item').forEach(function (b) {
      b.addEventListener('click', function () {
        document.querySelectorAll('#worker-sort .segmented-item').forEach(function (x) { x.classList.remove('active'); });
        this.classList.add('active');
        wSort = this.dataset.sort;
        uWk();
      });
    });
    var ws = $('worker-search');
    if (ws) ws.addEventListener('input', function () { uWk(); });

    document.querySelectorAll('#cmp-mode .segmented-item').forEach(function (b) {
      b.addEventListener('click', function () { setCmpMode(this.dataset.mode); });
    });

    var fi = $('file-input');
    if (fi) fi.addEventListener('change', function (e) { handleFiles(e.target.files || []); });
    var ib = $('import-btn');
    if (ib) {
      document.body.addEventListener('dragover', function (e) { e.preventDefault(); ib.classList.add('drag'); });
      document.body.addEventListener('dragleave', function (e) { if (!e.relatedTarget || !document.body.contains(e.relatedTarget)) ib.classList.remove('drag'); });
      document.body.addEventListener('drop', function (e) { e.preventDefault(); ib.classList.remove('drag'); handleFiles(e.dataTransfer && e.dataTransfer.files || []); });
    }

    var dates = (DB.hours_overview || []).map(function (d) { return d.date; }).sort();
    if (dates.length && ds && de) { ds.min = dates[0]; ds.max = dates[dates.length - 1]; de.min = dates[0]; de.max = dates[dates.length - 1]; }

    uOv();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
