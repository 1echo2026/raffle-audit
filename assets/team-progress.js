/* team-progress.js — 高经团队周进度核算
 * 输入: ① 高经团队月目标表(团队 + 月目标) ② 有数看板 GMV 表(高经 + 转化金额)
 * 口径: 周目标 = 月目标 ÷ 当月天数 × 本期天数(本期天数可手动改)
 *       目标完成度 = 该团队本期转化金额 ÷ 月目标 × 100%
 *       按完成度降序,第一名高亮;未设目标的团队排最后
 * 全部在浏览器本地计算,文件不上传服务器。
 */
(function () {
  'use strict';

  var GOAL = null;   // {fileName, sheetName, headerIdx, map, dataRows}
  var GMV = null;    // 同上
  var RESULT = null; // {rows, monthDays, periodDays, period, totals}

  // ---------- 基础 ----------
  function str(v) { return v === null || v === undefined ? '' : String(v).trim(); }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
      return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;'}[c];
    });
  }
  function toNum(v) {
    if (typeof v === 'number') return isFinite(v) ? v : NaN;
    var s = str(v).replace(/[,，\s]/g, '').replace(/[¥￥%]/g, '');
    if (!s) return NaN;
    var n = parseFloat(s);
    return isFinite(n) ? n : NaN;
  }
  function money(n) {
    var v = Math.round((Number(n) || 0) * 100) / 100;
    var parts = v.toFixed(2).split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return parts[1] === '00' ? parts[0] : parts.join('.');
  }
  function cleanTeam(v) {
    var s = str(v).replace(/\s+/g, '');
    if (!s) return '';
    if (/^(合计|总计|小计|汇总|总计:)$/.test(s)) return '';
    return s;
  }
  function cleanBase(v) { return str(v).replace(/\s+/g, ''); }
  function log(msg, cls) {
    var el = document.getElementById('log');
    var line = document.createElement('div');
    line.className = cls || '';
    line.textContent = '[' + new Date().toLocaleTimeString() + '] ' + msg;
    el.appendChild(line);
    el.scrollTop = el.scrollHeight;
  }

  // ---------- 表头识别 ----------
  var TEAM_ALIASES = ['高经团队', '高经', '团队', '高经姓名', '高经名称', '团队名称', '负责人'];
  var GOAL_ALIASES = ['月目标', '月度目标', '销售直推目标', '月目标值', '销售目标', 'GMV目标', '目标值', '目标'];
  var GMV_ALIASES = ['转化金额', '销售直推业绩', '转化gmv', '成交金额', '支付金额', '销售金额', '业绩', 'GMV'];
  var BASE_ALIASES = ['基地', '基地名称', '所属基地', '城市'];
  var MGR_ALIASES = ['经理', '销售经理'];
  var SALES_ALIASES = ['销售姓名', '销售', '姓名'];

  function normHeader(v) {
    return str(v)
      .replace(/^求和项[:：]/, '')
      .replace(/^计数项[:：]/, '')
      .replace(/[（(][^）)]*[）)]/g, '')
      .replace(/\s+/g, '');
  }
  function headerHitExact(cell, aliases, excludes) {
    var h = normHeader(cell).toLowerCase();
    if (!h) return false;
    if (excludes && excludes.some(function (x) { return h.indexOf(x) >= 0; })) return false;
    return aliases.some(function (a) { return h === a.toLowerCase(); });
  }
  function headerHitLoose(cell, aliases, excludes) {
    var h = normHeader(cell).toLowerCase();
    if (!h) return false;
    if (excludes && excludes.some(function (x) { return h.indexOf(x) >= 0; })) return false;
    return aliases.some(function (a) { return h.indexOf(a.toLowerCase()) >= 0; });
  }
  // 先找完全相等的列名(避免「基地负责人」被当成「高经团队」),再退化为包含匹配
  function findCol(headers, aliases, excludes) {
    var i;
    for (i = 0; i < headers.length; i++) {
      if (headerHitExact(headers[i], aliases, excludes)) return i;
    }
    for (i = 0; i < headers.length; i++) {
      if (headerHitLoose(headers[i], aliases, excludes)) return i;
    }
    return -1;
  }
  function detectTable(sheets, kind) {
    var best = null;
    sheets.forEach(function (sh) {
      var rows = sh.rows || [];
      var maxH = Math.min(9, rows.length - 1);
      for (var h = 0; h <= maxH; h++) {
        var headers = (rows[h] || []).map(function (x) { return str(x); });
        if (!headers.join('')) continue;
        var teamIdx = findCol(headers, TEAM_ALIASES, ['完成度', '排名', '目标']);
        var valIdx = kind === 'goal'
          ? findCol(headers, GOAL_ALIASES, ['完成度', '达成率', '排名', '业绩', '转化', '实际'])
          : findCol(headers, GMV_ALIASES, ['完成度', '达成率', '排名', '目标', '率']);
        if (teamIdx < 0 || valIdx < 0) continue;
        var dataRows = rows.slice(h + 1);
        var score = 0;
        dataRows.forEach(function (r) {
          if (!r) return;
          if (cleanTeam(r[teamIdx]) && isFinite(toNum(r[valIdx]))) score++;
        });
        if (!best || score > best.score) {
          best = {
            sheetName: sh.name, headerIdx: h, score: score, dataRows: dataRows,
            map: {
              team: teamIdx,
              value: valIdx,
              base: findCol(headers, BASE_ALIASES),
              manager: findCol(headers, MGR_ALIASES),
              sales: findCol(headers, SALES_ALIASES, ['经理'])
            },
            header: headers
          };
        }
      }
    });
    return best;
  }
  async function readWorkbook(file) {
    var buf = await file.arrayBuffer();
    var wb = XLSX.read(buf, {cellDates: true});
    return wb.SheetNames.map(function (sn) {
      return {name: sn, rows: XLSX.utils.sheet_to_json(wb.Sheets[sn], {header: 1, raw: true, defval: ''})};
    });
  }

  // ---------- 载入 ----------
  function colText(tbl, idx) {
    if (idx == null || idx < 0) return '';
    return str((tbl.header || [])[idx]);
  }
  async function handleGoal(file) {
    if (!/\.xlsx?$/i.test(file.name)) { log('只支持 .xlsx / .xls: ' + file.name, 'err'); return; }
    var sheets = await readWorkbook(file);
    var tbl = detectTable(sheets, 'goal');
    if (!tbl || !tbl.score) {
      log('未识别到「高经团队 + 月目标」表头: ' + file.name + '。请确认表头含“高经团队/高经”与“月目标/销售直推目标”', 'err');
      return;
    }
    tbl.fileName = file.name;
    GOAL = tbl;
    var n = Object.keys(buildTargets(tbl)).length;
    document.getElementById('goalFiles').innerHTML = '<div>✅ ' + esc(file.name) + '</div>';
    var box = document.getElementById('goalDetected');
    box.style.display = 'block';
    box.innerHTML = '识别工作表：<b>' + esc(tbl.sheetName) + '</b>（表头第 ' + (tbl.headerIdx + 1) + ' 行）<br>' +
      '团队列：<b>' + esc(colText(tbl, tbl.map.team)) + '</b> · 月目标列：<b>' + esc(colText(tbl, tbl.map.value)) + '</b>' +
      (tbl.map.base >= 0 ? ' · 基地列：<b>' + esc(colText(tbl, tbl.map.base)) + '</b>' : '') +
      '<br>有效团队 ' + n + ' 个';
    log('已载入月目标表: ' + file.name + '（' + tbl.sheetName + '，团队 ' + n + ' 个）', 'ok');
  }
  async function handleGmv(file) {
    if (!/\.xlsx?$/i.test(file.name)) { log('只支持 .xlsx / .xls: ' + file.name, 'err'); return; }
    var sheets = await readWorkbook(file);
    var tbl = detectTable(sheets, 'gmv');
    if (!tbl || !tbl.score) {
      log('未识别到「高经 + 转化金额」表头: ' + file.name + '。请确认表头含“高经/高经团队”与“转化金额/销售直推业绩”', 'err');
      return;
    }
    tbl.fileName = file.name;
    GMV = tbl;
    var agg = buildGmv(tbl);
    var names = Object.keys(agg);
    document.getElementById('gmvFiles').innerHTML = '<div>✅ ' + esc(file.name) + '</div>';
    var box = document.getElementById('gmvDetected');
    box.style.display = 'block';
    box.innerHTML = '识别工作表：<b>' + esc(tbl.sheetName) + '</b>（表头第 ' + (tbl.headerIdx + 1) + ' 行）<br>' +
      '高经列：<b>' + esc(colText(tbl, tbl.map.team)) + '</b> · 金额列：<b>' + esc(colText(tbl, tbl.map.value)) + '</b>' +
      ' · 明细 ' + tbl.score + ' 行<br>高经团队 ' + names.length + ' 个';
    log('已载入 GMV 表: ' + file.name + '（' + tbl.sheetName + '，明细 ' + tbl.score + ' 行）', 'ok');
  }

  // ---------- 聚合 ----------
  function buildTargets(tbl) {
    var map = {};
    (tbl.dataRows || []).forEach(function (r) {
      if (!r) return;
      var team = cleanTeam(r[tbl.map.team]);
      var v = toNum(r[tbl.map.value]);
      if (!team || !isFinite(v)) return;
      if (!map[team]) map[team] = {team: team, target: 0, base: tbl.map.base >= 0 ? cleanBase(r[tbl.map.base]) : ''};
      map[team].target += v;
      if (!map[team].base && tbl.map.base >= 0) map[team].base = cleanBase(r[tbl.map.base]);
    });
    return map;
  }
  function buildGmv(tbl) {
    var agg = {};
    (tbl.dataRows || []).forEach(function (r) {
      if (!r) return;
      var team = cleanTeam(r[tbl.map.team]);
      var amt = toNum(r[tbl.map.value]);
      if (!team || !isFinite(amt)) return;
      if (!agg[team]) agg[team] = {team: team, amount: 0, count: 0, bases: {}, detail: []};
      agg[team].amount += amt;
      agg[team].count++;
      var base = tbl.map.base >= 0 ? cleanBase(r[tbl.map.base]) : '';
      if (base) agg[team].bases[base] = 1;
      agg[team].detail.push([
        base,
        team,
        tbl.map.manager >= 0 ? str(r[tbl.map.manager]) : '',
        tbl.map.sales >= 0 ? str(r[tbl.map.sales]) : '',
        amt
      ]);
    });
    return agg;
  }

  function compute() {
    var monthDays = Math.max(1, Number(document.getElementById('monthDays').value) || 0);
    var periodDays = Math.max(1, Number(document.getElementById('periodDays').value) || 0);
    var period = str(document.getElementById('period').value);
    var targets = buildTargets(GOAL);
    var gmv = buildGmv(GMV);
    var names = {};
    Object.keys(targets).forEach(function (k) { names[k] = 1; });
    Object.keys(gmv).forEach(function (k) { names[k] = 1; });

    var rows = Object.keys(names).map(function (name) {
      var tg = targets[name] || null;
      var g = gmv[name] || null;
      var target = tg ? Math.round(tg.target * 100) / 100 : 0;
      var amount = g ? Math.round(g.amount * 100) / 100 : 0;
      var weekly = Math.round(target / monthDays * periodDays * 100) / 100;
      var rate = target > 0 ? amount / target : null;   // 完成度 = 转化 ÷ 月目标
      var base = (tg && tg.base) || (g ? Object.keys(g.bases).join('、') : '');
      return {
        team: name, base: base, target: target, weekly: weekly, amount: amount,
        rate: rate, count: g ? g.count : 0, detail: g ? g.detail : []
      };
    });
    rows.sort(function (a, b) {
      if ((a.rate == null) !== (b.rate == null)) return a.rate == null ? 1 : -1;
      if (a.rate != null) return b.rate - a.rate;
      return b.amount - a.amount;
    });
    var rank = 0;
    rows.forEach(function (r) { if (r.rate != null) { rank++; r.rank = rank; } else { r.rank = null; } });

    var totals = {target: 0, amount: 0, weekly: 0, teams: rows.length, withGoal: 0, top: rows[0] || null};
    rows.forEach(function (r) {
      totals.target += r.target;
      totals.amount += r.amount;
      totals.weekly += r.weekly;
      if (r.rate != null) totals.withGoal++;
    });
    return {rows: rows, totals: totals, monthDays: monthDays, periodDays: periodDays, period: period};
  }

  // ---------- 渲染 ----------
  function rateClass(rate) {
    if (rate == null) return 'rate-none';
    if (rate >= 1) return 'rate-hi';
    if (rate >= 0.6) return 'rate-mid';
    return 'rate-low';
  }
  function rateText(rate) { return rate == null ? '未设目标' : (rate * 100).toFixed(1) + '%'; }
  function renderTable() {
    var q = str(document.getElementById('searchTeam').value).toLowerCase();
    var rows = RESULT.rows.filter(function (r) {
      if (!q) return true;
      return (r.team + ' ' + r.base).toLowerCase().indexOf(q) >= 0;
    });
    var h = ['<table class="grid" id="teamTable"><tr><th style="width:64px">排名</th><th style="width:90px">基地</th>'
      + '<th style="width:180px">高经团队</th><th style="width:130px">月目标</th>'
      + '<th style="width:130px">周目标（本期 ' + RESULT.periodDays + ' 天）</th>'
      + '<th style="width:140px">本期转化金额</th><th style="width:180px">目标完成度</th></tr>'];
    rows.forEach(function (r, i) {
      var isTop = r.rank === 1;
      var rateCls = rateClass(r.rate);
      var w = r.rate == null ? 0 : Math.min(100, Math.max(0, r.rate * 100));
      h.push('<tr' + (isTop ? ' class="top-team"' : '') + '>'
        + '<td>' + (r.rank == null ? '—' : (isTop ? '<b>1</b>' : r.rank)) + '</td>'
        + '<td>' + esc(r.base || '—') + '</td>'
        + '<td class="team">' + esc(r.team) + (isTop ? '<span class="top-badge">🏆 TOP1</span>' : '') + '</td>'
        + '<td>' + (r.target > 0 ? '¥' + money(r.target) : '<span class="rate-none">未设目标</span>') + '</td>'
        + '<td>' + (r.target > 0 ? '¥' + money(r.weekly) : '—') + '</td>'
        + '<td><b>¥' + money(r.amount) + '</b></td>'
        + '<td><div class="' + rateCls + '">' + rateText(r.rate) + '</div>'
        + '<div class="bar-wrap"><div class="bar-fill' + (r.rate == null ? ' nogoal' : (r.rate < 0.6 ? ' warn' : '')) + '" style="width:' + w + '%"></div></div></td>'
        + '</tr>');
    });
    if (!rows.length) {
      h.push('<tr><td colspan="7" class="muted">没有匹配的记录</td></tr>');
    }
    h.push('</table>');
    document.getElementById('teamTable').innerHTML = h.join('');
  }
  function render() {
    var t = RESULT.totals;
    function card(n, label) {
      return '<div class="card"><div class="card-n">' + n + '</div><div class="card-t">' + label + '</div></div>';
    }
    document.getElementById('stats').innerHTML =
      card(t.teams, '高经团队') + card('¥' + money(t.amount), '本期转化合计') +
      card('¥' + money(t.target), '月目标合计') + card(t.withGoal, '已设目标团队');
    var top = t.top;
    var note = document.getElementById('topNote');
    if (top && top.rate != null) {
      note.style.display = 'block';
      note.innerHTML = '🏆 本期第一名：<b>' + esc(top.team) + '</b>（' + esc(top.base || '—') + '）· 完成度 <b>' +
        rateText(top.rate) + '</b> · 转化 ¥' + money(top.amount) + ' / 月目标 ¥' + money(top.target);
    } else {
      note.style.display = 'none';
    }
    renderTable();
    document.getElementById('results').style.display = 'block';
  }

  // ---------- 导出 ----------
  function exportExcel() {
    if (!RESULT) { log('请先开始核算', 'err'); return; }
    var BORDER = {top: {style: 'thin', color: {rgb: 'FFBFBFBF'}}, bottom: {style: 'thin', color: {rgb: 'FFBFBFBF'}},
                  left: {style: 'thin', color: {rgb: 'FFBFBFBF'}}, right: {style: 'thin', color: {rgb: 'FFBFBFBF'}}};
    var head = ['排名', '基地', '高经团队', '月目标', '周目标（本期' + RESULT.periodDays + '天）', '本期转化金额', '目标完成度'];
    var aoa = [head];
    RESULT.rows.forEach(function (r) {
      aoa.push([r.rank == null ? '' : r.rank, r.base || '', r.team, r.target, r.weekly, r.amount, r.rate]);
    });
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{wch: 6}, {wch: 10}, {wch: 18}, {wch: 14}, {wch: 18}, {wch: 16}, {wch: 12}];
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (var R = 0; R <= range.e.r; R++) {
      for (var C = 0; C <= range.e.c; C++) {
        var addr = XLSX.utils.encode_cell({r: R, c: C});
        if (!ws[addr]) ws[addr] = {t: 's', v: ''};
        ws[addr].s = R === 0
          ? {fill: {fgColor: {rgb: 'FFDDE7FF'}}, font: {bold: true}, alignment: {horizontal: 'center', vertical: 'center'}, border: BORDER}
          : {alignment: {horizontal: C === 2 ? 'left' : 'center', vertical: 'center'}, border: BORDER};
        if (R > 0 && (C === 3 || C === 4 || C === 5)) ws[addr].z = '#,##0.00';
        if (R > 0 && C === 6 && typeof ws[addr].v === 'number') ws[addr].z = '0.0%';
      }
    }
    var wb = XLSX.utils.book_new();
    var title = '高经团队周进度' + (RESULT.period ? '（' + RESULT.period + '）' : '');
    XLSX.utils.book_append_sheet(wb, ws, '高经团队周进度');

    if (GMV && GMV.map.base >= 0) {
      var det = [['基地', '高经团队', '经理', '销售姓名', '转化金额']];
      var agg = buildGmv(GMV);
      Object.keys(agg).forEach(function (k) {
        agg[k].detail.forEach(function (d) { det.push(d); });
      });
      var ws2 = XLSX.utils.aoa_to_sheet(det);
      ws2['!cols'] = [{wch: 10}, {wch: 18}, {wch: 12}, {wch: 12}, {wch: 14}];
      XLSX.utils.book_append_sheet(wb, ws2, '转化明细');
    }
    XLSX.writeFile(wb, title + '.xlsx');
    log('已导出: ' + title + '.xlsx', 'ok');
  }

  // ---------- 绑定 ----------
  function runCompute() {
    if (!GOAL) { log('请先上传「高经团队月目标表」', 'err'); return null; }
    if (!GMV) { log('请先上传「有数看板 GMV 表」', 'err'); return null; }
    try {
      RESULT = compute();
      render();
      log('核算完成: ' + RESULT.rows.length + ' 个团队，本期转化合计 ¥' + money(RESULT.totals.amount) +
        '，第一名 ' + (RESULT.totals.top ? RESULT.totals.top.team + ' ' + rateText(RESULT.totals.top.rate) : '—'), 'ok');
      var noGoal = RESULT.rows.filter(function (r) { return r.rate == null; });
      if (noGoal.length) log('有 ' + noGoal.length + ' 个团队未在月目标表中找到，已排在最后: ' + noGoal.map(function (r) { return r.team; }).join('、'), 'err');
      return RESULT;
    } catch (e) {
      log('核算出错: ' + e.message, 'err');
      return null;
    }
  }
  function bindDrop(dropId, inputId, handler) {
    var drop = document.getElementById(dropId);
    var input = document.getElementById(inputId);
    drop.addEventListener('click', function () { input.click(); });
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('drag'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('drag'); });
    drop.addEventListener('drop', function (e) {
      e.preventDefault(); drop.classList.remove('drag');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) handler(e.dataTransfer.files[0]);
    });
    input.addEventListener('change', function () {
      if (input.files && input.files[0]) handler(input.files[0]);
      input.value = '';
    });
  }

  document.addEventListener('DOMContentLoaded', function () {
    var now = new Date();
    var dim = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
    document.getElementById('monthDays').value = dim;
    bindDrop('goalDrop', 'goalInput', handleGoal);
    bindDrop('gmvDrop', 'gmvInput', handleGmv);

    document.getElementById('btnRun').addEventListener('click', runCompute);
    document.getElementById('btnDownload').addEventListener('click', exportExcel);
    document.getElementById('searchTeam').addEventListener('input', function () { if (RESULT) renderTable(); });

    // 供自动化验证/二次开发调用
    window.TeamProgress = {
      loadGoal: handleGoal,
      loadGmv: handleGmv,
      run: runCompute,
      getResult: function () { return RESULT; }
    };
  });
})();
