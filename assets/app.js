/* app.js — 网页 UI + Excel 读写(xlsx-js-style)。所有计算在浏览器本地完成,文件不上传服务器。 */
(function () {
  'use strict';

  var SLOTS = [
    {key: 'qjRaw', label: '强基订单池(原始)', need: 'order'},
    {key: 'sxRaw', label: '升学订单池(原始)', need: 'order'},
    {key: 'jjDist', label: '进阶中奖名单(奖品发放表)', need: 'dist'},
    {key: 'dfDist', label: '巅峰中奖名单(奖品发放表)', need: 'dist'},
    {key: 'jjForm', label: '进阶问卷(表单填写)', need: 'form'},
    {key: 'dfForm', label: '巅峰问卷(表单填写)', need: 'form'}
  ];
  var files = {};      // key -> {name, rows}
  var lastResult = null;

  // ---------- 异常话术模板(独立页面维护,存 localStorage;常量与读写见 templates-core.js) ----------
  var templates = (window.TplCore ? TplCore.load() : []);

  function fillScript(tpl, c) {
    var name = c.real_name || c.nickname || '';
    var map = {
      '{姓名}': name, '{手机号}': c.phone, '{奖池}': c.pool,
      '{应属奖池}': c.expected_pool || '', '{奖品}': c.prize_name || '',
      '{原因}': c.reason, '{备注}': c.remark,
      '{期数}': c.qishu || '', '{周期}': c.period || '',
      '{出单时间}': c.order_time || '', '{单数}': c.peak_orders || '',
      '{提交时间}': c.submit_time || ''
    };
    var s = tpl.script || '';
    Object.keys(map).forEach(function (k) { s = s.split(k).join(map[k] || '未填写'); });
    return s;
  }
  function scriptFor(c) {
    var t = templates.filter(function (t) { return t.type === c.abnormal_type; })[0];
    if (!t) t = templates.filter(function (t) { return t.type === '其他异常'; })[0];
    return t ? fillScript(t, c) : '';
  }

  // ---------- 样式 ----------
  var BORDER = {top: {style: 'thin', color: {rgb: 'FF7F7F7F'}}, bottom: {style: 'thin', color: {rgb: 'FF7F7F7F'}},
                left: {style: 'thin', color: {rgb: 'FF7F7F7F'}}, right: {style: 'thin', color: {rgb: 'FF7F7F7F'}}};
  function sHeader() { return {fill: {fgColor: {rgb: 'FFD9E1F2'}}, font: {bold: true}, alignment: {horizontal: 'center', vertical: 'center', wrapText: true}, border: BORDER}; }
  function sCell(align) { return {alignment: {horizontal: align || 'left', vertical: 'center', wrapText: true}, border: BORDER}; }
  function sRed() { return {font: {color: {rgb: 'FFC00000'}, bold: true}, alignment: {horizontal: 'left', vertical: 'center', wrapText: true}, border: BORDER}; }
  var MONEY = '#,##0.00';

  function aoa(rows) {
    var ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = rows._cols || [];
    return ws;
  }
  function styleAll(ws, ncols, headerRow) {
    headerRow = headerRow || 0;
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (var R = 0; R <= range.e.r; R++) {
      for (var C = 0; C <= ncols - 1; C++) {
        var addr = XLSX.utils.encode_cell({r: R, c: C});
        var cell = ws[addr];
        if (!cell) { cell = {t: 's', v: ''}; ws[addr] = cell; }
        cell.s = R === headerRow ? sHeader() : sCell();
      }
    }
  }
  function setColWidths(ws, widths) { ws['!cols'] = widths.map(function (w) { return {wch: w}; }); }

  function levelText(s) { return s.prize_level ? (['一','二','三','四','五','六','七','八'][s.prize_level - 1] + '等奖') : ''; }
  function fmtPeak(d) { return d ? AuditCore.fmtDt(d instanceof Date ? d : AuditCore.parseDt(d)) : '未达标'; }

  // ---------- 读取 / 识别 ----------
  function classify(rows, name) {
    var c0 = String((rows[0] && rows[0][0]) || '').trim();
    var h = (rows[1] || []).map(function (x) { return String(x || '').trim(); });
    var h0 = (rows[0] || []).map(function (x) { return String(x || '').trim(); });
    // 订单池: 有"订单池导出结果"标题行,或行0/行1含订单表头特征(账号ID/订单ID)
    var isOrder = c0 === '订单池导出结果' || h0.indexOf('账号ID') >= 0 || h0.indexOf('订单ID') >= 0 || h.indexOf('账号ID') >= 0;
    if (isOrder) {
      if (name.indexOf('强基') >= 0) return 'qjRaw';
      if (name.indexOf('升学') >= 0) return 'sxRaw';
      return h.length > 20 ? 'sxRaw' : 'qjRaw';
    }
    if (h0.indexOf('用户昵称') >= 0 || h.indexOf('用户昵称') >= 0) {
      return name.indexOf('巅峰') >= 0 ? 'dfForm' : 'jjForm';
    }
    // 奖品发放表
    return name.indexOf('巅峰') >= 0 ? 'dfDist' : 'jjDist';
  }

  // 内容特征打分:用于在多个 sheet 中挑选该文件真正有用的那个
  // (腾讯问卷导出的文件常含说明/透视 sheet,真正的表单可能在后续 sheet)
  function scoreSheet(rows, fname) {
    var hh = [];
    (rows[0] || []).concat(rows[1] || []).forEach(function (x) { hh.push(String(x || '').trim()); });
    var j = hh.join(',');
    var s = 0;
    if (j.indexOf('账号ID') >= 0) s += 10;
    if (j.indexOf('订单ID') >= 0) s += 5;
    if (j.indexOf('用户昵称') >= 0) s += 10;
    if (j.indexOf('出单') >= 0) s += 5;
    if (j.indexOf('奖项名称') >= 0) s += 8;
    if (j.indexOf('中奖者昵称') >= 0) s += 8;
    if (j.indexOf('中奖者头像') >= 0) s += 6;
    if (j.indexOf('提交时间') >= 0) s += 4;
    if (fname.indexOf('问卷') >= 0 && j.indexOf('用户昵称') >= 0) s += 15;
    if (fname.indexOf('名单') >= 0 && (j.indexOf('奖项名称') >= 0 || j.indexOf('中奖者昵称') >= 0)) s += 10;
    if (fname.indexOf('订单池') >= 0 && j.indexOf('账号ID') >= 0) s += 15;
    return s;
  }

  async function readFile(file) {
    var buf = await file.arrayBuffer();
    var wb = XLSX.read(buf, {cellDates: true});
    var best = null, bestScore = -1, bestSn = wb.SheetNames[0];
    wb.SheetNames.forEach(function (sn) {
      var ws = wb.Sheets[sn];
      var rows = XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: ''});
      var s = scoreSheet(rows, file.name);
      if (s > bestScore) { bestScore = s; best = rows; bestSn = sn; }
    });
    return best;
  }

  async function handleFiles(fileList) {
    var arr = Array.prototype.slice.call(fileList);
    for (var i = 0; i < arr.length; i++) {
      var f = arr[i];
      if (!/\.xlsx?$/i.test(f.name)) continue;
      try {
        var rows = await readFile(f);
        var key = classify(rows, f.name);
        files[key] = {name: f.name, rows: rows};
      } catch (e) {
        log('读取失败: ' + f.name + ' — ' + e.message, 'err');
      }
    }
    renderSlots();
  }

  function renderSlots() {
    var box = document.getElementById('slots');
    box.innerHTML = '';
    SLOTS.forEach(function (slot) {
      var cur = files[slot.key];
      var div = document.createElement('div');
      div.className = 'slot' + (cur ? ' ok' : '');
      var sel = SLOTS.map(function (s) {
        return '<option value="' + s.key + '"' + (s.key === slot.key ? ' selected' : '') + '>' + s.label + '</option>';
      }).join('');
      div.innerHTML =
        '<div class="slot-label">' + slot.label + '</div>' +
        '<div class="slot-file">' + (cur ? ('✅ ' + cur.name) : '<span class="muted">未选择(自动识别或手动指定)</span>') + '</div>' +
        (cur ? '<select data-file="' + slot.key + '"><option value="">移动到…</option>' +
          SLOTS.filter(function (s) { return s.key !== slot.key; })
            .map(function (s) { return '<option value="' + s.key + '">' + s.label + '</option>'; }).join('') +
          '</select>' : '');
      box.appendChild(div);
    });
    Array.prototype.forEach.call(box.querySelectorAll('select'), function (sel) {
      sel.onchange = function () {
        var from = sel.getAttribute('data-file'), to = sel.value;
        if (to && files[from]) { files[to] = files[from]; delete files[from]; renderSlots(); }
      };
    });
  }

  function log(msg, cls) {
    var el = document.getElementById('log');
    var line = document.createElement('div');
    line.className = cls || '';
    line.textContent = '[' + new Date().toLocaleTimeString() + '] ' + msg;
    el.appendChild(line);
    el.scrollTop = el.scrollHeight;
  }

  // ---------- 运行 ----------
  function run() {
    var missing = SLOTS.filter(function (s) { return !files[s.key]; }).map(function (s) { return s.label; });
    if (missing.length) { alert('还缺少文件:\n' + missing.join('\n')); return; }
    log('开始审核…');
    var result;
    try {
      result = AuditCore.audit({
        rawOrders: {qiangji: files.qjRaw.rows, shengxue: files.sxRaw.rows},
        distribution: {'进阶': files.jjDist.rows, '巅峰': files.dfDist.rows},
        questionnaire: {'进阶': files.jjForm.rows, '巅峰': files.dfForm.rows},
        params: {
          channelPrefix: document.getElementById('prefix').value.trim() || 'grow_xcg_zhuanjs_xiaoshou',
          minAmount: parseFloat(document.getElementById('minAmount').value) || 999,
          threshold: parseFloat(document.getElementById('threshold').value) || 20000,
          period: document.getElementById('period').value.trim(),
          qishu: document.getElementById('qishu').value.trim(),
          formStart: document.getElementById('formStart').value,
          formEnd: document.getElementById('formEnd').value,
          orderStart: document.getElementById('orderStart').value,
          orderEnd: document.getElementById('orderEnd').value
        }
      });
    } catch (e) {
      log('审核出错: ' + e.message, 'err');
      throw e;
    }
    lastResult = result;
    var tInfo = result.params.formStart || result.params.formEnd || result.params.orderStart || result.params.orderEnd
      ? ' (已启用时间范围过滤)' : '';
    log('完成:有效 ' + result.stats.valid + '(进阶' + result.stats.validJJ + '/巅峰' + result.stats.validDF +
        '),无效 ' + result.stats.invalid + ',重复剔除 ' + result.stats.duplicates +
        ';有效订单 强基' + result.validPools['强基'].stats.kept + '/升学' + result.validPools['升学'].stats.kept + tInfo, 'ok');
    renderStats(result);
    renderImagePreview(result);
    renderDetail(result);
    renderInvalid(result);
    renderComm(result);
    document.getElementById('results').style.display = 'block';
  }

  function renderStats(r) {
    var html = '' +
      card('进阶有效', r.stats.validJJ, 'var(--blue)', 'showPoolDetail', '进阶') +
      card('巅峰有效', r.stats.validDF, 'var(--gold)', 'showPoolDetail', '巅峰') +
      card('无效名单', r.stats.invalid, 'var(--red)', 'showInvalidDetail', '') +
      card('重复剔除', r.stats.duplicates, 'var(--gray)', 'showDupDetail', '') +
      card('强基有效订单', r.validPools['强基'].stats.kept, 'var(--green)', 'showPoolOrder', '强基') +
      card('升学有效订单', r.validPools['升学'].stats.kept, 'var(--green)', 'showPoolOrder', '升学');
    document.getElementById('stats').innerHTML = html;
  }
  function card(t, n, c, fn, arg) {
    return '<div class="card clickable" onclick="' + fn + '(\'' + arg + '\')"><div class="card-n" style="color:' + c + '">' + n + '</div><div class="card-t">' + t + '</div></div>';
  }

  // ---------- 统计数字详情弹窗(点击卡片数字弹出,统一列) ----------
  var DETAIL_COLS = ['基地','姓名','工号','奖品奖项与奖品名称','业绩归属渠道','业绩归属时间','订单转化金额','异常情况说明（如有）'];

  function escDetail(v) { return esc(v == null || v === '' ? '—' : v); }
  function prizeLabel(s) {
    var lv = levelText(s);
    return (lv && lv !== '奖品' ? lv + ' ' : '') + (s.prize_name || '');
  }
  function openDetailModal(title, rows) {
    var h = ['<table class="grid"><tr>'];
    DETAIL_COLS.forEach(function (c) { h.push('<th>' + c + '</th>'); });
    h.push('</tr>');
    rows.forEach(function (x, i) {
      h.push('<tr><td style="text-align:left">' + (i + 1) + '</td>');
      DETAIL_COLS.forEach(function (c, j) {
        h.push('<td style="text-align:left">' + escDetail(x[j]) + '</td>');
      });
      h.push('</tr>');
    });
    h.push('</table>');
    document.getElementById('modalTitle').textContent = title;
    document.getElementById('modalBody').innerHTML = h.join('');
    document.getElementById('detailModal').style.display = 'flex';
  }
  function closeDetailModal() {
    document.getElementById('detailModal').style.display = 'none';
  }

  // 弹窗函数供卡片 inline onclick 调用
  window.showPoolDetail = showPoolDetail;
  window.showInvalidDetail = showInvalidDetail;
  window.showDupDetail = showDupDetail;
  window.showPoolOrder = showPoolOrder;

  // 奖池有效中奖人(进阶/巅峰): 已匹配且无异常原因
  function showPoolDetail(pool) {
    var r = lastResult;
    var rows = r.validList.filter(function (s) { return s.pool === pool; }).map(function (s) {
      return [s.base, s.real_name || s.nickname, s.gonghao, prizeLabel(s), s.channel || '—',
              s.order_time || '—', AuditCore.fmtMoney(s.total_amount), s.remark || ''];
    });
    openDetailModal(pool + '奖池 有效中奖人（' + rows.length + ' 人）', rows);
  }

  // 无效名单
  function showInvalidDetail() {
    var r = lastResult;
    var rows = r.invalidList.map(function (s) {
      return [s.base, s.real_name || s.nickname, s.gonghao, prizeLabel(s), s.channel || '—',
              s.order_time || '—', AuditCore.fmtMoney(s.total_amount), s.invalid_reasons || ''];
    });
    openDetailModal('无效名单（' + rows.length + ' 人）', rows);
  }

  // 重复剔除
  function showDupDetail() {
    var r = lastResult;
    var rows = r.duplicates.map(function (s) {
      return [s.base, s.real_name || s.nickname, s.gonghao, prizeLabel(s), s.channel || '—',
              s.order_time || '—', AuditCore.fmtMoney(s.total_amount),
              '同手机号多抽，已剔除（保留最早提交）' + (s.invalid_reasons ? '；' + s.invalid_reasons : '')];
    });
    openDetailModal('重复剔除（' + rows.length + ' 人）', rows);
  }

  // 有效订单池明细(强基/升学)
  function showPoolOrder(poolName) {
    var r = lastResult;
    var pool = r.validPools[poolName];
    var rows = pool.kept.map(function (rr) {
      // 列序: 基地/姓名/工号/奖品/业绩归属渠道/业绩归属时间/订单转化金额/异常情况说明
      return ['', rr[1] || '', rr[25] || '', '', rr[20] || '',
              rr[5] || '', AuditCore.fmtMoney(rr[4]), ''];
    });
    openDetailModal(poolName + '有效订单池（' + rows.length + ' 条，仅销售直推渠道）', rows);
  }

  // ---------- 图片格式预览 ----------
  function sectionTable(title, groups) {
    var h = ['<table class="winner-table">'];
    h.push('<tr><th colspan="4" class="title-row">' + esc(title) + '</th></tr>');
    h.push('<tr><th class="h-pink">奖品等级</th><th class="h-pink">奖品</th><th class="h-pink">基地</th><th class="h-pink">姓名</th></tr>');
    groups.forEach(function (g) {
      g.rows.forEach(function (r, i) {
        h.push('<tr>');
        if (i === 0) {
          h.push('<td class="lv" rowspan="' + g.rows.length + '">' + esc(g.level_text) + '</td>');
          h.push('<td class="pz" rowspan="' + g.rows.length + '">' + esc(g.prize) + '</td>');
        }
        h.push('<td class="base">' + esc(r.base) + '</td><td class="nm">' + esc(r.name) + '</td></tr>');
      });
    });
    h.push('</table>');
    return h.join('');
  }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }

  function renderImagePreview(r) {
    var t1 = document.getElementById('titleDF').value.trim() || '巅峰奖池获奖名单';
    var t2 = document.getElementById('titleJJ').value.trim() || '进阶奖池获奖名单';
    document.getElementById('previewImage').innerHTML =
      sectionTable(t1, r.groups['巅峰']) + sectionTable(t2, r.groups['进阶']);
  }

  function renderDetail(r) {
    var rows = ['<table class="grid"><tr><th>序号</th><th>奖池</th><th>基地</th><th>姓名</th><th>工号</th><th>出单手机号</th><th>奖项</th><th>奖品</th></tr>'];
    var idx = 0;
    [['巅峰', r.groups['巅峰']], ['进阶', r.groups['进阶']]].forEach(function (pp) {
      pp[1].forEach(function (g) {
        g.rows.forEach(function (x) {
          idx++;
          rows.push('<tr><td>' + idx + '</td><td>' + pp[0] + '</td><td>' + esc(x.base) + '</td><td>' + esc(x.name) +
            '</td><td>' + esc(x.gonghao) + '</td><td>' + esc(x.phone) + '</td><td>' + esc(g.level_text) +
            '</td><td>' + esc(g.prize) + '</td></tr>');
        });
      });
    });
    rows.push('</table>');
    document.getElementById('previewDetail').innerHTML = rows.join('');
  }

  function renderInvalid(r) {
    var rows = ['<table class="grid"><tr><th>序号</th><th>奖池</th><th>昵称</th><th>基地</th><th>姓名</th><th>手机号</th><th>奖品</th><th>应属奖池</th><th>无效原因</th></tr>'];
    r.invalidList.forEach(function (s, i) {
      rows.push('<tr><td>' + (i + 1) + '</td><td>' + esc(s.pool) + '</td><td>' + esc(s.nickname) +
        '</td><td>' + esc(s.base) + '</td><td>' + esc(s.real_name) + '</td><td>' + esc(s.phone) +
        '</td><td>' + esc(s.prize_name) + '</td><td>' + esc(s.expected_pool || '—') +
        '</td><td class="reason">' + esc(s.invalid_reasons) + '</td></tr>');
    });
    rows.push('</table>');
    document.getElementById('previewInvalid').innerHTML = rows.join('');
  }

  function renderComm(r) {
    var rows = ['<table class="grid"><tr><th>序号</th><th>奖池</th><th>姓名</th><th>基地</th><th>手机号</th><th>异常类型</th><th>处理话术</th><th>备注</th></tr>'];
    r.communication.forEach(function (c, i) {
      var name = c.real_name || c.nickname || '';
      rows.push('<tr><td>' + (i + 1) + '</td><td>' + esc(c.pool) + '</td><td>' + esc(name) +
        '</td><td>' + esc(c.base) + '</td><td>' + esc(c.phone) + '</td><td>' + esc(c.abnormal_type) +
        '</td><td class="reason" style="text-align:left">' + esc(scriptFor(c)) +
        '</td><td style="text-align:left">' + esc(c.reason || c.remark) + '</td></tr>');
    });
    rows.push('</table>');
    document.getElementById('previewComm').innerHTML = rows.join('');
  }

  // ---------- Excel 导出 ----------
  function downloadValidPool(poolName, pool) {
    var wb = XLSX.utils.book_new();
    var rows = [pool.title, pool.header].concat(pool.kept);
    var ws = aoa(rows);
    styleAll(ws, (pool.header || []).length || 27);
    setColWidths(ws, new Array((pool.header || []).length || 27).fill(14));
    XLSX.utils.book_append_sheet(wb, ws, '订单池导出结果');
    XLSX.writeFile(wb, poolName + '有效订单池.xlsx');
  }

  var AUDIT_COLS = ['序号','奖池(实际)','问卷昵称','中奖名单昵称','基地','真实姓名','出单手机号','奖项等级','奖品名称',
    '提交时间','出单账号ID','订单归属人姓名','归属人工号','工号有效订单行数','工号累计有效金额','工号达2万时间',
    '本手机号订单应属奖池','工号单笔最高金额','核销情况','备注'];

  function recordRow(s, i) {
    return [i, s.pool, s.nickname, s.win_nickname, s.base, s.real_name, s.phone, levelText(s), s.prize_name,
      s.submit_time_raw, s.account, s.sales_name, s.gonghao, s.order_count, s.total_amount, fmtPeak(s.peak_since),
      s.expected_pool, s.max_single_order, s.win_redeem, s.remark || ''];
  }

  function sheetFromRecords(records, widths, reasonColIdx) {
    var rows = [AUDIT_COLS].concat(records.map(recordRow));
    var ws = aoa(rows);
    styleAll(ws, AUDIT_COLS.length);
    setColWidths(ws, widths);
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (var R = 1; R <= range.e.r; R++) {
      [14, 17].forEach(function (C) {
        var c = ws[XLSX.utils.encode_cell({r: R, c: C})];
        if (c) { c.t = 'n'; c.z = MONEY; }
      });
      if (reasonColIdx != null) {
        var addr = XLSX.utils.encode_cell({r: R, c: reasonColIdx});
        if (ws[addr]) ws[addr].s = sRed();
      }
    }
    ws['!freeze'] = {xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen'};
    return ws;
  }

  var AUDIT_W = [5,9,22,22,10,10,14,8,24,17,36,12,11,12,14,19,14,14,9,30];

  // 图片格式 sheet(合并单元格 + 粉头)
  function imageSheet(titleDF, titleJJ, groupsDF, groupsJJ) {
    var aoaRows = [], merges = [], r = 0;
    function addSection(title, groups) {
      aoaRows.push([title, '', '', '']);
      merges.push({s: {r: r, c: 0}, e: {r: r, c: 3}});
      r++;
      aoaRows.push(['奖品等级', '奖品', '基地', '姓名']);
      r++;
      groups.forEach(function (g) {
        var start = r;
        g.rows.forEach(function (x, i) {
          if (i === 0) aoaRows.push([g.level_text, g.prize, x.base, x.name]);
          else aoaRows.push(['', '', x.base, x.name]);
          r++;
        });
        var end = r - 1;
        if (end > start) {
          merges.push({s: {r: start, c: 0}, e: {r: end, c: 0}});
          merges.push({s: {r: start, c: 1}, e: {r: end, c: 1}});
        }
      });
    }
    addSection(titleDF, groupsDF);
    addSection(titleJJ, groupsJJ);
    var ws = aoa(aoaRows);
    ws['!merges'] = merges;
    setColWidths(ws, [10, 26, 10, 12]);
    // 样式:重画并定位标题行/表头行
    var range = XLSX.utils.decode_range(ws['!ref']);
    var R0 = 0;
    // 找出两个 section 的起始行
    var titleRows = {}, headerRows = {};
    for (var R = 0; R <= range.e.r; R++) {
      var v = ws[XLSX.utils.encode_cell({r: R, c: 0})];
      var vv = v ? v.v : '';
      if (vv === titleDF || vv === titleJJ) titleRows[R] = 1;
      if (vv === '奖品等级') headerRows[R] = 1;
    }
    for (R = 0; R <= range.e.r; R++) {
      for (var C = 0; C < 4; C++) {
        var cell = ws[XLSX.utils.encode_cell({r: R, c: C})];
        if (!cell) { cell = {t: 's', v: ''}; ws[XLSX.utils.encode_cell({r: R, c: C})] = cell; }
        cell.s = {border: BORDER, alignment: {horizontal: 'center', vertical: 'center', wrapText: true}};
        if (titleRows[R]) {
          cell.s.fill = {fgColor: {rgb: 'FFF8CBAD'}};
          cell.s.font = {bold: true, sz: 13};
        } else if (headerRows[R]) {
          cell.s.fill = {fgColor: {rgb: 'FFFCE4D6'}};
          cell.s.font = {bold: true};
        }
      }
    }
    return ws;
  }

  // 明细表 sheet:基地/姓名/奖品奖项与名/工号/出单手机号/业绩归属时间/累计订单金额(金额倒序)
  function detailSheet(groupsDF, groupsJJ) {
    var rows = [['序号','奖池','基地','姓名','奖品奖项与名','工号','出单手机号','业绩归属时间','累计订单金额']];
    var all = [];
    [['巅峰', groupsDF], ['进阶', groupsJJ]].forEach(function (pp) {
      pp[1].forEach(function (g) {
        g.rows.forEach(function (x) {
          all.push({
            pool: pp[0], base: x.base, name: x.name, gonghao: x.gonghao, phone: x.phone,
            prize: (g.level_text && g.level_text !== '奖品' ? g.level_text + ' ' : '') + g.prize,
            order_time: x.order_time || '', amount: x.total_amount || 0
          });
        });
      });
    });
    all.sort(function (a, b) { return b.amount - a.amount; });
    all.forEach(function (x, i) {
      rows.push([i + 1, x.pool, x.base, x.name, x.prize, x.gonghao, x.phone, x.order_time, x.amount]);
    });
    var ws = aoa(rows);
    styleAll(ws, 9);
    setColWidths(ws, [6, 8, 12, 10, 30, 12, 14, 20, 15]);
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (var R = 1; R <= range.e.r; R++) {
      var c8 = ws[XLSX.utils.encode_cell({r: R, c: 8})];
      if (c8) { c8.t = 'n'; c8.z = MONEY; }
    }
    return ws;
  }

  function duplicateSheet(dups) {
    var rows = [['序号','奖池','问卷昵称','基地','真实姓名','出单手机号','所中奖项','提交时间','处理方式','无效/备注']];
    dups.forEach(function (s, i) {
      rows.push([i + 1, s.pool, s.nickname, s.base, s.real_name, s.phone, s.prize_name,
        s.submit_time_raw, '同手机号多抽，已剔除（保留最早提交）', s.invalid_reasons || '']);
    });
    var ws = aoa(rows);
    styleAll(ws, 10);
    setColWidths(ws, [5, 6, 22, 10, 10, 14, 24, 17, 30, 40]);
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (var R = 1; R <= range.e.r; R++) {
      ws[XLSX.utils.encode_cell({r: R, c: 8})].s = sRed();
    }
    return ws;
  }

  function nameSheet(rowsGh) {
    var rows = [['序号','订单归属人姓名','工号','关联出单手机号（一个工号可对应多个）','手机号数量','姓名对应工号数','工号累计有效金额','工号达2万时间','重名标记']];
    rowsGh.forEach(function (x, i) {
      rows.push([i + 1, x.name, x.gonghao, x.phones.join('、'), x.phones.length, x.gh_count,
        x.total, fmtPeak(x.peak_since), x.is_dup ? '重名：同名对应多个工号' : '']);
    });
    var ws = aoa(rows);
    styleAll(ws, 9);
    setColWidths(ws, [5, 14, 12, 56, 10, 13, 15, 19, 26]);
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (var R = 1; R <= range.e.r; R++) {
      var c7 = ws[XLSX.utils.encode_cell({r: R, c: 6})];
      if (c7) { c7.t = 'n'; c7.z = MONEY; }
      var idx = R - 1;
      if (rowsGh[idx] && rowsGh[idx].is_dup) {
        [1, 5, 8].forEach(function (C) {
          var cell = ws[XLSX.utils.encode_cell({r: R, c: C})];
          cell.s = {font: {color: {rgb: 'FFC00000'}, bold: true}, fill: {fgColor: {rgb: 'FFFFC7CE'}},
                    alignment: {horizontal: 'center', vertical: 'center', wrapText: true}, border: BORDER};
        });
      }
    }
    return ws;
  }

  // 异常名单沟通话术 sheet
  function commSheet(comm) {
    var rows = [['序号','奖池','姓名','基地','出单手机号','异常类型','处理话术','备注']];
    comm.forEach(function (c, i) {
      rows.push([i + 1, c.pool, c.real_name || c.nickname || '', c.base, c.phone,
        c.abnormal_type, scriptFor(c), c.reason || c.remark]);
    });
    var ws = aoa(rows);
    styleAll(ws, 8);
    setColWidths(ws, [5, 8, 12, 10, 14, 16, 66, 40]);
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (var R = 1; R <= range.e.r; R++) {
      var c6 = ws[XLSX.utils.encode_cell({r: R, c: 6})];
      if (c6) c6.s = sRed();
    }
    return ws;
  }

  // 异常名单 sheet:无效 + 重复 + 含退款/换课订单,异常类型列红色
  function abnormalSheet(comm) {
    var rows = [['序号','奖池','姓名','基地','出单手机号','奖品','异常类型','原因','备注']];
    comm.forEach(function (c, i) {
      rows.push([i + 1, c.pool, c.real_name || c.nickname || '', c.base, c.phone,
        c.prize_name, c.abnormal_type, c.reason || c.remark, c.remark]);
    });
    var ws = aoa(rows);
    styleAll(ws, 9);
    setColWidths(ws, [5, 8, 12, 10, 14, 26, 16, 56, 24]);
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (var R = 1; R <= range.e.r; R++) {
      var c6 = ws[XLSX.utils.encode_cell({r: R, c: 6})];
      if (c6) c6.s = sRed();
    }
    return ws;
  }

  function downloadAll() {
    if (!lastResult) return;
    var r = lastResult;
    var t1 = document.getElementById('titleDF').value.trim() || '巅峰奖池获奖名单';
    var t2 = document.getElementById('titleJJ').value.trim() || '进阶奖池获奖名单';
    downloadValidPool('强基', r.validPools['强基']);
    downloadValidPool('升学', r.validPools['升学']);

    // 最终产物: 1 个 Excel,4 个子工作表(图片格式获奖名单 / 中奖人员明细表 / 异常名单 / 异常名单沟通话术)
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, imageSheet(t1, t2, r.groups['巅峰'], r.groups['进阶']), '获奖名单（图片格式）');
    XLSX.utils.book_append_sheet(wb, detailSheet(r.groups['巅峰'], r.groups['进阶']), '中奖人员明细表');
    XLSX.utils.book_append_sheet(wb, abnormalSheet(r.communication), '异常名单');
    XLSX.utils.book_append_sheet(wb, commSheet(r.communication), '异常名单沟通话术');
    XLSX.writeFile(wb, '中奖名单审核结果.xlsx');
    log('已导出: 中奖名单审核结果.xlsx(含 获奖名单(图片格式) / 中奖人员明细表 / 异常名单 / 异常名单沟通话术 4 个子表) + 强基/升学有效订单池');
  }

  // ---------- 应用周期到标题 ----------
  function applyPeriodToTitles() {
    var p = (document.getElementById('period').value || '').trim();
    if (!p) return;
    var rule = /\（[^（）]*\）/;
    ['titleDF', 'titleJJ'].forEach(function (id) {
      var el = document.getElementById(id);
      var v = el.value.trim();
      if (v && rule.test(v)) el.value = v.replace(rule, '（' + p + '）');
    });
  }

  // ---------- 绑定 ----------
  document.addEventListener('DOMContentLoaded', function () {
    renderSlots();
    var zone = document.getElementById('dropzone');
    var input = document.getElementById('fileInput');
    input.addEventListener('change', function () { handleFiles(input.files); });
    zone.addEventListener('dragover', function (e) { e.preventDefault(); zone.classList.add('drag'); });
    zone.addEventListener('dragleave', function () { zone.classList.remove('drag'); });
    zone.addEventListener('drop', function (e) {
      e.preventDefault();
      zone.classList.remove('drag');
      handleFiles(e.dataTransfer.files);
    });
    zone.addEventListener('click', function () { input.click(); });
    document.getElementById('btnRun').addEventListener('click', run);
    document.getElementById('btnDownload').addEventListener('click', downloadAll);
    // 详情弹窗:关闭按钮 + 点击遮罩关闭
    document.getElementById('modalClose').addEventListener('click', closeDetailModal);
    document.getElementById('detailModal').addEventListener('click', function (e) {
      if (e.target === this) closeDetailModal();
    });
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeDetailModal(); });
    document.getElementById('period').addEventListener('input', applyPeriodToTitles);
    ['titleDF', 'titleJJ'].forEach(function (id) {
      document.getElementById(id).addEventListener('input', function () { if (lastResult) renderImagePreview(lastResult); });
    });

    // 暴露给自动化自测使用
    window.App = {
      handleFiles: handleFiles,
      run: run,
      downloadAll: downloadAll,
      getFiles: function () {
        var out = {};
        Object.keys(files).forEach(function (k) { out[k] = files[k].name; });
        return out;
      },
      getResult: function () { return lastResult; },
      getTemplates: function () { return templates; },
      setTemplates: function (arr) { templates = arr; if (window.TplCore) TplCore.save(templates); }
    };
  });
})();
