/* app.js — 网页 UI + Excel 读写(xlsx-js-style)。所有计算在浏览器本地完成,文件不上传服务器。 */
(function () {
  'use strict';

  var SLOTS = [
    {key: 'qjRaw', label: '强基订单池', need: 'order'},
    {key: 'sxRaw', label: '升学订单池', need: 'order'},
    {key: 'jjForm', label: '进阶奖池问卷', need: 'form'},
    {key: 'dfForm', label: '巅峰奖池问卷', need: 'form'}
  ];
  var files = {};      // key -> {name, rows}
  var lastResult = null;

  // 人工裁决层(绿色通道): form_i -> {green_code};跨"重新审核"持续生效
  var overrides = {};
  var pendingEdit = null;    // 修改弹窗正在处理的 form_i
  var HIST_KEY = 'raffle_history_v1';
  var GREEN_KEY = 'raffle_green_codes_v1';  // 首页生成的绿色通道码(填在问卷"出单手机号/ID"栏即放行)

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
  // 一个文件识别为一个角色(4 文件流程: 2 订单池 + 2 问卷):
  //  订单池文件 → qjRaw / sxRaw
  //  问卷文件 → jjForm / dfForm(取"表单填写"表;发放情况表不再需要,问卷表单已含基地/姓名/手机号等审核所需信息)
  function classifySheets(sheets, name) {
    var out = {};
    var orderRows = null;
    var distFound = false;
    sheets.forEach(function (sh) {
      var hh = [];
      (sh.rows[0] || []).concat(sh.rows[1] || []).forEach(function (x) { hh.push(String(x || '').trim()); });
      var j = hh.join(',');
      if (!orderRows && (j.indexOf('账号ID') >= 0 || j.indexOf('订单ID') >= 0)) {
        orderRows = sh.rows;
        return;
      }
      if (j.indexOf('用户昵称') >= 0) {
        out[name.indexOf('巅峰') >= 0 ? 'dfForm' : 'jjForm'] = sh.rows;
      } else if (j.indexOf('奖项名称') >= 0 || j.indexOf('中奖者昵称') >= 0 ||
                 /[一二三四五六七八九十]等奖[：:、\s].*?×/.test(j)) {
        distFound = true;
      }
    });
    if (orderRows) {
      if (name.indexOf('强基') >= 0) out.qjRaw = orderRows;
      else if (name.indexOf('升学') >= 0) out.sxRaw = orderRows;
      else out[(orderRows[0] || []).length > 20 ? 'sxRaw' : 'qjRaw'] = orderRows;
    }
    return {roles: out, distFound: distFound};
  }

  async function readFile(file) {
    var buf = await file.arrayBuffer();
    var wb = XLSX.read(buf, {cellDates: true});
    var sheets = [];
    wb.SheetNames.forEach(function (sn) {
      var ws = wb.Sheets[sn];
      sheets.push({sn: sn, rows: XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: ''})});
    });
    return sheets;
  }

  async function handleFiles(fileList) {
    var arr = Array.prototype.slice.call(fileList);
    for (var i = 0; i < arr.length; i++) {
      var f = arr[i];
      if (!/\.xlsx?$/i.test(f.name)) continue;
      try {
        var sheets = await readFile(f);
        var res = classifySheets(sheets, f.name);
        var roles = res.roles;
        if (!Object.keys(roles).length) {
          log('未识别: ' + f.name + ' — ' + (res.distFound
            ? '只找到"发放情况"表(4 文件流程不需要该表),未找到"表单填写"表'
            : '未找到"表单填写"表(需含"用户昵称"表头)或订单池表(需含"账号ID/订单ID"表头)'), 'err');
        }
        Object.keys(roles).forEach(function (k) { files[k] = {name: f.name, rows: roles[k]}; });
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
        if (!to || !files[from]) return;
        files[to] = files[from];
        delete files[from];
        renderSlots();
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
      result = AuditCore.audit(auditInput());
    } catch (e) {
      log('审核出错: ' + e.message, 'err');
      throw e;
    }
    lastResult = applyOverrides(result);
    var tInfo = lastResult.params.formStart || lastResult.params.formEnd || lastResult.params.orderStart || lastResult.params.orderEnd
      ? ' (已启用时间范围过滤)' : '';
    log('完成:有效 ' + lastResult.stats.valid + '(进阶' + lastResult.stats.validJJ + '/巅峰' + lastResult.stats.validDF +
        '),无效 ' + lastResult.stats.invalid + ',重复剔除 ' + lastResult.stats.duplicates +
        ';有效订单 强基' + lastResult.validPools['强基'].stats.kept + '/升学' + lastResult.validPools['升学'].stats.kept + tInfo, 'ok');
    renderAll(lastResult);
    document.getElementById('results').style.display = 'block';
    saveHistory();
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
  // 业绩归属渠道保留订单池里查到的原值
  function cleanChannel(c) { return c || ''; }
  // 问卷奖品查询:在进阶/巅峰两个奖池问卷(全部提交记录)中,按掩码手机号/真实姓名查找客户所中奖品
  function buildPrizeLookup() {
    var r = lastResult;
    var byPhone = {}, byName = {};
    function add(s) {
      var lbl = prizeLabel(s);
      if (!lbl) return;
      var p = s.phone ? AuditCore.maskPhone(s.phone) : '';
      if (p) (byPhone[p] = byPhone[p] || []).push(lbl);
      var nm = s.real_name || '';
      if (nm) (byName[nm] = byName[nm] || []).push(lbl);
    }
    (r.validList || []).forEach(add);
    (r.invalidList || []).forEach(add);
    (r.duplicates || []).forEach(add);
    return {byPhone: byPhone, byName: byName};
  }
  function lookupPrize(lu, maskedPhone, name) {
    var arr = (maskedPhone && lu.byPhone[maskedPhone]) || (name && lu.byName[name]) || [];
    var seen = {}, uniq = [];
    arr.forEach(function (p) { if (!seen[p]) { seen[p] = 1; uniq.push(p); } });
    return uniq.join('、');
  }
  function openDetailModal(title, rows, cols) {
    var COLS = cols || DETAIL_COLS;
    var h = ['<table class="grid"><tr><th>序号</th>'];
    COLS.forEach(function (c) { h.push('<th>' + c + '</th>'); });
    h.push('</tr>');
    rows.forEach(function (x, i) {
      h.push('<tr><td style="text-align:left">' + (i + 1) + '</td>');
      COLS.forEach(function (c, j) {
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
      return [s.base, s.real_name || s.nickname, s.gonghao, prizeLabel(s), cleanChannel(s.channel) || '—',
              s.order_time || '—', AuditCore.fmtMoney(s.total_amount), s.remark || ''];
    });
    openDetailModal(pool + '奖池 有效中奖人（' + rows.length + ' 人）', rows);
  }

  // 无效名单
  function showInvalidDetail() {
    var r = lastResult;
    var rows = r.invalidList.map(function (s) {
      return [s.base, s.real_name || s.nickname, s.gonghao, prizeLabel(s), cleanChannel(s.channel) || '—',
              s.order_time || '—', AuditCore.fmtMoney(s.total_amount), s.invalid_reasons || ''];
    });
    openDetailModal('无效名单（' + rows.length + ' 人）', rows);
  }

  // 重复剔除
  function showDupDetail() {
    var r = lastResult;
    var rows = r.duplicates.map(function (s) {
      return [s.base, s.real_name || s.nickname, s.gonghao, prizeLabel(s), cleanChannel(s.channel) || '—',
              s.order_time || '—', AuditCore.fmtMoney(s.total_amount),
              '同手机号多抽，已剔除（保留最早提交）' + (s.invalid_reasons ? '；' + s.invalid_reasons : '')];
    });
    openDetailModal('重复剔除（' + rows.length + ' 人）', rows);
  }

  // 有效订单池明细(强基/升学)
  // 订单池列: c1=真实姓名 c4=订单金额 c5=支付时间 c7=业绩归属时间 c20=订单归属渠道 c25=工号 c26=组织架构(含城市)
  var BASE_CITIES = ['北京','武汉','成都','西安','合肥','新乡','郑州'];
  function baseFromOrg(org) {
    var s = String(org || '');
    for (var i = 0; i < BASE_CITIES.length; i++) {
      if (s.indexOf(BASE_CITIES[i]) >= 0) return BASE_CITIES[i];
    }
    return '';
  }
  function showPoolOrder(poolName) {
    var r = lastResult;
    var pool = r.validPools[poolName];
    var lu = buildPrizeLookup();
    // 出单手机号:按掩码从问卷(全部提交记录)反查全号,查不到则显示订单池掩码号
    var phoneFull = {};
    (r.validList || []).concat(r.invalidList || []).concat(r.duplicates || []).forEach(function (s) {
      if (!s.phone) return;
      var m = AuditCore.maskPhone(s.phone);
      if (!phoneFull[m]) phoneFull[m] = s.phone;
    });
    var cols = ['基地','姓名','工号','出单手机号','奖品','业绩归属渠道','业绩归属时间','订单转化金额','异常情况说明（如有）'];
    var rows = pool.kept.map(function (rr) {
      // 奖品在进阶/巅峰两个奖池问卷中按掩码手机号(优先)/真实姓名查找
      var masked = AuditCore.maskPhone(String(rr[2] || '').trim());
      var prize = lookupPrize(lu, masked, String(rr[1] || '').trim());
      // 列序: 基地/姓名/工号/出单手机号/奖品/业绩归属渠道/业绩归属时间/订单转化金额/异常情况说明
      return [baseFromOrg(rr[26]), rr[1] || '', rr[25] || '', phoneFull[masked] || rr[2] || '',
              prize, cleanChannel(rr[20]), rr[7] || rr[5] || '', AuditCore.fmtMoney(rr[4]), ''];
    });
    openDetailModal(poolName + '有效订单池（' + rows.length + ' 条，仅销售直推渠道）', rows, cols);
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

  // 获奖名单标题:活动周期名称与中奖期数即标题来源,自动生成,无需单独填写
  function awardTitle(poolName) {
    var p = (document.getElementById('period').value || '').trim();
    return poolName + '奖池' + (p ? '（' + p + '）' : '') + '获奖名单';
  }

  function renderImagePreview(r) {
    var t1 = awardTitle('巅峰');
    var t2 = awardTitle('进阶');
    document.getElementById('previewImage').innerHTML =
      sectionTable(t1, r.groups['巅峰']) + sectionTable(t2, r.groups['进阶']);
  }

  // ---------- 查找过滤 + 高亮 ----------
  function searchKws(id) {
    var el = document.getElementById(id);
    if (!el) return [];
    return el.value.trim().toLowerCase().split(/\s+/).filter(function (k) { return !!k; });
  }
  function kwMatch(kws, values) {
    if (!kws || !kws.length) return true;
    return kws.every(function (k) {
      return values.some(function (v) { return String(v == null ? '' : v).toLowerCase().indexOf(k) >= 0; });
    });
  }
  function hl(text, kws) {
    var s = esc(text);
    if (!kws || !kws.length) return s;
    kws.forEach(function (kw) {
      var rx = new RegExp('(' + kw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'gi');
      s = s.replace(rx, '<mark>$1</mark>');
    });
    return s;
  }
  function detailFields(d) {
    return [d.pool, d.base, d.name, d.gh, d.phone, d.lvText, d.prize, d.green || '',
            d.conv == null ? '' : AuditCore.fmtMoney(d.conv),
            d.count == null ? '' : String(d.count),
            d.total == null ? '' : AuditCore.fmtMoney(d.total)];
  }
  function invalidFields(s) {
    return [s.pool, s.nickname, s.base, s.real_name, s.phone, s.prize_name, s.expected_pool, s.invalid_reasons, s.abnormal_type || ''];
  }
  function rankMode() {
    return !!(document.getElementById('rankByCount') && document.getElementById('rankByCount').checked);
  }
  // 可编辑表格单元格(明细表 基地/姓名/工号/出单手机号/奖品;无效名单 无效原因)
  function edCell(fi, f, value, kws) {
    return '<span class="ed" contenteditable="true" spellcheck="false" data-fi="' + fi + '" data-f="' + f + '">' + hl(value, kws) + '</span>';
  }
  // 明细条目:排行模式=按工号出单单数从多到少;默认=原分组顺序
  function detailItems(r) {
    var items = [];
    if (rankMode()) {
      (r.validList || []).slice().sort(function (a, b) {
        return (b.order_count || 0) - (a.order_count || 0) || (b.total_amount || 0) - (a.total_amount || 0);
      }).forEach(function (s, i) {
        items.push({rank: i + 1, pool: s.pool, base: s.base, name: s.real_name || s.nickname,
          gh: s.gonghao || '', phone: s.phone, lvText: levelText(s) || '', prize: s.prize_name || '',
          conv: s.convert_amount || 0, count: s.order_count || 0, total: s.total_amount || 0, green: s.green_code || '', fi: s.form_i});
      });
    } else {
      [['巅峰', r.groups['巅峰']], ['进阶', r.groups['进阶']]].forEach(function (pp) {
        pp[1].forEach(function (g) {
          g.rows.forEach(function (x) {
            items.push({pool: pp[0], base: x.base, name: x.name, gh: x.gonghao || '', phone: x.phone,
              lvText: g.level_text, prize: g.prize, conv: x.convert_amount || 0, count: x.order_count || 0,
              total: null, green: x.green_code || '', fi: x.form_i});
          });
        });
      });
    }
    return items;
  }

  function renderDetail(r) {
    var kws = searchKws('searchDetail');
    var rank = rankMode();
    var all = detailItems(r);
    var items = all.filter(function (d) { return kwMatch(kws, detailFields(d)); });
    var hasGreen = (r.validList || []).some(function (s) { return s.green_code; });
    var h = ['<div class="muted search-count">共 ' + all.length + ' 人'
      + (kws.length ? ' · 匹配 ' + items.length + ' 人' : '') + ' · 点击 基地/姓名/工号/出单手机号/奖品 单元格可直接修改</div>'];
    h.push('<table class="grid"><tr><th>序号</th>');
    if (rank) h.push('<th>排行</th>');
    h.push('<th>奖池</th><th>基地</th><th>姓名</th><th>工号</th><th>出单手机号</th><th>转化金额</th><th>单量</th><th>奖项</th><th>奖品</th>');
    if (rank) h.push('<th>累计金额</th>');
    if (hasGreen) h.push('<th>绿色通道</th>');
    h.push('</tr>');
    items.forEach(function (x, i) {
      h.push('<tr><td>' + (i + 1) + '</td>');
      if (rank) h.push('<td>' + x.rank + '</td>');
      h.push('<td>' + hl(x.pool, kws) + '</td>'
        + '<td>' + edCell(x.fi, 'base', x.base, kws) + '</td>'
        + '<td>' + edCell(x.fi, 'name', x.name, kws) + '</td>'
        + '<td>' + edCell(x.fi, 'gh', x.gh, kws) + '</td>'
        + '<td>' + edCell(x.fi, 'phone', x.phone, kws) + '</td>'
        + '<td>' + hl(AuditCore.fmtMoney(x.conv), kws) + '</td>'
        + '<td>' + (x.count || 0) + '</td>'
        + '<td>' + hl(x.lvText, kws) + '</td>'
        + '<td>' + edCell(x.fi, 'prize', x.prize, kws) + '</td>');
      if (rank) h.push('<td>' + AuditCore.fmtMoney(x.total) + '</td>');
      if (hasGreen) h.push('<td>' + (x.green ? '<span class="green-tag">🟢 ' + esc(x.green) + '</span>' : '') + '</td>');
      h.push('</tr>');
    });
    h.push('</table>');
    document.getElementById('previewDetail').innerHTML = h.join('');
  }

  function renderInvalid(r) {
    var kws = searchKws('searchInvalid');
    var list = (r.invalidList || []).filter(function (s) { return kwMatch(kws, invalidFields(s)); });
    var h = ['<div class="muted search-count">无效共 ' + (r.invalidList || []).length + ' 人'
      + (kws.length ? ' · 匹配 ' + list.length + ' 人' : '') + ' · 点击"无效原因"单元格可直接修改</div>'];
    h.push('<table class="grid"><tr><th>序号</th><th>奖池</th><th>昵称</th><th>基地</th><th>姓名</th><th>手机号</th><th>奖品</th><th>应属奖池</th><th>无效原因</th><th>操作</th></tr>');
    list.forEach(function (s, i) {
      h.push('<tr><td>' + (i + 1) + '</td><td>' + hl(s.pool, kws) + '</td><td>' + hl(s.nickname, kws) + '</td>'
        + '<td>' + hl(s.base, kws) + '</td><td>' + hl(s.real_name, kws) + '</td><td>' + hl(s.phone, kws) + '</td>'
        + '<td>' + hl(s.prize_name, kws) + '</td><td>' + hl(s.expected_pool || '—', kws) + '</td>'
        + '<td class="reason">' + edCell(s.form_i == null ? '' : s.form_i, 'reason', s.invalid_reasons, kws) + '</td>'
        + '<td class="ops"><button type="button" class="btn plain btn-min pass-btn" data-fi="' + (s.form_i == null ? '' : s.form_i) + '">✅ 通过</button> '
        + '<button type="button" class="btn plain btn-min edit-btn" data-fi="' + (s.form_i == null ? '' : s.form_i) + '">✏️ 修改</button></td></tr>');
    });
    h.push('</table>');
    document.getElementById('previewInvalid').innerHTML = h.join('');
  }

  // ---------- 人工裁决层:绿色通道 + 修改重匹配 + 统一重渲染 ----------
  function findSub(fi) {
    var r = lastResult;
    if (!r || fi == null || fi === '') return null;
    var all = (r.validList || []).concat(r.invalidList || []).concat(r.duplicates || []);
    for (var i = 0; i < all.length; i++) if (all[i].form_i === fi) return all[i];
    return null;
  }

  // 按 validList 重建 groups(绿色通道转入的人也要出现在获奖名单/明细/导出)
  function rebuildGroups(r) {
    var CN = ['一','二','三','四','五','六','七','八'];
    function tg(pool, maxLevel) {
      var buckets = {}, ungraded = [];
      for (var lv = 1; lv <= maxLevel; lv++) buckets[lv] = [];
      (r.validList || []).forEach(function (s) {
        if (s.pool !== pool) return;
        if (s.prize_level && buckets[s.prize_level]) buckets[s.prize_level].push(s);
        else ungraded.push(s);
      });
      function toRow(s) {
        var b = s.base || '';
        if (b.lastIndexOf('基地') === b.length - 2) b = b.slice(0, -2);
        return {base: b, name: s.real_name, gonghao: s.gonghao || '', phone: s.phone,
                nickname: s.nickname, redeem: s.win_redeem || '',
                order_time: s.order_time || '', total_amount: s.total_amount || 0,
                green_code: s.green_code || '', form_i: s.form_i, account: s.account || '',
                convert_amount: s.convert_amount || 0, order_count: s.order_count || 0};
      }
      var out = [];
      for (var l = 1; l <= maxLevel; l++) {
        if (!buckets[l] || !buckets[l].length) continue;
        out.push({level: l, level_text: CN[l - 1] + '等奖', prize: buckets[l][0].prize_name, rows: buckets[l].map(toRow)});
      }
      if (ungraded.length) {
        var byName = {}, order = [];
        ungraded.forEach(function (s) {
          var k = s.prize_name || '未填写奖品';
          if (!byName[k]) { byName[k] = []; order.push(k); }
          byName[k].push(s);
        });
        order.forEach(function (k) {
          out.push({level: null, level_text: '奖品', prize: k, rows: byName[k].map(toRow)});
        });
      }
      return out;
    }
    r.groups = {'巅峰': tg('巅峰', 4), '进阶': tg('进阶', 8)};
  }

  function applyOverrides(r) {
    var moved = 0;
    var keepInvalid = [];
    (r.invalidList || []).forEach(function (s) {
      var o = overrides[s.form_i];
      if (o) {
        s.green_code = o.green_code;
        s.green_note = s.invalid_reasons || '';
        s.invalid_reasons = '';
        s.abnormal_type = '';
        s.remark = '绿色通道通过(' + o.green_code + ')';
        r.validList.push(s);
        moved++;
      } else {
        keepInvalid.push(s);
      }
    });
    r.invalidList = keepInvalid;
    if (moved) rebuildGroups(r);
    r.stats = r.stats || {};
    r.stats.valid = r.validList.length;
    r.stats.validJJ = r.validList.filter(function (s) { return s.pool === '进阶'; }).length;
    r.stats.validDF = r.validList.filter(function (s) { return s.pool === '巅峰'; }).length;
    r.stats.invalid = r.invalidList.length;
    r.communication = (r.communication || []).filter(function (c) { return !overrides[c.form_i]; });
    return r;
  }

  function renderAll(r) {
    renderStats(r);
    renderImagePreview(r);
    renderDetail(r);
    renderInvalid(r);
    renderComm(r);
  }

  function auditInput() {
    return {
      rawOrders: {qiangji: files.qjRaw.rows, shengxue: files.sxRaw.rows},
      distribution: {'进阶': [], '巅峰': []},
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
        orderEnd: document.getElementById('orderEnd').value,
        greenCodes: loadGreenCodes()
      }
    };
  }

  function rerunAudit() {
    var result = AuditCore.audit(auditInput());
    lastResult = applyOverrides(result);
    log('已重新匹配:有效 ' + lastResult.stats.valid + '(进阶' + lastResult.stats.validJJ + '/巅峰' + lastResult.stats.validDF +
        '),无效 ' + lastResult.stats.invalid + ',重复剔除 ' + lastResult.stats.duplicates, 'ok');
    renderAll(lastResult);
    return lastResult;
  }

  // ---------- 绿色通道 ----------
  // 首页生成随机码:问卷"出单手机号/ID"栏填写该码 → 审核自动放行(见 audit-core GREEN_SET)
  function loadGreenCodes() {
    try { return JSON.parse(localStorage.getItem(GREEN_KEY)) || []; } catch (e) { return []; }
  }
  function saveGreenCodes(arr) {
    try { localStorage.setItem(GREEN_KEY, JSON.stringify(arr)); } catch (e) {}
  }
  function genGreenCode() {
    var s = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789', out = '';
    for (var i = 0; i < 8; i++) out += s.charAt(Math.floor(Math.random() * s.length));
    var codes = loadGreenCodes();
    codes.push(out);
    saveGreenCodes(codes);
    renderGreenCodes();
    log('已生成绿色通道码: ' + out + '(共 ' + codes.length + ' 个)。将码填在问卷"出单用户手机号/ID"栏即可直接通过', 'ok');
    return out;
  }
  function renderGreenCodes() {
    var box = document.getElementById('greenCodesBox');
    if (!box) return;
    var codes = loadGreenCodes();
    if (!codes.length) {
      box.innerHTML = '<div class="muted">暂无绿色通道码。点"生成绿色通道码"后,把码告知抽奖人填写在问卷"出单用户手机号/ID"栏,审核时自动放行并出现在中奖名单。</div>';
      return;
    }
    box.innerHTML = '<div class="muted">绿色通道码(填在问卷"出单手机号/ID"栏即直接通过):</div><div class="gc-list">' +
      codes.map(function (c) {
        return '<span class="gc-chip"><b>' + esc(c) + '</b>' +
          '<button type="button" class="btn plain btn-min gc-copy" data-c="' + esc(c) + '">复制</button>' +
          '<button type="button" class="btn plain btn-min gc-del" data-c="' + esc(c) + '">删除</button></span>';
      }).join('') +
      '<button type="button" class="btn plain btn-min" id="gcClear">清空全部</button></div>';
  }
  // 无效名单行内"通过"按钮:直接转入中奖名单,无需复制粘贴码(码后台自动生成仅留凭证)
  function greenPass(fi) {
    var s = findSub(fi);
    if (!s) { alert('未找到该记录'); return; }
    var code = 'GC' + Date.now().toString(36).toUpperCase().slice(-6);
    overrides[fi] = {green_code: code};
    applyOverrides(lastResult);
    renderAll(lastResult);
    log('✅ 通过: ' + (s.real_name || s.nickname) + '(' + (s.phone || '') + ') 已转入中奖名单,凭证码 ' + code, 'ok');
    return code;
  }

  // ---------- 无效名单手动修改(写回原始问卷行,重新匹配后可转有效) ----------
  function openEditModal(fi) {
    var s = findSub(fi);
    if (!s) { alert('未找到该记录'); return; }
    if (!s.raw_row) { alert('该记录来自历史存档,未绑定原始问卷行;请重新上传文件审核后再修改'); return; }
    pendingEdit = fi;
    document.getElementById('editName').value = s.real_name || '';
    document.getElementById('editPhone').value = s.phone || '';
    document.getElementById('editBase').value = s.base || '';
    document.getElementById('editInfo').textContent =
      '当前: ' + (s.real_name || s.nickname || '') + ' · ' + (s.phone || '') + ' · ' + (s.base || '') + ' · ' + (s.invalid_reasons || '');
    document.getElementById('editModal').style.display = 'flex';
  }
  function saveEdit() {
    var s = findSub(pendingEdit);
    if (!s) { alert('未找到该记录'); return; }
    if (!s.raw_row) { alert('该记录来自历史存档,未绑定原始问卷行,无法改写;请重新上传文件审核后再修改。'); return; }
    var nm = (document.getElementById('editName').value || '').trim();
    var ph = (document.getElementById('editPhone').value || '').trim();
    var bs = (document.getElementById('editBase').value || '').trim();
    var cp = AuditCore.cleanPhone(ph);
    if (ph && !cp.ok) { alert('手机号格式不正确(需 1 开头 11 位数字)'); return; }
    // 写回原始问卷行:重新审核即按新信息匹配订单,能匹配且无异常则转入有效名单
    s.raw_row[5] = nm;
    s.raw_row[6] = ph;
    s.raw_row[4] = bs;
    document.getElementById('editModal').style.display = 'none';
    rerunAudit();
    log('已修改并重新匹配: ' + (nm || s.nickname) + ' 手机号=' + ph + ' 基地=' + bs);
  }

  // ---------- 历史记录(localStorage 快照) ----------
  function loadHistory() {
    try { return JSON.parse(localStorage.getItem(HIST_KEY)) || []; } catch (e) { return []; }
  }
  function miniSub(s) {
    return {
      form_i: s.form_i, pool: s.pool, nickname: s.nickname, win_nickname: s.win_nickname,
      base: s.base, real_name: s.real_name, phone: s.phone, phone_format_ok: s.phone_format_ok,
      prize_level: s.prize_level, prize_name: s.prize_name, submit_time_raw: s.submit_time_raw,
      expected_pool: s.expected_pool, invalid_reasons: s.invalid_reasons, remark: s.remark,
      abnormal_type: s.abnormal_type, green_code: s.green_code, green_note: s.green_note,
      is_duplicate: !!s.is_duplicate, channel: s.channel, account: s.account, sales_name: s.sales_name,
      gonghao: s.gonghao, order_count: s.order_count, total_amount: s.total_amount,
      max_single_order: s.max_single_order, peak_since: s.peak_since, order_time: s.order_time,
      peak_orders: s.peak_orders, win_redeem: s.win_redeem
    };
  }
  function snapResult(r) {
    return {
      validList: (r.validList || []).map(miniSub),
      invalidList: (r.invalidList || []).map(miniSub),
      duplicates: (r.duplicates || []).map(miniSub),
      communication: r.communication || [],
      nameGhRows: r.nameGhRows || [],
      ghStats: r.ghStats || {},
      groups: r.groups || {'进阶': [], '巅峰': []},
      params: r.params || {},
      stats: r.stats || {},
      validPools: r.validPools || {}
    };
  }
  function saveHistory() {
    var r = lastResult;
    if (!r) { alert('还没有审核结果可保存'); return; }
    var h = loadHistory();
    h.unshift({
      id: Date.now(),
      ts: new Date().toLocaleString(),
      period: (r.params && r.params.period) || '',
      qishu: (r.params && r.params.qishu) || '',
      stats: {
        submitted: r.stats.submitted, valid: r.stats.valid, validJJ: r.stats.validJJ,
        validDF: r.stats.validDF, invalid: r.stats.invalid, duplicates: r.stats.duplicates,
        qj: (r.validPools['强基'] && r.validPools['强基'].stats.kept) || 0,
        sx: (r.validPools['升学'] && r.validPools['升学'].stats.kept) || 0
      },
      result: snapResult(r)
    });
    h = h.slice(0, 15);
    try {
      localStorage.setItem(HIST_KEY, JSON.stringify(h));
    } catch (e) {
      log('历史记录保存失败: ' + e.message, 'err');
      return h;
    }
    log('已保存历史记录(共 ' + h.length + ' 条)', 'ok');
    renderHistory(h);
    return h;
  }
  function renderHistory(h) {
    var el = document.getElementById('historyList');
    if (!el) return;
    if (!h || !h.length) {
      el.innerHTML = '<div class="muted">暂无历史记录。点击「开始审核」自动存档,或点上方「保存当前审核结果」手动存档。</div>';
      return;
    }
    var top = '<div class="hist-actions-top"><button type="button" class="btn plain" id="histClear">🗑 清空全部</button></div>';
    var html = h.map(function (x) {
      return '<div class="hist-item"><div class="hist-main">🕘 ' + esc(x.ts) + ' · 第' + esc(String(x.qishu || '-')) + '期 '
        + esc(x.period || '') + '</div><div class="hist-stats">提交 ' + x.stats.submitted
        + ' · 有效 ' + x.stats.valid + '(进阶' + x.stats.validJJ + '/巅峰' + x.stats.validDF + ') · 无效 '
        + x.stats.invalid + ' · 重复 ' + x.stats.duplicates + ' · 强基订单 ' + x.stats.qj + ' · 升学订单 ' + x.stats.sx + '</div>'
        + '<div class="hist-actions"><button type="button" class="btn plain hist-load" data-id="' + x.id + '">📂 加载</button>'
        + '<button type="button" class="btn plain hist-del" data-id="' + x.id + '">🗑 删除</button></div></div>';
    }).join('');
    el.innerHTML = top + html;
  }

  // 本地记录按钮:不用上传文件即可回看历史
  function showHistoryPanel() {
    document.getElementById('results').style.display = 'block';
    if (!lastResult) {
      document.getElementById('stats').innerHTML =
        '<div class="muted" style="padding:10px 0">尚未审核本轮数据。以下为本地保存的历次审核记录,点「加载」可回看当时结果。</div>';
    }
    document.querySelectorAll('.tab').forEach(function (x) { x.classList.toggle('active', x.dataset.tab === 't-history'); });
    document.querySelectorAll('.tabpane').forEach(function (x) { x.classList.toggle('active', x.id === 't-history'); });
    renderHistory(loadHistory());
  }

  // ---------- 编辑单元格保存 ----------
  function saveDetailEdit(el) {
    var fi = parseInt(el.getAttribute('data-fi'), 10);
    var f = el.getAttribute('data-f');
    if (isNaN(fi)) return;
    var s = findSub(fi);
    if (!s) return;
    var v = (el.textContent || '').trim();
    var old = '';
    if (f === 'base') { old = s.base; s.base = v; if (s.raw_row) s.raw_row[4] = v; }
    else if (f === 'name') { old = s.real_name; s.real_name = v; if (s.raw_row) s.raw_row[5] = v; }
    else if (f === 'gh') { old = s.gonghao || ''; s.gonghao = v; }
    else if (f === 'phone') { old = s.phone; s.phone = v; if (s.raw_row) s.raw_row[6] = v; }
    else if (f === 'prize') { old = s.prize_name; s.prize_name = v; }
    else return;
    if (v === (old || '')) return;
    rebuildGroups(lastResult);
    renderStats(lastResult);
    renderDetail(lastResult);
    renderImagePreview(lastResult);
    log('已修改中奖明细(' + ({base: '基地', name: '姓名', gh: '工号', phone: '出单手机号', prize: '奖品'}[f] || f) + '): ' +
        (old || '—') + ' → ' + (v || '—'));
  }
  function saveInvalidEdit(el) {
    var fi = parseInt(el.getAttribute('data-fi'), 10);
    if (isNaN(fi)) return;
    var s = findSub(fi);
    if (!s) return;
    var v = (el.textContent || '').trim();
    if (v === (s.invalid_reasons || '')) return;
    s.invalid_reasons = v;
    if (AuditCore.classifyAbnormal) s.abnormal_type = AuditCore.classifyAbnormal(s);
    (lastResult.communication || []).forEach(function (c) { if (c.form_i === fi) c.reason = v; });
    renderInvalid(lastResult);
    renderComm(lastResult);
    log('已修改无效原因: ' + (s.real_name || s.nickname) + ' → ' + (v || '(已清空)'));
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
    var rows = (pool.title && pool.title.length ? [pool.title] : []).concat([pool.header], pool.kept);
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

  // 明细表 sheet:在出单手机号后加 转化金额/单量(按工号数透)列;金额按累计订单金额倒序;含绿色通道时追加一列
  function detailSheet(groupsDF, groupsJJ) {
    var all = [];
    [['巅峰', groupsDF], ['进阶', groupsJJ]].forEach(function (pp) {
      pp[1].forEach(function (g) {
        g.rows.forEach(function (x) {
          all.push({
            pool: pp[0], base: x.base, name: x.name, gonghao: x.gonghao, phone: x.phone,
            prize: (g.level_text && g.level_text !== '奖品' ? g.level_text + ' ' : '') + g.prize,
            convert: x.convert_amount || 0, count: x.order_count || 0,
            order_time: x.order_time || '', amount: x.total_amount || 0, green: x.green_code || ''
          });
        });
      });
    });
    all.sort(function (a, b) { return b.amount - a.amount; });
    var hasGreen = all.some(function (x) { return x.green; });
    var head = ['序号','奖池','基地','姓名','奖品奖项与名','工号','出单手机号','转化金额','单量','业绩归属时间','累计订单金额'];
    if (hasGreen) head.push('绿色通道');
    var rows = [head];
    all.forEach(function (x, i) {
      var rr = [i + 1, x.pool, x.base, x.name, x.prize, x.gonghao, x.phone, x.convert, x.count, x.order_time, x.amount];
      if (hasGreen) rr.push(x.green);
      rows.push(rr);
    });
    var ws = aoa(rows);
    styleAll(ws, hasGreen ? 12 : 11);
    setColWidths(ws, hasGreen ? [6, 8, 12, 10, 30, 12, 14, 13, 8, 20, 15, 14] : [6, 8, 12, 10, 30, 12, 14, 13, 8, 20, 15]);
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (var R = 1; R <= range.e.r; R++) {
      [7, 10].forEach(function (C) {
        var c = ws[XLSX.utils.encode_cell({r: R, c: C})];
        if (c) { c.t = 'n'; c.z = MONEY; }
      });
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
    downloadValidPool('强基', r.validPools['强基']);
    downloadValidPool('升学', r.validPools['升学']);
    downloadAuditResult();
  }

  // 期数标签:文件名用(第【】期)。取中奖期数,空则用活动周期,再空则"本"
  function periodTag() {
    var q = (document.getElementById('qishu').value || '').trim();
    if (q) return q;
    var p = (document.getElementById('period').value || '').trim();
    return p || '本';
  }

  // 最终产物: 1 个 Excel,5 个子工作表(图片格式获奖名单 / 中奖人员明细表 / 异常名单 / 异常名单沟通话术 / 单量审核名单)
  function auditBook() {
    var r = lastResult;
    var t1 = awardTitle('巅峰');
    var t2 = awardTitle('进阶');
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, imageSheet(t1, t2, r.groups['巅峰'], r.groups['进阶']), '获奖名单（图片格式）');
    XLSX.utils.book_append_sheet(wb, detailSheet(r.groups['巅峰'], r.groups['进阶']), '中奖人员明细表');
    XLSX.utils.book_append_sheet(wb, abnormalSheet(r.communication), '异常名单');
    XLSX.utils.book_append_sheet(wb, commSheet(r.communication), '异常名单沟通话术');
    XLSX.utils.book_append_sheet(wb, rankSheet(r), '单量审核名单');
    return wb;
  }
  function downloadAuditResult() {
    XLSX.writeFile(auditBook(), '第' + periodTag() + '期中奖名单审核结果.xlsx');
    log('已导出: 第' + periodTag() + '期中奖名单审核结果.xlsx(5 个子表)');
  }

  // 强基+升学有效订单池合并一个文件(两个子表),文件名带期数
  function poolSheet(pool) {
    var rows = [pool.title, pool.header].concat(pool.kept);
    var ws = aoa(rows);
    styleAll(ws, (pool.header || []).length || 27);
    setColWidths(ws, new Array((pool.header || []).length || 27).fill(14));
    return ws;
  }
  function downloadPools() {
    var r = lastResult;
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, poolSheet(r.validPools['强基']), '强基有效订单池');
    XLSX.utils.book_append_sheet(wb, poolSheet(r.validPools['升学']), '升学有效订单池');
    XLSX.writeFile(wb, '第' + periodTag() + '期有效订单池.xlsx');
    log('已导出: 第' + periodTag() + '期有效订单池.xlsx(强基+升学两个子表)');
  }

  // 单量审核名单:按工号汇总出单单数从多到少(作为审核结果的子表)
  function rankSheet(r) {
    var rows = [['序号','订单归属人姓名','工号','关联出单手机号','手机号数量','出单单数','工号累计有效金额','工号单笔最高金额','工号达2万时间']];
    var list = (r.nameGhRows || []).map(function (x) {
      var st = (r.ghStats || {})[x.gonghao] || {};
      return {name: x.name, gh: x.gonghao, phones: x.phones.join('、'), phn: x.phones.length,
              count: st.row_count || 0, total: st.total || x.total || 0, max: st.max_single || 0,
              peak: st.peak_since || x.peak_since || null};
    });
    list.sort(function (a, b) { return b.count - a.count || b.total - a.total; });
    list.forEach(function (x, i) {
      rows.push([i + 1, x.name, x.gh, x.phones, x.phn, x.count, x.total, x.max, fmtPeak(x.peak)]);
    });
    var ws = aoa(rows);
    styleAll(ws, 9);
    setColWidths(ws, [5, 14, 12, 56, 9, 9, 15, 15, 19]);
    var range = XLSX.utils.decode_range(ws['!ref']);
    for (var R = 1; R <= range.e.r; R++) {
      var c6 = ws[XLSX.utils.encode_cell({r: R, c: 6})];
      if (c6) { c6.t = 'n'; c6.z = MONEY; }
      var c7 = ws[XLSX.utils.encode_cell({r: R, c: 7})];
      if (c7) { c7.t = 'n'; c7.z = MONEY; }
    }
    return ws;
  }

  // 第N期奖品底表:奖池/奖项等级/奖品/份数(中奖人数) + 用户id(按出单手机号在订单池匹配到的账号ID)
  function prizeBaseSheet(r) {
    var rows = [['序号','奖池','奖项等级','奖品名称','份数(中奖人数)','用户id(按出单手机号匹配)','备注']];
    var idx = 0;
    [['巅峰', r.groups['巅峰']], ['进阶', r.groups['进阶']]].forEach(function (pp) {
      pp[1].forEach(function (g) {
        idx++;
        rows.push([idx, pp[0], g.level_text, g.prize, g.rows.length,
                   g.rows.map(function (x) { return x.account || ''; }).join('、'), '']);
      });
    });
    var ws = aoa(rows);
    styleAll(ws, 7);
    setColWidths(ws, [6, 8, 12, 32, 14, 56, 20]);
    return ws;
  }
  function downloadPrizeBase() {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, prizeBaseSheet(lastResult), '奖品底表');
    XLSX.writeFile(wb, '第' + periodTag() + '期奖品底表.xlsx');
    log('已导出: 第' + periodTag() + '期奖品底表.xlsx');
  }

  // ---------- 导出选择弹窗 ----------
  function openExportModal() {
    if (!lastResult) { alert('请先点击「开始审核」'); return; }
    var n = periodTag();
    document.getElementById('expPoolsLabel').textContent = '第' + n + '期有效订单池.xlsx(强基+升学两个子表)';
    document.getElementById('expAuditLabel').textContent = '第' + n + '期中奖名单审核结果.xlsx(5 个子表:获奖名单图片格式/中奖人员明细表/异常名单/异常名单沟通话术/单量审核名单)';
    document.getElementById('expPrizeLabel').textContent = '第' + n + '期奖品底表.xlsx(奖池/奖项等级/奖品/份数/用户id)';
    document.getElementById('exportModal').style.display = 'flex';
  }
  function doExport() {
    var sel = [];
    document.querySelectorAll('.exp-item').forEach(function (c) { if (c.checked) sel.push(c.value); });
    if (!sel.length) { alert('请至少勾选一项导出内容'); return; }
    document.getElementById('exportModal').style.display = 'none';
    sel.forEach(function (v) {
      if (v === 'pools') downloadPools();
      else if (v === 'audit') downloadAuditResult();
      else if (v === 'prize') downloadPrizeBase();
    });
    log('已导出 ' + sel.length + ' 项文件', 'ok');
  }

  // ---------- 时间范围按钮 + 弹窗 ----------
  function fmtT(v) { return v ? v.replace('T', ' ') : '不限'; }
  function updateTimeBtns() {
    var fBtn = document.getElementById('btnTimeForm');
    var oBtn = document.getElementById('btnTimeOrder');
    if (fBtn) {
      fBtn.textContent = '⏱ ' + fmtT(document.getElementById('formStart').value) + ' ~ ' + fmtT(document.getElementById('formEnd').value);
      fBtn.classList.toggle('set', !!(document.getElementById('formStart').value || document.getElementById('formEnd').value));
    }
    if (oBtn) {
      oBtn.textContent = '⏱ ' + fmtT(document.getElementById('orderStart').value) + ' ~ ' + fmtT(document.getElementById('orderEnd').value);
      oBtn.classList.toggle('set', !!(document.getElementById('orderStart').value || document.getElementById('orderEnd').value));
    }
  }
  function openTimeModal(kind) {
    updateTimeBtns();
    document.getElementById(kind === 'form' ? 'timeFormModal' : 'timeOrderModal').style.display = 'flex';
  }
  function saveTimeModal(kind) {
    document.getElementById(kind === 'form' ? 'timeFormModal' : 'timeOrderModal').style.display = 'none';
    updateTimeBtns();
    log('时间范围已更新:' + (kind === 'form' ? '问卷提交' : '订单支付') + ' ' +
        document.getElementById('btnTime' + (kind === 'form' ? 'Form' : 'Order')).textContent);
  }
  function clearTimeModal(kind) {
    var a = document.getElementById(kind === 'form' ? 'formStart' : 'orderStart');
    var b = document.getElementById(kind === 'form' ? 'formEnd' : 'orderEnd');
    a.value = '';
    b.value = '';
    document.getElementById(kind === 'form' ? 'timeFormModal' : 'timeOrderModal').style.display = 'none';
    updateTimeBtns();
    log('已清除' + (kind === 'form' ? '问卷提交' : '订单支付') + '时间范围(不再限制)');
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
    document.getElementById('btnDownload').addEventListener('click', openExportModal);
    // 详情弹窗:关闭按钮 + 点击遮罩关闭
    document.getElementById('modalClose').addEventListener('click', closeDetailModal);
    document.getElementById('detailModal').addEventListener('click', function (e) {
      if (e.target === this) closeDetailModal();
    });
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeDetailModal(); });
    document.getElementById('period').addEventListener('input', function () { if (lastResult) renderImagePreview(lastResult); });

    // 时间范围按钮 + 弹窗
    document.getElementById('btnTimeForm').addEventListener('click', function () { openTimeModal('form'); });
    document.getElementById('btnTimeOrder').addEventListener('click', function () { openTimeModal('order'); });
    document.getElementById('timeFormSave').addEventListener('click', function () { saveTimeModal('form'); });
    document.getElementById('timeOrderSave').addEventListener('click', function () { saveTimeModal('order'); });
    document.getElementById('timeFormClear').addEventListener('click', function () { clearTimeModal('form'); });
    document.getElementById('timeOrderClear').addEventListener('click', function () { clearTimeModal('order'); });
    document.getElementById('timeFormClose').addEventListener('click', function () { document.getElementById('timeFormModal').style.display = 'none'; });
    document.getElementById('timeOrderClose').addEventListener('click', function () { document.getElementById('timeOrderModal').style.display = 'none'; });
    document.getElementById('timeFormModal').addEventListener('click', function (e) {
      if (e.target === this) this.style.display = 'none';
    });
    document.getElementById('timeOrderModal').addEventListener('click', function (e) {
      if (e.target === this) this.style.display = 'none';
    });
    updateTimeBtns();

    // 导出选择弹窗
    document.getElementById('expConfirm').addEventListener('click', doExport);
    document.getElementById('expAll').addEventListener('click', function () {
      document.querySelectorAll('.exp-item').forEach(function (c) { c.checked = true; });
    });
    document.getElementById('expNone').addEventListener('click', function () {
      document.querySelectorAll('.exp-item').forEach(function (c) { c.checked = false; });
    });
    document.getElementById('exportClose').addEventListener('click', function () {
      document.getElementById('exportModal').style.display = 'none';
    });
    document.getElementById('exportModal').addEventListener('click', function (e) {
      if (e.target === this) this.style.display = 'none';
    });

    // 绿色通道(首页生成码 + 码列表操作)
    document.getElementById('btnGenGreen').addEventListener('click', genGreenCode);
    renderGreenCodes();
    document.getElementById('greenCodesBox').addEventListener('click', function (e) {
      var el = e.target;
      while (el && el !== this && el.tagName !== 'BUTTON') el = el.parentElement;
      if (!el || el === this) return;
      if (el.id === 'gcClear') {
        if (confirm('确定清空全部绿色通道码?')) { saveGreenCodes([]); renderGreenCodes(); log('已清空绿色通道码'); }
        return;
      }
      var c = el.getAttribute('data-c');
      if (!c) return;
      if (el.classList.contains('gc-copy')) {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(c).then(function () { log('绿色通道码已复制: ' + c); }, function () {});
        }
      } else if (el.classList.contains('gc-del')) {
        saveGreenCodes(loadGreenCodes().filter(function (x) { return x !== c; }));
        renderGreenCodes();
        log('已删除绿色通道码: ' + c);
      }
    });

    // 无效名单弹窗(修改)
    document.getElementById('editSave').addEventListener('click', saveEdit);
    document.getElementById('editCancel').addEventListener('click', function () {
      document.getElementById('editModal').style.display = 'none';
    });
    document.getElementById('editModal').addEventListener('click', function (e) {
      if (e.target === this) this.style.display = 'none';
    });

    // 无效名单每行操作按钮 + 无效原因单元格编辑
    document.getElementById('previewInvalid').addEventListener('click', function (e) {
      var el = e.target;
      while (el && el !== this && el.tagName !== 'BUTTON') el = el.parentElement;
      if (!el || el === this) return;
      var fiRaw = el.getAttribute('data-fi');
      if (fiRaw == null || fiRaw === '') return;
      var fi = parseInt(fiRaw, 10);
      if (el.classList.contains('pass-btn')) greenPass(fi);
      else if (el.classList.contains('edit-btn')) openEditModal(fi);
    });
    document.getElementById('previewInvalid').addEventListener('focusout', function (e) {
      var el = e.target;
      if (el && el.classList && el.classList.contains('ed') && el.getAttribute('data-f') === 'reason') saveInvalidEdit(el);
    });
    // 明细表单元格编辑
    document.getElementById('previewDetail').addEventListener('focusout', function (e) {
      var el = e.target;
      if (el && el.classList && el.classList.contains('ed') && el.getAttribute('data-f') !== 'reason') saveDetailEdit(el);
    });

    // 历史记录按钮(步骤2 开始审核旁)
    document.getElementById('btnHistory').addEventListener('click', showHistoryPanel);

    // 查找过滤(输入即重新渲染)
    document.getElementById('searchDetail').addEventListener('input', function () { if (lastResult) renderDetail(lastResult); });
    document.getElementById('searchInvalid').addEventListener('input', function () { if (lastResult) renderInvalid(lastResult); });
    // 单量排行选项
    document.getElementById('rankByCount').addEventListener('change', function () { if (lastResult) renderDetail(lastResult); });

    // 历史记录
    document.getElementById('btnSaveHistory').addEventListener('click', saveHistory);
    document.getElementById('historyList').addEventListener('click', function (e) {
      var el = e.target;
      while (el && el !== this && el.tagName !== 'BUTTON') el = el.parentElement;
      if (!el || el === this) return;
      if (el.id === 'histClear') {
        if (confirm('确定清空全部历史记录?')) {
          try { localStorage.removeItem(HIST_KEY); } catch (err) {}
          renderHistory([]);
          log('已清空历史记录');
        }
        return;
      }
      var id = el.getAttribute('data-id');
      if (el.classList.contains('hist-load')) {
        var h = loadHistory(), item = null;
        h.forEach(function (x) { if (String(x.id) === String(id)) item = x; });
        if (item && item.result) {
          lastResult = item.result;
          renderAll(lastResult);
          document.getElementById('results').style.display = 'block';
          log('已加载历史记录: ' + item.ts + ' (第' + item.qishu + '期 ' + item.period + ')', 'ok');
        }
      } else if (el.classList.contains('hist-del')) {
        var h2 = loadHistory().filter(function (x) { return String(x.id) !== String(id); });
        try { localStorage.setItem(HIST_KEY, JSON.stringify(h2)); } catch (err) {}
        renderHistory(h2);
        log('已删除一条历史记录,剩余 ' + h2.length + ' 条');
      }
    });
    renderHistory(loadHistory());

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
      setTemplates: function (arr) { templates = arr; if (window.TplCore) TplCore.save(templates); },
      getOverrides: function () { return overrides; },
      greenApprove: function (fi) {
        var code = 'GC-TEST' + String(Math.random()).slice(2, 6).toUpperCase();
        overrides[fi] = {green_code: code};
        applyOverrides(lastResult);
        renderAll(lastResult);
        return code;
      },
      setGreenCodes: function (arr) { saveGreenCodes(arr || []); },
      getGreenCodes: loadGreenCodes,
      genGreenCode: genGreenCode,
      rerunAudit: rerunAudit,
      auditBook: function () { return lastResult ? auditBook() : null; },
      poolsBook: function () {
        if (!lastResult) return null;
        var wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, poolSheet(lastResult.validPools['强基']), '强基有效订单池');
        XLSX.utils.book_append_sheet(wb, poolSheet(lastResult.validPools['升学']), '升学有效订单池');
        return wb;
      },
      periodTag: periodTag,
      editEntry: function (fi, phone, name, base) {
        var s = findSub(fi);
        if (!s) throw new Error('entry not found: ' + fi);
        if (!s.raw_row) throw new Error('entry has no raw_row');
        if (phone != null) s.raw_row[6] = phone;
        if (name != null) s.raw_row[5] = name;
        if (base != null) s.raw_row[4] = base;
        rerunAudit();
        return lastResult;
      },
      searchCounts: function (kw) {
        var kws = kw ? String(kw).trim().toLowerCase().split(/\s+/).filter(function (k) { return !!k; }) : [];
        return {
          detail: lastResult ? detailItems(lastResult).filter(function (d) { return kwMatch(kws, detailFields(d)); }).length : 0,
          invalid: lastResult ? (lastResult.invalidList || []).filter(function (s) { return kwMatch(kws, invalidFields(s)); }).length : 0
        };
      },
      prizeBaseRows: function () { return lastResult ? prizeBaseSheet(lastResult) : null; },
      rankRows: function () { return lastResult ? rankSheet(lastResult) : null; },
      getHistory: loadHistory,
      saveHistoryNow: saveHistory,
      updateTimeBtns: updateTimeBtns,
      awardTitle: function (poolName) { return awardTitle(poolName); }
    };
  });
})();
