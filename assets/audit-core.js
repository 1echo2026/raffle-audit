/*!
 * audit-core.js  中奖名单审核核心逻辑(纯 JS,无第三方依赖)
 * 与 Python 版 scripts/audit_winners.py + filter_orders.py 口径一致:
 *  1) 匹配池: 全部渠道且订单金额 >= 999(渠道不再剔除,渠道差异在资格判定阶段归类异常)
 *     有效订单(判池/工号累计/订单池导出): 仅渠道前缀 grow_xcg_zhuanjs_xiaoshou
 *  2) 头像匹配中奖名单 ↔ 问卷;加密手机号(前3位****后4位)匹配订单
 *  3) 一个手机号只保留提交时间最早的一次抽奖
 *  4) 按工号累计有效出单逐单判池:
 *     - 单笔订单金额 >= 2万:该订单直接归巅峰
 *     - 多手机号累计:跨2万的那笔及之后出单归巅峰,未满2万时归进阶
 *  5) 中奖手机号关联订单归哪个池,本次抽奖就必须在哪个池
 *  6) 渠道异常归类: 匹配到的订单全部非销售直推渠道时,渠道前缀为
 *     grow_xcg_zhuanjs_fudao 的按"非销售直推"处理,其它渠道按"非转介绍"处理
 *
 * 列映射层(2026-10 改造): 所有读列都经 resolveColumnMap 解析表头别名得到列下标,
 * 不再使用写死的数字下标;必填列缺失时抛出带字段名的错误并中止。
 */
(function (global) {
  'use strict';

  var CN_NUM = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8};
  var NUM_CN = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八'};
  var formSeq = 0;  // 提交记录的稳定编号(进阶问卷在前),跨重新审核保持一致,供前端绿色通道/修改定位
  var HEADER_SCAN = 10;  // 表头行扫描范围(前 10 行)

  function str(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim();
  }

  function toFloat(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var n = parseFloat(String(v).replace(/,/g, ''));
    return isFinite(n) ? n : 0;
  }

  function parseDt(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    var s = String(v).trim();
    if (!s) return null;
    var m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
    if (m) {
      return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0));
    }
    var d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  function fmtDt(d) {
    if (!d) return '';
    function p(n) { return (n < 10 ? '0' : '') + n; }
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) +
      ' ' + p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
  }

  function fmtMoney(n) {
    return toFloat(n).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
  }

  function cleanPhone(raw) {
    if (raw === null || raw === undefined) return {phone: '', ok: false};
    var m = String(raw).match(/1\d{10}/);
    return m ? {phone: m[0], ok: true} : {phone: String(raw).trim(), ok: false};
  }

  function maskPhone(p) {
    return p.length >= 7 ? p.slice(0, 3) + '****' + p.slice(-4) : p;
  }

  function parsePrizeText(text) {
    if (!text) return {level: null, name: ''};
    var s = String(text).trim();
    // 兼容 "一等奖：冰箱" "一等奖 冰箱" "1等奖·冰箱" "一等奖冰箱" 等格式
    var m = s.match(/^([一二三四五六七八九十\d]+)等奖[：:、\s·]*(.*)$/);
    if (!m) return {level: null, name: s};
    var lv = /^\d+$/.test(m[1]) ? parseInt(m[1], 10) : (CN_NUM[m[1]] || 0);
    var name = (m[2] || '').trim();
    return {level: lv, name: name || s};
  }

  // ==================== 列映射 ====================
  // 表头归一化:去所有空白、去全角/半角括号与标点、转小写
  function normHeader(v) {
    return String(v === null || v === undefined ? '' : v)
      .replace(/[\s\u3000\u00a0]+/g, '')
      .replace(/[（）()【】\[\]〈〉《》「」『』{}〔〕]/g, '')
      .replace(/[:：、,，.。;；/／\\|｜\-—－_~～*＊'‘’"“”!！?？+＋=＝#＃$＄%％&＆@＠^＾`·•●○◎※×√°\u2022]/g, '')
      .toLowerCase();
  }

  // 可扩展别名表(前 10 行内按此识别表头;完全相等优先,其次包含)
  var SCHEMAS = {
    // 订单池导出
    order: [
      {field: 'account', label: '账号ID', required: true, aliases: ['账号ID', '账号']},
      {field: 'cust_name', label: '真实姓名', required: true, aliases: ['真实姓名', '客户姓名', '学员姓名', '姓名']},
      {field: 'masked', label: '加密手机号', required: true, aliases: ['加密手机号', '联系方式', '手机号(掩码)', '手机号（掩码）', '掩码手机号']},
      {field: 'order_id', label: '订单ID', required: true, aliases: ['订单ID', '订单号']},
      {field: 'amount', label: '订单金额', required: true, aliases: ['订单金额', '实付金额', '支付金额', '金额']},
      {field: 'pay_time', label: '支付时间', required: true, aliases: ['支付时间', '付款时间', '下单时间']},
      {field: 'status', label: '订单状态', required: false, aliases: ['订单状态', '状态']},
      {field: 'perf_time', label: '业绩归属时间', required: false, aliases: ['业绩归属时间']},
      {field: 'product_id', label: '商品ID', required: false, aliases: ['商品ID', '商品编号', '产品ID']},
      {field: 'channel', label: '渠道', required: true, aliases: ['订单归属渠道', '渠道', '渠道前缀', '来源渠道']},
      {field: 'sales_name', label: '销售姓名', required: false, aliases: ['订单归属人姓名', '销售姓名', '销售', '顾问姓名']},
      {field: 'gonghao', label: '工号', required: true, aliases: ['订单归属人工号', '工号', '销售工号', '员工工号', '员工编号']},
      {field: 'org', label: '订单归属组织架构', required: false, aliases: ['订单归属组织架构', '组织架构']}
    ],
    // 抽奖助手 表单填写(问卷)
    form: [
      {field: 'nickname', label: '用户昵称', required: false, aliases: ['用户昵称', '昵称']},
      {field: 'avatar', label: '用户头像', required: false, aliases: ['用户头像', '头像', '头像图片']},
      {field: 'prize', label: '所中奖项', required: true, aliases: ['所中奖项', '中奖奖品', '奖品', '奖项']},
      {field: 'submit_time', label: '提交时间', required: true, aliases: ['提交时间', '填写时间']},
      {field: 'base', label: '基地', required: false, aliases: ['您的真实基地', '真实基地', '基地', '所属基地']},
      {field: 'real_name', label: '真实姓名', required: true, aliases: ['您的真实名字', '真实名字', '真实姓名', '姓名']},
      {field: 'phone', label: '出单手机号', required: true, aliases: ['出单用户手机号/ID', '出单用户手机号', '出单手机号', '手机号']}
    ],
    // 中奖名单 奖品发放情况(发放表,当前不参与判定,仅用于识别/兼容)
    dist: [
      {field: 'prize_name', label: '奖项名称', required: false, aliases: ['奖项名称', '奖品名称']},
      {field: 'nickname', label: '中奖者昵称', required: false, aliases: ['中奖者昵称', '用户昵称', '昵称']},
      {field: 'avatar', label: '中奖者头像', required: false, aliases: ['中奖者头像', '商品链接', '链接', '奖品图片']},
      {field: 'redeem', label: '核销情况', required: false, aliases: ['核销情况', '兑换码', '备注', '核销码']}
    ]
  };

  function fieldLabel(def) { return def.label || def.field; }
  function findDef(schema, field) {
    for (var i = 0; i < schema.length; i++) if (schema[i].field === field) return schema[i];
    return null;
  }
  function findDefByLabel(schema, label) {
    var n = normHeader(label);
    for (var i = 0; i < schema.length; i++) {
      if (normHeader(fieldLabel(schema[i])) === n) return schema[i];
    }
    return null;
  }

  /**
   * 解析表头 → 字段列映射
   * @param {Array} headerRow 表头行(一维数组)
   * @param {Array} schema [{field, label, aliases:[], required}]
   * @returns {{map, missing:[], missingRequired:[], warnings:[], fields:[], header:[]}}
   * 规则:归一化后逐字段按别名优先级做「完全相等」优先,其次「包含」;
   *      两个字段命中同一列时按字段定义顺序优先(先到先得)并记 warning。
   */
  function resolveColumnMap(headerRow, schema) {
    schema = schema || [];
    var heads = [];
    (headerRow || []).forEach(function (v) { heads.push(str(v)); });
    var norm = heads.map(normHeader);
    var claimed = {};     // colIndex -> field
    var map = {}, missing = [], warnings = [], fields = [];

    schema.forEach(function (def) {
      var col = -1, byExact = false, conflicts = [];
      function scan(useExact) {
        for (var a = 0; a < def.aliases.length && col < 0; a++) {
          var na = normHeader(def.aliases[a]);
          if (!na) continue;
          for (var c = 0; c < norm.length; c++) {
            if (!norm[c]) continue;
            var hit = useExact ? (norm[c] === na) : (norm[c].indexOf(na) >= 0);
            if (!hit) continue;
            if (claimed[c]) { if (conflicts.indexOf(c) < 0) conflicts.push(c); continue; }
            col = c;
            byExact = !!useExact;
            return;
          }
        }
      }
      scan(true);
      if (col < 0) scan(false);

      if (col >= 0) { claimed[col] = def.field; map[def.field] = col; }
      else missing.push(fieldLabel(def));

      conflicts.forEach(function (c) {
        var owner = findDef(schema, claimed[c]);
        warnings.push('列「' + (heads[c] || ('第' + (c + 1) + '列')) + '」同时匹配「' +
          fieldLabel(def) + '」与「' + (owner ? fieldLabel(owner) : claimed[c]) +
          '」,按字段定义顺序优先归给「' + (owner ? fieldLabel(owner) : claimed[c]) + '」');
      });

      fields.push({
        field: def.field, label: fieldLabel(def), required: !!def.required,
        col: col < 0 ? null : col, headerText: col < 0 ? '' : heads[col],
        matched: col >= 0, byExact: byExact, conflictCols: conflicts
      });
    });

    var missingRequired = fields.filter(function (f) { return f.required && !f.matched; })
      .map(function (f) { return f.label; });
    return {map: map, missing: missing, missingRequired: missingRequired,
            warnings: warnings, fields: fields, header: heads};
  }

  // 在前 10 行内找表头行:命中任一关键(必填)字段别名即视为候选;无必填字段的表(如发放表)要求至少命中 2 个字段
  function findHeaderRow(rows, schema) {
    var best = -1, bestScore = -1, bestInfo = null;
    var limit = Math.min(rows.length || 0, HEADER_SCAN);
    for (var i = 0; i < limit; i++) {
      var info = resolveColumnMap(rows[i] || [], schema);
      var matched = info.fields.filter(function (f) { return f.matched; }).length;
      var reqHit = info.fields.filter(function (f) { return f.required && f.matched; }).length;
      if (reqHit <= 0 && matched < 2) continue;   // 既没命中关键字段、命中又太少 → 不是表头行
      var score = matched * 10 + reqHit;          // 命中越多越像表头行
      if (score > bestScore) { bestScore = score; best = i; bestInfo = info; }
    }
    return best < 0 ? null : {idx: best, info: bestInfo};
  }

  /**
   * 识别并解析一张工作表
   * @param {Array} rows 工作表二维数组
   * @param {String} kind 'order' | 'form' | 'dist'
   * @param {Object} [overrides] {field: 表头文字} 手动指定的列(按表头文字定位)
   */
  function inspectSheet(rows, kind, overrides) {
    var schema = SCHEMAS[kind] || [];
    var found = findHeaderRow(rows, schema);
    if (!found) {
      return {
        kind: kind, headerIdx: null, header: [], title: [], startRow: 1, map: {},
        missing: schema.map(fieldLabel),
        missingRequired: schema.filter(function (d) { return d.required; }).map(fieldLabel),
        warnings: ['未在前 ' + HEADER_SCAN + ' 行内找到表头行(未命中任何关键字段别名)'],
        fields: [], ok: false
      };
    }
    var info = found.info;
    info.kind = kind;
    info.headerIdx = found.idx;
    info.header = info.header || [];
    info.title = found.idx > 0 ? (rows[found.idx - 1] || []) : [];
    info.startRow = found.idx + 1;

    if (overrides) {
      Object.keys(overrides).forEach(function (f) {
        var want = str(overrides[f]);
        if (!want) return;
        var nw = normHeader(want);
        var hit = -1;
        info.header.forEach(function (h, i) { if (hit < 0 && normHeader(h) === nw) hit = i; });
        if (hit < 0) {
          info.warnings.push('手动指定的列「' + want + '」在此表头中找不到,已忽略');
          return;
        }
        info.map[f] = hit;
        info.fields.forEach(function (fd) {
          if (fd.field !== f) return;
          fd.col = hit; fd.headerText = info.header[hit]; fd.matched = true; fd.manual = true;
        });
      });
      info.missing = info.fields.filter(function (fd) { return !fd.matched; }).map(fieldLabel);
      info.missingRequired = info.fields.filter(function (fd) { return fd.required && !fd.matched; }).map(fieldLabel);
    }
    info.ok = info.missingRequired.length === 0;
    return info;
  }

  // 内容识别:命中账号ID/订单ID→订单池;命中用户昵称+提交时间→问卷;命中奖项名称/中奖者昵称→发放表
  function guessKind(rows) {
    function hits(kind) {
      var found = findHeaderRow(rows, SCHEMAS[kind]);
      var h = {};
      if (found) found.info.fields.forEach(function (fd) { if (fd.matched) h[fd.field] = 1; });
      return h;
    }
    var ho = hits('order'), hf = hits('form'), hd = hits('dist');
    if (ho.account || ho.order_id) return 'order';
    if (hf.nickname && hf.submit_time) return 'form';
    if (hd.prize_name || hd.nickname) return 'dist';
    // 弱证据兜底(严格阈值,避免透视/汇总表被误判成订单池或问卷)
    var so = Object.keys(ho).length, sf = Object.keys(hf).length, sd = Object.keys(hd).length;
    if (so >= 4) return 'order';
    if (sf >= 4) return 'form';
    if (sd >= 2) return 'dist';
    return null;
  }

  // ==================== 读取结构(输入均为二维数组) ====================
  function loadDistribution(rows, info) {
    var map = (info && info.map) || {};
    var iPz = map.prize_name != null ? map.prize_name : 0;
    var iNk = map.nickname, iAv = map.avatar, iRd = map.redeem;
    var out = {};
    var curLevel = null;
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i] || [];
      var c0 = str(row[iPz]);
      var m = c0.match(/^([一二三四五六七八九十]+)等奖[：:]/);
      if (m && c0.indexOf('×') >= 0) { curLevel = CN_NUM[m[1]]; continue; }
      if (c0 === '奖项名称' || c0 === '奖品名称') continue;
      if (iNk == null || iAv == null) continue;
      if (row[iNk] && row[iAv] && str(row[iAv]).indexOf('http') === 0) {
        out[str(row[iAv])] = {
          level: curLevel,
          prize_name: c0,
          nickname: str(row[iNk]),
          redeem: iRd == null ? '' : str(row[iRd])
        };
      }
    }
    return out;
  }

  function loadQuestionnaire(rows, pool, info) {
    var map = (info && info.map) || {};
    var start = (info && info.startRow) || 1;
    var iNk = map.nickname, iAv = map.avatar, iPz = map.prize, iSt = map.submit_time,
        iBs = map.base, iNm = map.real_name, iPh = map.phone;
    var out = [];
    for (var i = start; i < rows.length; i++) {
      var row = rows[i] || [];
      var rawPhone = iPh == null ? '' : row[iPh];
      if (!rawPhone) continue;
      var pp = parsePrizeText(iPz == null ? '' : row[iPz]);
      var cp = cleanPhone(rawPhone);
      var rawSubmit = iSt == null ? '' : row[iSt];
      var dt = parseDt(rawSubmit);
      out.push({
        pool: pool,
        nickname: iNk == null ? '' : str(row[iNk]),
        avatar: iAv == null ? '' : str(row[iAv]),
        prize_level: pp.level,
        prize_name: pp.name,
        submit_time_raw: dt ? fmtDt(dt) : str(rawSubmit),
        submit_dt: dt,
        base: iBs == null ? '' : str(row[iBs]),
        real_name: iNm == null ? '' : str(row[iNm]),
        phone: cp.phone,
        phone_format_ok: cp.ok,
        form_i: formSeq++,
        raw_row: row
      });
    }
    return out;
  }

  // ---------- 异常类型分类 ----------
  // 依据无效原因/重复/备注,归类为可套用话术模板的异常类型
  function classifyAbnormal(s) {
    if (s.is_duplicate) return '重复抽奖';
    var r = s.invalid_reasons || '';
    if (r.indexOf('问卷提交时间') === 0 && r.indexOf('不在本期活动时间范围内') > 0) return '提交时间超范围';
    if (r.indexOf('出单时间不符合当前抽奖日期') >= 0) return '出单时间不符合当前抽奖日期';
    if (r.indexOf('手机号格式不正确') >= 0) return '手机号格式错误';
    if (r.indexOf('手机号在订单池中查找不到') >= 0) return '手机号未查到';
    if (r.indexOf('非转介绍') >= 0) return '非转介绍';
    if (r.indexOf('非销售直推') >= 0) return '非销售直推';
    if (r.indexOf('未在有效订单池查找到') >= 0) return '订单未达有效条件';
    if (r.indexOf('问卷头像与中奖名单无法匹配') >= 0) return '头像无法匹配';
    if (r.indexOf('奖池抽奖') >= 0 && (r.indexOf('应属') >= 0 || r.indexOf('应直接') >= 0 || r.indexOf('应在') >= 0)) {
      return s.expected_pool === '巅峰' ? '抽错奖池（应属巅峰）' : '抽错奖池（应属进阶）';
    }
    if (s.abnormal_status) return '含退款/换课订单';
    return '其他异常';
  }

  // ---------- 主流程 ----------
  function audit(input) {
    var params = input.params || {};
    var overrides = input.columnOverrides || {};
    formSeq = 0;
    var PREFIX = params.channelPrefix || 'grow_xcg_zhuanjs_xiaoshou';
    var MIN_AMT = toFloat(params.minAmount != null ? params.minAmount : 999);
    var THRESHOLD = toFloat(params.threshold != null ? params.threshold : 20000);
    // 可配置的时间范围(为空表示不限制)
    var PERIOD = str(params.period);
    var QISHU = str(params.qishu);
    var FORM_START = parseDt(params.formStart), FORM_END = parseDt(params.formEnd);      // 问卷提交时间范围
    var ORDER_START = parseDt(params.orderStart), ORDER_END = parseDt(params.orderEnd);  // 订单支付时间范围
    // 绿色通道码:问卷"出单手机号/ID"栏填写该码即直接通过审核(人工放行,跳过全部匹配与时间检查)
    var GREEN_SET = {};
    (params.greenCodes || []).forEach(function (c) {
      var k = String(c || '').trim().toUpperCase();
      if (k) GREEN_SET[k] = 1;
    });

    function inRange(dt, start, end) {
      if (!dt) return true;
      if (start && dt < start) return false;
      if (end && dt > end) return false;
      return true;
    }

    // 0) 列映射(表头别名解析 + 手动覆盖),必填列缺失直接报错中止
    var colQj = inspectSheet(input.rawOrders.qiangji || [], 'order', overrides.qjRaw);
    var colSx = inspectSheet(input.rawOrders.shengxue || [], 'order', overrides.sxRaw);
    var colJJ = inspectSheet(input.questionnaire['进阶'] || [], 'form', overrides.jjForm);
    var colDF = inspectSheet(input.questionnaire['巅峰'] || [], 'form', overrides.dfForm);
    var colDistJJ = inspectSheet(input.distribution['进阶'] || [], 'dist', null);
    var colDistDF = inspectSheet(input.distribution['巅峰'] || [], 'dist', null);
    var colProblems = [];
    [['强基订单池', colQj], ['升学订单池', colSx], ['进阶奖池问卷', colJJ], ['巅峰奖池问卷', colDF]]
      .forEach(function (x) {
        if (x[1].missingRequired.length) {
          colProblems.push(x[0] + ' 缺少必需列:' + x[1].missingRequired.join('、'));
        }
      });
    if (colProblems.length) {
      throw new Error('列映射失败,无法读取数据 — ' + colProblems.join(';') +
        '。请检查表头名称,或在上传后用「列映射确认」手动指定对应列。');
    }

    // 1) 过滤有效订单(含支付时间范围)
    function filterPool(rows, info) {
      var m = info.map;
      var iAcct = m.account, iAmt = m.amount, iPay = m.pay_time;
      // 匹配池保留全部渠道(渠道差异在资格判定阶段归类为异常)
      var kept = [], rejAmt = 0, rejTime = 0;
      for (var i = info.startRow; i < rows.length; i++) {
        var r = rows[i];
        if (!r || !r[iAcct]) continue;
        var amt = toFloat(r[iAmt]);
        if (amt < MIN_AMT) { rejAmt++; continue; }
        if (!inRange(parseDt(r[iPay]), ORDER_START, ORDER_END)) { rejTime++; continue; }
        kept.push(r);
      }
      return {title: info.title || [], header: info.header || [], map: m, info: info,
              kept: kept, rejAmount: rejAmt, rejTime: rejTime, total: kept.length + rejAmt + rejTime};
    }

    var fQj = filterPool(input.rawOrders.qiangji || [], colQj);
    var fSx = filterPool(input.rawOrders.shengxue || [], colSx);

    function toRec(r, source, m) {
      return {
        source: source,
        account: str(r[m.account]),
        cust_name: str(r[m.cust_name]),
        masked: str(r[m.masked]),
        order_id: str(r[m.order_id]),
        amount: toFloat(r[m.amount]),
        pay_time: str(r[m.pay_time]),
        dt: parseDt(r[m.pay_time]),
        // 可选列缺失时回退:业绩归属时间缺失 → 用支付时间
        perf_time: m.perf_time == null ? str(r[m.pay_time]) : str(r[m.perf_time]),
        perf_dt: m.perf_time == null ? parseDt(r[m.pay_time]) : parseDt(r[m.perf_time]),
        status: m.status == null ? '' : str(r[m.status]),
        product_id: m.product_id == null ? '' : str(r[m.product_id]),
        sales_name: m.sales_name == null ? '' : str(r[m.sales_name]),
        gonghao: str(r[m.gonghao]),
        channel: str(r[m.channel]),
        org: m.org == null ? '' : str(r[m.org])
      };
    }

    // 匹配池: 全部渠道(≥999),用于匹配中奖名单;渠道差异在资格判定阶段归类为异常
    var allRows = [];
    fQj.kept.forEach(function (r) { allRows.push(toRec(r, '强基', fQj.map)); });
    fSx.kept.forEach(function (r) { allRows.push(toRec(r, '升学', fSx.map)); });
    // 有效订单: 仅销售直推渠道(grow_xcg_zhuanjs_xiaoshou),用于判池与工号累计
    function isXiaoshou(r) { return r.channel.indexOf(PREFIX) === 0; }
    var validRows = allRows.filter(isXiaoshou);

    // (订单ID,商品ID) 去重,防两池重复导出
    var seenKey = {}, deduped = [];
    allRows.forEach(function (r) {
      var k = r.order_id + '|' + r.product_id;
      if (seenKey[k]) return;
      seenKey[k] = 1;
      deduped.push(r);
    });
    allRows = deduped;
    validRows = allRows.filter(isXiaoshou);

    // 2) 中奖名单 + 问卷
    var winners = {
      '进阶': loadDistribution(input.distribution['进阶'] || [], colDistJJ),
      '巅峰': loadDistribution(input.distribution['巅峰'] || [], colDistDF)
    };
    var submissions = []
      .concat(loadQuestionnaire(input.questionnaire['进阶'] || [], '进阶', colJJ))
      .concat(loadQuestionnaire(input.questionnaire['巅峰'] || [], '巅峰', colDF));

    // 3) 工号时间线 + 逐单判池
    var ghRowCnt = {}, ghNameVotes = {}, orderAgg = {};
    function inc(obj, k, n) { obj[k] = (obj[k] || 0) + (n || 1); }
    validRows.forEach(function (r) {
      if (!r.gonghao) return;
      inc(ghRowCnt, r.gonghao);
      ghNameVotes[r.gonghao] = ghNameVotes[r.gonghao] || {};
      inc(ghNameVotes[r.gonghao], r.sales_name);
      if (orderAgg[r.order_id]) {
        orderAgg[r.order_id].amount += r.amount;
      } else {
        orderAgg[r.order_id] = {
          amount: r.amount, dt: r.dt, masked: r.masked, gonghao: r.gonghao,
          pool: '进阶', cum_after: 0, direct_peak: false
        };
      }
    });

    var ghOrders = {};
    Object.keys(orderAgg).forEach(function (oid) {
      var gh = orderAgg[oid].gonghao;
      (ghOrders[gh] = ghOrders[gh] || {})[oid] = [orderAgg[oid].dt, orderAgg[oid].amount];
    });

    var ghStats = {};
    Object.keys(ghOrders).forEach(function (gh) {
      var entries = Object.keys(ghOrders[gh]).map(function (oid) {
        return [oid, ghOrders[gh][oid][0], ghOrders[gh][oid][1]];
      });
      entries.sort(function (a, b) { return (a[1] ? a[1].getTime() : 8.64e15) - (b[1] ? b[1].getTime() : 8.64e15); });
      var total = 0, maxSingle = 0, cum = 0, peakSince = null, peakOrders = 0;
      entries.forEach(function (e) {
        var oid = e[0], dt = e[1], amt = e[2];
        total += amt;
        if (amt > maxSingle) maxSingle = amt;
        cum += amt;
        var isPeak = amt >= THRESHOLD || cum >= THRESHOLD;
        var o = orderAgg[oid];
        o.pool = isPeak ? '巅峰' : '进阶';
        o.cum_after = Math.round(cum * 100) / 100;
        o.direct_peak = amt >= THRESHOLD;
        if (isPeak && !peakSince) peakSince = dt;
        if (isPeak) peakOrders++;
      });
      var votes = ghNameVotes[gh], bestName = '', bestN = -1;
      Object.keys(votes).forEach(function (n) { if (votes[n] > bestN) { bestN = votes[n]; bestName = n; } });
      ghStats[gh] = {
        sales_name: bestName,
        total: Math.round(total * 100) / 100,
        max_single: Math.round(maxSingle * 100) / 100,
        row_count: ghRowCnt[gh] || 0,
        peak_since: peakSince,
        peak_orders: peakOrders
      };
    });

    // 4) 掩码 -> 匹配池订单行(全部渠道,≥999)
    var validIndex = {};
    allRows.forEach(function (r) {
      if (!r.masked) return;
      (validIndex[r.masked] = validIndex[r.masked] || []).push(r);
    });

    // 5) 匹配 + 期望奖池
    submissions.forEach(function (s) {
      // 未提供"发放情况"(中奖名单)表时,跳过头像匹配校验(问卷表单已含基地/姓名/出单手机号,不再依赖发放表)
      var win = winners[s.pool][s.avatar];
      s.avatar_matched = Object.keys(winners[s.pool]).length ? !!win : true;
      s.win_nickname = win ? win.nickname : '';
      s.win_redeem = win ? win.redeem : '';
      s.match = null;
      s.total_amount = 0; s.order_count = 0; s.max_single_order = 0;
      s.convert_amount = 0;   // 转化金额:该出单手机号匹配到的订单金额合计
      s.peak_since = null; s.expected_pool = ''; s.peak_reason = '';
      s.account = ''; s.sales_name = ''; s.gonghao = ''; s.abnormal_status = ''; s.channel = '';
      if (!s.phone_format_ok) return;
      var rows = validIndex[maskPhone(s.phone)] || [];
      var acctSet = {};
      rows.forEach(function (r) { acctSet[r.account] = 1; });
      if (!rows.length || Object.keys(acctSet).length > 1) return;
      s.match = 'ok';
      // 渠道归类: 匹配到的订单全部非销售直推渠道时,记为渠道异常(非销售直推/非转介绍)
      var xsRows = rows.filter(isXiaoshou);
      if (!xsRows.length) {
        s.channel_bad = {channel: rows[rows.length - 1].channel};
        s.channel = s.channel_bad.channel;
        var tmax = null;
        rows.forEach(function (r) { if (r.dt && (!tmax || r.dt > tmax)) tmax = r.dt; });
        s.order_time = tmax ? fmtDt(tmax) : '';
        return;
      }
      rows = xsRows;
      s.account = rows[0].account;
      var cSum = 0;
      rows.forEach(function (r) { cSum += r.amount; });
      s.convert_amount = Math.round(cSum * 100) / 100;
      // 业绩归属渠道(匹配到的有效订单渠道,去重合并)
      var chSet = {};
      rows.forEach(function (r) { if (r.channel) chSet[r.channel] = 1; });
      s.channel = Object.keys(chSet).join('、');
      var nv = {}, gv = {};
      rows.forEach(function (r) { inc(nv, r.sales_name); inc(gv, r.gonghao); });
      s.sales_name = Object.keys(nv).sort(function (a, b) { return nv[b] - nv[a]; })[0] || '';
      s.gonghao = Object.keys(gv).sort(function (a, b) { return gv[b] - gv[a]; })[0] || '';
      var st = ghStats[s.gonghao];
      if (st) {
        s.total_amount = st.total;
        s.order_count = st.row_count;
        s.max_single_order = st.max_single;
        s.peak_since = st.peak_since;
        s.peak_orders = st.peak_orders || 0;
      }
      s.order_time = '';
      rows.forEach(function (r) {
        // 业绩归属时间取订单池"业绩归属时间"列;无该列时回退支付时间
        var t = r.perf_dt || r.dt;
        if (t && (!s.order_time || t > s.order_time)) s.order_time = t;
      });
      s.order_time = s.order_time ? fmtDt(s.order_time) : '';
      var oidMap = {}, relInfos = [];
      rows.forEach(function (r) { if (r.order_id) oidMap[r.order_id] = 1; });
      Object.keys(oidMap).forEach(function (oid) { if (orderAgg[oid]) relInfos.push(orderAgg[oid]); });
      if (relInfos.length) {
        var hasPeak = relInfos.some(function (o) { return o.pool === '巅峰'; });
        s.expected_pool = hasPeak ? '巅峰' : '进阶';
        if (hasPeak) {
          var direct = relInfos.filter(function (o) { return o.direct_peak; });
          if (direct.length) {
            var o = direct.sort(function (a, b) { return b.amount - a.amount; })[0];
            s.peak_reason = '该出单手机号存在单笔订单金额' + fmtMoney(o.amount) + '元≥2万元' +
              '（支付于' + fmtDt(o.dt) + '），应直接在巅峰奖池抽奖';
          } else {
            var o2 = relInfos.filter(function (o) { return o.pool === '巅峰'; })[0];
            s.peak_reason = '该手机号订单支付于' + fmtDt(o2.dt) + '时，工号累计有效出单金额已达' +
              fmtMoney(o2.cum_after) + '元（≥2万），累计达2万之后的出单应在巅峰奖池抽奖';
          }
        } else {
          var last = relInfos.sort(function (a, b) { return b.cum_after - a.cum_after; })[0];
          s.peak_reason = '该手机号订单支付于' + fmtDt(last.dt) + '时，工号累计有效出单金额仅' +
            fmtMoney(last.cum_after) + '元（<2万），应在进阶奖池抽奖';
        }
      }
      var ab = {};
      rows.forEach(function (r) { if (r.status === '完全退款' || r.status === '换课原订单') ab[r.status] = 1; });
      s.abnormal_status = Object.keys(ab).join('、');
    });

    // 6) 一手机号一单去重(保留提交时间最早的一条)
    var groups = {};
    submissions.forEach(function (s, i) { if (s.phone_format_ok) (groups[s.phone] = groups[s.phone] || []).push(i); });
    var duplicates = [];
    Object.keys(groups).forEach(function (phone) {
      var idxs = groups[phone];
      if (idxs.length <= 1) return;
      idxs.sort(function (a, b) {
        var da = submissions[a].submit_dt, db = submissions[b].submit_dt;
        return (da ? da.getTime() : 8.64e15) - (db ? db.getTime() : 8.64e15);
      });
      // 同一手机号在最早提交时刻有多条记录时(例如进阶全量表单与巅峰表单包含
      // 完全相同的提交),优先保留巅峰奖池那条 —— 巅峰表单是该奖池的精确来源,
      // 避免巅峰提交被进阶副本按数组顺序误判为重复而整池丢失。
      var earliest = idxs[0];
      var firstT = submissions[earliest].submit_dt ? submissions[earliest].submit_dt.getTime() : 8.64e15;
      idxs.forEach(function (i) {
        var t = submissions[i].submit_dt ? submissions[i].submit_dt.getTime() : 8.64e15;
        if (t === firstT && submissions[i].pool === '巅峰' && submissions[earliest].pool !== '巅峰') earliest = i;
      });
      // 同一手机号在进阶/巅峰都抽过,且其订单已累计达2万(应属巅峰):
      // 只保留巅峰奖池奖品,不再保留进阶奖池奖品(先抽进阶、后续累计达标的情形)
      if (submissions[earliest].expected_pool === '巅峰') {
        var peakIdx = -1;
        idxs.forEach(function (i) { if (submissions[i].pool === '巅峰') peakIdx = i; });
        if (peakIdx >= 0) earliest = peakIdx;
      }
      idxs.forEach(function (i) {
        if (i === earliest) return;
        submissions[i].is_duplicate = true;
        duplicates.push(submissions[i]);
      });
    });
    submissions.forEach(function (s) { s.is_duplicate = !!s.is_duplicate; });

    // 7) 资格判定
    var validList = [], invalidList = [];
    var rawAll = [];
    (input.rawOrders.qiangji || []).forEach(function (r) { rawAll.push([r, fQj.map]); });
    (input.rawOrders.shengxue || []).forEach(function (r) { rawAll.push([r, fSx.map]); });

    submissions.forEach(function (s) {
      var reasons = [], remark = '';
      // 绿色通道:问卷"出单手机号/ID"栏填写绿色通道码 → 直接通过(跳过全部检查)
      var gcKey = String(s.phone || '').trim().toUpperCase();
      if (gcKey && GREEN_SET[gcKey]) {
        s.green_code = gcKey;
        s.green_note = '绿色通道码放行';
        s.invalid_reasons = '';
        s.abnormal_type = '';
        s.remark = '';
        if (!s.is_duplicate) validList.push(s);
        return;
      }
      if (!inRange(s.submit_dt, FORM_START, FORM_END)) {
        reasons.push('问卷提交时间' + (s.submit_time_raw || '未知') + '不在本期活动时间范围内');
      }
      if (!s.phone_format_ok) {
        reasons.push('出单手机号格式不正确，在订单池中查找不到，疑似输入错误');
      } else if (s.channel_bad) {
        // 渠道前缀码:取订单池渠道值去掉 &&null&&null 脏后缀,展示给人工核对
        var cbCh = String(s.channel_bad.channel || '').split('&&')[0];
        s.channel_code = cbCh;
        if (cbCh.indexOf('grow_xcg_zhuanjs_fudao') === 0) {
          reasons.push('未在有效订单池查找到（订单归属渠道非销售直推渠道：' + cbCh + '，属于辅导leads接单）');
        } else {
          reasons.push('未在有效订单池查找到（订单归属渠道非转介绍渠道：' + cbCh + '，不属于转介绍出单）');
        }
      } else if (!s.match) {
        var mp = maskPhone(s.phone);
        var rawRows = [];
        rawAll.forEach(function (x) {
          var m = x[1];
          if (m.masked != null && str(x[0][m.masked]) === mp) rawRows.push([x[0], m]);
        });
        if (rawRows.length) {
          // 业绩归属时间(订单支付时间)是否在本期活动时间范围内;未配置时间范围时不限制
          var inWin = rawRows.some(function (x) { return inRange(parseDt(x[0][x[1].pay_time]), ORDER_START, ORDER_END); });
          if (!inWin && (ORDER_START || ORDER_END)) {
            var t0 = rawRows.map(function (x) { return parseDt(x[0][x[1].pay_time]); })
                            .filter(function (d) { return !!d; }).sort(function (a, b) { return a - b; })[0];
            s.order_time = t0 ? fmtDt(t0) : '';
            reasons.push('出单时间不符合当前抽奖日期（订单池匹配到的业绩归属时间' +
              (s.order_time || str(rawRows[0][0][rawRows[0][1].pay_time]) || '未知') + '不在本期活动时间范围内）');
          } else {
            // 手机号在原始订单池能找到但未进入匹配池(未达≥999):要么金额不足,要么匹配到多个账号
            var hasQualified = rawRows.some(function (x) {
              return toFloat(x[0][x[1].amount]) >= MIN_AMT &&
                     inRange(parseDt(x[0][x[1].pay_time]), ORDER_START, ORDER_END);
            });
            if (hasQualified) {
              reasons.push('未在有效订单池查找到（手机号在订单池匹配到多个账号的达标订单，无法唯一归属）');
            } else {
              reasons.push('未在有效订单池查找到（订单金额不足' + Math.round(MIN_AMT) + '元，未达活动有效条件）');
            }
          }
        } else {
          reasons.push('手机号在订单池中查找不到，疑似输入错误');
        }
      }
      if (!s.avatar_matched) reasons.push('问卷头像与中奖名单无法匹配');
      if (s.match && !s.channel_bad && s.expected_pool && s.pool !== s.expected_pool) {
        reasons.push(s.peak_reason + '，但本次在' + s.pool + '奖池抽奖');
      }
      if (s.abnormal_status) remark += '含' + s.abnormal_status + '订单，请核查';
      s.invalid_reasons = reasons.join('；');
      s.remark = remark;
      s.abnormal_type = classifyAbnormal(s);
      if (s.is_duplicate) return;
      (reasons.length ? invalidList : validList).push(s);
    });

    // 8) 姓名↔工号对照
    var nameGh = {}, ghPhones = {};
    validRows.forEach(function (r) {
      if (r.sales_name && r.gonghao) (nameGh[r.sales_name] = nameGh[r.sales_name] || {})[r.gonghao] = 1;
    });
    validList.concat(invalidList).forEach(function (s) {
      if (s.match === 'ok' && s.gonghao) (ghPhones[s.gonghao] = ghPhones[s.gonghao] || {})[s.phone] = 1;
    });
    var dupNames = {};
    Object.keys(nameGh).forEach(function (n) { if (Object.keys(nameGh[n]).length > 1) dupNames[n] = 1; });

    var nameGhRows = [];
    Object.keys(nameGh).sort().forEach(function (name) {
      Object.keys(nameGh[name]).sort().forEach(function (gh) {
        var st = ghStats[gh] || {};
        nameGhRows.push({
          name: name,
          gonghao: gh,
          phones: Object.keys(ghPhones[gh] || {}).sort(),
          gh_count: Object.keys(nameGh[name]).length,
          total: st.total || 0,
          peak_since: st.peak_since || null,
          is_dup: !!dupNames[name]
        });
      });
    });

    // 9) 图片格式分组(仅有效、去重后;保持问卷顺序)
    function tierGroups(pool, maxLevel) {
      var buckets = {};
      for (var lv = 1; lv <= maxLevel; lv++) buckets[lv] = [];
      var ungraded = [];
      validList.forEach(function (s) {
        if (s.pool !== pool) return;
        if (s.prize_level && buckets[s.prize_level]) buckets[s.prize_level].push(s);
        else ungraded.push(s);
      });
      function toRow(s) {
        var b = s.base;
        if (b.lastIndexOf('基地') === b.length - 2) b = b.slice(0, -2);
        return {base: b, name: s.real_name, gonghao: s.gonghao, phone: s.phone,
                nickname: s.nickname, redeem: s.win_redeem,
                order_time: s.order_time || '', total_amount: s.total_amount || 0,
                green_code: s.green_code || '', green_pass_manual: !!s.green_pass_manual,
                form_i: s.form_i, account: s.account || '',
                convert_amount: s.convert_amount || 0, order_count: s.order_count || 0};
      }
      var out = [];
      for (var l = 1; l <= maxLevel; l++) {
        if (!buckets[l] || !buckets[l].length) continue;
        out.push({
          level: l,
          level_text: NUM_CN[l] + '等奖',
          prize: buckets[l][0].prize_name,
          rows: buckets[l].map(toRow)
        });
      }
      // 问卷奖品文本无法解析出等级的行:不丢弃,按奖品名称聚合展示
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

    // 10) 异常沟通名单(无效+重复+含异常订单备注;供套用话术模板)
    var communication = [];
    function commBase(s) {
      return {
        pool: s.pool, nickname: s.nickname, real_name: s.real_name, base: s.base,
        phone: s.phone, prize_name: s.prize_name, prize_level: s.prize_level,
        gonghao: s.gonghao || '', channel_code: s.channel_code || '',
        submit_time: s.submit_time_raw, abnormal_type: s.abnormal_type,
        reason: s.invalid_reasons || '', remark: s.remark, is_duplicate: !!s.is_duplicate,
        expected_pool: s.expected_pool, period: PERIOD, qishu: QISHU,
        order_time: s.order_time || '', peak_orders: s.peak_orders || 0,
        audit_result: s.audit_result || '', form_i: s.form_i
      };
    }
    invalidList.forEach(function (s) { communication.push(commBase(s)); });
    duplicates.forEach(function (s) { communication.push(commBase(s)); });
    validList.forEach(function (s) { if (s.abnormal_status) communication.push(commBase(s)); });

    // 有效订单池导出仅保留销售直推渠道行
    function xiaoKept(rows, info) {
      var iCh = info.map.channel;
      return rows.filter(function (r) { return str(r[iCh]).indexOf(PREFIX) === 0; });
    }
    var keptQj = xiaoKept(fQj.kept, colQj), keptSx = xiaoKept(fSx.kept, colSx);

    return {
      params: {channelPrefix: PREFIX, minAmount: MIN_AMT, threshold: THRESHOLD,
               period: PERIOD, qishu: QISHU, formStart: FORM_START, formEnd: FORM_END,
               orderStart: ORDER_START, orderEnd: ORDER_END},
      columns: {qj: colQj, sx: colSx, jjForm: colJJ, dfForm: colDF},
      validPools: {
        '强基': {title: fQj.title, header: fQj.header, map: fQj.map, kept: keptQj,
                 stats: {total: fQj.total, rejAmount: fQj.rejAmount, rejTime: fQj.rejTime, kept: keptQj.length}},
        '升学': {title: fSx.title, header: fSx.header, map: fSx.map, kept: keptSx,
                 stats: {total: fSx.total, rejAmount: fSx.rejAmount, rejTime: fSx.rejTime, kept: keptSx.length}}
      },
      stats: {
        submitted: submissions.length,
        valid: validList.length,
        validJJ: validList.filter(function (s) { return s.pool === '进阶'; }).length,
        validDF: validList.filter(function (s) { return s.pool === '巅峰'; }).length,
        invalid: invalidList.length,
        duplicates: duplicates.length,
        ghCount: Object.keys(ghStats).length,
        ghPeak: Object.keys(ghStats).filter(function (g) { return ghStats[g].peak_since; }).length
      },
      validList: validList,
      invalidList: invalidList,
      duplicates: duplicates,
      communication: communication,
      nameGhRows: nameGhRows,
      ghStats: ghStats,
      groups: {'巅峰': tierGroups('巅峰', 4), '进阶': tierGroups('进阶', 8)}
    };
  }

  var api = {
    audit: audit,
    classifyAbnormal: classifyAbnormal,
    resolveColumnMap: resolveColumnMap,
    inspectSheet: inspectSheet,
    guessKind: guessKind,
    SCHEMAS: SCHEMAS,
    normHeader: normHeader,
    parseDt: parseDt,
    fmtDt: fmtDt,
    fmtMoney: fmtMoney,
    cleanPhone: cleanPhone,
    maskPhone: maskPhone,
    parsePrizeText: parsePrizeText
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  global.AuditCore = api;
})(typeof window !== 'undefined' ? window : this);
