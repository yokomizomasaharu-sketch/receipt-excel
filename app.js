/* ============================================================
   領収書スキャン → Excel ／ 青色申告・税務チェック
   ・写真/カメラ → ブラウザ内OCR(Tesseract.js) → 項目自動抽出
   ・一覧管理(localStorage・年度別) → 本物の .xlsx 出力(SheetJS)
   ・タブ: 領収書(経費帳) / 旅行費用 / 税務診断 / e-Taxガイド
   ※ 個人情報は一切コードに持たない。データは端末内のみ。
   ============================================================ */
(function () {
  'use strict';

  /* ---------- 定数 ---------- */
  var STORE_KEY = 'pickle.receipt2excel.v3';
  var LEGACY_KEY = 'pickle.receipt2excel.v2';
  var DEFAULT_RATE = 0.104;

  /* e-Tax 青色申告決算書1ページ目「経費」の科目（実際の並び） */
  var FORM_CATEGORIES = ['租税公課', '荷造運賃', '水道光熱費', '旅費交通費', '通信費', '広告宣伝費', '接待交際費', '損害保険料', '修繕費', '消耗品費', '減価償却費', '福利厚生費', '給料賃金', '外注工賃', '利子割引料', '地代家賃', '貸倒金', '専従者給与', '雑費'];
  /* アプリ独自の科目 → 決算書では別の欄にまとめる */
  var FORM_MAP = { '会議費': '雑費', '新聞図書費': '雑費', '車両費': '旅費交通費', '支払手数料': '雑費' };

  var TAX_CATEGORIES = [
    '消耗品費', '旅費交通費', '通信費', '会議費', '接待交際費', '水道光熱費', '新聞図書費', '支払手数料',
    '外注工賃', '地代家賃', '車両費', '雑費',
    '租税公課', '荷造運賃', '広告宣伝費', '損害保険料', '修繕費', '減価償却費', '福利厚生費', '給料賃金', '利子割引料', '貸倒金', '専従者給与'
  ];
  var TRIP_CATEGORIES = ['航空券', 'ホテル', '交通費', '食費・買い物', 'その他'];
  var TAX_TYPES = ['', '10%', '8%軽減', '非課税', '対象外'];

  var TRIP_PRESETS = [
    { date: '2026-11-14', vendor: 'Trip.com（WE514＋TW307）', category: '航空券', currency: 'JPY', amount: 36760, taxType: '', memo: '航空券（往復）／パラタ航空 WE514 関空17:35→仁川19:35、トリニティ航空 TW307 仁川20:20→関空22:10' },
    { date: '2026-11-14', vendor: '東横INN仁川富平（免税返金後）', category: 'ホテル', currency: 'KRW', amount: 64450, taxType: '', memo: '1泊・1名・朝食付き／67,450₩ − 免税返金3,000₩' },
    { date: '2026-11-14', vendor: '南海電鉄 ⇔ 関西空港（往復）', category: '交通費', currency: 'JPY', amount: 2000, taxType: '', memo: '往復（推定）' },
    { date: '2026-11-14', vendor: 'AREX直通（往路・Klook）', category: '交通費', currency: 'JPY', amount: 1359, taxType: '', memo: '仁川空港T1 → ソウル駅' },
    { date: '2026-11-15', vendor: 'AREX直通（復路・Klook）', category: '交通費', currency: 'JPY', amount: 1359, taxType: '', memo: 'ソウル駅 → 仁川空港T1' },
    { date: '2026-11-14', vendor: '気候同行カード 2日券', category: '交通費', currency: 'KRW', amount: 8000, taxType: '', memo: 'ソウル市内の地下鉄・バス乗り放題（実物カード代3,000₩は含まず）' },
    { date: '2026-11-14', vendor: '地下鉄1号線 ソウル駅⇔富平', category: '交通費', currency: 'KRW', amount: 3600, taxType: '', memo: '往復 1,800₩×2（富平は気候同行カード対象エリア外）' }
  ];

  /* ---------- 状態 ---------- */
  var state = {
    mode: 'tax',
    year: new Date().getFullYear(),
    rate: DEFAULT_RATE,
    trip: [],
    ledgers: {},     /* { '2026': [entry] } */
    sales: {},       /* { '2026': 41683 } */
    blue: {},        /* { '2026': 650000 } */
    slips: {},       /* { '2026': [slip] } */
    deductions: {},  /* { '2026': 123456 } ※未設定なら源泉徴収票③から自動 */
    checks: {},      /* { '2026': { prep: true } } */
    presetApplied: false,
    drafts: []
  };
  var uidSeq = 0;
  var activeDraftId = null;
  var lastImport = null;

  function uid() { uidSeq += 1; return 'r' + Date.now().toString(36) + uidSeq.toString(36); }
  function $(id) { return document.getElementById(id); }

  /* ============================================================
     ストレージ
     ============================================================ */
  function yearOf(s) {
    var m = String(s == null ? '' : s).match(/(\d{4})/);
    return m ? +m[1] : 0;
  }

  function load() {
    state.year = new Date().getFullYear();
    var raw = null;
    try { raw = localStorage.getItem(STORE_KEY); } catch (e) { raw = null; }

    if (raw) {
      try {
        var d = JSON.parse(raw);
        if (d && typeof d === 'object') {
          if (typeof d.rate === 'number' && d.rate > 0) state.rate = d.rate;
          if (Array.isArray(d.trip)) state.trip = d.trip;
          if (d.ledgers && typeof d.ledgers === 'object') state.ledgers = d.ledgers;
          if (d.sales && typeof d.sales === 'object') state.sales = d.sales;
          if (d.blue && typeof d.blue === 'object') state.blue = d.blue;
          if (d.slips && typeof d.slips === 'object') state.slips = d.slips;
          if (d.deductions && typeof d.deductions === 'object') state.deductions = d.deductions;
          if (d.checks && typeof d.checks === 'object') state.checks = d.checks;
          state.presetApplied = !!d.presetApplied;
          if (d.mode === 'tax' || d.mode === 'trip' || d.mode === 'risk' || d.mode === 'guide') state.mode = d.mode;
          if (typeof d.year === 'number' && d.year >= 2000 && d.year <= 2100) state.year = d.year;
        }
        return;
      } catch (e) { /* 壊れていたら下の移行へ */ }
    }

    /* 旧バージョン(v2)からの移行：経費を日付の年ごとに振り分ける */
    try {
      var old = localStorage.getItem(LEGACY_KEY);
      if (!old) return;
      var d2 = JSON.parse(old);
      if (Array.isArray(d2.trip)) state.trip = d2.trip;
      if (typeof d2.rate === 'number' && d2.rate > 0) state.rate = d2.rate;
      state.presetApplied = !!d2.presetApplied;
      (d2.tax || []).forEach(function (e) {
        var y = yearOf(e.date) || state.year;
        var k = String(y);
        if (!state.ledgers[k]) state.ledgers[k] = [];
        state.ledgers[k].push(e);
      });
      persist();
    } catch (e3) { /* noop */ }
  }

  function persist() {
    var payload = {
      v: 3, mode: state.mode, year: state.year, rate: state.rate, trip: state.trip,
      ledgers: state.ledgers, sales: state.sales, blue: state.blue,
      slips: state.slips, deductions: state.deductions, checks: state.checks,
      presetApplied: state.presetApplied
    };
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify(payload));
      return true;
    } catch (e) {
      try {
        var strip = function (arr) { return arr.map(function (x) { var c = Object.assign({}, x); delete c.thumb; return c; }); };
        var led = {};
        Object.keys(state.ledgers).forEach(function (k) { led[k] = strip(state.ledgers[k]); });
        payload.trip = strip(state.trip);
        payload.ledgers = led;
        localStorage.setItem(STORE_KEY, JSON.stringify(payload));
        return true;
      } catch (e2) {
        toast('端末の保存領域がいっぱいのため保存できませんでした。不要な領収書を削除してください。', true);
        return false;
      }
    }
  }

  /* ---------- 年度ごとのアクセサ ---------- */
  function ykey(y) { return String(y || state.year); }
  function taxList(y) {
    var k = ykey(y);
    if (!Array.isArray(state.ledgers[k])) state.ledgers[k] = [];
    return state.ledgers[k];
  }
  function salesOf(y) { return Number(state.sales[ykey(y)] || 0); }
  function blueOf(y) {
    var v = state.blue[ykey(y)];
    return (v === null || v === undefined || v === '') ? 650000 : Number(v);
  }
  function slipsOf(y) {
    var k = ykey(y);
    if (!Array.isArray(state.slips[k])) state.slips[k] = [];
    return state.slips[k];
  }
  function checksOf(y) {
    var k = ykey(y);
    if (!state.checks[k] || typeof state.checks[k] !== 'object') state.checks[k] = {};
    return state.checks[k];
  }

  /* ---------- ユーティリティ ---------- */
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (m) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
    });
  }
  function yen(n) { return '¥' + Math.round(Number(n) || 0).toLocaleString('ja-JP'); }
  function yenSigned(n) {
    var v = Math.round(Number(n) || 0);
    return (v < 0 ? '▲¥' : '¥') + Math.abs(v).toLocaleString('ja-JP');
  }
  function won(n) { return Number(n || 0).toLocaleString('ja-JP') + '₩'; }
  function num(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var s = String(v == null ? '' : v).replace(/[^0-9.\-]/g, '');
    var f = parseFloat(s);
    return isNaN(f) ? 0 : f;
  }
  function yenOf(item) {
    var a = num(item.amount);
    return item.currency === 'KRW' ? Math.round(a * state.rate) : Math.round(a);
  }
  function todayStr() {
    var d = new Date();
    var p = function (n) { return (n < 10 ? '0' : '') + n; };
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
  }
  function stamp() {
    var d = new Date();
    var p = function (n) { return (n < 10 ? '0' : '') + n; };
    return '' + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate());
  }
  function toast(msg, isErr) {
    var t = $('toast');
    t.textContent = msg;
    t.className = isErr ? 'show err' : 'show';
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.className = ''; }, 3800);
  }
  function withTimeout(promise, ms, message) {
    return new Promise(function (resolve, reject) {
      var done = false;
      var timer = setTimeout(function () {
        if (done) return;
        done = true;
        reject(new Error(message || 'タイムアウトしました'));
      }, ms);
      Promise.resolve(promise).then(function (v) {
        if (done) return;
        done = true; clearTimeout(timer); resolve(v);
      }, function (e) {
        if (done) return;
        done = true; clearTimeout(timer); reject(e);
      });
    });
  }
  function currentList() { return state.mode === 'trip' ? state.trip : taxList(); }
  function categoriesFor(mode) { return mode === 'trip' ? TRIP_CATEGORIES : TAX_CATEGORIES; }
  function copyText(text, okMsg) {
    var done = function () { toast(okMsg || 'コピーしました'); };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(done, function () { fallbackCopy(text, done); });
    } else {
      fallbackCopy(text, done);
    }
  }
  function fallbackCopy(text, cb) {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    document.body.removeChild(ta);
    if (ok) { if (cb) cb(); } else { toast('コピーできませんでした。長押しで選択してコピーしてください。', true); }
  }
  function signal(msg) {
    try {
      if (window.pickle && pickle.widget && typeof pickle.widget.signal === 'function') pickle.widget.signal(msg);
    } catch (e) { /* noop */ }
  }

  /* ---------- 年度ユーティリティ ---------- */
  function reiwa(y) { return y - 2018; }
  function yearLabel(y) { return y + '年（令和' + reiwa(y) + '年）'; }
  function availableYears() {
    var set = {};
    var now = new Date().getFullYear();
    var lo = Math.min(2024, now - 1);
    var hi = Math.max(2030, now + 6);
    for (var y = lo; y <= hi; y += 1) set[y] = 1;
    [state.ledgers, state.sales, state.blue, state.slips, state.deductions, state.checks].forEach(function (m) {
      Object.keys(m || {}).forEach(function (k) { if (/^\d{4}$/.test(k)) set[+k] = 1; });
    });
    set[state.year] = 1;
    return Object.keys(set).map(Number).sort(function (a, b) { return a - b; });
  }
  function deadlineOf(y) {
    var d1 = new Date(y + 1, 1, 16);
    var d2 = new Date(y + 1, 2, 15);
    var w = ['日', '月', '火', '水', '木', '金', '土'];
    return {
      start: fmtJa(y + 1, 2, 16) + '（' + w[d1.getDay()] + '）',
      end: fmtJa(y + 1, 3, 15) + '（' + w[d2.getDay()] + '）'
    };
  }
  function fmtJa(y, m, d) { return y + '年' + m + '月' + d + '日'; }

  /* ============================================================
     画面描画（共通）
     ============================================================ */
  function renderAll() {
    renderModeChrome();
    renderYearBar();
    renderKPIs();
    renderCategories();
    renderEntries();
    renderListFooter();
    renderRiskInputs();
    renderRiskResults();
    renderGuide();
  }

  function renderModeChrome() {
    var m = state.mode;
    var trip = m === 'trip';
    var titles = {
      tax: ['🧾 領収書スキャン', '撮影するだけで自動で表に整理・Excel出力'],
      trip: ['🇰🇷 ソウル旅行 費用', '11/14〜15 1泊2日｜現地レシートを読み取って総額を自動集計'],
      risk: ['🛡️ 税務リスク診断', '調査で指摘されやすい項目を自動で洗い出し・源泉徴収票と損益通算'],
      guide: ['📋 e-Tax 青色申告ガイド', 'スマホを見ながらそのまま操作・数字の転記先がわかる']
    };
    var t = titles[m] || titles.tax;
    $('appTitle').textContent = t[0];
    $('appSubtitle').textContent = t[1];

    $('viewLedger').classList.toggle('hidden', m === 'risk' || m === 'guide');
    $('viewRisk').classList.toggle('hidden', m !== 'risk');
    $('viewGuide').classList.toggle('hidden', m !== 'guide');
    $('stickyBar').classList.toggle('hidden', m === 'risk' || m === 'guide');
    $('yearBar').classList.toggle('hidden', trip);

    $('rateCard').classList.toggle('hidden', !trip);
    $('catCard').classList.toggle('hidden', !trip);
    $('listTitle').textContent = trip ? '旅行の費用明細' : '登録済みの領収書';
    $('exportTitle').textContent = trip ? '📥 ソウル旅行のExcel出力' : '📥 Excel出力';
    $('exportNote').innerHTML = trip
      ? 'シート「ソウル旅行費用明細」＋「カテゴリ別集計」。円建て・ウォン建てを円換算して合算します。'
      : 'シート「領収書明細」＋「科目別集計」＋「税務チェック」＋「確定申告サマリー」。金額はSUM数式入りで、そのまま計算・並べ替えできます。';
    $('stickyLabel').textContent = trip ? 'ソウル旅行 総合計' : '合計金額（税込）';

    Array.prototype.forEach.call(document.querySelectorAll('.mode-tab'), function (b) {
      var on = b.getAttribute('data-mode') === m;
      b.classList.toggle('active', on);
      b.setAttribute('aria-selected', on ? 'true' : 'false');
    });

    var sel = $('langSelect');
    if (trip && sel.value === 'jpn+eng') sel.value = 'kor+eng';
    if (!trip && sel.value === 'kor+eng') sel.value = 'jpn+eng';
  }

  function renderYearBar() {
    var sel = $('yearSelect');
    if (!sel) return;
    var ys = availableYears();
    var cur = String(state.year);
    var html = ys.map(function (y) {
      return '<option value="' + y + '"' + (String(y) === cur ? ' selected' : '') + '>' + yearLabel(y) + '</option>';
    }).join('');
    sel.innerHTML = html;
    sel.value = cur;
  }

  /* ============================================================
     KPI・一覧（タブ1・2）
     ============================================================ */
  function renderKPIs() {
    var g = $('kpiGrid');
    if (!g) return;
    var list = currentList();

    if (state.mode !== 'trip') {
      var total = 0, withInvoice = 0, months = {};
      list.forEach(function (e) {
        var a = yenOf(e);
        total += a;
        if (e.invoiceNo) withInvoice += 1;
        var mo = (e.date || '').slice(0, 7);
        if (mo) months[mo] = (months[mo] || 0) + a;
      });
      var mcount = Object.keys(months).length;
      var avg = mcount ? Math.round(total / mcount) : 0;
      var sales = salesOf();
      g.innerHTML =
        '<div class="kpi wide"><div class="kpi-label">💰 経費 合計（税込）</div>' +
        '<div class="kpi-value">' + yen(total) + '</div>' +
        '<div class="kpi-sub">' + yearLabel(state.year) + '分｜' + list.length + '件・' + Object.keys(months).length + 'か月分｜売上 ' + yen(sales) + '</div></div>' +
        '<div class="kpi"><div class="kpi-label">🧾 登録件数</div><div class="kpi-value">' + list.length + '<span style="font-size:12px;font-weight:600;color:#6b7280;"> 件</span></div>' +
        '<div class="kpi-sub">インボイス登録済 ' + withInvoice + '件</div></div>' +
        '<div class="kpi"><div class="kpi-label">📅 月平均の経費</div><div class="kpi-value green">' + yen(avg) + '</div>' +
        '<div class="kpi-sub">' + (mcount ? mcount + 'か月で割った平均' : 'データなし') + '</div></div>';
      $('stickyValue').textContent = yen(total);
      return;
    }

    // 韓国旅行
    var sum = 0, jpySum = 0, krwSum = 0, krwYen = 0;
    var cat = { '航空券': 0, 'ホテル': 0, '交通費': 0, '食費・買い物': 0, 'その他': 0 };
    list.forEach(function (e) {
      var y = yenOf(e);
      sum += y;
      if (e.currency === 'KRW') { krwSum += num(e.amount); krwYen += y; }
      else { jpySum += y; }
      cat[e.category || 'その他'] = (cat[e.category] || 0) + y;
    });
    var receipts = list.filter(function (e) { return e.source !== 'preset'; }).length;

    g.innerHTML =
      '<div class="kpi wide"><div class="kpi-label">🇰🇷 ソウル旅行 総合計</div>' +
      '<div class="kpi-value">' + yen(sum) + '</div>' +
      '<div class="kpi-sub">円建て ' + yen(jpySum) + ' ＋ ウォン ' + won(krwSum) + '（' + yen(krwYen) + '）｜レート 1₩=' + state.rate + '円</div></div>' +
      '<div class="kpi"><div class="kpi-label">✈️ 航空券</div><div class="kpi-value" style="font-size:16px;">' + yen(cat['航空券']) + '</div><div class="kpi-sub">往復・Trip.com</div></div>' +
      '<div class="kpi"><div class="kpi-label">🏨 ホテル</div><div class="kpi-value" style="font-size:16px;">' + yen(cat['ホテル']) + '</div><div class="kpi-sub">免税返金後</div></div>' +
      '<div class="kpi"><div class="kpi-label">🚃 交通費</div><div class="kpi-value" style="font-size:16px;">' + yen(cat['交通費']) + '</div><div class="kpi-sub">南海・AREX・地下鉄</div></div>' +
      '<div class="kpi"><div class="kpi-label">🍽️ 食費・買い物</div><div class="kpi-value" style="font-size:16px;">' + yen(cat['食費・買い物'] + cat['その他']) + '</div><div class="kpi-sub">現地レシート ' + receipts + '枚</div></div>';
    $('stickyValue').textContent = yen(sum);
  }

  function renderCategories() {
    if (state.mode !== 'trip') return;
    var list = state.trip;
    var cat = { '航空券': 0, 'ホテル': 0, '交通費': 0, '食費・買い物': 0, 'その他': 0 };
    var total = 0;
    list.forEach(function (e) {
      var y = yenOf(e);
      total += y;
      cat[e.category || 'その他'] = (cat[e.category] || 0) + y;
    });
    $('catCount').textContent = list.length + '件';
    $('catList').innerHTML = ['航空券', 'ホテル', '交通費', '食費・買い物', 'その他'].map(function (k) {
      var v = cat[k] || 0;
      var pct = total > 0 ? Math.round(v / total * 100) : 0;
      return '<div class="cat-row"><span class="cat-name">' + k + '</span>' +
        '<span class="cat-bar"><i style="width:' + Math.min(100, pct) + '%"></i></span>' +
        '<span class="cat-amt">' + yen(v) + '<span style="display:block;font-size:9px;color:#9ca3af;font-weight:500;">' + pct + '%</span></span></div>';
    }).join('');
  }

  function sortedEntries(list) {
    return list.slice().sort(function (a, b) {
      var d = String(a.date || '').localeCompare(String(b.date || ''));
      if (d !== 0) return d;
      return String(a.createdAt || '').localeCompare(String(b.createdAt || ''));
    });
  }

  function renderEntries() {
    var list = currentList();
    var box = $('entries');
    if (!box) return;
    $('listCount').textContent = list.length + '件';

    if (!list.length) {
      box.innerHTML = '<div class="empty">' + (state.mode === 'trip'
        ? 'レシートを読み取ると、ここに旅行の支出が追加されていきます。'
        : yearLabel(state.year) + '分の登録はまだありません。<br>上のボタンから領収書を撮影するか、税務診断タブでCSVを取り込めます。') + '</div>';
      return;
    }

    var sorted = sortedEntries(list).reverse();
    box.innerHTML = sorted.map(function (e) {
      var local = e.currency === 'KRW' ? won(e.amount) : yen(e.amount);
      var sub = e.currency === 'KRW' ? yenOf(e) : '';
      var meta = [];
      if (e.date) meta.push(esc(e.date));
      if (state.mode !== 'trip' && e.taxType) meta.push(esc(e.taxType));
      if (e.invoiceNo) meta.push(esc(e.invoiceNo));
      if (e.source === 'import') meta.push('取込');
      if (e.memo) meta.push(esc(e.memo).slice(0, 40));
      return '<div class="entry" data-id="' + e.id + '">' +
        '<div class="entry-main">' +
        '<div class="entry-top"><span class="entry-vendor">' + esc(e.vendor || '（店名なし）') + '</span>' +
        (e.category ? '<span class="chip">' + esc(e.category) + '</span>' : '') + '</div>' +
        '<div class="entry-sub">' + meta.join(' · ') + '</div></div>' +
        '<div class="entry-amt"><b>' + local + '</b>' + (sub ? '<span>≈' + sub + '</span>' : '') + '</div>' +
        '<button class="icon-btn" data-act="edit" aria-label="編集">✏️</button>' +
        '<button class="icon-btn" data-act="del" aria-label="削除">🗑️</button>' +
        '</div>';
    }).join('');
  }

  function renderListFooter() {
    var has = currentList().length > 0;
    $('listFooter').classList.toggle('hidden', !(has || state.mode === 'trip'));
    $('btnRestore').classList.toggle('hidden', state.mode !== 'trip');
    $('btnClear').classList.toggle('hidden', !has);
  }

  /* ============================================================
     入力カード（OCR結果 / 手入力 / 編集）
     ============================================================ */
  function blankFields() {
    return {
      date: todayStr(),
      vendor: '',
      category: state.mode === 'trip' ? '食費・買い物' : '雑費',
      currency: state.mode === 'trip' ? 'KRW' : 'JPY',
      amount: '',
      taxType: '',
      invoiceNo: '',
      memo: ''
    };
  }

  function selectOptions(list, selected) {
    return list.map(function (c) {
      return '<option value="' + esc(c) + '"' + (c === selected ? ' selected' : '') + '>' + esc(c || '（未設定）') + '</option>';
    }).join('');
  }

  function fieldsHTML(d) {
    var f = d.fields;
    var trip = state.mode === 'trip';
    var html = '<div class="fields">';
    html += '<div class="grid2">';
    html += '<label class="field"><span>日付</span><input class="input" data-f="date" type="date" value="' + esc(f.date) + '"></label>';
    html += '<label class="field"><span>' + (trip ? '金額' : '金額（税込）') + '</span><input class="input" data-f="amount" type="text" inputmode="numeric" placeholder="0" value="' + esc(f.amount) + '"></label>';
    html += '</div>';
    html += '<label class="field"><span>' + (trip ? '店名・内容' : '支払先・店名') + '</span><input class="input" data-f="vendor" type="text" placeholder="' + (trip ? '例：GS25 明洞店' : '例：セブン-イレブン 駅前店') + '" value="' + esc(f.vendor) + '"></label>';

    if (trip) {
      html += '<div class="grid2"><div class="field"><span>通貨</span><div class="cur-toggle" data-cur="' + esc(f.currency) + '">' +
        '<button type="button" class="cur-btn' + (f.currency === 'JPY' ? ' active' : '') + '" data-cur-val="JPY">🇯🇵 円</button>' +
        '<button type="button" class="cur-btn' + (f.currency === 'KRW' ? ' active' : '') + '" data-cur-val="KRW">🇰🇷 ウォン</button>' +
        '</div></div>' +
        '<label class="field"><span>カテゴリ</span><select class="input" data-f="category">' + selectOptions(TRIP_CATEGORIES, f.category) + '</select></label></div>';
      html += '<div class="yen-preview" data-preview></div>';
    } else {
      html += '<div class="grid2">' +
        '<label class="field"><span>勘定科目</span><select class="input" data-f="category">' + selectOptions(TAX_CATEGORIES, f.category) + '</select></label>' +
        '<label class="field"><span>税区分</span><select class="input" data-f="taxType">' + selectOptions(TAX_TYPES, f.taxType) + '</select></label></div>';
      html += '<label class="field"><span>インボイス番号</span><input class="input" data-f="invoiceNo" type="text" placeholder="T1234567890123" value="' + esc(f.invoiceNo) + '"></label>';
    }
    html += '<label class="field"><span>メモ</span><textarea class="input" data-f="memo" rows="2" placeholder="用途・同行者など（按分した場合は「按分30%」など根拠も）">' + esc(f.memo) + '</textarea></label>';
    html += '</div>';
    if (d.text) {
      html += '<details class="ocr-text"><summary>読み取ったテキストを確認</summary><pre>' + esc(d.text.slice(0, 1500)) + '</pre></details>';
    }
    return html;
  }

  function previewUpdate(card, d) {
    var pv = card.querySelector('[data-preview]');
    if (!pv) return;
    var amountEl = card.querySelector('[data-f="amount"]');
    var a = num(amountEl ? amountEl.value : d.fields.amount);
    if (d.fields.currency === 'KRW') pv.textContent = '≈ ' + yen(Math.round(a * state.rate)) + '（1₩=' + state.rate + '円）';
    else pv.textContent = '';
  }

  function buildDraftCard(d) {
    var card = document.createElement('div');
    card.className = 'draft';
    card.id = 'draft-' + d.id;

    var head = '<div class="draft-head">' +
      (d.thumb
        ? '<img class="thumb" src="' + d.thumb + '" alt="領収書の写真">'
        : '<div class="thumb thumb-empty">' + (d.status === 'manual' ? '✍️' : '🧾') + '</div>') +
      '<div class="draft-head-main">' +
      '<div class="draft-status" data-status>' + statusHTML(d) + '</div>' +
      '<div class="draft-msg" data-msg>' + esc(d.message || '') + '</div>' +
      '<div class="bar" data-bar><i style="width:' + Math.round((d.progress || 0) * 100) + '%"></i></div>' +
      '</div></div>';

    var tools = '<div class="draft-tools">' +
      (d.status === 'manual' ? '' : '<button type="button" class="mini-btn" data-tool="rotate">↻ 90°回して再読取</button>' +
        '<button type="button" class="mini-btn" data-tool="retry">🔄 再読取</button>') +
      '<button type="button" class="mini-btn" data-tool="cancel">✕ 破棄</button>' +
      '</div>';

    card.innerHTML = head + tools + fieldsHTML(d) +
      '<div class="save-row">' +
      '<button type="button" class="btn-save" data-tool="save"' + (d.status === 'running' ? ' disabled' : '') + '>' +
      (d.editingId ? '💾 変更を保存' : '✅ 表に追加') + '</button>' +
      '<button type="button" class="btn-cancel" data-tool="dismiss">' + (d.editingId ? 'やめる' : '閉じる') + '</button>' +
      '</div>';

    Array.prototype.forEach.call(card.querySelectorAll('[data-cur-val]'), function (btn) {
      btn.addEventListener('click', function () {
        d.fields.currency = btn.getAttribute('data-cur-val');
        Array.prototype.forEach.call(card.querySelectorAll('[data-cur-val]'), function (b) {
          b.classList.toggle('active', b.getAttribute('data-cur-val') === d.fields.currency);
        });
        previewUpdate(card, d);
      });
    });
    var amtEl = card.querySelector('[data-f="amount"]');
    if (amtEl) amtEl.addEventListener('input', function () { previewUpdate(card, d); });

    Array.prototype.forEach.call(card.querySelectorAll('[data-tool]'), function (btn) {
      btn.addEventListener('click', function () {
        var tool = btn.getAttribute('data-tool');
        if (tool === 'save') saveDraft(d, card);
        else if (tool === 'dismiss' || tool === 'cancel') discardDraft(d);
        else if (tool === 'retry') { d.rotation = d.rotation || 0; runOCR(d); }
        else if (tool === 'rotate') { d.rotation = ((d.rotation || 0) + 90) % 360; runOCR(d); }
      });
    });

    previewUpdate(card, d);
    if (d.status === 'running') card.querySelector('[data-bar]').classList.remove('hidden');
    return card;
  }

  function statusHTML(d) {
    if (d.status === 'running') return '<span class="spin"></span><span data-status-text>読み取り中…</span>';
    if (d.status === 'error') return '⚠️ <span>読み取れませんでした</span>';
    if (d.status === 'manual') return '✍️ <span>手入力</span>';
    if (d.editingId) return '✏️ <span>編集中</span>';
    return '✅ <span>読み取り完了</span>';
  }

  function refreshDraftCard(d) {
    var card = $('draft-' + d.id);
    if (!card) return;
    var st = card.querySelector('[data-status]');
    if (st) st.innerHTML = statusHTML(d);
    var msg = card.querySelector('[data-msg]');
    if (msg) {
      msg.textContent = d.message || '';
      msg.className = 'draft-msg' + (d.status === 'error' ? ' err' : '');
    }
    var bar = card.querySelector('[data-bar]');
    if (bar) {
      var fill = bar.querySelector('i');
      if (fill) fill.style.width = Math.round((d.progress || 0) * 100) + '%';
      bar.classList.toggle('done', d.status === 'done');
    }
    var saveBtn = card.querySelector('[data-tool="save"]');
    if (saveBtn) saveBtn.disabled = d.status === 'running';
    if (d.status === 'error') card.classList.add('error');
    else card.classList.remove('error');
  }

  function fillFields(d) {
    var card = $('draft-' + d.id);
    if (!card) return;
    ['date', 'vendor', 'amount', 'category', 'taxType', 'invoiceNo', 'memo'].forEach(function (k) {
      var input = card.querySelector('[data-f="' + k + '"]');
      if (input && d.fields[k] != null && d.fields[k] !== '') input.value = d.fields[k];
    });
    Array.prototype.forEach.call(card.querySelectorAll('[data-cur-val]'), function (b) {
      b.classList.toggle('active', b.getAttribute('data-cur-val') === d.fields.currency);
    });
    previewUpdate(card, d);
  }

  function mountDraft(d) {
    var host = $('drafts');
    var card = buildDraftCard(d);
    host.insertBefore(card, host.firstChild);
    if (d.text) {
      var old = card.querySelector('.ocr-text');
      if (!old) {
        var wrap = document.createElement('div');
        wrap.innerHTML = '<details class="ocr-text"><summary>読み取ったテキストを確認</summary><pre>' + esc(d.text.slice(0, 1500)) + '</pre></details>';
        card.appendChild(wrap.firstChild);
      }
    }
  }

  function addManualDraft() {
    var d = {
      id: uid(), thumb: '', status: 'manual', progress: 1,
      message: '日付・金額・店名を入れて「表に追加」してください。',
      fields: blankFields(), text: '', rotation: 0
    };
    state.drafts.push(d);
    mountDraft(d);
    var card = $('draft-' + d.id);
    if (card) {
      card.scrollIntoView({ behavior: 'smooth', block: 'center' });
      var v = card.querySelector('[data-f="vendor"]');
      if (v) setTimeout(function () { v.focus(); }, 250);
    }
  }

  function saveDraft(d, card) {
    var f = {};
    ['date', 'vendor', 'amount', 'category', 'taxType', 'invoiceNo', 'memo'].forEach(function (k) {
      var input = card.querySelector('[data-f="' + k + '"]');
      f[k] = input ? input.value : (d.fields[k] || '');
    });
    f.currency = state.mode === 'trip' ? d.fields.currency : 'JPY';
    var amount = num(f.amount);

    if (!amount && amount !== 0) { toast('金額を入力してください', true); return; }
    if (amount <= 0) { toast('金額は1円以上で入力してください', true); return; }

    var entry = {
      id: d.editingId || uid(),
      date: f.date || todayStr(),
      vendor: (f.vendor || '').trim(),
      category: f.category || (state.mode === 'trip' ? 'その他' : '雑費'),
      currency: f.currency,
      amount: amount,
      taxType: f.taxType || '',
      invoiceNo: (f.invoiceNo || '').trim(),
      memo: (f.memo || '').trim(),
      thumb: d.thumb || '',
      source: d.status === 'manual' ? 'manual' : 'ocr',
      createdAt: d.createdAt || new Date().toISOString()
    };

    var list = currentList();
    if (d.editingId) {
      var idx = list.findIndex(function (x) { return x.id === d.editingId; });
      if (idx >= 0) list[idx] = entry; else list.push(entry);
    } else {
      list.push(entry);
    }

    persist();
    removeDraft(d);
    renderAll();
    toast(state.mode === 'trip'
      ? '「' + (entry.vendor || entry.category) + '」を追加しました（' + (entry.currency === 'KRW' ? won(entry.amount) + ' ≈ ' + yen(yenOf(entry)) : yen(entry.amount)) + '）'
      : '「' + (entry.vendor || entry.category) + '」を ' + yen(entry.amount) + ' で追加しました');

    signal((state.mode === 'trip' ? 'ソウル旅行に' : '経費帳に') + '「' + (entry.vendor || entry.category) + '」' + yenOf(entry) + ' を登録しました。');
  }

  function removeDraft(d) {
    var card = $('draft-' + d.id);
    if (card && card.parentNode) card.parentNode.removeChild(card);
    state.drafts = state.drafts.filter(function (x) { return x.id !== d.id; });
  }

  function discardDraft(d) {
    if (d.editingId && d.backup) {
      var list = currentList();
      if (!list.some(function (x) { return x.id === d.editingId; })) list.push(d.backup);
      persist();
    }
    removeDraft(d);
    renderAll();
  }

  function startEdit(entry) {
    var list = currentList();
    var idx = list.findIndex(function (x) { return x.id === entry.id; });
    if (idx < 0) return;
    var backup = Object.assign({}, list[idx]);
    list.splice(idx, 1);
    persist();

    var d = {
      id: uid(), thumb: entry.thumb || '', status: 'done', progress: 1,
      message: '値を直して「変更を保存」を押してください。',
      fields: {
        date: entry.date || todayStr(),
        vendor: entry.vendor || '',
        category: entry.category || '',
        currency: entry.currency || 'JPY',
        amount: entry.amount,
        taxType: entry.taxType || '',
        invoiceNo: entry.invoiceNo || '',
        memo: entry.memo || ''
      },
      text: '', rotation: 0, editingId: entry.id, backup: backup, createdAt: entry.createdAt
    };
    state.drafts.push(d);
    renderAll();
    mountDraft(d);
    var card = $('draft-' + d.id);
    if (card) card.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  /* ============================================================
     画像 → OCR
     ============================================================ */
  var worker = null, workerLang = null, workerSource = 'local';

  function absUrl(p) {
    try { return new URL(p, window.location.href).href; } catch (e) { return p; }
  }

  var PATHS = {
    local: {
      workerPath: absUrl('./vendor/worker.min.js'),
      corePath: absUrl('./vendor'),
      langPath: absUrl('./vendor/tessdata')
    },
    cdn: {
      workerPath: 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/worker.min.js',
      corePath: 'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.1.1',
      langPath: 'https://tessdata.projectnaptha.com/4.0.0_fast'
    }
  };

  var STATUS_JA = {
    'loading tesseract core': 'OCRエンジンを準備中',
    'initializing tesseract': 'OCRエンジンを初期化中',
    'loading language traineddata': '辞書データを読み込み中',
    'initializing api': '読み取りの準備中',
    'recognizing text': '読み取り中'
  };

  function onProgress(m) {
    if (!activeDraftId) return;
    var d = state.drafts.find(function (x) { return x.id === activeDraftId; });
    if (!d) return;
    var label = STATUS_JA[m.status] || '処理中';
    d.progress = typeof m.progress === 'number' ? m.progress : 0;
    d.message = label + ' ' + Math.round((m.progress || 0) * 100) + '%';
    refreshDraftCard(d);
    var st = $('draft-' + d.id);
    if (st) {
      var t = st.querySelector('[data-status-text]');
      if (t) t.textContent = label + '…';
    }
  }

  async function ensureWorker(lang) {
    if (worker && workerLang === lang) return worker;
    if (worker) {
      try { await worker.terminate(); } catch (e) { /* noop */ }
      worker = null; workerLang = null;
    }
    var build = async function (paths) {
      return await Tesseract.createWorker(lang, 1, {
        workerPath: paths.workerPath,
        corePath: paths.corePath,
        langPath: paths.langPath,
        cacheMethod: 'none',
        gzip: true,
        logger: onProgress,
        errorHandler: function (e) { console.warn('[tesseract]', e); }
      });
    };

    try {
      worker = await withTimeout(build(PATHS.local), 90000, 'OCRエンジンの読み込みがタイムアウトしました');
      workerSource = 'local';
    } catch (e1) {
      console.warn('local OCR engine failed', e1);
      worker = null;
      try {
        worker = await withTimeout(build(PATHS.cdn), 90000, 'OCRエンジンの読み込みがタイムアウトしました');
        workerSource = 'cdn';
        toast('端末内のOCRエンジンが読めなかったため、オンライン版で読み取りました');
      } catch (e2) {
        worker = null;
        workerLang = null;
        throw new Error('OCRエンジンを起動できませんでした。通信環境を確かめて再試行するか、手入力で追加してください。');
      }
    }
    workerLang = lang;
    try {
      await worker.setParameters({ preserve_interword_spaces: '1', user_defined_dpi: '300' });
    } catch (e) { /* noop */ }
    return worker;
  }

  function loadImage(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () {
        var img = new Image();
        img.onload = function () { resolve(img); };
        img.onerror = function () { reject(new Error('画像を読み込めませんでした')); };
        img.src = reader.result;
      };
      reader.onerror = function () { reject(new Error('画像を読み込めませんでした')); };
      reader.readAsDataURL(file);
    });
  }

  function makeThumb(img) {
    var w = 200, h = Math.round(img.height * (200 / img.width));
    var c = document.createElement('canvas');
    c.width = w; c.height = h;
    c.getContext('2d').drawImage(img, 0, 0, w, h);
    try { return c.toDataURL('image/jpeg', 0.62); } catch (e) { return ''; }
  }

  function prepareCanvas(img, rotation) {
    var w = img.width, h = img.height;
    var scale = Math.min(1, 1800 / Math.max(w, h));
    if (Math.min(w, h) < 700) scale = Math.min(2, 1200 / Math.min(w, h));
    var cw = Math.max(1, Math.round(w * scale));
    var ch = Math.max(1, Math.round(h * scale));
    var rot = ((rotation || 0) + 360) % 360;
    var canvas = document.createElement('canvas');
    if (rot === 90 || rot === 270) { canvas.width = ch; canvas.height = cw; }
    else { canvas.width = cw; canvas.height = ch; }
    var ctx = canvas.getContext('2d', { willReadFrequently: true });
    ctx.save();
    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate(rot * Math.PI / 180);
    ctx.drawImage(img, -cw / 2, -ch / 2, cw, ch);
    ctx.restore();

    try {
      var data = ctx.getImageData(0, 0, canvas.width, canvas.height);
      var d = data.data, i, len = d.length, sum = 0;
      for (i = 0; i < len; i += 4) sum += d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114;
      var mean = sum / (len / 4);
      var contrast = 1.35;
      var bias = (128 - mean) * 0.45;
      for (i = 0; i < len; i += 4) {
        var g = (d[i] * 0.299 + d[i + 1] * 0.587 + d[i + 2] * 0.114) * contrast + 128 * (1 - contrast) + bias;
        g = g < 0 ? 0 : g > 255 ? 255 : g;
        d[i] = d[i + 1] = d[i + 2] = g;
      }
      ctx.putImageData(data, 0, 0);
    } catch (e) { /* 画像処理に失敗してもOCR本体は続行 */ }
    return canvas;
  }

  var ocrQueue = [];
  var ocrBusy = false;

  function enqueueOCR(id) {
    ocrQueue.push(id);
    pumpOCR();
  }

  async function pumpOCR() {
    if (ocrBusy) return;
    ocrBusy = true;
    while (ocrQueue.length) {
      var id = ocrQueue.shift();
      var d = state.drafts.find(function (x) { return x.id === id; });
      if (!d || !d.img) continue;
      await runOCR(d);
    }
    ocrBusy = false;
  }

  async function runOCR(d) {
    var card = $('draft-' + d.id);
    d.status = 'running';
    d.progress = 0;
    d.message = '読み取りを準備しています…';
    d.error = '';
    activeDraftId = d.id;
    refreshDraftCard(d);
    if (card) card.scrollIntoView({ behavior: 'smooth', block: 'center' });

    var lang = $('langSelect').value || 'jpn+eng';
    try {
      var w = await ensureWorker(lang);
      var canvas = prepareCanvas(d.img, d.rotation || 0);
      var res = await withTimeout(w.recognize(canvas), 120000, '読み取りが時間内に終わりませんでした');
      var text = (res && res.data && res.data.text) ? res.data.text : '';
      d.text = text;
      if (!text || text.replace(/\s/g, '').length < 4) {
        throw new Error('文字をうまく読み取れませんでした。明るい場所で撮り直すか、手入力で追加してください。');
      }
      var parsed = parseReceipt(text, state.mode);
      d.fields = Object.assign({}, d.fields, parsed.fields);
      d.status = 'done';
      d.progress = 1;
      d.message = parsed.hint || ('読み取り完了。内容を確認して「表に追加」を押してください。');
      fillFields(d);
      var c2 = $('draft-' + d.id);
      if (c2 && !c2.querySelector('.ocr-text')) {
        var det = document.createElement('details');
        det.className = 'ocr-text';
        det.innerHTML = '<summary>読み取ったテキストを確認</summary><pre>' + esc(text.slice(0, 1500)) + '</pre>';
        c2.appendChild(det);
      }
    } catch (e) {
      d.status = 'error';
      d.progress = 0;
      d.error = String((e && e.message) || e);
      d.message = '読み取れませんでした：' + d.error + '（手入力でも追加できます）';
      if (workerSource === 'cdn') d.message += ' ※オンラインOCRで試行済み';
      if (/時間内に終わり|タイムアウト/.test(d.error) && worker) {
        try { await worker.terminate(); } catch (e2) { /* noop */ }
        worker = null; workerLang = null;
      }
    } finally {
      activeDraftId = null;
      refreshDraftCard(d);
      var c3 = $('draft-' + d.id);
      if (c3) {
        var saveBtn = c3.querySelector('[data-tool="save"]');
        if (saveBtn) saveBtn.disabled = false;
        var tools = c3.querySelector('.draft-tools');
        if (tools && !tools.querySelector('[data-tool="retry"]')) {
          tools.innerHTML = '<button type="button" class="mini-btn" data-tool="rotate">↻ 90°回して再読取</button>' +
            '<button type="button" class="mini-btn" data-tool="retry">🔄 再読取</button>' +
            '<button type="button" class="mini-btn" data-tool="cancel">✕ 破棄</button>';
          Array.prototype.forEach.call(tools.querySelectorAll('[data-tool]'), function (btn) {
            btn.addEventListener('click', function () {
              var tool = btn.getAttribute('data-tool');
              if (tool === 'retry') runOCR(d);
              else if (tool === 'rotate') { d.rotation = ((d.rotation || 0) + 90) % 360; runOCR(d); }
              else discardDraft(d);
            });
          });
        }
      }
    }
  }

  async function handleFiles(files) {
    if (!files || !files.length) return;
    var list = Array.prototype.slice.call(files);
    for (var i = 0; i < list.length; i += 1) {
      var file = list[i];
      if (!/^image\//.test(file.type)) { toast('画像ファイルを選んでください', true); continue; }
      try {
        var img = await loadImage(file);
        var d = {
          id: uid(),
          img: img,
          thumb: makeThumb(img),
          status: 'running',
          progress: 0,
          message: '読み取り待ち…',
          fields: blankFields(),
          text: '',
          rotation: 0,
          createdAt: new Date().toISOString()
        };
        state.drafts.push(d);
        mountDraft(d);
        enqueueOCR(d.id);
      } catch (e) {
        toast('画像を読み込めませんでした：' + ((e && e.message) || e), true);
      }
    }
  }

  /* ============================================================
     領収書テキスト解析
     ============================================================ */
  function normalize(s) {
    if (!s) return '';
    return String(s)
      .replace(/[\uFF01-\uFF5E]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
      .replace(/\u3000/g, ' ')
      .replace(/[，]/g, ',')
      .replace(/￥/g, '¥')
      .replace(/[：]/g, ':')
      .replace(/[ \t]+/g, ' ');
  }

  function numbersIn(line) {
    var out = [];
    var re = /([0-9]{1,3}(?:,[0-9]{3})+|[0-9]+)(?:\.[0-9]+)?/g;
    var m;
    while ((m = re.exec(line)) !== null) out.push({ v: num(m[1]), raw: m[1], idx: m.index });
    return out;
  }

  var JUNK_LINE = /(TEL|전화|전화번호|사업자|승인\s?번호|카드\s?번호|번호|No\.|바코드|시리얼|거래\s?번호|회원|ポイント番号|会員番号|電話|登録番号|伝票|レシート番号)/i;
  var PHONE_LIKE = /[0-9]{2,4}[\-‐][0-9]{2,4}[\-‐][0-9]{2,4}/;
  var DATE_LIKE = /(20[0-9]{2}|令和[0-9元]{1,3})[年\/.\-][0-9]{1,2}[月\/.\-][0-9]{1,2}/;

  function looksLikeJunk(n, line) {
    if (n.v < 10) return true;
    if (/^0[0-9]/.test(n.raw)) return true;
    var dg = String(n.raw).replace(/[^0-9]/g, '');
    if (dg.length >= 8 && n.raw.indexOf(',') < 0) return true;
    if (JUNK_LINE.test(line)) return true;
    if (PHONE_LIKE.test(line)) return true;
    if (DATE_LIKE.test(line)) return true;
    return false;
  }

  var JP_TOTAL_LABELS = [
    ['合計', 10], ['総合計', 12], ['税込合計', 11], ['合計金額', 11], ['ご利用金額', 9], ['利用金額', 8],
    ['お会計', 9], ['お買上げ', 8], ['お買上', 7], ['領収金額', 10], ['ご請求金額', 9], ['請求金額', 8], ['総額', 8], ['金額', 3]
  ];
  var JP_EXCLUDE = /(小計|消費税|税額|お預|預り|預かり|釣|おつり|お釣|現金|ポイント|残高|枚数|点数|割引|値引|対象額|内税|外税)/;
  var KR_TOTAL_LABELS = [
    ['합계', 10], ['총액', 9], ['결제금액', 11], ['결제 금액', 11], ['승인금액', 10], ['승인 금액', 10],
    ['받을금액', 8], ['합계금액', 11], ['판매금액', 7], ['금액', 3]
  ];
  var KR_EXCLUDE = /(부가세|세액|공급가액|거스름|받은금액|현금|카드|포인트|할인|잔액|수량|개)/;

  function parseAmount(lines, isKRDoc) {
    var best = null;
    var pool = [];

    lines.forEach(function (rawLine, li) {
      var line = normalize(rawLine).trim();
      if (!line) return;
      var isKR = isKRDoc || /(원|₩|KRW|won)/i.test(line);
      var labels = isKR ? KR_TOTAL_LABELS : JP_TOTAL_LABELS;
      var exclude = isKR ? KR_EXCLUDE : JP_EXCLUDE;
      var marked = /[¥₩]|円|원/.test(line);

      var hit = null;
      labels.forEach(function (pair) {
        if (hit) return;
        var idx = line.indexOf(pair[0]);
        if (idx >= 0) hit = { label: pair[0], score: pair[1], idx: idx };
      });

      var nums = numbersIn(line).filter(function (n) { return !looksLikeJunk(n, line); });
      if (!nums.length) return;

      if (hit && !exclude.test(line)) {
        var after = nums.filter(function (n) { return n.idx >= hit.idx; });
        var pick = after.length ? after[0] : nums[nums.length - 1];
        var cand = { v: pick.v, cur: isKR ? 'KRW' : 'JPY', score: hit.score + (li / 1000) };
        if (!best || cand.score > best.score) best = cand;
      }
      if (!exclude.test(line) && li > 1) {
        nums.forEach(function (n) { pool.push({ v: n.v, cur: isKR ? 'KRW' : 'JPY', marked: marked }); });
      }
    });

    if (best) return { amount: best.v, currency: best.cur };
    var markedPool = pool.filter(function (p) { return p.marked; });
    var use = markedPool.length ? markedPool : pool;
    if (use.length) {
      var max = use.reduce(function (a, b) { return b.v > a.v ? b : a; });
      return { amount: max.v, currency: max.cur };
    }
    return { amount: 0, currency: '' };
  }

  function parseDate(text) {
    var s = normalize(text);
    var m;
    if ((m = s.match(/(20\d{2})[年\/.\-](\d{1,2})[月\/.\-](\d{1,2})/))) {
      var y = +m[1], mo = +m[2], da = +m[3];
      if (mo >= 1 && mo <= 12 && da >= 1 && da <= 31) return fmtDate(y, mo, da);
    }
    if ((m = s.match(/令和(\d{1,2}|元)年(\d{1,2})月(\d{1,2})日/))) {
      var ry = m[1] === '元' ? 1 : +m[1];
      return fmtDate(2018 + ry, +m[2], +m[3]);
    }
    if ((m = s.match(/(?:R|r)(\d{1,2})[.\/](\d{1,2})[.\/](\d{1,2})/))) {
      return fmtDate(2018 + (+m[1]), +m[2], +m[3]);
    }
    if ((m = s.match(/(20\d{2})(\d{2})(\d{2})/))) {
      var yy = +m[1], mm = +m[2], dd = +m[3];
      if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) return fmtDate(yy, mm, dd);
    }
    if ((m = s.match(/(\d{1,2})月(\d{1,2})日/))) {
      var base = state.mode === 'trip' ? 2026 : state.year;
      return fmtDate(base, +m[1], +m[2]);
    }
    if ((m = s.match(/(20\d{2})[.\-](\d{1,2})[.\-](\d{1,2})/u))) {
      return fmtDate(+m[1], +m[2], +m[3]);
    }
    return '';
  }

  function fmtDate(y, m, d) {
    if (!y || !m || !d) return '';
    return y + '-' + (m < 10 ? '0' + m : m) + '-' + (d < 10 ? '0' + d : d);
  }

  function guessVendor(rawText) {
    var lines = normalize(rawText).split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean);
    var isKR = /[가-힣]/.test(rawText);
    var bad = isKR
      ? /(영수증|매출전표|합계|소계|부가세|세액|승인|카드|현금|거스름|일시|전화|사업자|대표|주소|ＴＥＬ|TEL|번호|감사합니다|고객|포인트|할인)/
      : /(領収書|領収証|レシート|合計|小計|税込|消費税|税額|お預|預り|釣|おつり|現金|クレジット|カード|TEL|電話|登録番号|インボイス|ご来店|ありがとう|お買上|点数|担当|レジ|No\.|日付|時刻|^\d|¥|円)/;
    var scored = [];
    lines.slice(0, 12).forEach(function (line, i) {
      var t = line.replace(/^[#*\-\u3000\s・.]+/, '').replace(/\s{2,}/g, ' ').trim();
      if (t.length < 2 || t.length > 36) return;
      if (!/[\u3040-\u30ff\u4e00-\u9fff가-힣A-Za-z]/.test(t)) return;
      if (bad.test(t)) return;
      var digits = (t.match(/[0-9]/g) || []).length;
      if (digits > t.length * 0.5) return;
      var score = 100 - i * 6 - Math.abs(t.length - 12);
      if (/(株式会社|店|ストア|マート|カフェ|珈琲|薬局|駅|호텔|마트|카페|식당|편의점)/.test(t)) score += 25;
      scored.push({ t: t, s: score });
    });
    scored.sort(function (a, b) { return b.s - a.s; });
    return scored.length ? scored[0].t : '';
  }

  var CAT_RULES = [
    ['旅費交通費', /(JR|鉄道|電車|地下鉄|バス|タクシー|ハイタク|航空|空港|切符|きっぷ|乗車券|運賃|特急|新幹線|ICOCA|PiTaPa|SUICA|PASMO|ANA|JAL|ピーチ|ジェットスター|スターフライヤー|Klook|AREX|交通|駐車|パーキング|ガソリン|給油|ENEOS|出光)/i],
    ['通信費', /(NTT|ドコモ|docomo|ソフトバンク|softbank|楽天モバイル|ahamo|モバイル|通信|インターネット|プロバイダ|Wi-?Fi|SIM|切手|はがき|郵便|NHK)/i],
    ['水道光熱費', /(電気|電力|ガス|水道|上下水|関西電力|大阪ガス)/],
    ['新聞図書費', /(書店|書籍|ブックス|本屋|新聞|雑誌|Kindle)/],
    ['接待交際費', /(居酒屋|焼肉|寿司|すし|レストラン|ダイニング|宴会|ビア|飲み放題|接待|バル)/],
    ['消耗品費', /(文具|文房具|事務用品|ボールペン|ペン|ノート|鉛筆|消しゴム|マーカー|用紙|封筒|ファイル|はさみ|テープ|インク|トナー|電池|充電池|マウス|キーボード|ケーブル|フィルム|ケース|洗剤|ティッシュ|マスク|医薬品|目薬|ホームセンター|コーナン|ダイソー|セリア|キャンドゥ|100円|ドラッグ|薬局|マツモト|ウエルシア|サンドラッグ|雑貨|Amazon|楽天|ヨドバシ|ビックカメラ)/i],
    ['会議費', /(カフェ|珈琲|コーヒー|スターバックス|スタバ|ドトール|タリーズ|喫茶|サンマルク|ランチ|モーニング|ベーカリー)/],
    ['支払手数料', /(手数料|振込|ATM|送金|決済サービス)/],
    ['地代家賃', /(家賃|賃料|共益費|管理費)/]
  ];

  var TRIP_RULES = [
    ['航空券', /(항공|航空|air|flight|trip\.com|we\d{3}|tw\d{3}|e-?チケット|チケット)/i],
    ['ホテル', /(호텔|hotel|東横|inn|숙박|게스트하우스|모텔|리조트)/i],
    ['交通費', /(지하철|버스|택시|교통|arex|공항철도|티머니|t-?money|기후동행|気候同行|승차권|교통카드|klook|철도|역|공항)/i],
    ['食費・買い物', /(식당|음식|카페|커피|편의점|마트|gs25|cu|세븐|이마트|롯데|스타벅스|빵|치킨|김밥|분식|쇼핑|면세|올리브영|다이소|명동|홍대|동대문)/i]
  ];

  function guessCategory(text, mode) {
    var t = normalize(text);
    var rules = mode === 'trip' ? TRIP_RULES : CAT_RULES;
    for (var i = 0; i < rules.length; i += 1) {
      if (rules[i][1].test(t)) return rules[i][0];
    }
    return mode === 'trip' ? '食費・買い物' : '雑費';
  }

  function parseReceipt(text, mode) {
    var s = normalize(text);
    var lines = s.split(/\r?\n/);
    var hasHangul = /[가-힣]/.test(s);
    var isKRDoc = hasHangul || /(원|₩|KRW)/.test(s);
    var amt = parseAmount(lines, isKRDoc);
    var date = parseDate(s);
    var vendor = guessVendor(text);
    var invoice = '';
    var inv = s.match(/T\s?([0-9]{13})/);
    if (inv) invoice = 'T' + inv[1];
    else {
      var inv2 = s.match(/登録番号[^0-9]{0,4}([0-9]{13})/);
      if (inv2) invoice = 'T' + inv2[1];
    }

    var taxType = '';
    if (/軽減税率?|8%/.test(s)) taxType = '8%軽減';
    else if (/非課税|不課税/.test(s)) taxType = '非課税';
    else if (/10%|消費税/.test(s)) taxType = '10%';

    var isKR = isKRDoc && !/円|¥/.test(s);
    var currency = mode === 'trip' ? (isKR || amt.currency === 'KRW' ? 'KRW' : 'JPY') : 'JPY';

    var category = guessCategory([vendor, text].join(' '), mode);

    var hint = '';
    var missing = [];
    if (!amt.amount) missing.push('金額');
    if (!date) missing.push('日付');
    if (!vendor) missing.push('店名');
    if (missing.length) hint = missing.join('・') + 'が読み取れませんでした。入力して「表に追加」してください。';

    return {
      fields: {
        date: date || (mode === 'trip' ? '2026-11-14' : todayStr()),
        vendor: vendor,
        amount: amt.amount || '',
        currency: currency,
        category: category,
        taxType: mode === 'trip' ? '' : taxType,
        invoiceNo: invoice,
        memo: ''
      },
      hint: hint
    };
  }

  /* ============================================================
     税務計算（経費集計・給与所得・所得税）
     ============================================================ */
  function formCategoryOf(cat) {
    if (FORM_MAP[cat]) return FORM_MAP[cat];
    if (FORM_CATEGORIES.indexOf(cat) >= 0) return cat;
    return '雑費';
  }

  /* 経費の集計（決算書の科目に寄せた形） */
  function totalsForYear(y) {
    var list = taxList(y);
    var expenses = 0, dokujuu = 0, byForm = {}, byRaw = {}, rawCount = {}, cnt = 0;
    list.forEach(function (e) {
      var a = yenOf(e);
      var cat = e.category || '雑費';
      if (cat === '専従者給与') { dokujuu += a; return; }
      expenses += a;
      cnt += 1;
      byRaw[cat] = (byRaw[cat] || 0) + a;
      rawCount[cat] = (rawCount[cat] || 0) + 1;
      var f = formCategoryOf(cat);
      byForm[f] = (byForm[f] || 0) + a;
    });
    return { expenses: expenses, dokujuu: dokujuu, byForm: byForm, byRaw: byRaw, rawCount: rawCount, count: cnt };
  }

  function salaryDeduction(pay) {
    if (!pay) return 0;
    if (pay <= 1625000) return 550000;
    if (pay <= 1800000) return Math.floor(pay * 0.4) - 100000;
    if (pay <= 3600000) return Math.floor(pay * 0.3) + 80000;
    if (pay <= 6600000) return Math.floor(pay * 0.2) + 440000;
    if (pay <= 8500000) return Math.floor(pay * 0.1) + 1100000;
    return 1950000;
  }

  var TAX_BRACKETS = [
    [1949000, 0.05, 0],
    [3299000, 0.10, 97500],
    [6949000, 0.20, 427500],
    [8999000, 0.23, 636000],
    [17999000, 0.33, 1536000],
    [39999000, 0.40, 2796000],
    [Infinity, 0.45, 4796000]
  ];
  function incomeTaxOf(taxable) {
    if (taxable <= 0) return 0;
    for (var i = 0; i < TAX_BRACKETS.length; i += 1) {
      if (taxable <= TAX_BRACKETS[i][0]) {
        return Math.max(0, Math.floor(taxable * TAX_BRACKETS[i][1] - TAX_BRACKETS[i][2]));
      }
    }
    return 0;
  }

  function slipsAutoDeduction(slips) {
    var v = 0;
    slips.forEach(function (s) { v = Math.max(v, num(s.deductions)); });
    return v;
  }

  /* 未入力の空欄の源泉徴収票は計算に入れない */
  function activeSlips(slips) {
    return slips.filter(function (s) {
      return num(s.payment) || num(s.afterDeduction) || num(s.deductions) || num(s.taxWithheld);
    });
  }

  function computeSim() {
    var y = state.year;
    var sales = salesOf(y);
    var tot = totalsForYear(y);
    var blue = blueOf(y);
    var biz = sales - tot.expenses - tot.dokujuu - blue;

    var slips = activeSlips(slipsOf(y));
    var pay = 0, afterSum = 0, wh = 0;
    slips.forEach(function (s) {
      pay += num(s.payment);
      afterSum += num(s.afterDeduction);
      wh += num(s.taxWithheld);
    });

    var salary;
    if (!slips.length) salary = 0;
    else if (slips.length === 1 && num(slips[0].afterDeduction) > 0) salary = num(slips[0].afterDeduction);
    else salary = Math.max(0, pay - salaryDeduction(pay));

    var totalIncome = salary + biz;

    var man = state.deductions[ykey(y)];
    var auto = slipsAutoDeduction(slips);
    var ded = (man === null || man === undefined || man === '') ? auto : Number(man);

    var taxableRaw = totalIncome - ded;
    var taxable = taxableRaw > 0 ? Math.floor(taxableRaw / 1000) * 1000 : 0;
    var incomeTax = incomeTaxOf(taxable);
    var fukko = y <= 2037 ? Math.floor(incomeTax * 0.021) : 0;
    var totalTax = incomeTax + fukko;
    var diff = totalTax - wh;

    return {
      year: y, sales: sales, expenses: tot.expenses, dokujuu: tot.dokujuu,
      count: tot.count, blue: blue, biz: biz,
      slipCount: slips.length, pay: pay, afterSum: afterSum, withheld: wh,
      salary: salary, totalIncome: totalIncome, deductions: ded, deductionsAuto: auto,
      deductionsManual: !(man === null || man === undefined || man === ''),
      taxableRaw: taxableRaw, taxable: taxable, incomeTax: incomeTax, fukko: fukko,
      totalTax: totalTax, diff: diff,
      residentTax: taxable > 0 ? Math.floor(taxable * 0.10) : 0
    };
  }

  /* ============================================================
     税務リスク診断
     ============================================================ */
  function textOf(e) {
    return [e.vendor, e.category, e.memo, e.invoiceNo].join(' ').toLowerCase();
  }
  function vendorKey(s) {
    return String(s || '').toLowerCase().replace(/[\s\u3000（）()【】\[\]・,.\-_/]/g, '').slice(0, 12);
  }

  /* クレジットカードの「引落そのもの」か（店舗での決済とは区別する） */
  function isCardPayment(text) {
    if (/(引落|引き落とし|口座振替|口座振込|自動払込|自動引落|利用代金|カード代金|クレジット代金|カード請求|カード返済)/.test(text)) return true;
    if (/(三井住友カード|楽天カード|セゾンカード|エポスカード|ビューカード|jcbカード|アメリカン・?エキスプレス|アメックス|ダイナース|ライフカード|オリコカード)/.test(text)) return true;
    return false;
  }

  var RED_RULES = [
    {
      id: 'debt', title: '借金返済・消費者金融',
      test: function (t) { return /(プロミス|アコム|smbcコンシューマー|三井住友銀行カードローン|レイク|アイフル|モビット|武富士|セディナ|オリックス・?クレジット|ジャックス|カードローン|貸金|借入|借り入れ|返済|リボ払|リボ手数料)/.test(t); },
      why: '借入金の返済は経費になりません。元本は資産の減少で、経費にできるのは利息（利子割引料）だけです。まずこの行を経費から外してください。'
    },
    {
      id: 'card', title: 'クレジットカード引落・口座振替そのもの',
      test: isCardPayment,
      why: 'カード会社への引落は「個別の買い物」をまとめた支払なので、そのまま経費にすると二重計上になります。カード明細の個別行を科目ごとに計上します。'
    },
    {
      id: 'emoney', title: '電子マネー・ICカードのチャージ',
      test: function (t) { return /(チャージ|charge|nanaco|waon|suica|スイカ|pasmo|icoca|edy|paypay残高|au ?wallet|モバイルsuica|交通系ic)/.test(t); },
      why: 'チャージは現金が電子マネーに変わっただけで、まだ経費ではありません。実際に使ったとき（利用明細）に計上します。'
    },
    {
      id: 'atm', title: 'ATM引出・出金',
      test: function (t) { return /(atm|お引出し|お引き出し|現金引出|現金払出|お出金|キャッシング)/.test(t); },
      why: 'ATMの出金は「経費を使った証拠」になりません。何に使ったのかの領収書・レシートが必要です。'
    }
  ];

  function analyzeRisks() {
    var y = state.year;
    var list = taxList(y);
    var flags = [];
    var redCount = 0, amberCount = 0, redAmt = 0, amberAmt = 0;

    function push(level, entry, title, why, ruleId) {
      var amt = entry ? yenOf(entry) : 0;
      flags.push({
        level: level, ruleId: ruleId, entryId: entry ? entry.id : '',
        title: title, why: why, amount: amt,
        date: entry ? (entry.date || '') : '', vendor: entry ? (entry.vendor || '（店名なし）') : '',
        category: entry ? (entry.category || '') : ''
      });
      if (level === 'red') { redCount += 1; redAmt += amt; }
      else { amberCount += 1; amberAmt += amt; }
    }

    list.forEach(function (e) {
      var t = textOf(e);
      var hitRules = {};
      RED_RULES.forEach(function (r) {
        if (hitRules[r.id]) return;
        if (r.test(t)) { hitRules[r.id] = 1; push('red', e, r.title, r.why, r.id); }
      });

      var amt = yenOf(e);
      var cat = e.category || '';
      var memo = String(e.memo || '') + ' ' + String(e.vendor || '');
      var memoLow = memo.toLowerCase();

      // 10万円以上：減価償却・一括償却資産
      var noDepTarget = ['地代家賃', '水道光熱費', '通信費', '旅費交通費', '外注工賃', '支払手数料', '給料賃金', '租税公課', '専従者給与'];
      if (amt >= 100000 && noDepTarget.indexOf(cat) < 0) {
        push('amber', e, '10万円以上の支出', '10万円以上は「消耗品費」で一括経費にできません。減価償却（10万円未満なら一括）か一括償却資産（20万円未満・3年均等）の検討が必要です。備品・工具・機材ならなおさらです。', 'capex');
      }

      // 家事按分の記録なし
      if ((cat === '通信費' || cat === '地代家賃' || cat === '水道光熱費') && !/(按分|事業|業務|仕事|事業用|事業分|％|%|割)/.test(memo)) {
        push('amber', e, '家事按分の記録がありません', '自宅の通信費・家賃・水道光熱費は、事業で使った割合だけが経費です。メモに「按分30%」のように割合と根拠（稼働日数・面積など）を残すと説明できます。', 'apportion');
      }

      // プライベート混在
      if (/(apple|itunes|app ?store|google ?play|google ?one|icloud|netflix|spotify|prime ?video|amazonプライム|steam|nintendo|playstation|パチンコ|スロット|競馬|映画|シネマ|ライブ|コンサート|チケット|旅行|観光|エステ|美容|ジム|フィットネス|サウナ|ゲーム|課金)/.test(t)) {
        push('amber', e, 'プライベート利用の疑い', '趣味・娯楽・サブスクは、事業に必要な範囲しか経費にできません。業務で使った理由（取材・制作・配信の素材など）をメモに残すか、私的利用分を除外してください。', 'private');
      }

      // 会議費・接待交際費で高額・人数不明
      if ((cat === '会議費' || cat === '接待交際費') && amt >= 5000 && !/(名|人|様|会食|打合せ|打ち合わせ|商談|面談)/.test(memo)) {
        push('amber', e, '会議費・交際費の相手と目的が不明', '誰と・何のために・何人で使ったかが分かるメモ（例：「取引先2名と打合せ」）がないと、私的飲食と疑われます。', 'meeting');
      }

      // 年度外の日付
      var yy = yearOf(e.date);
      if (yy && yy !== y) {
        push('amber', e, '年度外の日付', y + '年分の帳簿に' + yy + '年の日付が入っています。年度を切り替えて記帳するか、日付を直してください。', 'outyear');
      }
    });

    // 二重・三重計上
    var groups = {};
    list.forEach(function (e) {
      var k = (e.date || '') + '|' + Math.round(yenOf(e)) + '|' + vendorKey(e.vendor);
      if (!groups[k]) groups[k] = [];
      groups[k].push(e);
    });
    Object.keys(groups).forEach(function (k) {
      var g = groups[k];
      if (g.length < 2) return;
      g.slice(1).forEach(function (e) {
        push('red', e, '二重・三重計上の疑い', '同じ日・同じ金額・同じ相手の記録が' + g.length + '件あります。明細を1件ずつ確認して、重複分を削除してください。', 'dup');
      });
    });

    /* 構造バランス（事業所得とみなされるか） */
    var structural = [];
    var tot = totalsForYear(y);
    var exp = tot.expenses + tot.dokujuu;
    var sales = salesOf(y);

    if (exp > 0 && sales === 0) {
      structural.push({
        level: 'red',
        title: '売上0で経費だけが計上されています',
        why: 'この状態だと「事業所得」ではなく「雑所得（趣味）」とみなされ、給与所得との損益通算も65万円控除も認められません。売上（入金）を集計して入れるか、売上が無いなら経費の計上自体を見直してください。'
      });
    } else if (sales > 0 && sales < 100000 && exp >= 300000) {
      structural.push({
        level: 'red',
        title: '売上に対して経費が極端に大きい（損益通算の否認リスク）',
        why: '売上' + yen(sales) + 'に対して経費' + yen(exp) + '。この形は「『事業所得』ではなく『雑所得・趣味』」とみなされ、赤字の損益通算や65万円控除が否認されるリスクがあります。売上の計上漏れが無いか、経費に私的支出が混ざっていないかを先に確認してください。'
      });
    } else if (exp > sales && exp > 0) {
      structural.push({
        level: 'amber',
        title: '経費が売上を上回っています（赤字）',
        why: '事業所得の赤字は給与所得と損益通算できますが、毎年続くと事業実態を問われます。売上の計上漏れと、私的支出の混入を確認しておくと安全です。'
      });
    }

    return {
      flags: flags, structural: structural,
      redCount: redCount, amberCount: amberCount,
      redAmt: redAmt, amberAmt: amberAmt,
      total: list.length
    };
  }

  function renderRiskInputs() {
    var yl = $('riskYearLabel');
    if (!yl) return;
    yl.textContent = yearLabel(state.year) + '分';
    var si = $('salesInput');
    if (si && document.activeElement !== si) si.value = salesOf() ? String(salesOf()) : '';
    var bs = $('blueSelect');
    if (bs) bs.value = String(blueOf());
    var di = $('deductionsInput');
    if (di && document.activeElement !== di) {
      var man = state.deductions[ykey()];
      var auto = slipsAutoDeduction(slipsOf());
      di.value = (man === null || man === undefined || man === '') ? (auto ? String(auto) : '') : String(man);
      di.placeholder = auto ? '源泉徴収票③から自動（' + yen(auto) + '）' : '源泉徴収票③から自動';
    }
    renderSlips();
    if (lastImport) {
      var ip = $('importPreview');
      if (ip) ip.innerHTML = importPreviewHTML(lastImport);
      var ic = $('importCount');
      if (ic) ic.textContent = lastImport.total + '件取込';
    }
  }

  function renderSlips() {
    var box = $('slipList');
    if (!box) return;
    var slips = slipsOf();
    if (!slips.length) {
      slips.push({ id: uid(), payer: '', payment: '', afterDeduction: '', deductions: '', taxWithheld: '' });
    }
    $('slipCount').textContent = slips.length + '枚';
    box.innerHTML = slips.map(function (s, i) {
      return '<div class="slip" data-idx="' + i + '">' +
        '<div class="slip-head"><b>源泉徴収票 ' + (i + 1) + (i === 0 ? '（年末調整済み）' : '（前職など）') + '</b>' +
        (slips.length > 1 ? '<button class="mini-btn" data-slip-del="' + i + '" style="color:#dc2626;border-color:#fecaca;min-height:28px;padding:3px 8px;">削除</button>' : '') +
        '</div>' +
        '<label class="field"><span>支払者の名称（メモ用）</span><input class="input" type="text" data-slip="payer" value="' + esc(s.payer || '') + '" placeholder="例：株式会社◯◯"></label>' +
        '<div class="grid2">' +
        '<label class="field"><span>① 支払金額</span><input class="input" type="number" inputmode="numeric" min="0" data-slip="payment" value="' + esc(s.payment || '') + '" placeholder="0"></label>' +
        '<label class="field"><span>② 給与所得控除後の金額</span><input class="input" type="number" inputmode="numeric" min="0" data-slip="afterDeduction" value="' + esc(s.afterDeduction || '') + '" placeholder="0"></label>' +
        '</div>' +
        '<div class="grid2">' +
        '<label class="field"><span>③ 所得控除の額の合計額</span><input class="input" type="number" inputmode="numeric" min="0" data-slip="deductions" value="' + esc(s.deductions || '') + '" placeholder="0"></label>' +
        '<label class="field"><span>④ 源泉徴収税額</span><input class="input" type="number" inputmode="numeric" min="0" data-slip="taxWithheld" value="' + esc(s.taxWithheld || '') + '" placeholder="0"></label>' +
        '</div>' +
        '</div>';
    }).join('');
  }

  function renderRiskResults() {
    var sum = $('riskSummary');
    var listBox = $('riskList');
    if (!sum || !listBox) return;

    var r = analyzeRisks();
    var sim = computeSim();

    $('riskCount').textContent = r.total + '件を診断';
    var judgment = (r.redCount > 0 || r.structural.some(function (s) { return s.level === 'red'; })) ? '要修正'
      : (r.amberCount > 0 || r.structural.length ? '要確認' : '問題なし');

    sum.innerHTML =
      '<div class="risk-chips">' +
      '<div class="rchip red"><div class="n">' + r.redCount + '件</div><div class="t">🔴 すぐ外す・直す</div><div class="t">' + yen(r.redAmt) + '</div></div>' +
      '<div class="rchip amber"><div class="n">' + r.amberCount + '件</div><div class="t">🟡 要確認・按分</div><div class="t">' + yen(r.amberAmt) + '</div></div>' +
      '<div class="rchip ok' + (judgment !== '問題なし' ? ' warn' : '') + '"><div class="n">' + (judgment === '問題なし' ? '✅ 問題なし' : (judgment === '要修正' ? '⚠️ 要修正' : '🔍 要確認')) + '</div>' +
      '<div class="t">' + (r.flags.length + r.structural.length) + '件の指摘｜経費 ' + yen(sim.expenses + sim.dokujuu) + '／売上 ' + yen(sim.sales) + '</div></div>' +
      '</div>';

    var html = '';
    r.structural.forEach(function (s) {
      html += '<div class="flag struct">' +
        '<div class="flag-body">' +
        '<div class="flag-top"><span class="flag-title">' + (s.level === 'red' ? '⚠️ ' : '🔍 ') + esc(s.title) + '</span></div>' +
        '<div class="flag-why">' + esc(s.why) + '</div>' +
        '</div></div>';
    });

    r.flags.forEach(function (f) {
      html += '<div class="flag ' + f.level + '">' +
        '<div class="flag-body">' +
        '<div class="flag-top"><span class="flag-title">' + (f.level === 'red' ? '🔴 ' : '🟡 ') + esc(f.title) + '</span>' +
        '<span class="flag-amt">' + yen(f.amount) + '</span></div>' +
        '<div class="flag-vendor">' + esc(f.vendor) + '</div>' +
        '<div class="flag-meta">' + esc(f.date || '日付なし') + (f.category ? ' · ' + esc(f.category) : '') + '</div>' +
        '<div class="flag-why">' + esc(f.why) + '</div>' +
        '<div class="flag-act">' +
        '<button class="mini-btn" data-act="edit" data-id="' + esc(f.entryId) + '">✏️ 直す</button>' +
        '<button class="mini-btn danger" data-act="del" data-id="' + esc(f.entryId) + '">🗑️ 経費から外す</button>' +
        '</div></div></div>';
    });

    if (!html) {
      html = '<div class="empty">指摘はありません。<br>赤・黄の項目が出たときに、ここで直していけば調査対策になります。</div>';
    }
    listBox.innerHTML = html;

    renderSim(sim);
  }

  function renderSim(sim) {
    var box = $('simResult');
    if (!box) return;

    var rows = [];
    rows.push(['事業売上（収入）', yenSigned(sim.sales), '']);
    rows.push(['必要経費' + (sim.dokujuu ? '（専従者給与別）' : ''), '−' + yen(sim.expenses), '']);
    if (sim.dokujuu) rows.push(['専従者給与', '−' + yen(sim.dokujuu), '']);
    rows.push(['青色申告特別控除', '−' + yen(sim.blue), 'blue']);
    rows.push(['＝ 事業所得', yenSigned(sim.biz), sim.biz < 0 ? 'neg' : '']);
    if (sim.slipCount) rows.push(['＋ 給与所得', yenSigned(sim.salary), '']);
    rows.push(['＝ 合計所得金額', yenSigned(sim.totalIncome), '']);
    rows.push(['− 所得控除の額の合計', '−' + yen(sim.deductions) + (sim.deductionsManual ? '（手入力）' : '（源泉徴収票③）'), '']);
    rows.push(['＝ 課税される所得金額', yenSigned(sim.taxable), 'blue']);
    rows.push(['所得税', yen(sim.incomeTax), '']);
    rows.push(['復興特別所得税（2.1%）', yen(sim.fukko), '']);
    rows.push(['所得税及び復興特別所得税', yen(sim.totalTax), '']);
    if (sim.slipCount) rows.push(['− 源泉徴収税額（前払い済み）', '−' + yen(sim.withheld), '']);

    var html = '<div>';
    rows.forEach(function (r, i) {
      var isBig = /^＝/.test(r[0]);
      html += '<div class="sim-row' + (isBig ? ' big' : '') + '"><span class="k">' + esc(r[0]) + '</span><span class="v ' + r[2] + '">' + r[1] + '</span></div>';
    });
    html += '</div>';

    if (sim.diff > 0) {
      html += '<div class="sim-box pay"><div class="lbl">納める税金の見込み</div><div class="amt">' + yen(sim.diff) + '</div></div>';
    } else if (sim.diff < 0) {
      html += '<div class="sim-box refund"><div class="lbl">還付される見込み</div><div class="amt">' + yen(-sim.diff) + '</div></div>';
    } else {
      html += '<div class="sim-box zero"><div class="lbl">納付も還付もなし</div><div class="amt">¥0</div></div>';
    }

    var notes = '';
    if (sim.sales === 0) {
      notes += '<div class="warn-box">売上が未入力です。売上を入れると課税所得と還付額が計算できます。</div>';
    } else if (sim.biz < 0 && sim.salary > 0) {
      notes += '<div class="info-box">事業所得が赤字（' + yenSigned(sim.biz) + '）なので、給与所得と損益通算して課税所得を下げています。この通算が認められるのは「事業所得」と認められた場合だけです（雑所得は通算できません）。</div>';
    }
    notes += '<div class="hint-ok">この試算は所得税の目安です。住民税は別に、課税所得の約10%＋均等割（自治体により約5,000円）が翌年にかかります（この年度の概算：' + yen(sim.residentTax) + '）。医療費控除・ふるさと納税・住宅ローン控除などは入っていないため、実際の税額は変わります。</div>';
    box.innerHTML = html + notes;
  }

  /* ============================================================
     CSV / Excel 取り込み
     ============================================================ */
  var HEAD_KEYS = {
    date: ['日付', '取引日', '年月日', '取引年月日', 'ご利用日', '利用日', '発生日', '支払日', 'date', 'transactiondate'],
    out: ['出金', '出金金額', '支出', '支出金額', 'お引出し', 'お引き出し', 'お引出し金額', 'お引き出し金額', '引出金額', '支払金額', '利用金額', '借方金額', 'debit', 'payment', 'amountout'],
    in: ['入金', '入金金額', '預り金額', '預かり金額', 'お預り', 'お預かり', 'お預り金額', 'お預かり金額', '収入金額', '貸方金額', 'credit', 'deposit', 'amountin'],
    amount: ['金額', '金額円', '税込金額', '合計金額', '取引金額', 'amount', 'value', '税込'],
    vendor: ['支払先', '取引先', '店名', '加盟店', '加盟店名', '相手先', '摘要', '内容', '品名', 'payee', 'vendor', 'merchant', 'description', 'memo'],
    memo: ['メモ', '備考', '摘要内容', 'comment', 'note', 'remarks'],
    category: ['勘定科目', '借方科目', '科目', '分類', 'category', 'account', '借方'],
    creditCat: ['貸方科目', '貸方', 'creditaccount']
  };

  function normHeader(s) {
    return String(s == null ? '' : s).toLowerCase()
      .replace(/[\s\u3000]/g, '')
      .replace(/[（）()【】\[\]「」:：]/g, '');
  }

  function headerField(cell) {
    var c = normHeader(cell);
    if (!c) return '';
    var order = ['date', 'in', 'out', 'amount', 'creditCat', 'category', 'vendor', 'memo'];
    for (var i = 0; i < order.length; i += 1) {
      var k = order[i];
      var arr = HEAD_KEYS[k];
      for (var j = 0; j < arr.length; j += 1) {
        if (c === normHeader(arr[j])) return k;
      }
    }
    for (var i2 = 0; i2 < order.length; i2 += 1) {
      var k2 = order[i2];
      var arr2 = HEAD_KEYS[k2];
      for (var j2 = 0; j2 < arr2.length; j2 += 1) {
        if (c.indexOf(normHeader(arr2[j2])) >= 0) return k2;
      }
    }
    return '';
  }

  function parseAnyDate(v) {
    if (v == null || v === '') return '';
    if (typeof v === 'number' && v > 20000 && v < 60000) {
      // Excelのシリアル値
      var d = new Date(Math.round((v - 25569) * 86400 * 1000));
      return d.getUTCFullYear() + '-' + p2(d.getUTCMonth() + 1) + '-' + p2(d.getUTCDate());
    }
    var s = String(v).trim();
    var m = s.match(/(20\d{2})\s*[年\/.\-]\s*(\d{1,2})\s*[月\/.\-]\s*(\d{1,2})/);
    if (m) return m[1] + '-' + p2(+m[2]) + '-' + p2(+m[3]);
    m = s.match(/(\d{1,2})\s*[\/.\-]\s*(\d{1,2})\s*[\/.\-]\s*(\d{2,4})/);
    if (m) {
      var yy = +m[3]; if (yy < 100) yy += 2000;
      return yy + '-' + p2(+m[1]) + '-' + p2(+m[2]);
    }
    m = s.match(/^(\d{1,2})\s*[\/.\-]\s*(\d{1,2})$/);
    if (m) return state.year + '-' + p2(+m[1]) + '-' + p2(+m[2]);
    m = s.match(/(\d{4})(\d{2})(\d{2})/);
    if (m) return m[1] + '-' + m[2] + '-' + m[3];
    return '';
  }
  function p2(n) { return (n < 10 ? '0' : '') + n; }

  function mapCategory(raw) {
    var s = String(raw == null ? '' : raw);
    if (!s) return '';
    for (var i = 0; i < TAX_CATEGORIES.length; i += 1) {
      if (s.indexOf(TAX_CATEGORIES[i]) >= 0) return TAX_CATEGORIES[i];
    }
    var alias = [
      ['消耗品', '消耗品費'], ['事務用品', '消耗品費'], ['備品', '消耗品費'], ['工具器具', '消耗品費'], ['文房具', '消耗品費'], ['雑貨', '消耗品費'],
      ['旅費', '旅費交通費'], ['交通費', '旅費交通費'], ['交通', '旅費交通費'], ['車両', '旅費交通費'],
      ['通信', '通信費'], ['電話', '通信費'], ['インターネット', '通信費'], ['サーバー', '通信費'],
      ['交際', '接待交際費'], ['会議', '会議費'], ['飲食', '会議費'],
      ['光熱', '水道光熱費'], ['電気', '水道光熱費'], ['ガス', '水道光熱費'], ['水道', '水道光熱費'],
      ['図書', '新聞図書費'], ['新聞', '新聞図書費'],
      ['手数料', '支払手数料'], ['振込', '支払手数料'],
      ['外注', '外注工賃'], ['家賃', '地代家賃'], ['賃借', '地代家賃'], ['地代', '地代家賃'],
      ['租税', '租税公課'], ['広告', '広告宣伝費'], ['保険', '損害保険料'], ['修繕', '修繕費'],
      ['減価', '減価償却費'], ['福利', '福利厚生費'], ['給料', '給料賃金'], ['人件費', '給料賃金'],
      ['利子', '利子割引料'], ['貸倒', '貸倒金']
    ];
    for (var j = 0; j < alias.length; j += 1) {
      if (s.indexOf(alias[j][0]) >= 0) return alias[j][1];
    }
    return '雑費';
  }

  function readImportFile(file) {
    var reader = new FileReader();
    reader.onload = function (ev) {
      try {
        var wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array', cellDates: false });
        var name = wb.SheetNames[0];
        var ws = wb.Sheets[name];
        var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
        applyImport(rows, file.name);
      } catch (e) {
        toast('読み込めませんでした：' + ((e && e.message) || e), true);
      }
    };
    reader.onerror = function () { toast('ファイルを読めませんでした', true); };
    reader.readAsArrayBuffer(file);
  }

  function applyImport(rows, filename) {
    if (!rows || !rows.length) { toast('データが空でした', true); return; }

    var headerIdx = -1, bestScore = 0;
    for (var i = 0; i < Math.min(rows.length, 15); i += 1) {
      var cells = (rows[i] || []).map(headerField);
      var uniq = {};
      cells.forEach(function (c) { if (c) uniq[c] = 1; });
      var score = Object.keys(uniq).length;
      if (score > bestScore) { bestScore = score; headerIdx = i; }
    }
    if (headerIdx < 0 || bestScore < 2) {
      toast('見出し行が見つかりませんでした（日付・金額などの列名が必要です）', true);
      return;
    }

    var cols = {};
    (rows[headerIdx] || []).forEach(function (c, idx) {
      var f = headerField(c);
      if (f && cols[f] === undefined) cols[f] = idx;
    });

    var added = [], skipped = 0, noAmount = 0, income = 0, incomeCount = 0, otherYear = {};
    var byYear = {};

    for (var r = headerIdx + 1; r < rows.length; r += 1) {
      var row = rows[r] || [];
      if (!row.length) continue;

      var date = parseAnyDate(cols.date !== undefined ? row[cols.date] : '');
      var vendor = String((cols.vendor !== undefined ? row[cols.vendor] : '') || '').trim();
      var memo = String((cols.memo !== undefined ? row[cols.memo] : '') || '').trim();
      var catRaw = String((cols.category !== undefined ? row[cols.category] : '') || '').trim();
      var creditCat = String((cols.creditCat !== undefined ? row[cols.creditCat] : '') || '').trim();

      var outV = cols.out !== undefined ? num(row[cols.out]) : 0;
      var inV = cols.in !== undefined ? num(row[cols.in]) : 0;
      var genV = cols.amount !== undefined ? num(row[cols.amount]) : 0;

      var isIncome = false;
      var amount = 0;
      if (inV > 0) { amount = inV; isIncome = true; }
      else if (outV > 0) { amount = outV; }
      else if (genV !== 0) { amount = genV; }

      if (/売上|売掛/.test(catRaw) || /売上|売掛/.test(creditCat)) isIncome = true;
      if (!outV && /(売上|売掛|入金|振込入金)/.test(vendor + ' ' + memo) && !/返金|払戻/.test(vendor + ' ' + memo)) isIncome = true;

      if (!amount) { noAmount += 1; continue; }
      amount = Math.abs(amount);

      if (isIncome) { income += amount; incomeCount += 1; continue; }

      var cat = mapCategory(catRaw || vendor + ' ' + memo);
      var y = yearOf(date) || state.year;
      var entry = {
        id: uid(), date: date || '', vendor: vendor, category: cat, currency: 'JPY',
        amount: amount, taxType: '', invoiceNo: '', memo: memo.slice(0, 120),
        thumb: '', source: 'import', importFile: filename,
        createdAt: new Date().toISOString()
      };

      if (!byYear[y]) byYear[y] = [];
      byYear[y].push(entry);
      if (y !== state.year) otherYear[y] = (otherYear[y] || 0) + 1;
    }

    var totalAdded = 0, totalAmt = 0;
    Object.keys(byYear).forEach(function (yy) {
      var list = taxList(+yy);
      var index = {};
      list.forEach(function (e) { index[(e.date || '') + '|' + Math.round(num(e.amount)) + '|' + vendorKey(e.vendor)] = 1; });
      byYear[yy].forEach(function (e) {
        var k = (e.date || '') + '|' + Math.round(e.amount) + '|' + vendorKey(e.vendor);
        if (index[k]) { skipped += 1; return; }
        index[k] = 1;
        list.push(e);
        totalAdded += 1;
        totalAmt += e.amount;
      });
    });

    var salesAdded = 0;
    if (incomeCount) {
      var ok = window.confirm('収入（売上）の行が ' + incomeCount + '件・' + yen(income) + ' あります。\n' + state.year + '年分の「年間の事業売上」に加算しますか？\n（今の売上：' + yen(salesOf()) + '）');
      if (ok) {
        state.sales[ykey()] = salesOf() + income;
        salesAdded = income;
      }
    }

    lastImport = {
      file: filename, total: totalAdded, amount: totalAmt, skipped: skipped,
      income: income, incomeCount: incomeCount, salesAdded: salesAdded,
      otherYear: otherYear, noAmount: noAmount, at: new Date().toISOString()
    };

    persist();
    renderAll();

    if (totalAdded) {
      toast(totalAdded + '件・' + yen(totalAmt) + ' を取り込みました' + (skipped ? '（重複' + skipped + '件は除外）' : ''));
      signal('経費帳に' + totalAdded + '件（' + yen(totalAmt) + '）を取り込み、税務診断を更新しました。');
    } else if (skipped) {
      toast('すべて登録済みでした（重複' + skipped + '件）');
    } else {
      toast('取り込める行がありませんでした', true);
    }
  }

  function importPreviewHTML(l) {
    var parts = [];
    parts.push('<div class="info-box" style="margin-top:10px;"><b>取り込み結果</b><br>' +
      esc(l.file) + '：' + l.total + '件・' + yen(l.amount) +
      (l.skipped ? '／重複除外 ' + l.skipped + '件' : '') +
      (l.incomeCount ? '<br>収入 ' + l.incomeCount + '件・' + yen(l.income) + (l.salesAdded ? '（売上に加算済み）' : '（売上には未加算）') : '') +
      (l.noAmount ? '<br>金額が空のためスキップ ' + l.noAmount + '行' : '') +
      '</div>');
    var oy = Object.keys(l.otherYear || {});
    if (oy.length) {
      parts.push('<div class="rate-note">' + oy.map(function (k) { return k + '年分に' + l.otherYear[k] + '件'; }).join('・') + 'を振り分けました。</div>');
    }
    return parts.join('');
  }

  /* ============================================================
     e-Taxガイド
     ============================================================ */
  function guideSteps() {
    var y = state.year;
    var yj = reiwa(y);
    var js = '令和' + yj + '年分';
    return [
      {
        id: 'prep', ico: '🧰', title: '事前準備', sub: 'マイナンバーカード・源泉徴収票・1年分の集計',
        items: [
          ['マイナンバーカードの電子証明書の有効期限を確認', '期限が切れていると送信できません。更新は市区町村の窓口で（無料）。'],
          ['読み取り方法を決める', 'NFC対応のスマホなら、マイナポータルアプリまたは作成コーナーの「スマホで読み取り」でカードリーダーは不要です。'],
          ['利用者識別番号を控えておく', '初回は作成コーナーの「利用者識別番号の発行」から取得します。2回目以降は入力するだけです。'],
          ['青色申告承認申請書が出ているか確認', '提出は開業から2か月以内、または前年3月15日まで。出ていないと青色申告（65万円控除）は使えません。'],
          ['源泉徴収票を用意（前職分も）', '①支払金額 ②給与所得控除後の金額 ③所得控除の額の合計額 ④源泉徴収税額 の4つを使います。'],
          ['1年分の売上と経費を締める', 'このアプリの「Excel出力」で書き出し、税務診断タブの赤・黄を潰しておきます。'],
          ['マイナポータル連携（任意だけどおすすめ）', '社会保険料・医療費・ふるさと納税の控除証明を自動で取得できます。一度連携すれば翌年も使えます。']
        ],
        tip: '12月のうちに領収書を読み終えておくと、2月は「画面に数字を写すだけ」で終わります。'
      },
      {
        id: 'step1', ico: '1️⃣', title: 'ステップ1：青色申告決算書をつくる', sub: '売上 → 経費 → 特別控除 → 所得金額',
        items: [
          ['「作成開始」→「所得税」→「' + js + '」→「青色申告決算書から作成」', '決算書から作ると、所得金額が確定申告書へ自動で引き継がれます。'],
          ['売上（収入）金額を入れる', '売上先が2件までなら直接入力。3件以上は「売上（収入）金額の内訳」を開いて入力します。'],
          ['売上原価を入れる', '期首商品棚卸高・仕入金額・期末商品棚卸高。モノを仕入れない（制作・サービス）なら0のままで大丈夫です。'],
          ['経費を科目ごとに入れる', 'e-Taxガイドの転記ボードの金額をそのまま写します。給料賃金・地代家賃・専従者給与は内訳の入力が求められます。'],
          ['青色申告特別控除額を選ぶ', '65万円／55万円／10万円。65万円は「e-Tax送信＋電子帳簿保存」が条件です。迷ったら55万円が安全です。'],
          ['所得金額が赤字でも0にしない', 'マイナスのまま進みます。ここを0にすると給与所得との損益通算ができなくなります。'],
          ['2ページ目を埋める', '月別の売上・仕入、減価償却費の計算、地代家賃の内訳、専従者給与の内訳。']
        ],
        tip: '決算書の所得金額と、確定申告書の「事業所得」が一致しているか最後に必ず見比べてください。'
      },
      {
        id: 'step2', ico: '2️⃣', title: 'ステップ2：所得税の確定申告書をつくる', sub: '給与所得＋事業所得＋所得控除',
        items: [
          ['給与所得を入れる', '源泉徴収票の「支払金額」と「所得控除後の金額」。2枚目（前職）は「前職分」の欄に入れます。'],
          ['社会保険料控除を入れる', '国民年金・国民健康保険・任意継続の保険料。マイナポータル連携していれば自動で入ります。'],
          ['その他の控除を入れる', '生命保険料・地震保険料・医療費・寄附金（ふるさと納税）・小規模企業共済等掛金。'],
          ['税額を確認する', '課税所得→所得税→復興特別所得税は自動計算。ここで還付か納付かが決まります。'],
          ['住民税・事業税に関する事項', '給与以外の所得があると、住民税の徴収方法（特別徴収／普通徴収）を選べます。'],
          ['事業所得が赤字なら通算される', '第一表の「事業所得（営業等）」にマイナスで入り、給与所得と相殺されます。これが損益通算です。']
        ],
        tip: '還付になる年でも、申告しないと還付は受けられません（還付申告は' + (y + 1) + '年1月から提出できます）。'
      },
      {
        id: 'step3', ico: '3️⃣', title: 'ステップ3：送信・控えの保存・納付', sub: 'マイナンバーカードで署名 → PDF保存 → 納付',
        items: [
          ['「入力内容の確認」で全項目を見る', '収入・所得・控除・税額。源泉徴収票の数字と1つずつ照合します。'],
          ['「申告書を送信」→カードで署名', 'マイナンバーカードを読み取り、利用者証明用電子証明書のパスワード（4桁）を入力します。'],
          ['受付番号と受付日時を控える', '送信結果の画面はPDFで保存。あとで「送信結果・照会」からも確認できます。'],
          ['帳票一覧からPDFを保存', '申告書・青色申告決算書をダウンロードして保存。所得税の帳簿書類は7年保存（欠損金の繰越は9年）。'],
          ['納付する', '振替納税（口座振替の登録は3月中旬まで）・e-Taxダイレクト・コンビニ納付（QR）・クレジットカード納付（手数料がかかります）。'],
          ['還付を待つ', '登録口座へ通常3〜4週間で振込。「還付金の処理状況」で確認できます。']
        ],
        tip: '送信は24時間できますが、期限ぎりぎりはサーバーが混みます。' + deadlineOf(y).end + 'の3日前までに送っておくと安心です。'
      }
    ];
  }

  function renderGuide() {
    var stepsBox = $('guideSteps');
    if (!stepsBox) return;

    // 期限
    var dl = deadlineOf(state.year);
    var db = $('deadlineBox');
    if (db) {
      db.innerHTML = '<div class="deadline"><div class="dl-ico">🗓️</div><div class="dl-main">' +
        '<div class="dl-title">' + yearLabel(state.year) + '分の申告期限</div>' +
        '<div class="dl-sub">' + dl.start + ' 〜 ' + dl.end + '<br>還付申告なら' + (state.year + 1) + '年1月から提出できます。期限を過ぎると無申告加算税・延滞税の対象になります。</div>' +
        '</div></div>';
    }

    var checks = checksOf();
    var steps = guideSteps();
    var doneCount = steps.filter(function (s) { return checks[s.id]; }).length;
    var h = '<div class="card" style="padding:11px 14px;"><div class="card-title" style="margin-bottom:6px;"><span>✅ 進捗</span><span class="count">' + doneCount + ' / ' + steps.length + 'ステップ</span></div>' +
      '<div class="bar" style="margin-top:0;"><i style="width:' + Math.round(doneCount / steps.length * 100) + '%"></i></div></div>';

    steps.forEach(function (s) {
      h += '<details class="step">' +
        '<summary>' +
        '<span class="step-ico">' + s.ico + '</span>' +
        '<span class="step-main"><span class="step-title">' + esc(s.title) + '</span><span class="step-sub">' + esc(s.sub) + '</span></span>' +
        '<button type="button" class="chk' + (checks[s.id] ? ' on' : '') + '" data-step="' + s.id + '" aria-label="完了にする">✓</button>' +
        '</summary>' +
        '<div class="step-body">' +
        s.items.map(function (it, i) {
          return '<div class="step-item"><span class="no">' + (i + 1) + '</span><span class="tx"><b>' + esc(it[0]) + '</b><span>' + esc(it[1]) + '</span></span></div>';
        }).join('') +
        '<div class="step-tip">💡 ' + esc(s.tip) + '</div>' +
        '</div></details>';
    });
    stepsBox.innerHTML = h;

    renderBoard();
  }

  function renderBoard() {
    var box = $('board');
    if (!box) return;
    var y = state.year;
    var sim = computeSim();
    var tot = totalsForYear();

    var sec1 = [];
    sec1.push({
      label: '売上（収入）金額',
      src: '税務診断タブの「年間の事業売上」',
      v: sim.sales
    });
    FORM_CATEGORIES.forEach(function (c) {
      if (c === '専従者給与') return;
      var v = tot.byForm[c] || 0;
      if (!v) return;
      var raws = Object.keys(tot.byRaw).filter(function (raw) { return formCategoryOf(raw) === c; });
      var n = raws.reduce(function (s, raw) { return s + (tot.rawCount[raw] || 0); }, 0);
      var src = '経費帳 ' + n + '件：' + raws.join('・');
      if (raws.length > 1) src += '（決算書では1欄にまとめます）';
      sec1.push({ label: c, src: src, v: v });
    });
    sec1.push({ label: '経費計', src: '経費の合計（専従者給与を除く）', v: sim.expenses, total: true });
    if (sim.dokujuu) sec1.push({ label: '専従者給与（事業専従者控除）', src: '経費計の下の別欄', v: sim.dokujuu });
    sec1.push({ label: '青色申告特別控除前の所得金額', src: '自動計算（売上−経費−専従者給与）', v: sim.sales - sim.expenses - sim.dokujuu });
    sec1.push({ label: '青色申告特別控除額', src: '選択中の控除額（65万／55万／10万）', v: sim.blue });
    sec1.push({ label: '所得金額', src: '自動計算。赤字ならマイナスのまま', v: sim.biz, total: true });

    var sec2 = [];
    if (sim.slipCount) {
      sec2.push({ label: '収入金額等「給与」', src: '源泉徴収票①の合計', v: sim.pay });
      sec2.push({ label: '所得金額等「給与所得」', src: '源泉徴収票②（合算後）', v: sim.salary });
    }
    sec2.push({ label: '所得金額等「事業所得（営業等）」', src: '決算書の所得金額と同じ数字', v: sim.biz });
    sec2.push({ label: '合計所得金額', src: '給与所得＋事業所得', v: sim.totalIncome, total: true });
    sec2.push({ label: '所得控除の額の合計', src: sim.deductionsManual ? '手入力の合計' : '源泉徴収票③', v: sim.deductions });
    sec2.push({ label: '課税される所得金額', src: '自動計算（1,000円未満切り捨て）', v: sim.taxable, total: true });
    sec2.push({ label: '所得税及び復興特別所得税', src: '税率を掛けて自動計算', v: sim.totalTax });
    if (sim.slipCount) sec2.push({ label: '源泉徴収税額', src: '源泉徴収票④の合計', v: sim.withheld });
    sec2.push({
      label: sim.diff >= 0 ? '納める税金' : '還付される税金',
      src: '税額−源泉徴収税額', v: Math.abs(sim.diff), total: true
    });

    function secHTML(title, rows) {
      return '<div class="board-sec"><h4>' + esc(title) + '</h4>' + rows.map(function (r) {
        return '<div class="board-row' + (r.total ? ' board-total' : '') + '">' +
          '<div class="bl"><b>' + esc(r.label) + '</b><span>' + esc(r.src || '') + '</span></div>' +
          '<div class="bv' + (r.v < 0 ? ' neg' : '') + '">' + (r.v < 0 ? '▲' : '') + Math.abs(Math.round(r.v)).toLocaleString('ja-JP') + '</div>' +
          '<button class="copy-btn" data-copy="' + Math.round(r.v) + '" aria-label="コピー">⧉</button>' +
          '</div>';
      }).join('') + '</div>';
    }

    box.innerHTML = secHTML('青色申告決算書 1ページ目', sec1) + secHTML('確定申告書B 第一表', sec2) +
      '<div class="rate-note">コピーされるのは数字だけです（例：41683）。金額欄にそのまま貼り付けられます。' +
      (Object.keys(tot.byForm).some(function (k) { return ['雑費', '旅費交通費'].indexOf(k) >= 0; }) ? '会議費・新聞図書費・支払手数料は決算書では「雑費」に、車両費は「旅費交通費」にまとめて表示しています。' : '') +
      '</div>';

    var bc = $('boardCount');
    if (bc) bc.textContent = (sec1.length + sec2.length) + '項目';
  }

  /* ============================================================
     Excel / TSV
     ============================================================ */
  function maskSheetName(n) { return n.replace(/[\\\/\?\*\[\]:]/g, ' ').slice(0, 31); }
  function styleWidths(ws, widths) { ws['!cols'] = widths.map(function (w) { return { wch: w }; }); }

  function buildTaxWorkbook() {
    var list = sortedEntries(taxList());
    var rows = [['No', '日付', '支払先・店名', '勘定科目', '内容・メモ', '金額(税込)', '税区分', 'インボイス番号']];
    var total = 0;
    list.forEach(function (e, i) {
      var a = yenOf(e);
      total += a;
      rows.push([i + 1, e.date || '', e.vendor || '', e.category || '', e.memo || '', a, e.taxType || '', e.invoiceNo || '']);
    });
    var lastDataRow = rows.length;
    rows.push(['', '', '合計', '', '', { t: 'n', f: 'SUM(F2:F' + lastDataRow + ')', v: total }, '', '']);

    var ws = XLSX.utils.aoa_to_sheet(rows);
    styleWidths(ws, [5, 12, 26, 14, 32, 12, 10, 18]);
    for (var r = 2; r <= lastDataRow + 1; r += 1) {
      var cell = ws['F' + r];
      if (cell) cell.z = '#,##0';
    }

    var map = {};
    list.forEach(function (e) {
      var k = e.category || '未分類';
      if (!map[k]) map[k] = { n: 0, sum: 0 };
      map[k].n += 1;
      map[k].sum += yenOf(e);
    });
    var keys = Object.keys(map).sort(function (a, b) { return map[b].sum - map[a].sum; });
    var rows2 = [['勘定科目', '件数', '合計金額(円)', '構成比']];
    keys.forEach(function (k) {
      rows2.push([k, map[k].n, map[k].sum, { t: 'n', v: total ? map[k].sum / total : 0, z: '0.0%' }]);
    });
    var last2 = rows2.length;
    rows2.push(['合計', { t: 'n', f: 'SUM(B2:B' + last2 + ')', v: list.length }, { t: 'n', f: 'SUM(C2:C' + last2 + ')', v: total }, { t: 'n', v: 1, z: '0.0%' }]);

    var ws2 = XLSX.utils.aoa_to_sheet(rows2);
    styleWidths(ws2, [18, 8, 16, 10]);
    for (var r2 = 2; r2 <= last2 + 1; r2 += 1) {
      var c2 = ws2['C' + r2];
      if (c2) c2.z = '#,##0';
    }

    /* 税務チェック */
    var rk = analyzeRisks();
    var rows3 = [['判定', 'リスク項目', '対象', '日付', '金額', '理由・対応']];
    rk.structural.forEach(function (s) {
      rows3.push([s.level === 'red' ? '赤' : '黄', s.title, '（全体）', '', '', s.why]);
    });
    rk.flags.forEach(function (f) {
      rows3.push([f.level === 'red' ? '赤' : '黄', f.title, f.vendor + (f.category ? ' / ' + f.category : ''), f.date, f.amount, f.why]);
    });
    if (rows3.length === 1) rows3.push(['—', '指摘なし', '', '', '', '']);
    var ws3 = XLSX.utils.aoa_to_sheet(rows3);
    styleWidths(ws3, [6, 26, 26, 12, 12, 60]);
    for (var r3 = 2; r3 <= rows3.length; r3 += 1) {
      if (ws3['E' + r3]) ws3['E' + r3].z = '#,##0';
    }

    /* 確定申告サマリー */
    var sim = computeSim();
    var rows4 = [['項目', '金額', '備考']];
    rows4.push(['年度', state.year + '年分（令和' + reiwa(state.year) + '年分）', '']);
    rows4.push(['事業売上（収入）', sim.sales, '']);
    rows4.push(['必要経費', sim.expenses, '']);
    if (sim.dokujuu) rows4.push(['専従者給与', sim.dokujuu, '']);
    rows4.push(['青色申告特別控除', sim.blue, '']);
    rows4.push(['事業所得', sim.biz, sim.biz < 0 ? '赤字（給与所得と損益通算）' : '']);
    rows4.push(['給与所得（源泉徴収票②）', sim.salary, sim.slipCount + '枚']);
    rows4.push(['合計所得金額', sim.totalIncome, '']);
    rows4.push(['所得控除の額の合計', sim.deductions, sim.deductionsManual ? '手入力' : '源泉徴収票③から自動']);
    rows4.push(['課税される所得金額', sim.taxable, '']);
    rows4.push(['所得税', sim.incomeTax, '']);
    rows4.push(['復興特別所得税', sim.fukko, '']);
    rows4.push(['所得税及び復興特別所得税', sim.totalTax, '']);
    rows4.push(['源泉徴収税額', sim.withheld, '']);
    rows4.push([sim.diff >= 0 ? '納める税金（見込み）' : '還付される税金（見込み）', Math.abs(sim.diff), '目安です']);
    var ws4 = XLSX.utils.aoa_to_sheet(rows4);
    styleWidths(ws4, [30, 16, 34]);
    for (var r4 = 2; r4 <= rows4.length; r4 += 1) {
      if (ws4['B' + r4] && typeof ws4['B' + r4].v === 'number') ws4['B' + r4].z = '#,##0';
    }

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, maskSheetName('領収書明細'));
    XLSX.utils.book_append_sheet(wb, ws2, maskSheetName('科目別集計'));
    XLSX.utils.book_append_sheet(wb, ws3, maskSheetName('税務チェック'));
    XLSX.utils.book_append_sheet(wb, ws4, maskSheetName('確定申告サマリー'));
    return {
      wb: wb,
      filename: '領収書一覧_' + state.year + '_' + stamp() + '.xlsx',
      total: total, count: list.length
    };
  }

  function buildTripWorkbook() {
    var list = sortedEntries(state.trip);
    var rows = [['No', '日付', '区分', '内容・支払先', '通貨', '金額（現地）', 'レート(円/₩)', '円換算', 'メモ']];
    var total = 0;
    list.forEach(function (e, i) {
      var y = yenOf(e);
      total += y;
      rows.push([
        i + 1, e.date || '', e.category || 'その他', e.vendor || '',
        e.currency === 'KRW' ? 'KRW' : 'JPY',
        num(e.amount),
        e.currency === 'KRW' ? state.rate : '',
        y,
        e.memo || ''
      ]);
    });
    var lastDataRow = rows.length;
    rows.push(['', '', '総合計', '', '', '', '', { t: 'n', f: 'SUM(H2:H' + lastDataRow + ')', v: total }, '']);

    var ws = XLSX.utils.aoa_to_sheet(rows);
    styleWidths(ws, [5, 12, 12, 30, 7, 13, 12, 12, 34]);
    for (var r = 2; r <= lastDataRow + 1; r += 1) {
      if (ws['F' + r]) ws['F' + r].z = '#,##0';
      if (ws['H' + r]) ws['H' + r].z = '#,##0';
    }

    var order = ['航空券', 'ホテル', '交通費', '食費・買い物', 'その他'];
    var map = {};
    order.forEach(function (k) { map[k] = { n: 0, sum: 0 }; });
    list.forEach(function (e) {
      var k = order.indexOf(e.category) >= 0 ? e.category : 'その他';
      map[k].n += 1;
      map[k].sum += yenOf(e);
    });
    var rows2 = [['区分', '件数', '円換算合計', '構成比']];
    order.forEach(function (k) {
      if (!map[k].n) return;
      rows2.push([k, map[k].n, map[k].sum, { t: 'n', v: total ? map[k].sum / total : 0, z: '0.0%' }]);
    });
    var last2 = rows2.length;
    rows2.push(['合計', { t: 'n', f: 'SUM(B2:B' + last2 + ')', v: list.length }, { t: 'n', f: 'SUM(C2:C' + last2 + ')', v: total }, { t: 'n', v: 1, z: '0.0%' }]);

    var jpySum = 0, krwSum = 0, krwYen = 0;
    list.forEach(function (e) {
      if (e.currency === 'KRW') { krwSum += num(e.amount); krwYen += yenOf(e); } else { jpySum += yenOf(e); }
    });

    rows2.push([]);
    rows2.push(['【旅行サマリー】', '', '', '']);
    rows2.push(['ソウル旅行 総合計（円）', '', total, '']);
    rows2.push(['うち 円建て', '', jpySum, '']);
    rows2.push(['うち ウォン建て（円換算）', krwSum + ' KRW', krwYen, '']);
    rows2.push(['適用レート（1₩=）', '', state.rate, '']);
    rows2.push(['期間', '2026-11-14 〜 2026-11-15（1泊2日）', '', '']);
    rows2.push(['備考', '航空券・ホテル・交通費は2026-11-16作成の費用内訳PDF、それ以外は現地レシートの読み取り分。', '', '']);
    rows2.push(['', '免税返金・実費ベース。実際のカード請求額は決済時レートにより変動。', '', '']);

    var ws2 = XLSX.utils.aoa_to_sheet(rows2);
    styleWidths(ws2, [26, 30, 16, 10]);
    for (var r2 = 2; r2 <= last2 + 1; r2 += 1) {
      if (ws2['C' + r2]) ws2['C' + r2].z = '#,##0';
    }

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, maskSheetName('ソウル旅行費用明細'));
    XLSX.utils.book_append_sheet(wb, ws2, maskSheetName('カテゴリ別集計'));
    return { wb: wb, filename: 'ソウル旅行費用明細_' + stamp() + '.xlsx', total: total, count: list.length };
  }

  function buildWorkbook() {
    return state.mode === 'trip' ? buildTripWorkbook() : buildTaxWorkbook();
  }

  function saveWorkbook(preferShare) {
    var built = buildWorkbook();
    if (!built.count) {
      toast('先に領収書を1件以上登録してください', true);
      return;
    }
    var out = XLSX.write(built.wb, { bookType: 'xlsx', type: 'array' });
    var blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    if (preferShare && typeof File !== 'undefined' && navigator.canShare) {
      var file = new File([blob], built.filename, { type: blob.type });
      if (navigator.canShare({ files: [file] })) {
        navigator.share({ files: [file], title: built.filename }).then(function () {
          toast('共有シートを開きました。「ファイルに保存」でExcelとして保存できます');
        }).catch(function (err) {
          if (err && err.name === 'AbortError') return;
          downloadBlob(blob, built.filename);
        });
        signalExport(built);
        return;
      }
    }
    downloadBlob(blob, built.filename);
    signalExport(built);
  }

  function downloadBlob(blob, filename) {
    try {
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.rel = 'noopener';
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(url); if (a.parentNode) a.parentNode.removeChild(a); }, 5000);
      toast('Excelファイルを出力しました：' + filename);
    } catch (e) {
      toast('ダウンロードできませんでした。「共有」から保存してみてください。', true);
    }
  }

  function signalExport(built) {
    signal((state.mode === 'trip' ? 'ソウル旅行の費用明細' : '領収書一覧') + '（' + built.count + '件・合計' + yen(built.total) + '）をExcel出力しました。');
  }

  function copyTSV() {
    var list = sortedEntries(currentList());
    if (!list.length) { toast('先に領収書を1件以上登録してください', true); return; }
    var lines = [];
    if (state.mode === 'trip') {
      lines.push(['No', '日付', '区分', '内容・支払先', '通貨', '金額（現地）', '円換算', 'メモ'].join('\t'));
      list.forEach(function (e, i) {
        lines.push([i + 1, e.date, e.category, e.vendor, e.currency, num(e.amount), yenOf(e), (e.memo || '').replace(/\t/g, ' ')].join('\t'));
      });
    } else {
      lines.push(['No', '日付', '支払先・店名', '勘定科目', '内容・メモ', '金額(税込)', '税区分', 'インボイス番号'].join('\t'));
      list.forEach(function (e, i) {
        lines.push([i + 1, e.date, e.vendor, e.category, (e.memo || '').replace(/\t/g, ' '), yenOf(e), e.taxType, e.invoiceNo].join('\t'));
      });
    }
    copyText(lines.join('\n'), 'スプレッドシートに貼り付けられる形でコピーしました');
  }

  function copyBoardTSV() {
    var sim = computeSim();
    var tot = totalsForYear();
    var lines = ['項目\t金額'];
    lines.push('事業売上\t' + sim.sales);
    FORM_CATEGORIES.forEach(function (c) {
      if (c === '専従者給与') return;
      var v = tot.byForm[c] || 0;
      if (v) lines.push(c + '\t' + v);
    });
    lines.push('経費計\t' + sim.expenses);
    if (sim.dokujuu) lines.push('専従者給与\t' + sim.dokujuu);
    lines.push('青色申告特別控除額\t' + sim.blue);
    lines.push('所得金額（事業所得）\t' + sim.biz);
    lines.push('給与所得\t' + sim.salary);
    lines.push('合計所得金額\t' + sim.totalIncome);
    lines.push('所得控除の額の合計\t' + sim.deductions);
    lines.push('課税される所得金額\t' + sim.taxable);
    lines.push('所得税及び復興特別所得税\t' + sim.totalTax);
    lines.push('源泉徴収税額\t' + sim.withheld);
    lines.push((sim.diff >= 0 ? '納める税金' : '還付される税金') + '\t' + Math.abs(sim.diff));
    copyText(lines.join('\n'), '転記ボードをまとめてコピーしました');
  }

  /* ============================================================
     初期化・イベント
     ============================================================ */
  function closeWidget() {
    try {
      if (window.pickle && pickle.widget && typeof pickle.widget.close === 'function') { pickle.widget.close(); return; }
    } catch (e) { /* noop */ }
    try {
      if (window.parent && window.parent !== window) {
        window.parent.postMessage({ type: 'pickle.widget.close' }, '*');
        window.parent.postMessage({ source: 'pickle-widget', action: 'close' }, '*');
        return;
      }
    } catch (e2) { /* noop */ }
    if (history.length > 1) history.back();
  }

  function applyPresets(force) {
    if (state.presetApplied && !force) return;
    var now = new Date().toISOString();
    state.trip = TRIP_PRESETS.map(function (p, i) {
      return {
        id: 'trip-preset-' + (i + 1),
        date: p.date, vendor: p.vendor, category: p.category, currency: p.currency,
        amount: p.amount, taxType: p.taxType || '', invoiceNo: '', memo: p.memo || '',
        thumb: '', source: 'preset', createdAt: now
      };
    });
    state.presetApplied = true;
    persist();
  }

  function switchMode(m) {
    if (m === state.mode) return;
    state.mode = m;
    persist();
    renderAll();
    window.scrollTo(0, 0);
  }

  function init() {
    load();
    applyPresets(false);

    $('rateInput').value = state.rate;

    Array.prototype.forEach.call(document.querySelectorAll('.mode-tab'), function (btn) {
      btn.addEventListener('click', function () { switchMode(btn.getAttribute('data-mode')); });
    });

    // 年度切替
    $('yearSelect').addEventListener('change', function () {
      var y = parseInt(this.value, 10);
      if (!y) return;
      state.year = y;
      persist();
      renderAll();
      toast(yearLabel(y) + '分に切り替えました');
    });
    $('btnAddYear').addEventListener('click', function () {
      var v = window.prompt('追加する年度を西暦で入力してください（例：2031）', String(state.year + 1));
      if (!v) return;
      var y = parseInt(String(v).replace(/[^0-9]/g, ''), 10);
      if (!y || y < 2000 || y > 2100) { toast('西暦4桁で入力してください', true); return; }
      state.year = y;
      persist();
      renderAll();
      toast(yearLabel(y) + '分に切り替えました');
    });

    // 写真・カメラ
    $('btnCamera').addEventListener('click', function () { $('fileCamera').click(); });
    $('btnLibrary').addEventListener('click', function () { $('fileLibrary').click(); });
    $('captureZone').addEventListener('click', function (ev) {
      if (ev.target.closest('button')) return;
      $('fileLibrary').click();
    });
    $('fileCamera').addEventListener('change', function (ev) {
      handleFiles(ev.target.files);
      ev.target.value = '';
    });
    $('fileLibrary').addEventListener('change', function (ev) {
      handleFiles(ev.target.files);
      ev.target.value = '';
    });
    $('btnManual').addEventListener('click', addManualDraft);

    // レート
    $('rateInput').addEventListener('input', function () {
      var v = parseFloat($('rateInput').value);
      if (!isNaN(v) && v > 0) {
        state.rate = v;
        persist();
        renderAll();
      }
    });
    $('rateInput').addEventListener('blur', function () {
      var v = parseFloat($('rateInput').value);
      if (isNaN(v) || v <= 0) { $('rateInput').value = state.rate; }
    });

    // 一覧操作
    $('entries').addEventListener('click', function (ev) {
      var btn = ev.target.closest('[data-act]');
      if (!btn) return;
      var row = btn.closest('.entry');
      var id = row.getAttribute('data-id');
      var list = currentList();
      var entry = list.find(function (x) { return x.id === id; });
      if (!entry) return;
      if (btn.getAttribute('data-act') === 'edit') startEdit(entry);
      if (btn.getAttribute('data-act') === 'del') {
        if (!window.confirm('「' + (entry.vendor || entry.category) + '」を削除しますか？')) return;
        var idx = list.findIndex(function (x) { return x.id === id; });
        if (idx >= 0) list.splice(idx, 1);
        persist();
        renderAll();
        toast('削除しました');
      }
    });

    $('btnClear').addEventListener('click', function () {
      var list = currentList();
      if (!list.length) return;
      if (!window.confirm(list.length + '件をすべて削除します。よろしいですか？（元に戻せません）')) return;
      if (state.mode === 'trip') { state.trip = []; state.presetApplied = true; }
      else { state.ledgers[ykey()] = []; }
      persist();
      renderAll();
      toast('すべて削除しました');
    });

    $('btnRestore').addEventListener('click', function () {
      if (!window.confirm('旅行の明細を初期データ（PDFの内訳7件）に戻します。追加した分は消えますがよろしいですか？')) return;
      state.presetApplied = false;
      applyPresets(true);
      renderAll();
      toast('初期データに戻しました');
    });

    // エクスポート
    $('btnExport').addEventListener('click', function () { saveWorkbook(true); });
    $('btnStickyExport').addEventListener('click', function () { saveWorkbook(true); });
    $('btnDownload').addEventListener('click', function () { saveWorkbook(false); });
    $('btnTsv').addEventListener('click', copyTSV);

    /* ---------- 税務診断タブ ---------- */
    $('salesInput').addEventListener('input', function () {
      var v = this.value === '' ? 0 : Math.max(0, Math.round(num(this.value)));
      state.sales[ykey()] = v;
      persist();
      renderRiskResults();
      renderKPIs();
    });
    $('blueSelect').addEventListener('change', function () {
      state.blue[ykey()] = Number(this.value);
      persist();
      renderRiskResults();
      renderBoard();
    });
    $('deductionsInput').addEventListener('input', function () {
      if (this.value === '') delete state.deductions[ykey()];
      else state.deductions[ykey()] = Math.max(0, Math.round(num(this.value)));
      persist();
      renderRiskResults();
      renderBoard();
    });

    $('slipList').addEventListener('input', function (ev) {
      var input = ev.target.closest('[data-slip]');
      if (!input) return;
      var wrap = input.closest('.slip');
      var idx = parseInt(wrap.getAttribute('data-idx'), 10);
      var slips = slipsOf();
      if (!slips[idx]) return;
      slips[idx][input.getAttribute('data-slip')] = input.value;
      persist();
      renderRiskResults();
      if (document.activeElement === $('deductionsInput')) return;
      var di = $('deductionsInput');
      if (di && state.deductions[ykey()] === undefined) {
        var auto = slipsAutoDeduction(slips);
        di.value = auto ? String(auto) : '';
        di.placeholder = auto ? '源泉徴収票③から自動（' + yen(auto) + '）' : '源泉徴収票③から自動';
      }
      renderBoard();
    });
    $('slipList').addEventListener('click', function (ev) {
      var btn = ev.target.closest('[data-slip-del]');
      if (!btn) return;
      var idx = parseInt(btn.getAttribute('data-slip-del'), 10);
      var slips = slipsOf();
      if (slips.length <= 1) return;
      slips.splice(idx, 1);
      persist();
      renderRiskInputs();
      renderRiskResults();
      renderBoard();
    });
    $('btnAddSlip').addEventListener('click', function () {
      slipsOf().push({ id: uid(), payer: '', payment: '', afterDeduction: '', deductions: '', taxWithheld: '' });
      persist();
      renderRiskInputs();
      renderRiskResults();
      renderBoard();
    });

    $('btnImport').addEventListener('click', function () { $('fileImport').click(); });
    $('fileImport').addEventListener('change', function (ev) {
      if (ev.target.files && ev.target.files[0]) readImportFile(ev.target.files[0]);
      ev.target.value = '';
    });

    $('riskList').addEventListener('click', function (ev) {
      var btn = ev.target.closest('[data-act]');
      if (!btn) return;
      var id = btn.getAttribute('data-id');
      var list = taxList();
      var entry = list.find(function (x) { return x.id === id; });
      if (!entry) return;
      if (btn.getAttribute('data-act') === 'edit') {
        state.mode = 'tax';
        persist();
        renderAll();
        startEdit(entry);
      } else if (btn.getAttribute('data-act') === 'del') {
        if (!window.confirm('「' + (entry.vendor || entry.category) + '」' + yen(entry.amount) + ' を経費から外しますか？\n（この端末の帳簿から削除されます）')) return;
        var idx = list.findIndex(function (x) { return x.id === id; });
        if (idx >= 0) list.splice(idx, 1);
        persist();
        renderAll();
        toast('経費から外しました');
      }
    });

    /* ---------- ガイドタブ ---------- */
    $('guideSteps').addEventListener('click', function (ev) {
      var btn = ev.target.closest('.chk');
      if (!btn) return;
      ev.preventDefault();
      ev.stopPropagation();
      var id = btn.getAttribute('data-step');
      var c = checksOf();
      if (c[id]) delete c[id]; else c[id] = true;
      btn.classList.toggle('on', !!c[id]);
      persist();
      // 進捗バーだけ更新
      var steps = guideSteps();
      var done = steps.filter(function (s) { return c[s.id]; }).length;
      var cnt = document.querySelector('#guideSteps .card .count');
      var bar = document.querySelector('#guideSteps .bar > i');
      if (cnt) cnt.textContent = done + ' / ' + steps.length + 'ステップ';
      if (bar) bar.style.width = Math.round(done / steps.length * 100) + '%';
    });

    $('board').addEventListener('click', function (ev) {
      var btn = ev.target.closest('[data-copy]');
      if (!btn) return;
      var v = btn.getAttribute('data-copy');
      copyText(v, v + ' をコピーしました');
      btn.classList.add('done');
      btn.textContent = '✓';
      setTimeout(function () { btn.classList.remove('done'); btn.textContent = '⧉'; }, 1200);
    });
    $('btnBoardTsv').addEventListener('click', copyBoardTSV);
    $('btnBoardPrint').addEventListener('click', function () { window.print(); });

    $('closeBtn').addEventListener('click', closeWidget);

    // 下部の固定バーは、出力カードが見えている間は隠して重なりを防ぐ
    try {
      if ('IntersectionObserver' in window) {
        var sticky = document.querySelector('.sticky-bar');
        var io = new IntersectionObserver(function (entries) {
          entries.forEach(function (en) { sticky.classList.toggle('fade-out', en.isIntersecting); });
        }, { threshold: 0.15 });
        io.observe($('exportCard'));
      }
    } catch (e) { /* noop */ }

    $('langSelect').value = 'jpn+eng';
    renderAll();
  }

  /* 連携用の最小API（アプリ本体や検証から参照するため） */
  window.ReceiptToExcel = {
    parseReceipt: parseReceipt,
    parseAmount: parseAmount,
    guessVendor: guessVendor,
    guessCategory: guessCategory,
    buildWorkbook: buildWorkbook,
    analyzeRisks: analyzeRisks,
    computeSim: computeSim,
    totals: totalsForYear,
    formCategoryOf: formCategoryOf,
    importRows: applyImport,
    setYear: function (y) { state.year = y; persist(); renderAll(); },
    setMode: function (m) { switchMode(m); },
    setSales: function (v) { state.sales[ykey()] = Number(v) || 0; persist(); renderAll(); },
    addEntry: function (e) {
      var x = Object.assign({ id: uid(), date: todayStr(), vendor: '', category: '雑費', currency: 'JPY', amount: 0, taxType: '', invoiceNo: '', memo: '', source: 'manual', createdAt: new Date().toISOString() }, e);
      taxList().push(x);
      persist(); renderAll();
      return x;
    },
    snapshot: function () {
      return {
        mode: state.mode, year: state.year, rate: state.rate,
        tax: taxList(), trip: state.trip,
        sales: salesOf(), blue: blueOf(), slips: slipsOf(),
        drafts: state.drafts.map(function (d) {
          return { id: d.id, status: d.status, progress: d.progress, message: d.message, fields: d.fields, text: (d.text || '').slice(0, 600) };
        })
      };
    }
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
