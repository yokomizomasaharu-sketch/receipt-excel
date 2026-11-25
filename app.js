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

  var TRIP_PRESETS = [];

  /* ---------- 状態 ---------- */
  var state = {
    mode: 'tax',
    year: new Date().getFullYear(),
    rate: DEFAULT_RATE,
    trip: [],
    ledgers: {},     /* { '2026': [entry] } */
    sales: {},       /* { '2026': 年間売上 } */
    blue: {},        /* { '2026': 650000 } */
    slips: {},       /* { '2026': [slip] } 複数枚OK */
    ded: {},         /* { '2026': { socialExtra, ideco, lifeGenNew, ... } } 所得控除の詳細 */
    deductions: {},  /* 旧：所得控除合計の手入力（v3以前。現在は使わない） */
    checks: {},      /* { '2026': { prep: true } } */
    assets: {},      /* { '2026': [減価償却資産] } 固定資産台帳 */
    revs: {},        /* { '2026': [{date, amount, memo}] } 売上の内訳（月別売上用） */
    sheet2026Applied: false,
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
          if (d.ded && typeof d.ded === 'object') state.ded = d.ded;
          if (d.checks && typeof d.checks === 'object') state.checks = d.checks;
          if (d.assets && typeof d.assets === 'object') state.assets = d.assets;
          if (d.revs && typeof d.revs === 'object') state.revs = d.revs;
          state.sheet2026Applied = !!d.sheet2026Applied;
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
      slips: state.slips, ded: state.ded, checks: state.checks,
      assets: state.assets, revs: state.revs, sheet2026Applied: state.sheet2026Applied,
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
  /* 減価償却リスト（固定資産台帳）*/ 
  function assetsOf(y) {
    var k = ykey(y);
    if (!Array.isArray(state.assets[k])) state.assets[k] = [];
    return state.assets[k];
  }
  /* 売上の内訳（月別売上の入力用） */
  function revsOf(y) {
    var k = ykey(y);
    if (!Array.isArray(state.revs[k])) state.revs[k] = [];
    return state.revs[k];
  }

  /* 所得控除の詳細（年度ごと）。数値はすべて「支払額」か「控除額」で入る */
  function dedOf(y) {
    var k = ykey(y);
    if (!state.ded[k] || typeof state.ded[k] !== 'object') state.ded[k] = {};
    return state.ded[k];
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
    [state.ledgers, state.sales, state.blue, state.slips, state.ded, state.checks].forEach(function (m) {
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
    renderDep();
    renderRiskResults();
    renderFilings();
    renderMyna();
    renderGuide();
  }

  function renderModeChrome() {
    var m = state.mode;
    var trip = m === 'trip';
    var titles = {
      tax: ['🧾 領収書スキャン', '撮影するだけで自動で表に整理・Excel出力'],
      trip: ['🇰🇷 ソウル旅行 費用', '11/14〜15 1泊2日｜現地レシートを読み取って総額を自動集計'],
      risk: ['🛡️ 税務リスク診断', '調査で指摘されやすい項目を自動で洗い出し・源泉徴収票と損益通算'],
      guide: ['🖥️ PC版e-Tax 入力ナビ', 'PCの画面どおりに進める・数字をコピー・マイナポータル連携']
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
      : 'シート「領収書明細」＋「科目別集計」＋「税務チェック」＋「税務診断・源泉」＋「確定申告サマリー」＋「減価償却資産台帳」。金額はSUM数式入りで、そのまま計算・並べ替えできます。';
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

  /* ============================================================
     減価償却リスト（固定資産台帳）
     ・e-Tax 青色申告決算書3ページ目「減価償却費の計算」と同じ列
     ・定額法（平成19年4月以後取得）／一括償却（3年）／少額特例（30万円未満）
     ・タックスナップ（スプレッドシート）の計上額と並べて差を見せる
     ============================================================ */
  var DEP_METHODS = ['定額法', '一括償却3年', '少額特例', '対象外（消耗品費）'];
  var DEP_RATES = {
    2: 0.500, 3: 0.333, 4: 0.250, 5: 0.200, 6: 0.167, 7: 0.143, 8: 0.125, 9: 0.112,
    10: 0.100, 11: 0.091, 12: 0.083, 13: 0.077, 14: 0.071, 15: 0.067, 16: 0.063, 17: 0.059,
    18: 0.056, 19: 0.053, 20: 0.050
  };
  function depRateOf(a) {
    var r = Number(a && a.rate);
    if (r > 0) return r;
    var life = Number(a && a.life);
    return DEP_RATES[life] || 0.200;
  }
  function depAcquired(a) {
    var m = String((a && a.acquired) || '').match(/(\d{4})\D{0,2}(\d{1,2})/);
    if (!m) return null;
    var mo = Number(m[2]);
    if (!(mo >= 1 && mo <= 12)) return null;
    return { year: Number(m[1]), month: mo };
  }
  function depMonthsOf(a, y) {
    var m = Number(a && a.months);
    if (m > 0) return Math.max(0, Math.min(12, m));
    var acq = depAcquired(a);
    if (acq) {
      if (acq.year === y) return 13 - acq.month;
      if (acq.year > y) return 0;
    }
    return 12;
  }
  function depFirstMonths(a, acq) {
    var m = Number(a && a.months);
    if (m > 0) return Math.max(0, Math.min(12, m));
    return acq ? (13 - acq.month) : 12;
  }
  function depAccumBefore(a, y, cost, rate) {
    var acq = depAcquired(a);
    if (!acq || cost <= 0) return 0;
    var total = 0;
    for (var yy = acq.year; yy < y; yy += 1) {
      var m = (yy === acq.year) ? (13 - acq.month) : 12;
      total += Math.floor(cost * rate * m / 12);
      if (total >= cost - 1) return Math.max(0, cost - 1);
    }
    return Math.min(total, Math.max(0, cost - 1));
  }
  function depCalc(a, y) {
    a = a || {};
    if (!y) y = state.year;
    var cost = num(a.cost);
    var method = DEP_METHODS.indexOf(a.method) >= 0 ? a.method : '定額法';
    var rate = depRateOf(a);
    var months = depMonthsOf(a, y);
    var acq = depAcquired(a);
    var pctRaw = a.bizPct;
    var pct = (pctRaw === '' || pctRaw === undefined || pctRaw === null) ? 100 : Math.max(0, Math.min(100, Number(pctRaw) || 0));
    var ordinary = 0, bookEnd = '', accum = 0;
    var opening = num(a.opening);
    var notes = [];

    if (!cost) {
      ordinary = 0;
      bookEnd = '';
      notes.push('取得価額が未入力のため、法定計算の金額を出せません（スプレッドシートの計上額を手入力で上書きできます）。');
    } else if (method === '対象外（消耗品費）') {
      ordinary = 0;
      bookEnd = cost;
      notes.push('減価償却せず、消耗品費で全額経費にします（事業割合を掛けた ' + yen(Math.floor(cost * pct / 100)) + '円）。この行はe-Taxの3ページ目には載せず、1ページ目の消耗品費に入れてください。取得年以外で処理するときは、事業に使い始めた年か処理の根拠を確認してください。');
    } else if (method === '少額特例') {
      ordinary = cost;
      bookEnd = 0;
      if (cost >= 300000) notes.push('少額特例（少額減価償却資産）は取得価額30万円未満が対象です。この資産は' + yen(cost) + 'なので対象外です。');
      else notes.push('少額特例：30万円未満の資産は取得年に全額を経費にできます（年間合計300万円まで）。');
    } else if (method === '一括償却3年') {
      if (cost >= 200000) notes.push('一括償却資産は取得価額20万円未満が対象です（20万円以上は定額法）。');
      var b4m = 0, b4amt = 0;
      if (acq) {
        for (var yy = acq.year; yy < y; yy += 1) {
          var mm = (yy === acq.year) ? (13 - acq.month) : 12;
          var take = Math.min(mm, 36 - b4m);
          if (take <= 0) break;
          b4amt += Math.floor(cost / 36 * take);
          b4m += take;
        }
      }
      b4amt = Math.min(b4amt, cost);
      ordinary = Math.min(Math.max(0, cost - b4amt), Math.floor(cost / 36 * months));
      bookEnd = Math.max(0, cost - b4amt - ordinary);
      notes.push('一括償却資産：3年間（36ヶ月）で均等に経費化。' + (months < 12 ? '本年は' + months + 'ヶ月分。12ヶ月なら' + yen(Math.floor(cost / 3)) + '円。' : ''));
    } else {
      if (opening > 0) {
        var rem2 = Math.min(cost, opening);
        accum = cost - rem2;
        ordinary = Math.min(rem2, Math.floor(cost * rate * months / 12));
        bookEnd = Math.max(1, rem2 - ordinary);
      } else {
        accum = depAccumBefore(a, y, cost, rate);
        var remain = Math.max(0, cost - accum);
        ordinary = Math.min(remain, Math.floor(cost * rate * months / 12));
        bookEnd = Math.max(1, cost - accum - ordinary);
      }
      if (bookEnd === 1) notes.push('償却可能限度額（備忘価額1円）に到達。以降は償却できません。');
      if (cost < 100000) notes.push('10万円未満なら全額を消耗品費で経費にできます（この行は定額法で計算中）。');
    }
    if (!acq && cost) notes.push('取得年月が未入力のため、本年中の償却期間を' + months + 'ヶ月として計算しています。');

    var autoExpense = Math.floor(ordinary * pct / 100);
    var hasOverride = !(a.override === '' || a.override === undefined || a.override === null);
    var expense = hasOverride ? num(a.override) : autoExpense;

    return {
      cost: cost, method: method, rate: rate, months: months, pct: pct, acq: acq,
      opening: opening, ordinary: ordinary, autoExpense: autoExpense, expense: expense, bookEnd: bookEnd,
      accumBefore: accum, override: hasOverride, notes: notes
    };
  }
  /* 比較用：同じ資産を3つの方法で計算したときの本年分 */
  function depCompare(a, y) {
    if (!y) y = state.year;
    var out = [];
    DEP_METHODS.forEach(function (m) {
      var c = depCalc(Object.assign({}, a, { method: m, override: '' }), y);
      var label = m === '対象外（消耗品費）' ? '消耗品費（全額）' : m;
      var acqC = depAcquired(a);
      if (acqC && acqC.year < y && (m === '少額特例' || m === '対象外（消耗品費）')) label += '（取得年のみ）';
      var amt = m === '対象外（消耗品費）' ? Math.floor(c.cost * c.pct / 100) : c.expense;
      out.push({ method: m, label: label, expense: amt, ordinary: c.ordinary });
    });
    return out;
  }
  function depNorm(s) { return String(s || '').toUpperCase().replace(/[\s　]/g, ''); }
  function depTaxsnapRows(y) {
    if (!y) y = state.year;
    var rows = [];
    taxList(y).forEach(function (e) {
      if ((e.category || '') !== '減価償却費') return;
      var t = String(e.vendor || e.memo || '');
      var m = t.match(/減価償却[（(]([^）)]*)[）)]/);
      rows.push({ name: m ? m[1] : t, amount: yenOf(e), entry: e });
    });
    return rows;
  }
  function depTaxsnapFor(a, rows) {
    var key = depNorm(a.matchKey || a.name);
    if (!key) return null;
    var hit = null;
    rows.forEach(function (r) {
      if (hit) return;
      var k = depNorm(r.name);
      if (!k) return;
      if (k === key || k.indexOf(key) >= 0 || key.indexOf(k) >= 0) { hit = r; return; }
      var toks = key.match(/[A-Z0-9\-]{5,}/g) || [];
      if (toks.some(function (tk) { return k.indexOf(tk) >= 0; })) hit = r;
    });
    return hit;
  }
  function depSummary(y) {
    if (!y) y = state.year;
    var list = assetsOf(y);
    var tsRows = depTaxsnapRows(y);
    var out = { count: list.length, cost: 0, ordinary: 0, expense: 0, taxsnap: 0, diff: 0, missing: [], rows: [], overridden: 0 };
    list.forEach(function (a) {
      var c = depCalc(a, y);
      var ts = depTaxsnapFor(a, tsRows);
      out.cost += c.cost;
      out.ordinary += c.ordinary;
      out.expense += c.expense;
      if (!c.cost) out.missing.push(a.name || '（名称未入力）');
      if (c.override) out.overridden += 1;
      out.rows.push({ a: a, c: c, taxsnap: ts });
    });
    tsRows.forEach(function (r) { out.taxsnap += r.amount; });
    out.diff = out.expense - out.taxsnap;
    return out;
  }
  /* 経費帳の地代家賃内訳（e-Tax 2ページ目「地代家賃の内訳」用） */
  function rentBreakdown(y) {
    if (!y) y = state.year;
    var map = {};
    taxList(y).forEach(function (e) {
      if ((e.category || '') !== '地代家賃') return;
      var k = e.vendor || '（支払先なし）';
      if (!map[k]) map[k] = { n: 0, sum: 0 };
      map[k].n += 1;
      map[k].sum += yenOf(e);
    });
    return Object.keys(map).map(function (k) { return { vendor: k, n: map[k].n, sum: map[k].sum }; }).sort(function (a, b) { return b.sum - a.sum; });
  }
  function monthlySales(y) {
    if (!y) y = state.year;
    var m = new Array(13).fill(0);
    revsOf(y).forEach(function (r) {
      var mo = Number(String(r.date || '').slice(5, 7));
      if (mo >= 1 && mo <= 12) m[mo] += num(r.amount);
    });
    return m;
  }

  /* ---------- 税制（年度別） ---------- */
  /* 給与所得控除
     - 令和8年分・令和9年分：最低保障額74万円　（令和8年度税制改正。国税庁 No.1410「令和8年分・令和9年分」）
     - 令和7年分：最低保障額65万円
     - 令和6年分以前：最低保障額55万円 */
  function salaryDedMin(y) {
    if (y >= 2026) return 740000;
    if (y >= 2025) return 650000;
    return 550000;
  }
  function salaryDeduction(pay, y) {
    pay = Math.max(0, Number(pay) || 0);
    if (!pay) return 0;
    if (y === undefined || y === null) y = state.year;
    if (y >= 2026) {
      /* 収入69万1,000円〜220万円未満の特例（給与所得＝収入−74万円、74万1,000円未満は給与所得なし） */
      if (pay < 741000) return pay;
      if (pay < 2200000) return 740000;
      if (pay <= 3600000) return Math.floor(pay * 0.3) + 80000;
      if (pay <= 6600000) return Math.floor(pay * 0.2) + 440000;
      if (pay <= 8500000) return Math.floor(pay * 0.1) + 1100000;
      return 1950000;
    }
    if (y >= 2025) {
      if (pay <= 1900000) return 650000;
      if (pay <= 3600000) return Math.floor(pay * 0.3) + 80000;
      if (pay <= 6600000) return Math.floor(pay * 0.2) + 440000;
      if (pay <= 8500000) return Math.floor(pay * 0.1) + 1100000;
      return 1950000;
    }
    if (pay <= 1625000) return 550000;
    if (pay <= 1800000) return Math.floor(pay * 0.4) - 100000;
    if (pay <= 3600000) return Math.floor(pay * 0.3) + 80000;
    if (pay <= 6600000) return Math.floor(pay * 0.2) + 440000;
    if (pay <= 8500000) return Math.floor(pay * 0.1) + 1100000;
    return 1950000;
  }

  /* 源泉徴収票1枚分の「給与所得」：②が手入力されていればそれを、無ければ①から自動計算 */
  function salaryIncomeOf(slip, y) {
    var pay = num(slip && slip.payment);
    if (!pay) return 0;
    if (slip && slip.afterManual) return Math.max(0, num(slip.after));
    return Math.max(0, pay - salaryDeduction(pay, y));
  }
  function salaryAfterAuto(slip, y) {
    var pay = num(slip && slip.payment);
    if (!pay) return 0;
    return Math.max(0, pay - salaryDeduction(pay, y));
  }

  /* 基礎控除（合計所得金額で決まる。国税庁 No.1199 の年度別表）
     - 令和8年分・令和9年分：合計所得489万円以下104万円／655万円以下67万円／2,350万円以下62万円
     - 令和7年分：132万円以下95万円／336万円以下88万円／489万円以下68万円／655万円以下63万円／2,350万円以下58万円
     - 令和6年分以前：2,400万円以下48万円 */
  function basicDeductionFor(y, income) {
    var inc = Math.max(0, Number(income) || 0);
    if (y >= 2028) {
      if (inc <= 1320000) return 990000;
      if (inc <= 23500000) return 620000;
      if (inc <= 24000000) return 480000;
      if (inc <= 24500000) return 320000;
      if (inc <= 25000000) return 160000;
      return 0;
    }
    if (y >= 2026) {
      if (inc <= 4890000) return 1040000;
      if (inc <= 6550000) return 670000;
      if (inc <= 23500000) return 620000;
      if (inc <= 24000000) return 480000;
      if (inc <= 24500000) return 320000;
      if (inc <= 25000000) return 160000;
      return 0;
    }
    if (y >= 2025) {
      if (inc <= 1320000) return 950000;
      if (inc <= 3360000) return 880000;
      if (inc <= 4890000) return 680000;
      if (inc <= 6550000) return 630000;
      if (inc <= 23500000) return 580000;
      if (inc <= 24000000) return 480000;
      if (inc <= 24500000) return 320000;
      if (inc <= 25000000) return 160000;
      return 0;
    }
    if (inc <= 24000000) return 480000;
    if (inc <= 24500000) return 320000;
    if (inc <= 25000000) return 160000;
    return 0;
  }
  function basicBandLabel(y, income) {
    var inc = Math.max(0, Number(income) || 0);
    if (inc > 25000000) return '合計所得2,500万円超（控除なし）';
    if (inc > 24500000) return '合計所得2,450万円超2,500万円以下';
    if (inc > 24000000) return '合計所得2,400万円超2,450万円以下';
    if (y >= 2028) return inc <= 1320000 ? '合計所得132万円以下' : '合計所得2,350万円以下';
    if (y >= 2026) return inc <= 4890000 ? '合計所得489万円以下' : (inc <= 6550000 ? '合計所得489万円超655万円以下' : '合計所得2,350万円以下');
    if (y >= 2025) {
      if (inc <= 1320000) return '合計所得132万円以下';
      if (inc <= 3360000) return '合計所得336万円以下';
      if (inc <= 4890000) return '合計所得489万円以下';
      if (inc <= 6550000) return '合計所得655万円以下';
      return '合計所得2,350万円以下';
    }
    return '合計所得2,400万円以下';
  }

  /* 生命保険料控除（新制度：一般・介護医療・個人年金 各4万／旧制度：各5万。新旧合計の上限12万円） */
  function lifeCalcNew(p) {
    p = Math.max(0, Number(p) || 0);
    if (p <= 20000) return p;
    if (p <= 40000) return Math.floor(p / 2) + 10000;
    if (p <= 80000) return Math.floor(p / 4) + 20000;
    return 40000;
  }
  function lifeCalcOld(p) {
    p = Math.max(0, Number(p) || 0);
    if (p <= 25000) return p;
    if (p <= 50000) return Math.floor(p / 2) + 12500;
    if (p <= 100000) return Math.floor(p / 4) + 25000;
    return 50000;
  }
  function lifeDeduction(d) {
    var newSum = Math.min(120000, lifeCalcNew(d.lifeGenNew) + lifeCalcNew(d.lifeCare) + lifeCalcNew(d.lifePenNew));
    var oldSum = Math.min(100000, lifeCalcOld(d.lifeGenOld) + lifeCalcOld(d.lifePenOld));
    return Math.min(120000, newSum + oldSum);
  }

  /* 地震保険料控除（地震保険料は5万円まで、旧長期損害保険料は最大1万円、合わせて5万円まで） */
  function quakeDeduction(d) {
    var a = Math.min(50000, Math.max(0, num(d.quake)));
    var p = Math.max(0, num(d.quakeOld));
    var b = p <= 0 ? 0 : (p <= 5000 ? p : (p <= 15000 ? Math.floor(p / 2) + 2500 : 10000));
    return Math.min(50000, a + b);
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

  /* 未入力の空欄の源泉徴収票は計算に入れない */
  function activeSlips(slips) {
    return slips.filter(function (s) {
      return num(s.payment) || num(s.after) || num(s.deductions) || num(s.taxWithheld) || num(s.social);
    });
  }

  /* 所得控除の内訳をまとめて計算する（源泉徴収票の社会保険料も合算） */
  function deductionBreakdown(totalIncome) {
    var y = state.year;
    var d = dedOf(y);
    var slips = activeSlips(slipsOf(y));
    var socialSlip = 0, dedSlip = 0;
    slips.forEach(function (s) { socialSlip += num(s.social); dedSlip += num(s.deductions); });

    var social = socialSlip + num(d.socialExtra);
    var ideco = num(d.ideco);
    var life = d.lifeMode === 'manual' ? num(d.lifeManual) : lifeDeduction(d);
    var quake = d.quakeMode === 'manual' ? num(d.quakeManual) : quakeDeduction(d);

    var medPaid = num(d.medPaid), medReimb = num(d.medReimb);
    var medFloor = Math.min(100000, Math.floor(Math.max(0, totalIncome) * 0.05));
    var medical = Math.max(0, Math.min(2000000, medPaid - medReimb - medFloor));

    var donPaid = num(d.donation);
    var donCap = Math.max(0, Math.floor(Math.max(0, totalIncome) * 0.4));
    var donation = donPaid > 0 ? Math.max(0, Math.min(donPaid, donCap) - 2000) : 0;

    var spouse = num(d.spouse);
    var depGeneral = num(d.depGeneral) * 380000;
    var depSpecific = num(d.depSpecific) * 630000;
    var depElder = num(d.depElder) * 480000;
    var dependents = depGeneral + depSpecific + depElder;
    var disability = num(d.disability);
    var singleParent = num(d.singleParent);

    var basicAuto = basicDeductionFor(y, totalIncome);
    var basicSet = !(d.basicManual === undefined || d.basicManual === null || d.basicManual === '');
    var basic = basicSet ? num(d.basicManual) : basicAuto;
    var other = num(d.other);

    var total = social + ideco + life + quake + medical + donation + spouse + dependents + disability + singleParent + basic + other;

    return {
      d: d, slips: slips, socialSlip: socialSlip, dedSlip: dedSlip,
      social: social, ideco: ideco, life: life, quake: quake,
      medPaid: medPaid, medReimb: medReimb, medFloor: medFloor, medical: medical,
      donPaid: donPaid, donCap: donCap, donation: donation,
      spouse: spouse, depGeneral: depGeneral, depSpecific: depSpecific, depElder: depElder,
      dependents: dependents, disability: disability, singleParent: singleParent,
      basic: basic, basicAuto: basicAuto, basicSet: basicSet, other: other,
      total: total
    };
  }

  function computeSim() {
    var y = state.year;
    var sales = salesOf(y);
    var tot = totalsForYear(y);
    var blue = blueOf(y);
    var beforeBlue = sales - tot.expenses - tot.dokujuu;
    /* 青色申告特別控除は所得金額を限度とする（控除で赤字は作れない） */
    var blueApplied = Math.max(0, Math.min(blue, beforeBlue));
    var biz = beforeBlue - blueApplied;

    var bp = deductionBreakdown(0); /* 所得が決まる前の暫定（医療費の足切りなどに使う） */
    var slips = bp.slips;
    var pay = 0, salary = 0, wh = 0;
    slips.forEach(function (s) {
      pay += num(s.payment);
      salary += salaryIncomeOf(s, y);
      wh += num(s.taxWithheld);
    });

    var totalIncome = salary + biz;
    var b = deductionBreakdown(totalIncome);
    var ded = b.total;

    var taxableRaw = totalIncome - ded;
    var taxable = taxableRaw > 0 ? Math.floor(taxableRaw / 1000) * 1000 : 0;
    var incomeTaxBase = incomeTaxOf(taxable);
    var housingWant = Math.max(0, num(b.d.housing));
    var housing = Math.min(housingWant, incomeTaxBase);
    var incomeTax = incomeTaxBase - housing;
    var fukko = y <= 2037 ? Math.floor(incomeTax * 0.021) : 0;
    var totalTax = incomeTax + fukko;
    var diff = totalTax - wh;

    /* 住民税（概算）：所得控除は所得税と同じだが基礎控除は43万円。所得割10% − 調整控除 − 住宅ローン控除の住民税分＋均等割 */
    var residentBasic = y >= 2026 ? 450000 : 430000;
    var residentDed = Math.max(0, ded - b.basic + Math.min(b.basic, residentBasic));
    var residentTaxable = Math.max(0, Math.floor((totalIncome - residentDed) / 1000) * 1000);
    var residentIncomePart = Math.max(0, Math.floor(residentTaxable * 0.10 / 100) * 100 - (residentTaxable > 0 ? 2500 : 0));
    var residentHousing = Math.min(136500, Math.max(0, housingWant - housing));
    var residentTax = Math.max(0, residentIncomePart - residentHousing) + (totalIncome > 0 ? 5000 : 0);

    var warnings = [];
    function warn(level, text) { warnings.push({ level: level, text: text }); }
    var dd = b.d;
    if (dd.lifeMode === 'manual' && num(dd.lifeManual) > 120000) warn('warn', '生命保険料控除が上限12万円を超えています。控除証明書の金額に直してください。');
    if (dd.quakeMode === 'manual' && num(dd.quakeManual) > 50000) warn('warn', '地震保険料控除が上限5万円を超えています。');
    if (num(dd.medPaid) > 0 && b.medical === 0) warn('info', '医療費の支払額が足切り額（10万円と総所得の5%の少ない方＝' + yen(b.medFloor) + '）以下のため、医療費控除は0円です。');
    if (b.donPaid > 0 && b.donPaid > b.donCap) warn('warn', '寄附金控除の対象は総所得金額等の40%（' + yen(b.donCap) + '）までです。超えた分は控除されません。');
    if (ded > totalIncome && totalIncome > 0) warn('warn', '所得控除の合計が所得を上回っています。控除しきれない分は翌年に繰り越せません（雑損控除を除く）。');
    slips.forEach(function (s, i) {
      if (num(s.payment) > 0 && num(s.social) > num(s.payment)) warn('warn', '源泉徴収票' + (i + 1) + '：社会保険料が支払金額を超えています。入力位置を確認してください。');
      if (s.afterManual && num(s.after) > num(s.payment)) warn('warn', '源泉徴収票' + (i + 1) + '：給与所得控除後の金額が支払金額より大きくなっています。');
    });
    if (slips.length > 1) warn('info', '源泉徴収票が' + slips.length + '枚あります。給与が2か所以上なら、年末調整済みでも確定申告が必要です（副業がばれる・ばれないの問題ではなく、法定です）。');
    if (b.dedSlip > 0 && Math.abs(b.dedSlip - ded) > 100000) {
      warn('info', '源泉徴収票③の所得控除合計（' + yen(b.dedSlip) + '）と、この画面の控除合計（' + yen(ded) + '）が' + yen(Math.abs(b.dedSlip - ded)) + '違います。事業の青色申告特別控除・医療費・ふるさと納税などを入れている場合は正常です。');
    }
    if (b.basicSet && Math.abs(b.basic - b.basicAuto) > 1000) {
      warn('info', '基礎控除を手入力（' + yen(b.basic) + '）で上書きしています。' + y + '年分の自動計算は' + yen(b.basicAuto) + 'です。');
    }
    if (housingWant > incomeTaxBase) warn('info', '住宅ローン控除が所得税を超えています。引ききれない分は住民税から控除されます（最大13万6,500円）。');
    if (num(dd.housing) > 0) warn('info', '住宅ローン控除は税額控除です。所得控除の合計には入れず、所得税から直接引いています（この画面の計算どおり）。');

    return {
      year: y, sales: sales, expenses: tot.expenses, dokujuu: tot.dokujuu,
      count: tot.count, blue: blue, blueApplied: blueApplied, beforeBlue: beforeBlue, biz: biz,
      slipCount: slips.length, pay: pay, withheld: wh,
      salary: salary, totalIncome: totalIncome,
      ded: ded, deductions: ded, deductionsManual: false,
      breakdown: b,
      taxableRaw: taxableRaw, taxable: taxable,
      incomeTaxBase: incomeTaxBase, housing: housing, housingWant: housingWant,
      incomeTax: incomeTax, fukko: fukko,
      totalTax: totalTax, diff: diff,
      residentTaxable: residentTaxable, residentTax: residentTax,
      warnings: warnings
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

      // 家事按分の記録なし（自宅家賃・水道光熱費の全額計上は否認されやすいので赤）
      var APPORTION = /(按分|事業|業務|仕事|事業用|事業分|％|%|割)/;
      if ((cat === '地代家賃' || cat === '水道光熱費') && !APPORTION.test(memo)) {
        push('red', e, '家事按分なしで全額を経費計上', '自宅の家賃・水道光熱費は、事業で使った割合だけが経費です。全額計上は税務調査で真っ先に否認されます。メモに「按分30%（面積◯㎡／◯㎡・稼働◯日）」のように根拠を残してください。', 'apportion-red');
      } else if (cat === '通信費' && !APPORTION.test(memo)) {
        push('amber', e, '家事按分の記録がありません', '自宅の通信費は、事業で使った割合だけが経費です。メモに「按分30%」のように割合と根拠（稼働日数・面積など）を残すと説明できます。', 'apportion');
      }

      // 旧居の退去精算（自宅兼事務所の按分根拠が問われる）
      if (/(旧居|退去|解約|明渡|明け渡し|原状回復|ハウスクリーニング|クリーニング代|敷金|礼金|更新料|退去精算)/.test(memoLow) &&
          ['地代家賃', '水道光熱費', '雑費', '修繕費', '支払手数料'].indexOf(cat) >= 0) {
        push('amber', e, '旧居の退去精算（自宅兼事務所の按分根拠）', '退去精算・原状回復・クリーニング代を経費にするなら、その住まいを事業に使っていた割合と根拠（事務所として使った面積・日数、按分率）を残してください。全額を経費にすると「私的な退去費用」と判断されます。', 'oldhome');
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
      if (exp >= 100000) {
        structural.push({
          level: 'red',
          title: '売上0で経費だけが計上されています',
          why: 'この状態だと「事業所得」ではなく「雑所得（趣味）」とみなされ、給与所得との損益通算も65万円控除も認められません。売上（入金）を集計して入れるか、売上が無いなら経費の計上自体を見直してください。'
        });
      } else {
        structural.push({
          level: 'amber',
          title: '売上が未入力です',
          why: '経費だけが登録され、売上が0のままです。売上を入れると損益通算と還付額まで計算できます（今の経費は' + yen(exp) + '）。'
        });
      }
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

    /* 給与＋赤字事業の損益通算で還付を受ける形（税務署が事業実態を確認する典型） */
    var sim = computeSim();
    if (sim.salary > 0 && sim.biz < 0) {
      var ratio = sim.salary ? Math.abs(sim.biz) / sim.salary : 0;
      if (sim.diff < 0) {
        structural.push({
          level: 'red',
          title: '赤字の損益通算で還付を受ける形（' + yen(-sim.diff) + '）',
          why: '給与所得' + yen(sim.salary) + 'に事業の赤字' + yen(Math.abs(sim.biz)) + 'をぶつけて、払いすぎた所得税の還付を受けようとしています。この形は税務署が「本当に事業か（趣味の範囲か）」を確認しに来る典型です。売上・経費の裏付（請求書・入金記録・領収書）を、日付と相手が分かる形で揃えておいてください。'
        });
      } else if (ratio > 0.2) {
        structural.push({
          level: 'amber',
          title: '事業赤字が給与の' + Math.round(ratio * 100) + '%に達しています',
          why: '赤字額が大きく、損益通算で課税所得をかなり下げています。毎年続くと事業実態を問われます。売上の計上漏れが無いか、私的支出が混ざっていないかを先に確認してください。'
        });
      }
    }

    /* 保険料控除・社会保険料の入れ忘れ（損をする側のリスク） */
    var bd = deductionBreakdown(sim.totalIncome);
    if (sim.slipCount && !bd.life && !bd.quake && !bd.medical) {
      structural.push({
        level: 'amber',
        title: '生命保険・地震保険・医療費の控除がまだ入っていません',
        why: '控除証明書が届いていれば、所得控除が増えて還付が増えます。入力を忘れるとあなたが損をするだけなので、書類が届いたらすぐ「所得控除の詳細」に入れてください。'
      });
    }
    if (sim.slipCount && !bd.social) {
      structural.push({
        level: 'amber',
        title: '社会保険料が未入力です',
        why: '源泉徴収票の「社会保険料等の金額」を入れると、その全額が所得控除になります。入っていないと課税所得がその分高く出ます。'
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
    renderSlips();
    renderDeductions();
    if (lastImport) {
      var ip = $('importPreview');
      if (ip) ip.innerHTML = importPreviewHTML(lastImport);
      var ic = $('importCount');
      if (ic) ic.textContent = lastImport.total + '件取込';
    }
  }

  /* ---------- 源泉徴収票（複数枚） ---------- */
  function renderSlips() {
    var box = $('slipList');
    if (!box) return;
    var slips = slipsOf();
    if (!slips.length) {
      slips.push({ id: uid(), payer: '', payment: '', after: '', afterManual: false, social: '', deductions: '', taxWithheld: '' });
    }
    $('slipCount').textContent = slips.length + '枚';
    box.innerHTML = slips.map(function (s, i) {
      var autoAfter = salaryAfterAuto(s, state.year);
      var pay = num(s.payment);
      return '<div class="slip" data-idx="' + i + '">' +
        '<div class="slip-head"><b>源泉徴収票 ' + (i + 1) + (i === 0 ? '（年末調整済み）' : '（前職・2か所目など）') + '</b>' +
        (slips.length > 1 ? '<button class="mini-btn" data-slip-del="' + i + '" style="color:#dc2626;border-color:#fecaca;min-height:28px;padding:3px 8px;">削除</button>' : '') +
        '</div>' +
        '<label class="field"><span>支払者の名称（メモ用）</span><input class="input" type="text" data-slip="payer" value="' + esc(s.payer || '') + '" placeholder="例：株式会社◯◯"></label>' +
        '<label class="field"><span>① 支払金額（給与収入）</span><input class="input" type="number" inputmode="numeric" min="0" data-slip="payment" value="' + esc(s.payment || '') + '" placeholder="0"></label>' +
        '<label class="field"><span>② 給与所得控除後の金額' + (s.afterManual ? '（手入力）' : '（空欄なら自動）') + '</span>' +
          '<input class="input' + (s.afterManual ? ' edited' : '') + '" type="number" inputmode="numeric" min="0" data-slip="after" value="' + esc(s.afterManual ? (s.after || '') : '') + '" placeholder="' + (pay ? esc(String(autoAfter)) : '0') + '">' +
        '</label>' +
        '<div class="slip-note" data-note="' + i + '">' + slipNoteHTML(s, i) + '</div>' +
        '<div class="grid2">' +
        '<label class="field"><span>社会保険料等の金額</span><input class="input" type="number" inputmode="numeric" min="0" data-slip="social" value="' + esc(s.social || '') + '" placeholder="0"></label>' +
        '<label class="field"><span>③ 所得控除の額の合計額</span><input class="input" type="number" inputmode="numeric" min="0" data-slip="deductions" value="' + esc(s.deductions || '') + '" placeholder="0"></label>' +
        '</div>' +
        '<label class="field"><span>④ 源泉徴収税額</span><input class="input" type="number" inputmode="numeric" min="0" data-slip="taxWithheld" value="' + esc(s.taxWithheld || '') + '" placeholder="0"></label>' +
        '<div class="slip-foot">' + slipFootHTML(s) + '</div>' +
        '</div>';
    }).join('');
    renderSlipTotals();
  }

  function slipNoteHTML(s, i) {
    var auto = salaryAfterAuto(s, state.year);
    var pay = num(s.payment);
    var html = '';
    if (s.afterManual) {
      html += '②は手入力中（自動計算は' + (pay ? yen(auto) : '—') + '）。';
      html += '<button class="link-btn" data-slip-auto="' + i + '" style="font-size:11px">自動に戻す</button>';
    } else {
      html += pay
        ? '①' + yen(pay) + ' − 給与所得控除' + yen(salaryDeduction(pay, state.year)) + ' ＝ <b>' + yen(auto) + '</b>（' + state.year + '年分の税法）'
        : '①を入れると、給与所得控除（' + state.year + '年分は最低' + yen(salaryDedMin(state.year)) + '）を引いた額を自動で出します。';
    }
    return html;
  }

  function slipFootHTML(s) {
    var pay = num(s.payment);
    var kofu = pay ? (s.afterManual ? Math.max(0, pay - num(s.after)) : salaryDeduction(pay, state.year)) : 0;
    return 'この票の給与所得 <b>' + yen(salaryIncomeOf(s, state.year)) + '</b>（控除 ' + yen(kofu) + '）';
  }

  function refreshSlipNote(i) {
    var slips = slipsOf();
    var s = slips[i];
    if (!s) return;
    var el = document.querySelector('[data-note="' + i + '"]');
    if (el) el.innerHTML = slipNoteHTML(s, i);
    var wrap = document.querySelector('.slip[data-idx="' + i + '"]');
    if (wrap) {
      var afterInput = wrap.querySelector('[data-slip="after"]');
      if (afterInput && !s.afterManual) afterInput.placeholder = num(s.payment) ? String(salaryAfterAuto(s, state.year)) : '0';
      var foot = wrap.querySelector('.slip-foot');
      if (foot) foot.innerHTML = slipFootHTML(s);
    }
  }

  function renderSlipTotals() {
    var box = $('slipTotals');
    if (!box) return;
    var slips = activeSlips(slipsOf());
    if (!slips.length) { box.innerHTML = ''; return; }
    var pay = 0, sal = 0, soc = 0, wh = 0, dedS = 0;
    slips.forEach(function (s) {
      pay += num(s.payment); sal += salaryIncomeOf(s, state.year);
      soc += num(s.social); wh += num(s.taxWithheld); dedS += num(s.deductions);
    });
    box.innerHTML =
      '<div class="st-row"><span>① 支払金額 合計</span><b>' + yen(pay) + '</b></div>' +
      '<div class="st-row"><span>給与所得 合計（②の合算）</span><b>' + yen(sal) + '</b></div>' +
      '<div class="st-row"><span>社会保険料等 合計</span><b>' + yen(soc) + '</b></div>' +
      '<div class="st-row"><span>③ 所得控除 合計</span><b>' + yen(dedS) + '</b></div>' +
      '<div class="st-row"><span>④ 源泉徴収税額 合計</span><b>' + yen(wh) + '</b></div>';
  }

  /* ---------- 所得控除の詳細入力 ---------- */
  function moneyInput(key, val, ph) {
    return '<input class="input" type="number" inputmode="numeric" min="0" step="1" data-ded="' + key + '" value="' + (val === undefined || val === null ? '' : esc(val)) + '" placeholder="' + esc(ph === undefined ? '0' : ph) + '">';
  }
  function out(key, text) { return '<span class="dout" data-out="' + key + '">' + text + '</span>'; }

  var dedOpenState = null;
  function syncDedOpen() {
    var box = $('dedDetail');
    if (!box) return;
    dedOpenState = [];
    Array.prototype.forEach.call(box.querySelectorAll('details'), function (dd) { dedOpenState.push(!!dd.open); });
  }

  function renderDeductions() {
    var box = $('dedDetail');
    if (!box) return;
    if (box.querySelector('details')) syncDedOpen();
    var d = dedOf();
    var y = state.year;
    var lifeAuto = d.lifeMode !== 'manual';
    var quakeAuto = d.quakeMode !== 'manual';
    var basicVal = (d.basicManual === undefined || d.basicManual === null) ? '' : d.basicManual;

    var spouseOpts = [['0', 'なし'], ['380000', '配偶者（特別）控除 38万円'],
      ['330000', '同 33万円'], ['310000', '同 31万円'], ['260000', '同 26万円'], ['210000', '同 21万円'],
      ['160000', '同 16万円'], ['110000', '同 11万円'], ['60000', '同 6万円'], ['30000', '同 3万円']];
    function opt(v, label, cur) {
      var val = String(v).replace(/s$/, '');
      return '<option value="' + val + '"' + (String(cur || 0) === val ? ' selected' : '') + '>' + esc(label) + '</option>';
    }

    box.innerHTML =
      '<details class="dsec" open>' +
        '<summary><span class="ds-ico">🏥</span><span class="ds-main"><b>社会保険料</b><span>国民年金・国保・任意継続など、源泉徴収票に入っていない分</span></span><span class="ds-amt" data-out="socialTotal">—</span></summary>' +
        '<div class="ds-body">' +
          '<label class="field"><span>上記以外の社会保険料（国民年金・国民健康保険・任意継続・後期高齢者医療など）</span>' + moneyInput('socialExtra', d.socialExtra, '0') + '</label>' +
          '<div class="ds-calc">源泉徴収票の社会保険料 ' + out('socialSlip', '¥0') + ' ＋ ここに入れた額 ＝ <b>' + out('socialTotal2', '¥0') + '</b></div>' +
        '</div>' +
      '</details>' +

      '<details class="dsec">' +
        '<summary><span class="ds-ico">🏦</span><span class="ds-main"><b>小規模企業共済等掛金控除</b><span>iDeCo・小規模企業共済の掛金（全額控除）</span></span><span class="ds-amt">' + out('idecoTotal', '¥0') + '</span></summary>' +
        '<div class="ds-body">' +
          '<label class="field"><span>年間の掛金合計（iDeCo・共済掛金）</span>' + moneyInput('ideco', d.ideco, '0') + '</label>' +
          '<div class="ds-calc">掛金は全額が所得控除になります（上限は掛金の制度側で決まります）。</div>' +
        '</div>' +
      '</details>' +

      '<details class="dsec"' + (lifeAuto ? '' : '') + '>' +
        '<summary><span class="ds-ico">🛡️</span><span class="ds-main"><b>生命保険料控除</b><span>一般・介護医療・個人年金／上限12万円</span></span><span class="ds-amt">' + out('lifeTotal', '¥0') + '</span></summary>' +
        '<div class="ds-body">' +
          '<div class="mode-toggle"><button class="mt-btn' + (lifeAuto ? ' on' : '') + '" data-ded-mode="life" data-mode="auto">支払額から自動計算</button><button class="mt-btn' + (!lifeAuto ? ' on' : '') + '" data-ded-mode="life" data-mode="manual">控除額を直接入力</button></div>' +
          '<div class="' + (lifeAuto ? '' : 'hidden') + '" data-life-panel="auto">' +
            '<label class="field"><span>新制度・一般生命保険料（2012年以降の契約）</span>' + moneyInput('lifeGenNew', d.lifeGenNew, '0') + '</label>' +
            '<label class="field"><span>旧制度・一般生命保険料（2011年以前の契約）</span>' + moneyInput('lifeGenOld', d.lifeGenOld, '0') + '</label>' +
            '<label class="field"><span>介護医療保険料</span>' + moneyInput('lifeCare', d.lifeCare, '0') + '</label>' +
            '<label class="field"><span>新制度・個人年金保険料</span>' + moneyInput('lifePenNew', d.lifePenNew, '0') + '</label>' +
            '<label class="field"><span>旧制度・個人年金保険料</span>' + moneyInput('lifePenOld', d.lifePenOld, '0') + '</label>' +
            '<div class="ds-calc">各区分の控除：' + out('lifeEach', '—') + '／合計（上限12万円）＝ <b>' + out('lifeTotal2', '¥0') + '</b></div>' +
          '</div>' +
          '<div class="' + (lifeAuto ? 'hidden' : '') + '" data-life-panel="manual">' +
            '<label class="field"><span>控除証明書の「控除額」（合計12万円まで）</span>' + moneyInput('lifeManual', d.lifeManual, '0') + '</label>' +
          '</div>' +
        '</div>' +
      '</details>' +

      '<details class="dsec">' +
        '<summary><span class="ds-ico">🌏</span><span class="ds-main"><b>地震保険料控除</b><span>地震保険＋旧長期損害保険／上限5万円</span></span><span class="ds-amt">' + out('quakeTotal', '¥0') + '</span></summary>' +
        '<div class="ds-body">' +
          '<div class="mode-toggle"><button class="mt-btn' + (quakeAuto ? ' on' : '') + '" data-ded-mode="quake" data-mode="auto">支払額から自動計算</button><button class="mt-btn' + (!quakeAuto ? ' on' : '') + '" data-ded-mode="quake" data-mode="manual">控除額を直接入力</button></div>' +
          '<div class="' + (quakeAuto ? '' : 'hidden') + '" data-quake-panel="auto">' +
            '<div class="grid2">' +
            '<label class="field"><span>地震保険料</span>' + moneyInput('quake', d.quake, '0') + '</label>' +
            '<label class="field"><span>旧長期損害保険料</span>' + moneyInput('quakeOld', d.quakeOld, '0') + '</label>' +
            '</div>' +
            '<div class="ds-calc">地震保険料は5万円まで、旧長期は最大1万円。合わせて <b>' + out('quakeTotal2', '¥0') + '</b>（上限5万円）</div>' +
          '</div>' +
          '<div class="' + (quakeAuto ? 'hidden' : '') + '" data-quake-panel="manual">' +
            '<label class="field"><span>控除証明書の控除額（5万円まで）</span>' + moneyInput('quakeManual', d.quakeManual, '0') + '</label>' +
          '</div>' +
        '</div>' +
      '</details>' +

      '<details class="dsec">' +
        '<summary><span class="ds-ico">💊</span><span class="ds-main"><b>医療費控除</b><span>支払額−補填額−足切り（10万円か所得の5%）</span></span><span class="ds-amt">' + out('medTotal', '¥0') + '</span></summary>' +
        '<div class="ds-body">' +
          '<div class="grid2">' +
          '<label class="field"><span>1年間の医療費（支払額）</span>' + moneyInput('medPaid', d.medPaid, '0') + '</label>' +
          '<label class="field"><span>補填額（保険金・高額療養費など）</span>' + moneyInput('medReimb', d.medReimb, '0') + '</label>' +
          '</div>' +
          '<div class="ds-calc">足切り ' + out('medFloor', '¥100,000') + ' を引いて <b>' + out('medTotal2', '¥0') + '</b>（上限200万円）。セルフメディケーション税制との併用はできません。</div>' +
        '</div>' +
      '</details>' +

      '<details class="dsec">' +
        '<summary><span class="ds-ico">🎁</span><span class="ds-main"><b>寄附金控除</b><span>ふるさと納税・認定NPOなど（寄附額−2,000円）</span></span><span class="ds-amt">' + out('donTotal', '¥0') + '</span></summary>' +
        '<div class="ds-body">' +
          '<label class="field"><span>1年間の寄附金の合計（ふるさと納税を含む）</span>' + moneyInput('donation', d.donation, '0') + '</label>' +
          '<div class="ds-calc">寄附額−2,000円 ＝ <b>' + out('donTotal2', '¥0') + '</b>（総所得の40%が限度）。ワンストップ特例を使った分も、確定申告するならここに入れて構いません。</div>' +
        '</div>' +
      '</details>' +

      '<details class="dsec">' +
        '<summary><span class="ds-ico">👪</span><span class="ds-main"><b>配偶者・扶養・障害者など</b><span>人的控除</span></span><span class="ds-amt">' + out('humanTotal', '¥0') + '</span></summary>' +
        '<div class="ds-body">' +
          '<label class="field"><span>配偶者控除・配偶者特別控除</span><select class="input" data-ded="spouse">' +
            spouseOpts.map(function (o) { return opt(o[0], o[1], d.spouse); }).join('') +
          '</select></label>' +
          '<div class="grid2">' +
          '<label class="field"><span>扶養控除：一般（38万円/人）</span>' + moneyInput('depGeneral', d.depGeneral, '0') + '</label>' +
          '<label class="field"><span>扶養控除：特定19〜22歳（63万円/人）</span>' + moneyInput('depSpecific', d.depSpecific, '0') + '</label>' +
          '</div>' +
          '<label class="field"><span>扶養控除：老人70歳以上（48万円/人）</span>' + moneyInput('depElder', d.depElder, '0') + '</label>' +
          '<label class="field"><span>障害者・ひとり親・寡婦</span><select class="input" data-ded="disability">' +
            opt('0', 'なし', d.disability) + opt('270000', '障害者控除 27万円', d.disability) + opt('400000', '特別障害者 40万円', d.disability) + opt('750000', '同居特別障害者 75万円', d.disability) +
          '</select></label>' +
          '<label class="field"><span>ひとり親控除・寡婦控除</span><select class="input" data-ded="singleParent">' +
            opt('0', 'なし', d.singleParent) + opt('350000', 'ひとり親控除 35万円', d.singleParent) + opt('270000', '寡婦控除 27万円', d.singleParent) +
          '</select></label>' +
          '<div class="ds-calc">人的控除の小計 <b>' + out('humanTotal2', '¥0') + '</b></div>' +
        '</div>' +
      '</details>' +

      '<details class="dsec">' +
        '<summary><span class="ds-ico">📄</span><span class="ds-main"><b>基礎控除・その他・税額控除</b><span>基礎控除は年度別に自動／住宅ローンは税額控除</span></span><span class="ds-amt">' + out('basicOut', '¥0') + '</span></summary>' +
        '<div class="ds-body">' +
          '<label class="field"><span>基礎控除（' + y + '年分の自動額 ' + out('basicAuto', '¥—') + '／空欄なら自動）</span>' + moneyInput('basicManual', basicVal, '自動') + '</label>' +
          '<label class="field"><span>その他の所得控除（雑損控除など）</span>' + moneyInput('other', d.other, '0') + '</label>' +
          '<label class="field"><span>住宅借入金等特別控除（税額控除・年末残高から計算した控除額）</span>' + moneyInput('housing', d.housing, '0') + '</label>' +
          '<div class="ds-calc">基礎控除 <b>' + out('basicOut2', '¥0') + '</b> を所得控除に、住宅ローン控除は所得税から直接引きます。</div>' +
        '</div>' +
      '</details>' +

      '<div class="ded-total"><span>所得控除の額の合計</span><b data-out="dedTotal">¥0</b></div>' +
      '<div class="rate-note">事業の経費・青色申告特別控除と、ここの控除は別物です。控除証明書は12月〜1月に届くので、届いた順に入れていけば申告書の数字がそのまま出来ます。</div>';

    if (dedOpenState) {
      Array.prototype.forEach.call(box.querySelectorAll('details'), function (dd, i) {
        if (dedOpenState[i] !== undefined) dd.open = dedOpenState[i];
      });
    }
    refreshDed();
  }

  function refreshDed() {
    var sim = computeSim();
    var b = sim.breakdown;
    var d = b.d;
    var y = state.year;
    var basicAutoText = yen(b.basicAuto);

    function set(key, text) {
      Array.prototype.forEach.call(document.querySelectorAll('[data-out="' + key + '"]'), function (el) { el.textContent = text; });
    }
    set('socialSlip', yen(b.socialSlip));
    set('socialTotal', yen(b.social));
    set('socialTotal2', yen(b.social) + '（うち源泉徴収票 ' + yen(b.socialSlip) + '）');
    set('idecoTotal', yen(b.ideco));
    set('lifeTotal', yen(b.life));
    set('lifeTotal2', yen(b.life));
    set('lifeEach', '一般（新）' + yen(lifeCalcNew(d.lifeGenNew)) + '・一般（旧）' + yen(lifeCalcOld(d.lifeGenOld)) +
      '・介護医療' + yen(lifeCalcNew(d.lifeCare)) + '・年金（新）' + yen(lifeCalcNew(d.lifePenNew)) + '・年金（旧）' + yen(lifeCalcOld(d.lifePenOld)));
    set('quakeTotal', yen(b.quake));
    set('quakeTotal2', yen(b.quake));
    set('medTotal', yen(b.medical));
    set('medFloor', yen(b.medFloor));
    set('medTotal2', yen(b.medical));
    set('donTotal', yen(b.donation));
    set('donTotal2', yen(b.donation));
    set('humanTotal', yen(b.spouse + b.dependents + b.disability + b.singleParent));
    set('humanTotal2', yen(b.spouse + b.dependents + b.disability + b.singleParent));
    set('basicOut', yen(b.basic));
    set('basicAuto', basicAutoText + '（' + basicBandLabel(y, sim.totalIncome) + '）');
    set('basicOut2', yen(b.basic));
    set('dedTotal', yen(b.total));
    var cnt = [b.social, b.ideco, b.life, b.quake, b.medical, b.donation, b.spouse, b.dependents, b.disability + b.singleParent, b.other].filter(function (v) { return v > 0; }).length + 1;
    var dc = $('dedCount');
    if (dc) dc.textContent = cnt + '項目';
    renderSlipTotals();
  }

  function renderRiskResults() {
    var sum = $('riskSummary');
    var listBox = $('riskList');
    if (!sum || !listBox) return;

    var r = analyzeRisks();
    var sim = computeSim();

    $('riskCount').textContent = r.total + '件を診断';
    var hasData = r.total > 0 || sim.sales > 0 || sim.slipCount > 0 || sim.expenses > 0;
    var judgment = !hasData ? '未入力'
      : ((r.redCount > 0 || r.structural.some(function (s) { return s.level === 'red'; })) ? '要修正'
        : (r.amberCount > 0 || r.structural.length ? '要確認' : '問題なし'));

    var chip = $('bnChip');
    if (chip) {
      chip.className = 'bn-chip ' + (judgment === '問題なし' ? 'ok' : (judgment === '要修正' ? 'red' : (judgment === '未入力' ? 'gray' : 'amber')));
      chip.textContent = yearLabel(state.year) + '分：' + judgment;
    }
    var bnSub = $('bnLast');
    if (bnSub) bnSub.textContent = '指摘 ' + (r.flags.length + r.structural.length) + '件（赤' + r.redCount + '・黄' + r.amberCount + '）｜経費 ' + yen(sim.expenses + sim.dokujuu) + '／売上 ' + yen(sim.sales);

    sum.innerHTML =
      '<div class="risk-chips">' +
      '<div class="rchip red"><div class="n">' + r.redCount + '件</div><div class="t">🔴 すぐ外す・直す</div><div class="t">' + yen(r.redAmt) + '</div></div>' +
      '<div class="rchip amber"><div class="n">' + r.amberCount + '件</div><div class="t">🟡 要確認・按分</div><div class="t">' + yen(r.amberAmt) + '</div></div>' +
      '<div class="rchip ok' + (judgment === '問題なし' ? '' : ' warn') + '"><div class="n">' + (judgment === '問題なし' ? '✅ 問題なし' : (judgment === '要修正' ? '⚠️ 要修正' : (judgment === '未入力' ? '— 未入力' : '🔍 要確認'))) + '</div>' +
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
    var b = sim.breakdown;

    var rows = [];
    function neg(v) { return v ? '−' + yen(v) : '¥0'; }
    rows.push(['事業売上（収入）', yenSigned(sim.sales), '']);
    rows.push(['必要経費' + (sim.dokujuu ? '（専従者給与別）' : ''), neg(sim.expenses), '']);
    if (sim.dokujuu) rows.push(['専従者給与', neg(sim.dokujuu), '']);
    rows.push(['青色申告特別控除', neg(sim.blueApplied) + (sim.blueApplied < sim.blue ? '（所得が少ないため減額）' : ''), 'blue']);
    rows.push(['＝ 事業所得', yenSigned(sim.biz), sim.biz < 0 ? 'neg' : '']);
    if (sim.slipCount) rows.push(['＋ 給与所得（' + sim.slipCount + '枚の合算）', yenSigned(sim.salary), '']);
    rows.push(['＝ 合計所得金額', yenSigned(sim.totalIncome), '']);
    rows.push(['所得控除の内訳（以下を合計）', '', '']);

    var detail = [
      ['　社会保険料控除', b.social],
      ['　小規模企業共済等掛金控除', b.ideco],
      ['　生命保険料控除', b.life],
      ['　地震保険料控除', b.quake],
      ['　医療費控除', b.medical],
      ['　寄附金控除', b.donation],
      ['　配偶者（特別）控除', b.spouse],
      ['　扶養控除', b.dependents],
      ['　障害者・ひとり親・寡婦控除', b.disability + b.singleParent],
      ['　基礎控除', b.basic],
      ['　その他', b.other]
    ];
    detail.forEach(function (r2) {
      if (!r2[1]) return;
      rows.push([r2[0], '−' + yen(r2[1]), '']);
    });
    rows.push(['所得控除の額の合計', neg(sim.ded), '']);
    rows.push(['＝ 課税される所得金額', yenSigned(sim.taxable), 'blue']);
    if (sim.housing) {
      rows.push(['所得税（累進税率）', yen(sim.incomeTaxBase), '']);
      rows.push(['住宅借入金等特別控除（税額控除）', neg(sim.housing), '']);
      rows.push(['所得税（控除後）', yen(sim.incomeTax), '']);
    } else {
      rows.push(['所得税（累進税率）', yen(sim.incomeTax), '']);
    }
    rows.push(['復興特別所得税（2.1%）', yen(sim.fukko), '']);
    rows.push(['所得税及び復興特別所得税', yen(sim.totalTax), '']);
    if (sim.slipCount) rows.push(['源泉徴収税額（前払い済み）', neg(sim.withheld), '']);

    var html = '<div>';
    rows.forEach(function (r, i) {
      var isBig = /^＝/.test(r[0]);
      html += '<div class="sim-row' + (isBig ? ' big' : '') + '"><span class="k">' + esc(r[0]) + '</span><span class="v ' + r[2] + '">' + r[1] + '</span></div>';
    });
    html += '</div>';

    if (sim.diff > 0) {
      html += '<div class="sim-box pay"><div class="lbl">追加で納める税金の見込み</div><div class="amt">' + yen(sim.diff) + '</div></div>';
    } else if (sim.diff < 0) {
      html += '<div class="sim-box refund"><div class="lbl">還付される見込み</div><div class="amt">' + yen(-sim.diff) + '</div></div>';
    } else {
      html += '<div class="sim-box zero"><div class="lbl">納付も還付もなし</div><div class="amt">¥0</div></div>';
    }

    /* 入力の不整合・上限チェック */
    if (sim.warnings.length) {
      html += '<div class="val-box"><div class="val-title">入力チェック（' + sim.warnings.length + '件）</div>';
      sim.warnings.forEach(function (w) {
        html += '<div class="val-item ' + w.level + '"><span class="vi">' + (w.level === 'warn' ? '⚠️' : 'ℹ️') + '</span><span>' + esc(w.text) + '</span></div>';
      });
      html += '</div>';
    }

    var notes = '';
    if (sim.sales === 0) {
      notes += '<div class="warn-box">売上が未入力です。売上を入れると課税所得と還付額が計算できます。</div>';
    } else if (sim.biz < 0 && sim.salary > 0) {
      notes += '<div class="info-box">事業所得が赤字（' + yenSigned(sim.biz) + '）なので、給与所得と損益通算して課税所得を下げています。この通算が認められるのは「事業所得」と認められた場合だけです（雑所得は通算できません）。</div>';
    }
    notes += '<div class="resident-box"><span>住民税の概算（翌年6月〜）</span><b>' + yen(sim.residentTax) + '</b><span class="rb-sub">所得割 ' + yen(Math.max(0, sim.residentTax - 5000)) + ' ＋ 均等割 約¥5,000（自治体により異なります）</span></div>';
    notes += '<div class="hint-ok">所得税は 課税所得×累進税率−控除額＋復興特別所得税2.1%、そこから住宅ローン控除と源泉徴収税額を引いた額です。基礎控除・給与所得控除は国税庁の年度別表（' + sim.year + '年分）で自動計算しています。住民税は基礎控除' + (sim.year >= 2026 ? '45万円' : '43万円') + 'として概算。最終確認はe-Taxの計算結果で行ってください。</div>';
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
    var isCsv = /\.csv$/i.test(file.name) || /csv/i.test(file.type || '');
    var reader = new FileReader();
    reader.onload = function (ev) {
      try {
        var wb;
        if (isCsv) {
          wb = XLSX.read(decodeCsvText(ev.target.result), { type: 'string' });
        } else {
          wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array', cellDates: false });
        }
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

  /* CSVはUTF-8/Shift_JISの両方を扱う（文字化けで見出しが読めなくなるのを防ぐ） */
  function decodeCsvText(buf) {
    var bytes = new Uint8Array(buf);
    if (bytes.length > 2 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
      return new TextDecoder('utf-8').decode(bytes.subarray(3));
    }
    var text = '';
    try { text = new TextDecoder('utf-8', { fatal: false }).decode(bytes); } catch (e) { text = ''; }
    if (text.indexOf('\uFFFD') < 0) return text;
    try { return new TextDecoder('shift_jis').decode(bytes); } catch (e2) { return text; }
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
  /* ============================================================
     減価償却リスト（固定資産台帳）の表示
     ============================================================ */
  function depField(a, f, type, cls, ph) {
    var v = a[f];
    if (v === undefined || v === null) v = '';
    if (type === 'number') {
      return '<input class="dep-in ' + cls + '" type="number" inputmode="numeric" min="0" step="1" data-dep="' + esc(a.id) + '" data-f="' + f + '" value="' + esc(v) + '" placeholder="' + esc(ph || '') + '">';
    }
    return '<input class="dep-in ' + cls + '" type="text" data-dep="' + esc(a.id) + '" data-f="' + f + '" value="' + esc(v) + '" placeholder="' + esc(ph || '') + '">';
  }

  function depNoteRowHTML(r) {
    var a = r.a, c = r.c, ts = r.taxsnap;
    var bits = [];
    if (ts) {
      var d = c.expense - ts.amount;
      bits.push('<span class="dep-badge ' + (d === 0 ? 'ok' : 'warn') + '">スプレッドシート ' + yen(ts.amount) + '／差 ' + (d === 0 ? '0円' : (d > 0 ? '+' : '−') + yen(Math.abs(d))) + '</span>');
    } else {
      bits.push('<span class="dep-badge info">スプレッドシート計上なし</span>');
    }
    if (c.override) bits.push('<span class="dep-badge warn">必要経費算入額を手入力中（自動 ' + yen(c.autoExpense) + '）</span>');
    if (c.cost > 0) {
      var cmp = depCompare(a, state.year).map(function (x) {
        return (x.label || x.method) + ' ' + yen(x.expense);
      }).join(' ／ ');
      bits.push('<span class="dep-badge info">方法比較：' + cmp + '</span>');
    }
    var notes = c.notes.slice();
    if (a.memo) notes.push('メモ：' + a.memo);
    return '<tr class="dep-note-row" data-noterow="' + esc(a.id) + '"><td colspan="12">' +
      '<div>' + bits.join(' ') + '</div>' +
      (notes.length ? '<div style="margin-top:3px;">' + notes.map(function (n) { return '・' + esc(n); }).join('<br>') + '</div>' : '') +
      '<div class="dep-note-in">' +
      '<span style="font-size:10px;color:#6b7280;">前年末未償却残高</span>' +
      '<input class="dep-in dep-ov" type="number" inputmode="numeric" min="0" data-dep="' + esc(a.id) + '" data-f="opening" value="' + esc(a.opening === undefined || a.opening === null ? '' : a.opening) + '" placeholder="' + (c.cost && c.accumBefore ? c.cost - c.accumBefore : '任意') + '">' +
      '<span style="font-size:10px;color:#6b7280;">算入額の手入力</span>' +
      '<input class="dep-in dep-ov" type="number" inputmode="numeric" min="0" data-dep="' + esc(a.id) + '" data-f="override" value="' + esc(a.override === undefined || a.override === null ? '' : a.override) + '" placeholder="' + (c.cost ? c.autoExpense : '—') + '">' +
      '<button class="mini-btn" data-depact="copy" data-id="' + esc(a.id) + '">⧉ この行をコピー</button>' +
      '<button class="mini-btn" data-depact="del" data-id="' + esc(a.id) + '" style="color:#b91c1c;border-color:#fecaca;">🗑️ 削除</button>' +
      '</div></td></tr>';
  }

  function depRowHTML(r) {
    var a = r.a, c = r.c;
    var methodSel = DEP_METHODS.map(function (m) {
      return '<option value="' + m + '"' + (c.method === m ? ' selected' : '') + '>' + m + '</option>';
    }).join('');
    return '<tr class="dep-row" data-row="' + esc(a.id) + '">' +
      '<td data-label="資産の名称等">' + depField(a, 'name', 'text', 'name', '例：ノートパソコン') + '</td>' +
      '<td data-label="取得年月日">' + depField(a, 'acquired', 'text', '', '2026-03') + '</td>' +
      '<td data-label="取得価額">' + depField(a, 'cost', 'number', 'money', '0') + '</td>' +
      '<td data-label="償却方法"><select class="dep-in" data-dep="' + esc(a.id) + '" data-f="method">' + methodSel + '</select></td>' +
      '<td data-label="耐用年数">' + depField(a, 'life', 'number', '', '5') + '</td>' +
      '<td data-label="償却率">' + depField(a, 'rate', 'number', '', '0.200') + '</td>' +
      '<td data-label="本年中の償却期間">' + depField(a, 'months', 'number', '', '12') + '</td>' +
      '<td data-label="本年分の普通償却費" class="num calc"><span class="dep-calc" data-outof="' + esc(a.id) + ':ordinary">' + (c.cost ? yen(c.ordinary) : '—') + '</span></td>' +
      '<td data-label="事業専用割合">' + depField(a, 'bizPct', 'number', '', '100') + '</td>' +
      '<td data-label="本年分の必要経費算入額" class="num calc"><span class="dep-calc" data-outof="' + esc(a.id) + ':expense" style="color:#0284c7;">' + (c.cost ? yen(c.expense) : '—') + '</span></td>' +
      '<td data-label="期末未償却残高" class="num calc"><span class="dep-calc" data-outof="' + esc(a.id) + ':bookEnd">' + (c.cost ? yen(c.bookEnd) : '—') + '</span></td>' +
      '<td data-label="メモ">' + depField(a, 'memo', 'text', 'memo', '摘要（例：措置法28条の2）') + '</td>' +
      '</tr>' + depNoteRowHTML(r);
  }

  function depSummaryHTML(s) {
    var diffTxt = s.diff === 0 ? '±0円' : ((s.diff > 0 ? '+' : '−') + yen(Math.abs(s.diff)));
    return '<div class="dep-sum" id="depSum">' +
      '<div class="dep-kpi"><div class="l">資産数</div><div class="v">' + s.count + '件</div><div class="n">取得価額 ' + yen(s.cost) + '</div></div>' +
      '<div class="dep-kpi"><div class="l">本年分の普通償却費</div><div class="v">' + yen(s.ordinary) + '</div><div class="n">事業専用割合を掛ける前</div></div>' +
      '<div class="dep-kpi"><div class="l">必要経費算入額（台帳合計）</div><div class="v" style="color:#0284c7;">' + yen(s.expense) + '</div><div class="n">e-Tax 3ページ目の合計</div></div>' +
      '<div class="dep-kpi' + (s.diff ? ' amber' : '') + '"><div class="l">経費帳の減価償却費</div><div class="v">' + yen(s.taxsnap) + '</div><div class="n">差 ' + diffTxt + (s.overridden ? '／手入力 ' + s.overridden + '件' : '') + '</div></div>' +
      (s.missing.length ? '<div class="dep-kpi full amber"><div class="l">取得価額が未入力</div><div class="v" style="font-size:12px;">' + s.missing.map(esc).join('、') + '</div><div class="n">分かっている範囲で入れると、法定計算と比較できます（空欄でも台帳は保存されます）。</div></div>' : '') +
      '</div>';
  }

  function renderDep() {
    var box = $('depBody');
    if (!box) return;
    var s = depSummary();
    var cnt = $('depCount');
    if (cnt) cnt.textContent = s.count + '件';
    var body = s.rows.map(depRowHTML).join('');
    if (!s.count) {
      body = '<tr><td colspan="12" style="text-align:center;padding:18px 8px;color:#9ca3af;font-size:12px;line-height:1.6;">まだ資産が登録されていません。<br>「＋ 資産を追加」から、e-Taxの3ページ目と同じ並びで入力できます。</td></tr>';
    }
    box.innerHTML = depSummaryHTML(s) +
      '<div class="dep-wrap"><table class="dep"><thead><tr>' +
      '<th>資産の名称等</th><th>取得年月日</th><th>取得価額</th><th>償却方法</th><th>耐用年数</th><th>償却率</th><th>本年中の<br>償却期間</th><th>本年分の<br>普通償却費</th><th>事業専用<br>割合(%)</th><th>本年分の必要<br>経費算入額</th><th>期末未<br>償却残高</th><th>メモ</th>' +
      '</tr></thead><tbody>' + body + '</tbody></table></div>' +
      '<div class="scroll-hint">📱 スマホでは横にスクロールできます（各項目の見出しは左に出ます）。</div>' +
      '<div class="dep-actions">' +
      '<button class="btn-ghost" data-depact="add">＋ 資産を追加</button>' +
      '<button class="btn-ghost" data-depact="copyall">⧉ 台帳をTSVで全部コピー</button>' +
      '<button class="btn-ghost" data-depact="toexpense">📥 この台帳をExcel（減価償却資産台帳シート）に出す</button>' +
      depLocalButtons() +
      '</div>' +
      '<div class="rate-note">e-Taxの画面では「資産の名称等 → 取得年月日 → 取得価額 → 償却方法 → 耐用年数 → 償却率 → 本年中の償却期間 → 本年分の普通償却費 → 事業専用割合 → 本年分の必要経費算入額 → 未償却残高」の順に並んでいます。各資産の「⧉ この行をコピー」で、その順のタブ区切りがコピーされます。金額は円未満切り捨ての概算です。<br>少額特例・消耗品費の一括計上は、原則として取得した年（事業に使い始めた年）の処理です。年をまたぐ資産は、前年末未償却残高と償却方法を確認してから申告してください。</div>';
  }

  function depLocalButtons() {
    return '';
  }

  function depFind(id) {
    return assetsOf().find(function (x) { return x.id === id; }) || null;
  }
  function depCopyRow(a) {
    var c = depCalc(a, state.year);
    var v = [a.name || '', a.acquired || '', c.cost || '', c.method, a.life || '', c.rate ? c.rate.toFixed(3) : '', String(c.months), String(c.ordinary), String(c.pct), String(c.expense), String(c.bookEnd === '' ? '' : c.bookEnd)];
    copyText(v.join('\t'), '「' + (a.name || '資産') + '」の行をコピーしました（e-Taxの同じ順に貼り付け）');
  }
  function depCopyAll() {
    var s = depSummary();
    if (!s.count) { toast('資産がまだありません', true); return; }
    var lines = [['資産の名称等', '取得年月日', '取得価額', '償却方法', '耐用年数', '償却率', '本年中の償却期間', '本年分の普通償却費', '事業専用割合', '本年分の必要経費算入額', '未償却残高', 'メモ'].join('\t')];
    s.rows.forEach(function (r) {
      var a = r.a, c = r.c;
      lines.push([a.name || '', a.acquired || '', c.cost || '', c.method, a.life || '', c.rate ? c.rate.toFixed(3) : '', c.months, c.ordinary, c.pct, c.expense, c.bookEnd === '' ? '' : c.bookEnd, (a.memo || '').replace(/\t/g, ' ')].join('\t'));
    });
    copyText(lines.join('\n'), '減価償却リストをTSVでコピーしました');
  }
  function depNewAsset() {
    var a = { id: 'dep-' + uid(), name: '', acquired: '', cost: '', method: '定額法', life: 5, rate: DEP_RATES[5], months: 12, bizPct: 100, memo: '', opening: '', override: '', source: 'manual' };
    assetsOf().push(a);
    persist();
    renderDep();
    var el = document.querySelector('[data-row="' + a.id + '"] [data-f="name"]');
    if (el) el.focus();
  }
  function depRestore() {
  }
  function refreshDepLive(id) {
    var a = depFind(id);
    if (!a) return;
    var y = state.year;
    var c = depCalc(a, y);
    function put(k, val) {
      var el = document.querySelector('[data-outof="' + id + ':' + k + '"]');
      if (el) el.textContent = c.cost ? val : '—';
    }
    put('ordinary', yen(c.ordinary));
    put('expense', yen(c.expense));
    put('bookEnd', yen(c.bookEnd));
    var s = depSummary();
    var sum = $('depSum');
    if (sum) {
      var tmp = document.createElement('div');
      tmp.innerHTML = depSummaryHTML(s);
      if (sum.parentNode) sum.parentNode.replaceChild(tmp.firstChild, sum);
    }
    var note = document.querySelector('[data-noterow="' + id + '"]');
    if (note && document.activeElement !== note.querySelector('[data-f="override"]')) {
      var r = { a: a, c: c, taxsnap: depTaxsnapFor(a, depTaxsnapRows(y)) };
      var tmp2 = document.createElement('tbody');
      tmp2.innerHTML = depNoteRowHTML(r);
      var fresh = tmp2.firstChild;
      if (fresh) note.parentNode.replaceChild(fresh, note);
    }
    var cnt = $('depCount');
    if (cnt) cnt.textContent = s.count + '件';
  }
  function depInputHandler(ev) {
    var el = ev.target.closest('[data-dep]');
    if (!el) return;
    var a = depFind(el.getAttribute('data-dep'));
    if (!a) return;
    var f = el.getAttribute('data-f');
    var isNum = el.type === 'number';
    a[f] = isNum ? (el.value === '' ? '' : Math.max(0, Number(el.value))) : el.value;
    if (f === 'life' && isNum && Number(el.value) > 0) {
      a.rate = DEP_RATES[Number(el.value)] || a.rate;
      var rateEl = document.querySelector('[data-dep="' + a.id + '"][data-f="rate"]');
      if (rateEl) rateEl.value = a.rate;
    }
    if (f === 'acquired') {
      var acq = depAcquired(a);
      if (acq && acq.year === state.year) {
        var mEl = document.querySelector('[data-dep="' + a.id + '"][data-f="months"]');
        if (mEl && !Number(a.months)) { /* keep */ }
      }
    }
    persist();
    refreshDepLive(a.id);
    if (f === 'method' || f === 'life') renderBoard();
  }

  /* ============================================================
     PC版e-Tax 画面どおり入力ナビ
     ============================================================ */
  function guideLocalNotes(id) {
    return '';
  }

  function pcSteps() {
    var y = state.year;
    var js = '令和' + reiwa(y) + '年分';
    return [
      {
        id: 'start', no: 1, ico: '🚪', title: '作成コーナーを開いて「作成開始」', sub: '決算書・収支内訳書（＋所得税）を選ぶところまで',
        items: [
          { t: '用意するもの', d: 'マイナンバーカード（利用者証明用パスワード4桁・署名用6〜16桁）、スマホのマイナポータルアプリ、源泉徴収票、控除証明書、昨年の申告データ（.data）', path: '' },
          { t: '「作成開始」を押す', d: '確定申告書等作成コーナー（keisan.nta.go.jp）のトップで「作成開始」→ 注意事項の画面で「次へ」。', path: '作成開始' },
          { t: '作るものを選ぶ', d: '「作成する申告書等の選択」で年分「' + js + '」→「決算書・収支内訳書（＋所得税）」。ここが大事です（「所得税」から入ると決算書が作れません）。', path: '決算書・収支内訳書（＋所得税）' },
          { t: '提出方法に答える', d: '「マイナンバーカードをお持ちですか」→はい／「スマホで読み取れますか」→はい／マイナポータル連携→「利用する」（任意）。', path: 'e-Tax（マイナンバーカード方式）' },
          { t: 'カードを読み取る', d: 'PC画面のQRコードをスマホで読むか、ICカードリーダーで読み取り。利用者証明用パスワード（4桁）を入力。初めてなら利用者識別番号を発行して控えます。', path: 'マイナンバーカード読取' },
          { t: '昨年のデータを引き継ぐ', d: 'トップの「保存データを利用して作成」から .data を選ぶと、住所・氏名・利用者識別番号・前年末未償却残高が入った状態で始められます。', path: '保存データを利用して作成' }
        ],
        tip: 'PCでは Ctrl＋F（Macは ⌘＋F）で画面内の項目名を検索できます。ブラウザの「戻る」は使わず、画面の「前へ」で戻ってください。',
        board: ''
      },
      {
        id: 'page1', no: 2, ico: '1️⃣', title: '青色申告決算書 1ページ目（損益計算書）', sub: '売上 → 経費19科目 → 青色申告特別控除',
        items: [
          { t: '売上（収入）金額', d: '売上先が2件までなら直接入力。3件以上は「売上（収入）金額の内訳」を開いて入力します。', path: '売上（収入）金額' },
          { t: '売上原価', d: '期首商品棚卸高・仕入金額・期末商品棚卸高。モノを仕入れないなら0のままでOK（仕入高がある年はその金額）。', path: '売上原価' },
          { t: '経費を科目ごとに転記', d: '下の「1ページ目」ボードの金額を、e-Taxの同じ科目名の欄にコピーします。会議費・新聞図書費・支払手数料は「雑費」、車両費は「旅費交通費」にまとまります。', path: '経費' },
          { t: '減価償却費', d: '3ページ目で入力した金額が自動で入ります。ここには手入力しません。', path: '減価償却費' },
          { t: '青色申告特別控除額', d: '65万円（e-Tax送信＋電子帳簿保存）／55万円（e-Taxのみ）／10万円。赤字で所得が少ない年は、所得金額が限度になります（自動で減額されます）。', path: '青色申告特別控除額' },
          { t: '所得金額', d: '赤字でもマイナスのまま進めます。0にすると給与所得との損益通算ができません。', path: '所得金額' }
        ],
        tip: '決算書の所得金額と、確定申告書の「事業所得」が一致しているか最後に見比べてください。',
        board: 'p1'
      },
      {
        id: 'page2', no: 3, ico: '2️⃣', title: '決算書 2〜3ページ目（月別売上・地代家賃・減価償却）', sub: '3ページ目の「減価償却費の計算」は台帳からそのままコピー',
        items: [
          { t: '月別売上（収入）金額（2ページ目）', d: '1月〜12月の売上を入れます。下の「月別売上」のとおりに入れると合計が売上金額と一致します。', path: '月別売上（収入）金額' },
          { t: '地代家賃の内訳（2ページ目）', d: '賃借人・年間賃借料・左のうち必要経費算入額。家事按分している場合は満額と按分後を分けて入れます。', path: '地代家賃の内訳' },
          { t: '専従者給与の内訳（2ページ目）', d: '家族に給与を払っていなければ空欄のままです。', path: '専従者給与の内訳' },
          { t: '減価償却費の計算（3ページ目）', d: '「資産の名称等」から順に、下の台帳の行をコピーして貼り付けます。行が足りなければ「明細を追加」で増やせます。', path: '減価償却費の計算' },
          { t: '少額特例・一括償却の扱い', d: '30万円未満の資産を少額特例で全額経費にする場合も、この明細に載せて償却方法と必要経費算入額を入れます。20万円未満なら一括償却（3年）も選べます。', path: '償却方法' },
          { t: '貸倒金・繰越損失', d: 'なければ空欄。前年からの赤字を繰り越す場合は「青色申告決算書」ではなく確定申告書の「繰越損失」で入力します。', path: '' }
        ],
        tip: '「本年中の償却期間」は、取得した年は 13−取得月（3月取得なら10ヶ月）、翌年以降は12ヶ月です。',
        board: 'dep'
      },
      {
        id: 'form', no: 4, ico: '3️⃣', title: '確定申告書 第一表・第二表', sub: '給与所得＋事業所得＋所得控除 → 税額',
        items: [
          { t: '決算書から引き継ぐ', d: '決算書の入力が終わったら「所得税の申告書作成はこちら」→「引継ぎ情報等の確認」→「所得税の申告書の作成画面へ」。事業所得が自動で入ります。', path: '所得税の申告書作成はこちら' },
          { t: '給与所得（第一表）', d: '源泉徴収票の「支払金額」と「所得控除後の金額」。前職があれば「前職分」の欄も。下の「確定申告書」ボードの金額と照合してください。', path: '収入金額等・所得金額等' },
          { t: '所得控除（第一表）', d: '社会保険料・生命保険料・地震保険料・医療費・寄附金・基礎控除。マイナポータル連携していれば自動で入ります。', path: '所得控除' },
          { t: '税額と還付口座', d: '課税所得→所得税→復興特別所得税は自動計算。還付になる場合は「還付される税金」と受取口座（本人名義）を確認します。', path: '税金の計算・還付金の受取口座' },
          { t: '第二表', d: '「所得の内訳」に給与（源泉徴収票の支払者）と事業（屋号・事業場所）を書きます。給与は源泉徴収票のとおりに転記。', path: '第二表' },
          { t: '住民税の徴収方法', d: '給与以外の所得があると「自分で納付（普通徴収）」か「給与から差し引き（特別徴収）」を選べます。', path: '住民税・事業税に関する事項' }
        ],
        tip: '還付申告は' + (y + 1) + '年1月から提出できます。期限前でも、源泉徴収票がそろっていれば先に出せます。',
        board: 'form'
      },
      {
        id: 'send', no: 5, ico: '4️⃣', title: '送信・控え保存・納付', sub: '送信前確認 → 署名 → 受付番号 → PDFで7年保存',
        items: [
          { t: '入力内容の確認', d: '「入力内容の確認」で帳票イメージを表示。収入・所得・控除・税額を源泉徴収票と1つずつ照合します。', path: '入力内容の確認' },
          { t: '申告書を送信', d: '「送信する」→「送信を実行する」。署名用電子証明書のパスワード（6〜16桁）を聞かれます。', path: '送信する' },
          { t: '受付番号を控える', d: '「送信完了」の画面と受付番号・受付日時をメモ。e-Taxソフト（WEB版）のメッセージボックスにも届きます。', path: '送信結果の確認' },
          { t: '控えPDFを保存', d: '「帳票一覧」から申告書・青色申告決算書をPDFで保存。所得税の帳簿書類は7年保存（欠損金の繰越は9年）。', path: '帳票一覧' },
          { t: '納付（納税がある場合）', d: '振替納税・ダイレクト納付・コンビニ納付（QR）・クレジットカード納付（手数料あり）。振替納税の登録は3月中旬までです。', path: '納付' },
          { t: '還付を待つ', d: '登録口座へ通常3〜4週間。「還付金の処理状況」で確認できます。', path: '' }
        ],
        tip: '送信は24時間できますが、期限ぎりぎりはサーバーが混みます。' + deadlineOf(y).end + 'の3日前までに送っておくと安心です。',
        board: ''
      }
    ];
  }

  function boardRowHTML(label, src, v, cls) {
    var val = Math.round(Number(v) || 0);
    return '<div class="board-row' + (cls === 'total' ? ' board-total' : '') + '">' +
      '<div class="bl"><b>' + esc(label) + '</b><span>' + esc(src || '') + '</span></div>' +
      '<div class="bv' + (val < 0 ? ' neg' : '') + '">' + (val < 0 ? '▲' : '') + Math.abs(val).toLocaleString('ja-JP') + '</div>' +
      '<button class="copy-btn" data-copy="' + val + '" aria-label="コピー">⧉</button>' +
      '</div>';
  }

  function boardP1HTML() {
    var sim = computeSim();
    var tot = totalsForYear();
    var rows = [];
    rows.push(boardRowHTML('売上（収入）金額', '税務診断タブの「年間の事業売上」', sim.sales));
    FORM_CATEGORIES.forEach(function (c) {
      if (c === '専従者給与') return;
      var v = tot.byForm[c] || 0;
      if (!v) return;
      var raws = Object.keys(tot.byRaw).filter(function (raw) { return formCategoryOf(raw) === c; });
      var n = raws.reduce(function (s, raw) { return s + (tot.rawCount[raw] || 0); }, 0);
      var src = '経費帳 ' + n + '件：' + raws.join('・');
      if (raws.length > 1) src += '（決算書では1欄にまとめます）';
      if (c === '減価償却費') {
        var ds = depSummary();
        if (ds.count) {
          src = '固定資産台帳の算入額合計 ' + yen(ds.expense) + '／経費帳は ' + yen(v) + '（差 ' + (ds.diff === 0 ? '±0' : (ds.diff > 0 ? '+' : '−') + yen(Math.abs(ds.diff))) + '）';
        }
      }
      rows.push(boardRowHTML(c, src, v));
    });
    rows.push(boardRowHTML('経費計', '経費の合計（専従者給与を除く）', sim.expenses, 'total'));
    if (sim.dokujuu) rows.push(boardRowHTML('専従者給与（事業専従者控除）', '経費計の下の別欄', sim.dokujuu));
    rows.push(boardRowHTML('青色申告特別控除前の所得金額', '自動計算（売上−経費−専従者給与）', sim.beforeBlue));
    rows.push(boardRowHTML('青色申告特別控除額', '選択中の控除額' + (sim.blueApplied < sim.blue ? '／所得を限度に減額' : ''), sim.blueApplied));
    rows.push(boardRowHTML('所得金額', '自動計算。赤字ならマイナスのまま', sim.biz, 'total'));
    var depWarn = '';
    var depS = depSummary();
    if (depS.count && depS.diff !== 0) {
      depWarn = '<div class="pc-note" style="background:#fffcf3;border-color:#f2e6c9;color:#7a5316;"><b>先に減価償却の方法を決めてください。</b>e-Taxでは3ページ目「減価償却費の計算」の合計が1ページ目の減価償却費に自動で入ります。このアプリの固定資産台帳では <b>' + yen(depS.expense) + '</b>、経費帳（スプレッドシート）では <b>' + yen(depS.taxsnap) + '</b> で、差が <b>' + (depS.diff > 0 ? '+' : '−') + yen(Math.abs(depS.diff)) + '</b> あります。台帳の各資産のメモ（4ヶ月/12ヶ月、少額特例、一括償却など）を見て決め、経費帳側の減価償却費も同じ額にそろえてください。</div>';
    }
    return '<div class="board-sec"><h4>青色申告決算書 1ページ目（損益計算書）</h4>' + depWarn + rows.join('') + '</div>';
  }

  function boardFormHTML() {
    var sim = computeSim();
    var rows = [];
    if (sim.slipCount) {
      rows.push(boardRowHTML('収入金額等「給与」', '源泉徴収票①の合計', sim.pay));
      rows.push(boardRowHTML('所得金額等「給与所得」', '源泉徴収票②（合算後）', sim.salary));
    }
    rows.push(boardRowHTML('所得金額等「事業所得（営業等）」', '決算書の所得金額と同じ数字', sim.biz));
    rows.push(boardRowHTML('合計所得金額', '給与所得＋事業所得', sim.totalIncome, 'total'));
    var b = sim.breakdown;
    [['社会保険料控除', '源泉徴収票の社会保険料＋国保・年金など', b.social],
     ['小規模企業共済等掛金控除', 'iDeCo・共済掛金', b.ideco],
     ['生命保険料控除', '一般・介護医療・個人年金（上限12万円）', b.life],
     ['地震保険料控除', '地震保険＋旧長期（上限5万円）', b.quake],
     ['医療費控除', '支払額−補填額−足切り' + yen(b.medFloor), b.medical],
     ['寄附金控除', 'ふるさと納税など（寄附額−2,000円）', b.donation],
     ['配偶者（特別）控除', 'この画面で選択した額', b.spouse],
     ['扶養控除', '一般38万・特定63万・老人48万', b.dependents],
     ['障害者・ひとり親・寡婦控除', '選択した区分の額', b.disability + b.singleParent],
     ['基礎控除', sim.year + '年分の自動計算' + (b.basicSet ? '（手入力）' : ''), b.basic],
     ['その他の所得控除', '雑損控除など', b.other]
    ].forEach(function (r) {
      if (!r[2]) return;
      rows.push(boardRowHTML(r[0], r[1], r[2]));
    });
    rows.push(boardRowHTML('所得控除の額の合計', '上の控除の合計', sim.ded));
    rows.push(boardRowHTML('課税される所得金額', '自動計算（1,000円未満切り捨て）', sim.taxable, 'total'));
    if (sim.housing) rows.push(boardRowHTML('住宅借入金等特別控除（税額控除）', '所得税から直接控除', sim.housing));
    rows.push(boardRowHTML('所得税及び復興特別所得税', '税率を掛けて自動計算' + (sim.housing ? '（住宅ローン控除後）' : ''), sim.totalTax));
    if (sim.slipCount) rows.push(boardRowHTML('源泉徴収税額', '源泉徴収票④の合計', sim.withheld));
    rows.push(boardRowHTML(sim.diff >= 0 ? '納める税金' : '還付される税金', '税額−源泉徴収税額', Math.abs(sim.diff), 'total'));

    var rows2 = [];
    if (sim.slipCount) rows2.push(boardRowHTML('所得の内訳「給与」', '支払者・支払金額（源泉徴収票のとおり）', sim.pay));
    rows2.push(boardRowHTML('所得の内訳「事業」', '屋号・事業場所・所得金額', sim.biz));
    return '<div class="board-sec"><h4>確定申告書B 第一表</h4>' + rows.join('') + '</div>' +
      '<div class="board-sec" style="margin-top:8px;"><h4>第二表（所得の内訳）</h4>' + rows2.join('') +
      '<div class="rate-note">第二表は「所得の内訳」のほか、社会保険料・生命保険料などの控除の内訳、住民税の徴収方法を書きます。金額は第一表と同じです。</div></div>';
  }

  function boardDepHTML() {
    var s = depSummary();
    if (!s.count) {
      return '<div class="board-sec"><h4>減価償却費の計算（3ページ目）</h4><div class="empty" style="padding:12px 4px;font-size:11.5px;">資産が未登録です。税務診断タブの「減価償却リスト」で追加するか、下の「＋ 資産を追加」から入れてください。</div></div>';
    }
    var head = '<thead><tr><th>資産の名称等</th><th>取得年月日</th><th>取得価額</th><th>償却方法</th><th>耐用年数</th><th>償却率</th><th>償却期間</th><th>普通償却費</th><th>事業割合</th><th>算入額</th><th>未償却残高</th><th></th></tr></thead>';
    var inTable = s.rows.filter(function (r) { return r.c.method !== '対象外（消耗品費）'; });
    var outside = s.rows.filter(function (r) { return r.c.method === '対象外（消耗品費）'; });
    var body = inTable.map(function (r) {
      var a = r.a, c = r.c;
      return '<tr>' +
        '<td>' + esc(a.name || '（名称なし）') + '</td>' +
        '<td>' + esc(a.acquired || '—') + '</td>' +
        '<td class="num">' + (c.cost ? Math.round(c.cost).toLocaleString('ja-JP') : '—') + '</td>' +
        '<td>' + esc(c.method) + '</td>' +
        '<td class="num">' + esc(a.life || '—') + '</td>' +
        '<td class="num">' + (c.rate ? c.rate.toFixed(3) : '—') + '</td>' +
        '<td class="num">' + esc(c.months) + '</td>' +
        '<td class="num">' + Math.round(c.ordinary).toLocaleString('ja-JP') + '</td>' +
        '<td class="num">' + esc(c.pct) + '%</td>' +
        '<td class="num" style="font-weight:700;color:#0284c7;">' + Math.round(c.expense).toLocaleString('ja-JP') + '</td>' +
        '<td class="num">' + (c.bookEnd === '' ? '—' : Math.round(c.bookEnd).toLocaleString('ja-JP')) + '</td>' +
        '<td><button class="copy-btn" data-copyrow="' + esc(a.id) + '" aria-label="この行をコピー">⧉</button></td>' +
        '</tr>';
    }).join('');
    var sumOrd = inTable.reduce(function (a, r) { return a + r.c.ordinary; }, 0);
    var sumExp = inTable.reduce(function (a, r) { return a + r.c.expense; }, 0);
    body += '<tr style="background:#f8fafc;font-weight:700;"><td colspan="7">合計</td><td class="num">' + sumOrd.toLocaleString('ja-JP') + '</td><td></td><td class="num" style="color:#0284c7;">' + sumExp.toLocaleString('ja-JP') + '</td><td colspan="2"></td></tr>';
    var outNote = outside.length ? '<div class="pc-note" style="background:#fffcf3;border-color:#f2e6c9;color:#7a5316;">' + outside.map(function (r) {
      return '<b>' + esc(r.a.name || '資産') + '</b> は消耗品費で処理（' + yen(Math.floor(r.c.cost * r.c.pct / 100)) + '）。3ページ目には載せず、1ページ目の消耗品費に入れてください。';
    }).join('<br>') + '</div>' : '';
    return '<div class="board-sec"><h4>減価償却費の計算（3ページ目）</h4>' +
      '<div class="dep-wrap"><table class="dep" style="min-width:860px;font-size:10.5px;">' + head + '<tbody>' + body + '</tbody></table></div>' + outNote +
      '<div class="rate-note">「⧉」は e-Tax の並び（資産名→取得年月→取得価額→償却方法→耐用年数→償却率→償却期間→普通償却費→事業割合→算入額→未償却残高）でコピーします。金額を直したいときは税務診断タブの台帳で編集してください。</div></div>';
  }

  function boardRentHTML() {
    var rb = rentBreakdown();
    var rows = [];
    rb.forEach(function (r) {
      rows.push(boardRowHTML(r.vendor, r.n + '件の合計（事業按分後）', r.sum));
    });
    if (!rb.length) rows.push('<div class="empty" style="padding:10px 4px;font-size:11.5px;">地代家賃の経費がありません。</div>');
    return '<div class="board-sec"><h4>地代家賃の内訳（2ページ目）</h4>' + rows.join('') + guideLocalNotes('page2') + '</div>';
  }

  function boardMonthHTML() {
    var m = monthlySales();
    var total = 0;
    for (var i = 1; i <= 12; i += 1) total += m[i];
    var cells = '';
    for (var j = 1; j <= 12; j += 1) {
      cells += '<div class="month-cell' + (m[j] ? '' : ' zero') + '" title="' + j + '月"><div class="m">' + j + '月</div><div class="a">' + (m[j] ? Math.round(m[j]).toLocaleString('ja-JP') : '0') + '</div></div>';
    }
    return '<div class="board-sec"><h4>月別売上（収入）金額（2ページ目）</h4>' +
      '<div class="month-grid">' + cells + '</div>' +
      '<div class="board-row board-total" style="margin-top:8px;"><div class="bl"><b>合計</b><span>1月〜12月の合計（売上金額と一致）</span></div><div class="bv">' + Math.round(total).toLocaleString('ja-JP') + '</div><button class="copy-btn" data-copy="' + Math.round(total) + '">⧉</button></div>' +
      (total ? '' : '<div class="rate-note">売上の内訳が未登録です。税務診断タブの売上金額だけでも合計として使えます。</div>') +
      '</div>';
  }

  function boardFor(key) {
    if (key === 'p1') return '<div class="step-board"><div class="step-board-title">⧉ 1ページ目へ転記する金額<span class="count">押すとコピー</span></div>' + boardP1HTML() + '</div>';
    if (key === 'dep') return '<div class="step-board"><div class="step-board-title">⧉ 2〜3ページ目へ転記する金額<span class="count">押すとコピー</span></div>' + boardDepHTML() + boardMonthHTML() + boardRentHTML() + '</div>';
    if (key === 'form') return '<div class="step-board"><div class="step-board-title">⧉ 確定申告書へ転記する金額<span class="count">押すとコピー</span></div>' + boardFormHTML() + '</div>';
    return '';
  }

  function renderGuide() {
    var stepsBox = $('guideSteps');
    if (!stepsBox) return;

    var dl = deadlineOf(state.year);
    var db = $('deadlineBox');
    if (db) {
      db.innerHTML = '<div class="deadline"><div class="dl-ico">🗓️</div><div class="dl-main">' +
        '<div class="dl-title">' + yearLabel(state.year) + '分の申告期限</div>' +
        '<div class="dl-sub">' + dl.start + ' 〜 ' + dl.end + '<br>還付申告なら' + (state.year + 1) + '年1月から提出できます。期限を過ぎると無申告加算税・延滞税の対象になります。</div>' +
        '</div></div>';
    }

    var checks = checksOf();
    var steps = pcSteps();
    var doneCount = steps.filter(function (s) { return checks[s.id]; }).length;

    var nav = steps.map(function (s) {
      return '<button type="button" class="etax-nav-item' + (checks[s.id] ? ' done' : '') + '" data-goto="gstep-' + s.id + '">' +
        '<span class="nn">' + (checks[s.id] ? '✓' : s.no) + '</span>' +
        '<span class="nt">' + esc(s.title) + '<span class="ns">' + esc(s.sub) + '</span></span></button>';
    }).join('') +
      '<div class="bar" style="margin-top:8px;"><i style="width:' + Math.round(doneCount / steps.length * 100) + '%"></i></div>' +
      '<div class="rate-note" style="margin-top:5px;">進捗 ' + doneCount + ' / ' + steps.length + 'ステップ</div>';
    var navBox = $('guideNav');
    if (navBox) navBox.innerHTML = nav;
    var navCount = $('guideNavCount');
    if (navCount) navCount.textContent = doneCount + '/' + steps.length;

    var h = '';
    steps.forEach(function (s) {
      h += '<details class="step" id="gstep-' + s.id + '"' + (s.id === 'start' ? ' open' : '') + '>' +
        '<summary>' +
        '<span class="step-ico">' + s.ico + '</span>' +
        '<span class="step-main"><span class="step-title">' + s.no + '. ' + esc(s.title) + '</span><span class="step-sub">' + esc(s.sub) + '</span></span>' +
        '<button type="button" class="chk' + (checks[s.id] ? ' on' : '') + '" data-step="' + s.id + '" aria-label="完了にする">✓</button>' +
        '</summary>' +
        '<div class="step-body">' +
        s.items.map(function (it, i) {
          return '<div class="step-item"><span class="no">' + (i + 1) + '</span><span class="tx"><b>' + esc(it.t) + '</b>' +
            (it.path ? '<span class="screen-path" style="display:inline-block;margin:3px 0;">PC画面：' + esc(it.path) + '</span>' : '') +
            '<span>' + esc(it.d) + '</span></span></div>';
        }).join('') +
        guideLocalNotes(s.id) +
        boardFor(s.board) +
        '<div class="step-tip">💡 ' + esc(s.tip) + '</div>' +
        '</div></details>';
    });
    stepsBox.innerHTML = h;

    var bc = $('boardCount');
    if (bc) {
      var p1 = FORM_CATEGORIES.filter(function (c) { return (totalsForYear().byForm[c] || 0) > 0; }).length;
      bc.textContent = (p1 + depSummary().count + 7) + '項目';
    }
  }

  /* 旧名の互換（外部から呼ばれたとき用） */
  function renderBoard() { renderGuide(); }

  /* ============================================================
     PC連携：JSON書き出し / 読み込み
     ============================================================ */
  function stripThumbs(list) {
    return (Array.isArray(list) ? list : []).map(function (x) {
      var c = Object.assign({}, x);
      delete c.thumb;
      return c;
    });
  }
  function jsonPayload() {
    var led = {};
    Object.keys(state.ledgers || {}).forEach(function (k) { led[k] = stripThumbs(state.ledgers[k]); });
    return {
      app: 'receipt-excel',
      schema: 4,
      exportedAt: new Date().toISOString(),
      year: state.year,
      mode: state.mode,
      rate: state.rate,
      ledgers: led,
      sales: state.sales,
      blue: state.blue,
      slips: state.slips,
      ded: state.ded,
      checks: state.checks,
      assets: state.assets,
      revs: state.revs,
      trip: stripThumbs(state.trip),
      sheet2026Applied: state.sheet2026Applied
    };
  }
  function exportJSON() {
    var payload = jsonPayload();
    var name = '税金データ_' + state.year + '_' + stamp().replace(/[^0-9A-Za-z]/g, '') + '.json';
    try {
      var blob = new Blob([JSON.stringify(payload)], { type: 'application/json' });
      downloadBlob(blob, name);
      toast('PC用データを書き出しました。パソコンで公開版を開いて「データ読み込み」で選んでください');
    } catch (e) {
      toast('書き出しに失敗しました', true);
    }
  }
  function applyJSON(data, filename) {
    if (!data || typeof data !== 'object') throw new Error('データの形式が違います');
    if (!data.ledgers && !data.sales && !data.assets && !data.slips && !data.trip) throw new Error('このファイルには税金データが入っていません');
    var counts = { led: 0, assets: 0 };
    var led = data.ledgers && typeof data.ledgers === 'object' ? data.ledgers : {};
    Object.keys(led).forEach(function (k) { if (Array.isArray(led[k])) counts.led += led[k].length; });
    if (data.assets && typeof data.assets === 'object') {
      Object.keys(data.assets).forEach(function (k) { if (Array.isArray(data.assets[k])) counts.assets += data.assets[k].length; });
    }
    var msg = '読み込むデータ：' + Object.keys(led).join('・') + '年分／経費 ' + counts.led + '件・減価償却 ' + counts.assets + '件\n' +
      'この端末の同じ年度のデータは、読み込んだ内容で置きかわります。よろしいですか？';
    if (!window.confirm(msg)) return;
    if (data.ledgers) state.ledgers = data.ledgers;
    if (data.sales) state.sales = data.sales;
    if (data.blue) state.blue = data.blue;
    if (data.slips) state.slips = data.slips;
    if (data.ded) state.ded = data.ded;
    if (data.checks) state.checks = data.checks;
    if (data.assets) state.assets = data.assets;
    if (data.revs) state.revs = data.revs;
    if (Array.isArray(data.trip)) state.trip = data.trip;
    if (typeof data.rate === 'number' && data.rate > 0) state.rate = data.rate;
    if (typeof data.year === 'number' && data.year >= 2000 && data.year <= 2100) state.year = data.year;
    if (typeof data.sheet2026Applied === 'boolean') state.sheet2026Applied = data.sheet2026Applied;
    persist();
    renderAll();
    var yl = $('yearSelect'); if (yl) yl.value = String(state.year);
    var ri = $('rateInput'); if (ri) ri.value = state.rate;
    toast('読み込みました（経費 ' + counts.led + '件・減価償却 ' + counts.assets + '件）');
  }
  function importJSONFile(file) {
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function () {
      try {
        var data = JSON.parse(String(reader.result || '{}'));
        applyJSON(data, file.name);
      } catch (e) {
        toast('読み込めませんでした：' + (e && e.message ? e.message : 'JSONファイルではありません'), true);
      }
    };
    reader.onerror = function () { toast('ファイルを読めませんでした', true); };
    reader.readAsText(file, 'UTF-8');
  }

  /* ============================================================
     データを最新に一新（リセット＆再同期）
     ============================================================ */
  function resetAllData() {
    var msg = 'この端末の登録データをすべて消して、最新の状態に作り直します。\n' +
      '（領収書・源泉徴収票・所得控除・減価償却リスト・チェックの進捗が消えます。元に戻せません）\n\n' +
      '必要なら先に「PC用データ書き出し（.json）」でバックアップしてください。続けますか？';
    if (!window.confirm(msg)) return;
    try { localStorage.removeItem(STORE_KEY); } catch (e) { /* noop */ }
    state.trip = [];
    state.ledgers = {};
    state.sales = {};
    state.blue = {};
    state.slips = {};
    state.ded = {};
    state.checks = {};
    state.assets = {};
    state.revs = {};
    state.sheet2026Applied = false;
    state.presetApplied = false;
    state.drafts = [];
    state.mode = 'tax';
    state.year = new Date().getFullYear();
    applyPresets(false);
    applyLocalPresets();
    persist();
    renderAll();
    var ys = $('yearSelect'); if (ys) ys.value = String(state.year);
    var ri = $('rateInput'); if (ri) ri.value = state.rate;
    toast('データを最新の状態に一新しました');
  }

  /* ============================================================
     各種届出・申請ナビ
     ============================================================ */
  function filingsLocalNote() {
    return '';
  }

  function renderFilings() {
    var box = $('filingsBody');
    if (!box) return;
    box.innerHTML =
      '<details class="step" open><summary><span class="step-ico">🏠</span><span class="step-main"><span class="step-title">① 住所変更（納税地の異動）</span><span class="step-sub">届出書は原則不要。ただしe-Taxの利用者情報は自分で直す</span></span></summary>' +
      '<div class="step-body">' +
      '<div class="step-item"><span class="no">1</span><span class="tx"><b>令和5年1月1日以後は「納税地の異動届出書」の提出は不要</b><span>確定申告書に異動後の納税地（新住所）を書いて提出すれば、税務署はそこから把握します。転居した年の申告書で新住所を入れてください。</span></span></div>' +
      '<div class="step-item"><span class="no">2</span><span class="tx"><b>e-Tax「利用者情報（基本情報）」の住所は自分で変更</b><span>e-Taxソフト（WEB版）にログイン →「利用者情報」→「基本情報」→ 住所・電話番号を訂正。管轄税務署（納税地）が変わる場合は「納税地」も新しい住所にします。ここを直さないと、還付通知・督促・証明書が旧住所に届きます。</span></span></div>' +
      '<div class="step-item"><span class="no">3</span><span class="tx"><b>年の途中で送付先を変えたいとき</b><span>「所得税・消費税の納税地の異動又は変更に関する申出書」をe-Taxまたは書面で提出できます（任意）。</span></span></div>' +
      '<div class="step-item"><span class="no">4</span><span class="tx"><b>振替納税・口座振替を使っている場合</b><span>金融機関に届け出た住所も変更が必要です。e-Taxソフト（WEB版）から「預貯金口座振替依頼書」を出し直すか、金融機関の窓口で住所変更をしてください。</span></span></div>' +
      '<div class="pc-note">🔗 <a href="https://www.nta.go.jp/taxes/shiraberu/taxanswer/shotoku/2091.htm" target="_blank" rel="noopener">国税庁 No.2091 個人事業者の納税地等に異動があった場合の届出関係</a>／<a href="https://www.e-tax.nta.go.jp/" target="_blank" rel="noopener">e-Taxソフト（WEB版）</a></div>' +
      filingsLocalNote() +
      '</div></details>' +

      '<details class="step"><summary><span class="step-ico">🏪</span><span class="step-main"><span class="step-title">② 個人事業の開業・廃業等届出書（開業届）</span><span class="step-sub">e-Taxソフト（WEB版）からオンライン提出</span></span></summary>' +
      '<div class="step-body">' +
      '<div class="step-item"><span class="no">1</span><span class="tx"><b>提出期限</b><span>開業した日の属する年分の確定申告期限まで。ただし青色申告の承認申請は「開業から2ヶ月以内」と別に期限があるので先に出します。</span></span></div>' +
      '<div class="step-item"><span class="no">2</span><span class="tx"><b>e-Taxソフト（WEB版）を開く</b><span><span class="screen-path">申告・申請・納税</span> → <span class="screen-path">新規作成</span> → <span class="screen-path">申請・届出等</span> → <span class="screen-path">個人事業の開業・廃業等届出書</span> の順に選びます。</span></span></div>' +
      '<div class="step-item"><span class="no">3</span><span class="tx"><b>入力する項目</b><span>納税地（住所）／氏名・マイナンバー／職業（例：飲食店の運営、撮影・映像制作、Web制作など）／屋号（任意・銀行口座名になります）／開業日／事業所の所在地と電話番号／給与支払事務所の有無／源泉徴収の有無／青色申告承認申請の有無。</span></span></div>' +
      '<div class="step-item"><span class="no">4</span><span class="tx"><b>提出したら控えを保存</b><span>送信結果の画面と「受付番号」をPDFで保存。e-Taxのメッセージボックスにも届きます。</span></span></div>' +
      '<div class="pc-note">🔗 <a href="https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/shinkoku/annai/04.htm" target="_blank" rel="noopener">国税庁 A1-5 個人事業の開業届出・廃業届出等手続</a>／<a href="https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/shinkoku/annai/42.htm" target="_blank" rel="noopener">開業する場合（提出書類一覧）</a></div>' +
      '</div></details>' +

      '<details class="step"><summary><span class="step-ico">🟦</span><span class="step-main"><span class="step-title">③ 所得税の青色申告承認申請書</span><span class="step-sub">開業から2ヶ月以内／その年3月15日まで</span></span></summary>' +
      '<div class="step-body">' +
      '<div class="step-item"><span class="no">1</span><span class="tx"><b>期限</b><span>その年分から青色にしたいなら3月15日まで。年の中途（1月16日以後）に開業した場合は、開業日から2ヶ月以内です。</span></span></div>' +
      '<div class="step-item"><span class="no">2</span><span class="tx"><b>e-Taxソフト（WEB版）から提出</b><span>ログイン → <span class="screen-path">マイページ</span> → <span class="screen-path">申請・届出等</span> → <span class="screen-path">所得税の青色申告承認申請書</span>。税目は「所得税」、提出年分を選びます。</span></span></div>' +
      '<div class="step-item"><span class="no">3</span><span class="tx"><b>入力する項目</b><span>提出年分／納税地／氏名・マイナンバー／職業・事業内容／開業日／簿記方式（複式簿記）／備忘価額（1円）／前年分の承認の有無／事業専従者の有無。</span></span></div>' +
      '<div class="step-item"><span class="no">4</span><span class="tx"><b>控除額の違い</b><span>65万円＝e-Tax送信＋電子帳簿保存（優良な電子帳簿）／55万円＝e-Tax送信のみ／10万円＝紙提出。65万円を狙うなら、帳簿はこのアプリのExcelと領収書PDFを7年間保存してください。</span></span></div>' +
      '<div class="pc-note">🔗 <a href="https://www.nta.go.jp/taxes/tetsuzuki/shinsei/annai/shinkoku/annai/09.htm" target="_blank" rel="noopener">国税庁 A1-8 所得税の青色申告承認申請手続</a>／<a href="https://www.nta.go.jp/taxes/shiraberu/taxanswer/shotoku/2070.htm" target="_blank" rel="noopener">No.2070 青色申告制度</a></div>' +
      '</div></details>' +

      '<details class="step"><summary><span class="step-ico">🧰</span><span class="step-main"><span class="step-title">④ その他よく使う申請・確認</span><span class="step-sub">振替納税・納税証明書・更正の請求の結果</span></span></summary>' +
      '<div class="step-body">' +
      '<div class="step-item"><span class="no">1</span><span class="tx"><b>振替納税（口座振替）</b><span>e-Taxソフト（WEB版）→「申請・届出等」→「所得税」→「振替納税に関する依頼書」。または金融機関に「預貯金口座振替依頼書」を出す。登録は3月中旬までに済ませると、その年の納付から口座振替になります。</span></span></div>' +
      '<div class="step-item"><span class="no">2</span><span class="tx"><b>納税証明書</b><span>e-Taxソフト（WEB版）からオンラインで交付請求できます（税目・年度・用途を選ぶ）。コンビニ交付や税務署窓口でも取得できます。</span></span></div>' +
      '<div class="step-item"><span class="no">3</span><span class="tx"><b>更正の請求の結果</b><span>e-Taxのメッセージボックスに「更正通知書」が届きます。還付は指定口座へ振込まれます（処理状況は「還付金の処理状況」で確認）。</span></span></div>' +
      '<div class="step-item"><span class="no">4</span><span class="tx"><b>消費税</b><span>開業した年と翌年は原則免税事業者です。インボイス発行事業者の登録をした場合だけ消費税の申告が必要になります。</span></span></div>' +
      '<div class="pc-note">🔗 <a href="https://www.e-tax.nta.go.jp/" target="_blank" rel="noopener">e-Taxソフト（WEB版）</a>／<a href="https://www.nta.go.jp/taxes/tetsuzuki/" target="_blank" rel="noopener">国税庁 申告・申請・納税</a></div>' +
      '</div></details>';
  }

  /* ============================================================
     マイナポータル連携ナビ＋控除証明書XMLの簡易読み取り
     ============================================================ */
  var mynaXmlValues = [];
  function mynaXmlGuess(tag) {
    var t = String(tag).toLowerCase();
    if (/(number|code|id$|^no$|year|date)/.test(t)) return null;
    if (/earthquake|quake|jishin|地震/.test(t)) return { field: 'quake', label: '地震保険料' };
    if (/care|kaigo|介護/.test(t)) return { field: 'lifeCare', label: '介護医療保険料' };
    if (/pension|nenkin|年金/.test(t)) return { field: 'lifePenNew', label: '個人年金保険料' };
    if (/ideco|kigyo|小規模企業共済/.test(t)) return { field: 'ideco', label: '小規模企業共済等掛金' };
    if (/social|kokuho|nenkin|社会保険|国民年金|国民健康保険/.test(t)) return { field: 'socialExtra', label: '社会保険料' };
    if (/life|seimei|生命|insurance|premium|hoken|保険/.test(t)) return { field: 'lifeGenNew', label: '生命保険料（一般）' };
    if (/donation|furusato|kifu|寄附/.test(t)) return { field: 'donPaid', label: '寄附金' };
    if (/medical|iryo|医療/.test(t)) return { field: 'medPaid', label: '医療費' };
    return null;
  }
  function renderMynaXml() {
    var box = $('mynaXmlResult');
    if (!box) return;
    if (!mynaXmlValues.length) {
      box.innerHTML = '';
      return;
    }
    var rows = mynaXmlValues.map(function (r, i) {
      return '<div class="board-row">' +
        '<div class="bl"><b>' + esc(r.label || r.tag) + '</b><span>' + esc(r.tag) + ' ＝ ' + esc(r.value) + '</span></div>' +
        '<div class="bv">' + Number(r.amount).toLocaleString('ja-JP') + '</div>' +
        '<button class="copy-btn" data-xmlcopy="' + i + '" aria-label="コピー">⧉</button>' +
        (r.field ? '<button class="mini-btn" data-xmlapply="' + i + '">控除へ反映</button>' : '') +
        '</div>';
    }).join('');
    box.innerHTML = '<div class="board-sec"><h4>読み取った控除証明書の内容（参考）</h4>' + rows +
      '<div class="rate-note">金額が「支払額」か「控除額」かは証明書の表示で確認してください。「控除へ反映」は所得控除の詳細の該当欄に入れます（上書きされます）。</div></div>';
  }
  function readMynaXml(file) {
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function () {
      try {
        var txt = String(reader.result || '');
        if (!txt.trim()) throw new Error('ファイルが空です');
        var doc = new DOMParser().parseFromString(txt, 'application/xml');
        if (doc.getElementsByTagName('parsererror').length) throw new Error('XMLとして読めませんでした（拡張子と中身を確認してください）');
        var leaves = [];
        (function walk(node) {
          for (var i = 0; i < node.childNodes.length; i += 1) {
            var ch = node.childNodes[i];
            if (ch.nodeType !== 1) continue;
            var kids = ch.childNodes, textOnly = true;
            for (var j = 0; j < kids.length; j += 1) { if (kids[j].nodeType === 1) { textOnly = false; break; } }
            var v = String(ch.textContent || '').trim();
            if (textOnly) { if (v) leaves.push({ tag: ch.nodeName, value: v }); }
            else walk(ch);
          }
        })(doc.documentElement);
        var rows = [];
        leaves.forEach(function (l) {
          var digits = l.value.replace(/[^0-9]/g, '');
          if (!digits || digits.length < 3) return;
          if (/^[0-9,，.]+$/.test(l.value) && !/(number|code|id$|^no$)/i.test(l.tag)) {
            var g = mynaXmlGuess(l.tag);
            rows.push({ tag: l.tag, value: l.value, amount: Number(digits), label: g ? g.label : '', field: g ? g.field : '' });
          }
        });
        if (!rows.length) throw new Error('金額らしい項目が見つかりませんでした');
        mynaXmlValues = rows.slice(0, 40);
        renderMynaXml();
        toast('控除証明書を読み取りました（' + mynaXmlValues.length + '項目）');
      } catch (e) {
        toast('読み取れませんでした：' + (e && e.message ? e.message : 'XMLではありません'), true);
      }
    };
    reader.onerror = function () { toast('ファイルを読めませんでした', true); };
    reader.readAsText(file, 'UTF-8');
  }
  function applyXmlDed(i) {
    var r = mynaXmlValues[i];
    if (!r || !r.field) return;
    var d = dedOf();
    d[r.field] = r.amount;
    persist();
    renderDeductions();
    renderRiskResults();
    toast(r.label + ' ' + yen(r.amount) + ' を控除の詳細に入れました');
  }
  function renderMyna() {
    /* カードの中身は index.html に静的においてある。ここではXML結果の初期化だけ。 */
    renderMynaXml();
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
    rows4.push(['青色申告特別控除前の所得金額', sim.beforeBlue, '']);
    rows4.push(['青色申告特別控除', sim.blueApplied, sim.blueApplied < sim.blue ? '所得を限度に減額' : '']);
    rows4.push(['事業所得', sim.biz, sim.biz < 0 ? '赤字（給与所得と損益通算）' : '']);
    rows4.push(['給与所得（源泉徴収票②の合算）', sim.salary, sim.slipCount + '枚']);
    rows4.push(['合計所得金額', sim.totalIncome, '']);
    rows4.push(['所得控除の額の合計', sim.ded, '所得控除の詳細から自動計算']);
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

    /* 税務診断・源泉（源泉徴収票の内訳＋所得控除の内訳＋税額試算＋診断） */
    var bd = sim.breakdown;
    var rows5 = [];
    rows5.push(['税務診断・源泉 ' + state.year + '年分（令和' + reiwa(state.year) + '年分）', '', '', '', '', '', '']);
    rows5.push([]);
    rows5.push(['■ 源泉徴収票の内訳', '', '', '', '', '', '']);
    rows5.push(['No', '支払者', '支払金額①', '給与所得控除後の金額②', '社会保険料等', '所得控除の額③', '源泉徴収税額④']);
    var slipsAll = activeSlips(slipsOf());
    var tPay = 0, tSal = 0, tSoc = 0, tDed = 0, tWh = 0;
    slipsAll.forEach(function (sl, i) {
      var sal = salaryIncomeOf(sl, state.year);
      tPay += num(sl.payment); tSal += sal; tSoc += num(sl.social); tDed += num(sl.deductions); tWh += num(sl.taxWithheld);
      rows5.push([i + 1, sl.payer || '（名称なし）', num(sl.payment), sal + (sl.afterManual ? 0 : 0), num(sl.social), num(sl.deductions), num(sl.taxWithheld)]);
    });
    if (!slipsAll.length) rows5.push(['—', '源泉徴収票の入力なし', 0, 0, 0, 0, 0]);
    rows5.push(['合計', slipsAll.length + '枚', tPay, tSal, tSoc, tDed, tWh]);
    rows5.push(['', '②は①−給与所得控除で自動計算（手入力があればその額）', '', '', '', '', '']);
    rows5.push([]);
    rows5.push(['■ 所得控除の内訳', '', '']);
    rows5.push(['項目', '金額', '算定・備考']);
    [
      ['社会保険料控除', bd.social, '源泉徴収票の社会保険料 ' + bd.socialSlip + ' ＋ 国保・年金など ' + num(bd.d.socialExtra)],
      ['小規模企業共済等掛金控除', bd.ideco, 'iDeCo・小規模企業共済の掛金'],
      ['生命保険料控除', bd.life, bd.d.lifeMode === 'manual' ? '控除額を直接入力' : '一般（新）' + lifeCalcNew(bd.d.lifeGenNew) + '／一般（旧）' + lifeCalcOld(bd.d.lifeGenOld) + '／介護医療' + lifeCalcNew(bd.d.lifeCare) + '／年金（新）' + lifeCalcNew(bd.d.lifePenNew) + '／年金（旧）' + lifeCalcOld(bd.d.lifePenOld) + '（上限12万円）'],
      ['地震保険料控除', bd.quake, '地震保険料 ' + num(bd.d.quake) + '／旧長期 ' + num(bd.d.quakeOld) + '（上限5万円）'],
      ['医療費控除', bd.medical, '支払 ' + bd.medPaid + ' − 補填 ' + bd.medReimb + ' − 足切り ' + bd.medFloor],
      ['寄附金控除', bd.donation, '寄附 ' + bd.donPaid + ' − 2,000円（限度 ' + bd.donCap + '）'],
      ['配偶者（特別）控除', bd.spouse, '選択した区分'],
      ['扶養控除', bd.dependents, '一般 ' + bd.depGeneral + '／特定 ' + bd.depSpecific + '／老人 ' + bd.depElder],
      ['障害者・ひとり親・寡婦控除', bd.disability + bd.singleParent, '障害者 ' + bd.disability + '／ひとり親・寡婦 ' + bd.singleParent],
      ['基礎控除', bd.basic, bd.basicSet ? '手入力（自動計算 ' + bd.basicAuto + '）' : state.year + '年分の自動計算'],
      ['その他の所得控除', bd.other, '雑損控除など']
    ].forEach(function (r) { if (r[1]) rows5.push([r[0], r[1], r[2]]); });
    rows5.push(['所得控除の額の合計', bd.total, '']);
    rows5.push([]);
    rows5.push(['■ 税額の試算', '', '']);
    rows5.push(['項目', '金額', '備考']);
    rows5.push(['事業売上（収入）', sim.sales, '']);
    rows5.push(['必要経費', sim.expenses, '']);
    rows5.push(['青色申告特別控除', sim.blueApplied, sim.blueApplied < sim.blue ? '所得を限度に減額' : '']);
    rows5.push(['事業所得', sim.biz, sim.biz < 0 ? '赤字（給与所得と損益通算）' : '']);
    rows5.push(['給与収入', sim.pay, slipsAll.length + '枚']);
    rows5.push(['給与所得', sim.salary, '②の合算']);
    rows5.push(['合計所得金額', sim.totalIncome, '給与所得＋事業所得']);
    rows5.push(['所得控除の額の合計', sim.ded, '']);
    rows5.push(['課税される所得金額', sim.taxable, '1,000円未満切り捨て']);
    rows5.push(['所得税（累進税率）', sim.incomeTaxBase, '']);
    rows5.push(['住宅借入金等特別控除', sim.housing, '税額控除']);
    rows5.push(['復興特別所得税', sim.fukko, '所得税の2.1%']);
    rows5.push(['所得税及び復興特別所得税', sim.totalTax, '']);
    rows5.push(['源泉徴収税額', sim.withheld, '前払い済み']);
    rows5.push([sim.diff >= 0 ? '追加納税の見込み' : '還付の見込み', Math.abs(sim.diff), '']);
    rows5.push(['住民税の概算（翌年）', sim.residentTax, '所得割＋均等割の目安']);
    rows5.push([]);
    rows5.push(['■ 税務診断', '', '']);
    rows5.push(['判定', 'リスク項目', '理由・対応']);
    rk.structural.forEach(function (st) { rows5.push([st.level === 'red' ? '赤' : '黄', st.title, st.why]); });
    rk.flags.forEach(function (f) { rows5.push([f.level === 'red' ? '赤' : '黄', f.title + '（' + f.vendor + (f.amount ? ' ' + f.amount + '円' : '') + '）', f.why]); });
    if (!rk.structural.length && !rk.flags.length) rows5.push(['—', '指摘なし', '']);
    rows5.push([]);
    rows5.push(['入力チェック', '', '']);
    if (sim.warnings.length) {
      sim.warnings.forEach(function (w) { rows5.push([w.level === 'warn' ? '要修正' : '参考', w.text, '']); });
    } else {
      rows5.push(['—', '不整合はありません', '']);
    }

    var ws5 = XLSX.utils.aoa_to_sheet(rows5);
    styleWidths(ws5, [22, 20, 20, 22, 16, 18, 16]);
    for (var r5 = 1; r5 <= rows5.length; r5 += 1) {
      for (var c5 = 2; c5 <= 7; c5 += 1) {
        var cell5 = ws5[XLSX.utils.encode_cell({ r: r5 - 1, c: c5 - 1 })];
        if (cell5 && typeof cell5.v === 'number') cell5.z = '#,##0';
      }
    }

    /* 減価償却資産台帳（e-Taxの3ページ目と同じ並び＋参考列） */
    var dep = depSummary();
    var rows6 = [['資産の名称等', '取得年月日', '取得価額', '償却方法', '耐用年数', '償却率', '本年中の償却期間', '本年分の普通償却費', '事業専用割合(%)', '本年分の必要経費算入額', '期末未償却残高', '前年末未償却残高', '摘要・メモ', '経費帳計上額', '差額']];
    dep.rows.forEach(function (r) {
      var a = r.a, c = r.c;
      var ts = r.taxsnap ? r.taxsnap.amount : null;
      rows6.push([
        a.name || '', a.acquired || '', c.cost || 0, c.method, Number(a.life) || 0, c.rate, c.months,
        c.ordinary, c.pct, c.expense, c.bookEnd === '' ? '' : c.bookEnd,
        a.opening === '' || a.opening === undefined ? '' : a.opening,
        (a.memo || '') + (c.notes.length ? (a.memo ? '／' : '') + c.notes.join('／') : ''),
        ts === null ? '' : ts, ts === null ? '' : c.expense - ts
      ]);
    });
    if (!dep.rows.length) rows6.push(['（登録なし）', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    rows6.push([]);
    rows6.push(['資産数', dep.count, '取得価額合計', dep.cost, '', '', '', '普通償却費合計', dep.ordinary, '算入額合計', dep.expense, '', '', '', '']);
    rows6.push(['経費帳の減価償却費', dep.taxsnap, '差額', dep.diff, '', '', '', '※金額は円未満切り捨ての概算。申告前に取得価額と事業割合を書類で確認してください。', '', '', '', '', '', '', '']);
    ['この台帳は e-Tax 青色申告決算書3ページ目「減価償却費の計算」と同じ並びです。',
     '定額法：本年分の普通償却費＝取得価額×償却率×本年中の償却期間÷12（償却可能限度額は備忘価額1円）。',
     '一括償却3年：取得価額÷36×本年中の償却期間。少額特例：30万円未満を取得年に全額（年間合計300万円まで）。',
     '事業専用割合は家事按分後の％です。必要経費算入額＝普通償却費×事業専用割合。'
    ].forEach(function (n) { rows6.push([n]); });
    var ws6 = XLSX.utils.aoa_to_sheet(rows6);
    styleWidths(ws6, [30, 13, 12, 12, 9, 8, 10, 13, 10, 14, 13, 13, 46, 12, 10]);
    for (var r6 = 2; r6 <= dep.rows.length + 1; r6 += 1) {
      [3, 8, 10, 11, 12, 14, 15].forEach(function (c) {
        var cell = ws6[XLSX.utils.encode_cell({ r: r6 - 1, c: c - 1 })];
        if (cell && typeof cell.v === 'number') cell.z = '#,##0';
      });
    }

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, maskSheetName('領収書明細'));
    XLSX.utils.book_append_sheet(wb, ws2, maskSheetName('科目別集計'));
    XLSX.utils.book_append_sheet(wb, ws3, maskSheetName('税務チェック'));
    XLSX.utils.book_append_sheet(wb, ws5, maskSheetName('税務診断・源泉'));
    XLSX.utils.book_append_sheet(wb, ws4, maskSheetName('確定申告サマリー'));
    XLSX.utils.book_append_sheet(wb, ws6, maskSheetName('減価償却資産台帳'));
    return {
      wb: wb,
      filename: '領収書一覧_' + state.year + '_' + stamp() + '.xlsx',
      total: total, count: list.length,
      hasTaxData: activeSlips(slipsOf()).length > 0 || salesOf() > 0 || dep.count > 0
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
    if (!built.count && !built.hasTaxData) {
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
    lines.push('青色申告特別控除額\t' + sim.blueApplied);
    lines.push('所得金額（事業所得）\t' + sim.biz);
    var dep = depSummary();
    if (dep.count) {
      lines.push('【減価償却費の計算（3ページ目）】\t資産名\t取得年月\t取得価額\t償却方法\t耐用年数\t償却率\t償却期間\t普通償却費\t事業割合\t算入額\t未償却残高');
      dep.rows.forEach(function (r) {
        var a = r.a, c = r.c;
        lines.push(['', a.name || '', a.acquired || '', c.cost || '', c.method, a.life || '', c.rate ? c.rate.toFixed(3) : '', c.months, c.ordinary, c.pct, c.expense, c.bookEnd === '' ? '' : c.bookEnd].join('\t'));
      });
    }
    lines.push('給与収入\t' + sim.pay);
    lines.push('給与所得\t' + sim.salary);
    lines.push('合計所得金額\t' + sim.totalIncome);
    var bb = sim.breakdown;
    [['社会保険料控除', bb.social], ['小規模企業共済等掛金控除', bb.ideco], ['生命保険料控除', bb.life],
     ['地震保険料控除', bb.quake], ['医療費控除', bb.medical], ['寄附金控除', bb.donation],
     ['配偶者（特別）控除', bb.spouse], ['扶養控除', bb.dependents],
     ['障害者・ひとり親・寡婦控除', bb.disability + bb.singleParent], ['基礎控除', bb.basic], ['その他', bb.other]
    ].forEach(function (r) { if (r[1]) lines.push(r[0] + '\t' + r[1]); });
    lines.push('所得控除の額の合計\t' + sim.ded);
    lines.push('課税される所得金額\t' + sim.taxable);
    if (sim.housing) lines.push('住宅借入金等特別控除\t' + sim.housing);
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

  function applySlipPresets2025() {
  }

  var SHEET_PRESET_2026 = null;

  function applySheetPresets2026() {
    if (!SHEET_PRESET_2026) return;
    if (state.sheet2026Applied) return;
    var key = String(SHEET_PRESET_2026.year);
    var existing = state.ledgers[key];
    if (Array.isArray(existing) && existing.length) { state.sheet2026Applied = true; persist(); return; }
    var now = new Date().toISOString();
    state.ledgers[key] = SHEET_PRESET_2026.entries.map(function (e, i) {
      return Object.assign({}, e, { id: e.id || ('sheet-' + i), createdAt: now });
    });
    state.sales[key] = SHEET_PRESET_2026.sales;
    state.blue[key] = SHEET_PRESET_2026.blue;
    state.revs[key] = (SHEET_PRESET_2026.revs || []).map(function (r) { return Object.assign({}, r); });
    if (!Array.isArray(state.assets[key]) || !state.assets[key].length) {
      state.assets[key] = SHEET_PRESET_2026.assets.map(function (a) { return Object.assign({}, a); });
    }
    state.sheet2026Applied = true;
    persist();
  }

  /* 個人向けの初期データ（公開版ビルドでは中身が空になる） */
  function applyLocalPresets() {
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
    applyLocalPresets();

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
      else {
        state.ledgers[ykey()] = [];
        if (String(state.year) === '2026') state.sheet2026Applied = true;
      }
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
    /* 所得控除の詳細（支払額→控除額を自動計算） */
    function dedInputHandler(ev) {
      var el = ev.target.closest('[data-ded]');
      if (!el) return;
      var key = el.getAttribute('data-ded');
      var d = dedOf();
      if (el.tagName === 'SELECT') {
        d[key] = Number(el.value) || 0;
      } else {
        d[key] = el.value === '' ? '' : Math.max(0, Math.round(num(el.value)));
      }
      persist();
      refreshDed();
      renderRiskResults();
      renderBoard();
    }
    $('dedDetail').addEventListener('toggle', syncDedOpen, true);
    $('dedDetail').addEventListener('input', dedInputHandler);
    $('dedDetail').addEventListener('change', dedInputHandler);
    $('dedDetail').addEventListener('click', function (ev) {
      var modeBtn = ev.target.closest('[data-ded-mode]');
      if (modeBtn) {
        var which = modeBtn.getAttribute('data-ded-mode');
        var d = dedOf();
        if (which === 'life') d.lifeMode = modeBtn.getAttribute('data-mode');
        if (which === 'quake') d.quakeMode = modeBtn.getAttribute('data-mode');
        persist();
        renderDeductions();
        renderRiskResults();
        renderBoard();
      }
    });

    $('slipList').addEventListener('input', function (ev) {
      var input = ev.target.closest('[data-slip]');
      if (!input) return;
      var wrap = input.closest('.slip');
      var idx = parseInt(wrap.getAttribute('data-idx'), 10);
      var slips = slipsOf();
      if (!slips[idx]) return;
      var key = input.getAttribute('data-slip');
      if (key === 'after') {
        if (input.value === '') { slips[idx].afterManual = false; slips[idx].after = ''; }
        else { slips[idx].afterManual = true; slips[idx].after = Math.max(0, Math.round(num(input.value))); }
      } else if (key === 'payer') {
        slips[idx].payer = input.value;
      } else {
        slips[idx][key] = input.value === '' ? '' : Math.max(0, Math.round(num(input.value)));
      }
      persist();
      refreshSlipNote(idx);
      renderSlipTotals();
      refreshDed();
      renderRiskResults();
      renderBoard();
    });
    $('slipList').addEventListener('click', function (ev) {
      var autoBtn = ev.target.closest('[data-slip-auto]');
      if (autoBtn) {
        var ai = parseInt(autoBtn.getAttribute('data-slip-auto'), 10);
        var sl = slipsOf()[ai];
        if (sl) { sl.afterManual = false; sl.after = ''; }
        persist();
        renderSlips();
        refreshDed();
        renderRiskResults();
        renderBoard();
        return;
      }
      var btn = ev.target.closest('[data-slip-del]');
      if (!btn) return;
      var idx = parseInt(btn.getAttribute('data-slip-del'), 10);
      var slips = slipsOf();
      if (slips.length <= 1) return;
      slips.splice(idx, 1);
      persist();
      renderRiskInputs();
      refreshDed();
      renderRiskResults();
      renderBoard();
    });
    $('btnAddSlip').addEventListener('click', function () {
      slipsOf().push({ id: uid(), payer: '', payment: '', after: '', afterManual: false, social: '', deductions: '', taxWithheld: '' });
      persist();
      renderRiskInputs();
      refreshDed();
      renderRiskResults();
      renderBoard();
      var box = $('slipList');
      var last = box && box.lastElementChild;
      if (last) { var f = last.querySelector('[data-slip="payment"]'); if (f) f.focus(); }
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
      // 進捗と左のナビだけ更新
      var steps = pcSteps();
      var done = steps.filter(function (s) { return c[s.id]; }).length;
      var navCount = $('guideNavCount');
      if (navCount) navCount.textContent = done + '/' + steps.length;
      var navBox = $('guideNav');
      if (navBox) {
        Array.prototype.forEach.call(navBox.querySelectorAll('.etax-nav-item'), function (b, i) {
          var on = !!c[steps[i] && steps[i].id];
          b.classList.toggle('done', on);
          var nn = b.querySelector('.nn');
          if (nn) nn.textContent = on ? '✓' : String(steps[i] ? steps[i].no : i + 1);
        });
        var bar = navBox.querySelector('.bar > i');
        if (bar) bar.style.width = Math.round(done / steps.length * 100) + '%';
      }
    });

    /* 転記ボードのコピー（ステップの中に埋め込まれている） */
    $('guideSteps').addEventListener('click', function (ev) {
      var rowBtn = ev.target.closest('[data-copyrow]');
      if (rowBtn) {
        var a = depFind(rowBtn.getAttribute('data-copyrow'));
        if (a) depCopyRow(a);
        return;
      }
      var btn = ev.target.closest('[data-copy]');
      if (!btn) return;
      var v = btn.getAttribute('data-copy');
      copyText(v, v + ' をコピーしました');
      btn.classList.add('done');
      btn.textContent = '✓';
      setTimeout(function () { btn.classList.remove('done'); btn.textContent = '⧉'; }, 1200);
    });

    /* 左ナビ：押したステップを開いて移動 */
    var navBox = $('guideNav');
    if (navBox) {
      navBox.addEventListener('click', function (ev) {
        var b = ev.target.closest('[data-goto]');
        if (!b) return;
        var el = document.getElementById(b.getAttribute('data-goto'));
        if (!el) return;
        el.open = true;
        try { el.scrollIntoView({ behavior: 'smooth', block: 'start' }); } catch (e) { el.scrollIntoView(); }
        Array.prototype.forEach.call(navBox.querySelectorAll('.etax-nav-item'), function (x) { x.classList.remove('on'); });
        b.classList.add('on');
      });
    }

    /* ---------- 減価償却リスト ---------- */
    var depBody = $('depBody');
    if (depBody) {
      depBody.addEventListener('input', depInputHandler);
      depBody.addEventListener('change', function (ev) {
        var el = ev.target.closest('[data-dep]');
        if (!el) return;
        if (el.tagName === 'SELECT') { renderDep(); renderBoard(); }
      });
      depBody.addEventListener('click', function (ev) {
        var b = ev.target.closest('[data-depact]');
        if (!b) return;
        var act = b.getAttribute('data-depact');
        if (act === 'add') { depNewAsset(); return; }
        if (act === 'copyall') { depCopyAll(); return; }
        if (act === 'restore') { depRestore(); return; }
        if (act === 'toexpense') { saveWorkbook(false); return; }
        var id = b.getAttribute('data-id');
        if (act === 'copy' && id) { var a = depFind(id); if (a) depCopyRow(a); return; }
        if (act === 'del' && id) {
          var t = depFind(id);
          if (!t) return;
          if (!window.confirm('「' + (t.name || 'この資産') + '」を台帳から削除しますか？')) return;
          var arr = assetsOf();
          var idx = arr.findIndex(function (x) { return x.id === id; });
          if (idx >= 0) arr.splice(idx, 1);
          persist();
          renderDep();
          renderBoard();
          toast('削除しました');
        }
      });
    }

    /* ---------- PC連携・マイナポータル・リセット ---------- */
    document.addEventListener('click', function (ev) {
      var j = ev.target.closest('[data-json]');
      if (!j) return;
      if (j.getAttribute('data-json') === 'export') exportJSON();
      else { var fi = $('fileJsonImport'); if (fi) fi.click(); }
    });
    var fji = $('fileJsonImport');
    if (fji) fji.addEventListener('change', function (ev) {
      if (ev.target.files && ev.target.files[0]) importJSONFile(ev.target.files[0]);
      ev.target.value = '';
    });
    var btnXml = $('btnXml');
    if (btnXml) {
      btnXml.addEventListener('click', function () { var f = $('fileXml'); if (f) f.click(); });
      var fx = $('fileXml');
      if (fx) fx.addEventListener('change', function (ev) {
        if (ev.target.files && ev.target.files[0]) readMynaXml(ev.target.files[0]);
        ev.target.value = '';
      });
    }
    var xmlRes = $('mynaXmlResult');
    if (xmlRes) {
      xmlRes.addEventListener('click', function (ev) {
        var c = ev.target.closest('[data-xmlcopy]');
        if (c) {
          var r = mynaXmlValues[Number(c.getAttribute('data-xmlcopy'))];
          if (r) copyText(String(r.amount), r.amount.toLocaleString('ja-JP') + ' をコピーしました');
          return;
        }
        var a = ev.target.closest('[data-xmlapply]');
        if (a) applyXmlDed(Number(a.getAttribute('data-xmlapply')));
      });
    }
    var btnReset = $('btnResetAll');
    if (btnReset) btnReset.addEventListener('click', resetAllData);

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
    setDed: function (patch) {
      var d = dedOf();
      Object.keys(patch || {}).forEach(function (k) { d[k] = patch[k]; });
      persist(); renderAll();
      return d;
    },
    getDed: function () { return dedOf(); },
    addSlip: function (sl) {
      var x = Object.assign({ id: uid(), payer: '', payment: '', after: '', afterManual: false, social: '', deductions: '', taxWithheld: '' }, sl || {});
      slipsOf().push(x);
      persist(); renderAll();
      return x;
    },
    sim: function () { return computeSim(); },
    dep: function () { return depSummary(); },
    assets: function () { return assetsOf(); },
    addAsset: function (a) {
      var x = Object.assign({ id: 'dep-' + uid(), name: '', acquired: '', cost: '', method: '定額法', life: 5, rate: 0.2, months: 12, bizPct: 100, memo: '', opening: '', override: '', source: 'manual' }, a || {});
      assetsOf().push(x);
      persist(); renderAll();
      return x;
    },
    json: function () { return jsonPayload(); },
    importJSON: function (data) { applyJSON(data, ''); },
    renderDep: renderDep,
    addEntry: function (e) {
      var x = Object.assign({ id: uid(), date: todayStr(), vendor: '', category: '雑費', currency: 'JPY', amount: 0, taxType: '', invoiceNo: '', memo: '', source: 'manual', createdAt: new Date().toISOString() }, e);
      taxList().push(x);
      persist(); renderAll();
      return x;
    },
    snapshot: function () {
      return {
        mode: state.mode, year: state.year, rate: state.rate,
        tax: taxList(), trip: state.trip, assets: assetsOf(), revs: revsOf(),
        dep: (function () { var d = depSummary(); return { count: d.count, cost: d.cost, ordinary: d.ordinary, expense: d.expense, taxsnap: d.taxsnap, diff: d.diff, missing: d.missing }; })(),
        sales: salesOf(), blue: blueOf(), slips: slipsOf(), ded: dedOf(),
        sim: (function () { var x = computeSim(); return { salary: x.salary, totalIncome: x.totalIncome, ded: x.ded, taxable: x.taxable, totalTax: x.totalTax, diff: x.diff, residentTax: x.residentTax, warnings: x.warnings }; })(),
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
