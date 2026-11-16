/* ============================================================
   領収書スキャン → Excel   (receipt-to-excel widget)
   ・写真/カメラ → ブラウザ内OCR(Tesseract.js) → 項目自動抽出
   ・一覧管理(localStorage) → 本物の .xlsx 出力(SheetJS)
   ・モード: 青色申告（経費帳） / 韓国旅行（JPY・KRW混在）
   ============================================================ */
(function () {
  'use strict';

  /* ---------- 定数 ---------- */
  var STORE_KEY = 'pickle.receipt2excel.v2';
  var DEFAULT_RATE = 0.104;

  var TAX_CATEGORIES = ['消耗品費', '旅費交通費', '通信費', '会議費', '接待交際費', '水道光熱費', '新聞図書費', '支払手数料', '外注工賃', '地代家賃', '車両費', '雑費'];
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
    rate: DEFAULT_RATE,
    tax: [],
    trip: [],
    presetApplied: false,
    drafts: []
  };
  var uidSeq = 0;
  var activeDraftId = null;

  function uid() { uidSeq += 1; return 'r' + Date.now().toString(36) + uidSeq.toString(36); }

  /* ---------- DOM ---------- */
  var el = {};
  function $(id) { return document.getElementById(id); }

  /* ---------- ストレージ ---------- */
  function load() {
    try {
      var raw = localStorage.getItem(STORE_KEY);
      if (!raw) return;
      var d = JSON.parse(raw);
      if (d && typeof d === 'object') {
        if (typeof d.rate === 'number' && d.rate > 0) state.rate = d.rate;
        if (Array.isArray(d.tax)) state.tax = d.tax;
        if (Array.isArray(d.trip)) state.trip = d.trip;
        state.presetApplied = !!d.presetApplied;
        if (d.mode === 'tax' || d.mode === 'trip') state.mode = d.mode;
      }
    } catch (e) { /* localStorage 不可でも動作は継続 */ }
  }

  function persist() {
    var payload = { mode: state.mode, rate: state.rate, tax: state.tax, trip: state.trip, presetApplied: state.presetApplied };
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify(payload));
      return true;
    } catch (e) {
      // 容量不足 → サムネイルを捨てて再挑戦
      try {
        var strip = function (arr) {
          return arr.map(function (x) { var c = Object.assign({}, x); delete c.thumb; return c; });
        };
        localStorage.setItem(STORE_KEY, JSON.stringify({ mode: state.mode, rate: state.rate, tax: strip(state.tax), trip: strip(state.trip), presetApplied: state.presetApplied }));
        return true;
      } catch (e2) {
        toast('端末の保存領域がいっぱいのため保存できませんでした。不要な領収書を削除してください。', true);
        return false;
      }
    }
  }

  /* ---------- ユーティリティ ---------- */
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (m) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
    });
  }
  function yen(n) { return '¥' + Math.round(Number(n) || 0).toLocaleString('ja-JP'); }
  function won(n) { return Number(n || 0).toLocaleString('ja-JP') + '₩'; }
  function num(v) {
    if (typeof v === 'number') return v;
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
    toast._t = setTimeout(function () { t.className = ''; }, 3600);
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
  function currentList() { return state.mode === 'trip' ? state.trip : state.tax; }
  function categoriesFor(mode) { return mode === 'trip' ? TRIP_CATEGORIES : TAX_CATEGORIES; }

  /* ============================================================
     画面描画
     ============================================================ */
  function renderAll() {
    renderModeChrome();
    renderKPIs();
    renderCategories();
    renderEntries();
    renderListFooter();
  }

  function renderModeChrome() {
    var trip = state.mode === 'trip';
    $('appTitle').textContent = trip ? '🇰🇷 ソウル旅行 費用' : '🧾 領収書スキャン';
    $('appSubtitle').textContent = trip
      ? '11/14〜15 1泊2日｜現地レシートを読み取って総額を自動集計'
      : '撮影するだけで自動で表に整理・Excel出力';
    $('rateCard').classList.toggle('hidden', !trip);
    $('catCard').classList.toggle('hidden', !trip);
    $('listTitle').textContent = trip ? '旅行の費用明細' : '登録済みの領収書';
    $('exportTitle').textContent = trip ? '📥 ソウル旅行のExcel出力' : '📥 Excel出力';
    $('exportNote').innerHTML = trip
      ? 'シート「ソウル旅行費用明細」＋「カテゴリ別集計」。円建て・ウォン建てを円換算して合算します。'
      : 'シート「領収書明細」＋「科目別集計」。金額はSUM数式入りで、そのまま計算・並べ替えできます。';
    $('stickyLabel').textContent = trip ? 'ソウル旅行 総合計' : '合計金額（税込）';

    Array.prototype.forEach.call(document.querySelectorAll('.mode-tab'), function (b) {
      var on = b.getAttribute('data-mode') === state.mode;
      b.classList.toggle('active', on);
      b.setAttribute('aria-selected', on ? 'true' : 'false');
    });

    var sel = $('langSelect');
    if (trip && sel.value === 'jpn+eng') sel.value = 'kor+eng';
    if (!trip && sel.value === 'kor+eng') sel.value = 'jpn+eng';
  }

  function renderKPIs() {
    var g = $('kpiGrid');
    var list = currentList();

    if (state.mode === 'tax') {
      var total = 0, withInvoice = 0, month = 0;
      var ym = todayStr().slice(0, 7);
      var cats = {};
      list.forEach(function (e) {
        var a = yenOf(e);
        total += a;
        if (e.invoiceNo) withInvoice += 1;
        if ((e.date || '').slice(0, 7) === ym) month += a;
        cats[e.category || '未分類'] = 1;
      });
      g.innerHTML =
        '<div class="kpi wide"><div class="kpi-label">💰 経費 合計（税込）</div>' +
        '<div class="kpi-value">' + yen(total) + '</div>' +
        '<div class="kpi-sub">' + list.length + '件・' + Object.keys(cats).length + '科目｜今月 ' + yen(month) + '</div></div>' +
        '<div class="kpi"><div class="kpi-label">🧾 登録件数</div><div class="kpi-value">' + list.length + '<span style="font-size:12px;font-weight:600;color:#6b7280;"> 件</span></div>' +
        '<div class="kpi-sub">インボイス登録済 ' + withInvoice + '件</div></div>' +
        '<div class="kpi"><div class="kpi-label">📅 今月の経費</div><div class="kpi-value green">' + yen(month) + '</div>' +
        '<div class="kpi-sub">' + ym.replace('-', '年') + '月分</div></div>';
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
    var rows = ['航空券', 'ホテル', '交通費', '食費・買い物', 'その他'].map(function (k) {
      var v = cat[k] || 0;
      var pct = total > 0 ? Math.round(v / total * 100) : 0;
      return '<div class="cat-row"><span class="cat-name">' + k + '</span>' +
        '<span class="cat-bar"><i style="width:' + Math.min(100, pct) + '%"></i></span>' +
        '<span class="cat-amt">' + yen(v) + '<span style="display:block;font-size:9px;color:#9ca3af;font-weight:500;">' + pct + '%</span></span></div>';
    }).join('');
    $('catList').innerHTML = rows;
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
    $('listCount').textContent = list.length + '件';

    if (!list.length) {
      box.innerHTML = '<div class="empty">' + (state.mode === 'trip'
        ? 'レシートを読み取ると、ここに旅行の支出が追加されていきます。'
        : 'まだ登録がありません。<br>上のボタンから領収書を撮影してください。') + '</div>';
      return;
    }

    var sorted = sortedEntries(list).reverse();
    box.innerHTML = sorted.map(function (e) {
      var local = e.currency === 'KRW' ? won(e.amount) : yen(e.amount);
      var sub = e.currency === 'KRW' ? yenOf(e) : '';
      var meta = [];
      if (e.date) meta.push(esc(e.date));
      if (state.mode === 'tax' && e.taxType) meta.push(esc(e.taxType));
      if (e.invoiceNo) meta.push(esc(e.invoiceNo));
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
    html += '<label class="field"><span>メモ</span><textarea class="input" data-f="memo" rows="2" placeholder="用途・同行者など">' + esc(f.memo) + '</textarea></label>';
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
    var trip = state.mode === 'trip';

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

    // 通貨トグル
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
      bar.classList.toggle('hidden-bar', d.status === 'manual' || !!d.editingId);
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
    // 再描画で text が増えた場合に備えて差し替え
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

    try {
      if (window.pickle && pickle.widget && typeof pickle.widget.signal === 'function') {
        pickle.widget.signal((state.mode === 'trip' ? 'ソウル旅行に' : '経費帳に') + '「' + (entry.vendor || entry.category) + '」' + yenOf(entry) + ' を登録しました。');
      }
    } catch (e) { /* noop */ }
  }

  function removeDraft(d) {
    var card = $('draft-' + d.id);
    if (card && card.parentNode) card.parentNode.removeChild(card);
    state.drafts = state.drafts.filter(function (x) { return x.id !== d.id; });
  }

  function discardDraft(d) {
    if (d.editingId && d.backup) {
      // 編集中だった元データを戻す
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

    // グレースケール＋コントラスト補正（レシートの薄い印字を拾いやすくする）
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
      // テキストが後から入った場合は折りたたみを追加
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
      // 固まったワーカーは捨てて、次の再試行で作り直せるようにする
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

  // 電話番号・承認番号・登録番号・日付などを金額と取り違えないための足切り
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
        // ラベルより後ろで最も近い数字を採用（「合計 3点 ¥1,980」などに対応）
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
    // ラベルが読めなかった場合：通貨記号付きの最大値 → 全体の最大値
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
      // 年月が読めないときは旅行期間(2026-11)を優先し、次に今年
      var base = state.mode === 'trip' ? 2026 : new Date().getFullYear();
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
      if (/(株式会社|店|ストア|ストア|マート|カフェ|珈琲|薬局|駅|호텔|마트|카페|식당|편의점)/.test(t)) score += 25;
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
    if (mode === 'trip') {
      for (var i = 0; i < TRIP_RULES.length; i += 1) {
        if (TRIP_RULES[i][1].test(t)) return TRIP_RULES[i][0];
      }
      return '食費・買い物';
    }
    for (var j = 0; j < CAT_RULES.length; j += 1) {
      if (CAT_RULES[j][1].test(t)) return CAT_RULES[j][0];
    }
    return '雑費';
  }

  function parseReceipt(text, mode) {
    var s = normalize(text);
    var lines = s.split(/\r?\n/);
    // ハングルが読めているレシートは韓国語レシートと判断
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

    // ハングルが読めているレシートはウォン建てと判断（원 の印字が無い場合の補完）
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
     Excel / TSV
     ============================================================ */
  function maskSheetName(n) { return n.replace(/[\\\/\?\*\[\]:]/g, ' ').slice(0, 31); }

  function styleWidths(ws, widths) { ws['!cols'] = widths.map(function (w) { return { wch: w }; }); }

  function buildTaxWorkbook() {
    var list = sortedEntries(state.tax);
    var rows = [['No', '日付', '支払先・店名', '勘定科目', '内容・メモ', '金額(税込)', '税区分', 'インボイス番号']];
    var total = 0;
    list.forEach(function (e, i) {
      var a = yenOf(e);
      total += a;
      rows.push([i + 1, e.date || '', e.vendor || '', e.category || '', e.memo || '', a, e.taxType || '', e.invoiceNo || '']);
    });
    var lastDataRow = rows.length; // Excel row index of last data row (1-based)
    rows.push(['', '', '合計', '', '', { t: 'n', f: 'SUM(F2:F' + lastDataRow + ')', v: total }, '', '']);

    var ws = XLSX.utils.aoa_to_sheet(rows);
    styleWidths(ws, [5, 12, 26, 14, 32, 12, 10, 18]);
    for (var r = 2; r <= lastDataRow + 1; r += 1) {
      var cell = ws['F' + r];
      if (cell) cell.z = '#,##0';
    }

    // 科目別集計
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

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, maskSheetName('領収書明細'));
    XLSX.utils.book_append_sheet(wb, ws2, maskSheetName('科目別集計'));
    return { wb: wb, filename: '領収書一覧_' + stamp() + '.xlsx', total: total, count: list.length };
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

    // カテゴリ別集計
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
    try {
      if (window.pickle && pickle.widget && typeof pickle.widget.signal === 'function') {
        pickle.widget.signal((state.mode === 'trip' ? 'ソウル旅行の費用明細' : '領収書一覧') + '（' + built.count + '件・合計' + yen(built.total) + '）をExcel出力しました。');
      }
    } catch (e) { /* noop */ }
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
    var text = lines.join('\n');
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () {
        toast('スプレッドシートに貼り付けられる形でコピーしました');
      }).catch(function () { fallbackCopy(text); });
    } else {
      fallbackCopy(text);
    }
  }

  function fallbackCopy(text) {
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
    toast(ok ? 'スプレッドシートに貼り付けられる形でコピーしました' : 'コピーできませんでした。Excel出力をご利用ください。', !ok);
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
        date: p.date,
        vendor: p.vendor,
        category: p.category,
        currency: p.currency,
        amount: p.amount,
        taxType: p.taxType || '',
        invoiceNo: '',
        memo: p.memo || '',
        thumb: '',
        source: 'preset',
        createdAt: now
      };
    });
    state.presetApplied = true;
    persist();
  }

  function init() {
    load();
    applyPresets(false);

    $('rateInput').value = state.rate;

    // モード切替
    Array.prototype.forEach.call(document.querySelectorAll('.mode-tab'), function (btn) {
      btn.addEventListener('click', function () {
        var m = btn.getAttribute('data-mode');
        if (m === state.mode) return;
        state.mode = m;
        persist();
        renderAll();
      });
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
      else { state.tax = []; }
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

    // 初期表示
    $('langSelect').value = 'jpn+eng';
    renderAll();
  }

  /* 連携用の最小API（アプリ本体や検証から解析・集計結果を参照するため） */
  window.ReceiptToExcel = {
    parseReceipt: parseReceipt,
    parseAmount: parseAmount,
    guessVendor: guessVendor,
    guessCategory: guessCategory,
    buildWorkbook: buildWorkbook,
    snapshot: function () {
      return {
        mode: state.mode,
        rate: state.rate,
        tax: state.tax,
        trip: state.trip,
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
