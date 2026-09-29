(function () {
  'use strict';

  var APP_VERSION = 'v7';
  var HISTORY_KEY = 'po.history.v1'; // 지난 발주 기록 (주문 키의 해시만 저장)
  var USED_KEY = 'po.used';
  var HISTORY_DAYS = 14;
  var SESSION = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

  var $ = function (id) { return document.getElementById(id); };
  var P = window.POParser;

  var state = {
    files: [],      // {id, name, sig, market, marketLabel, count, error, notice, demo}
    rows: [],       // 변환된 주문 행 (+ uid, fileId, override, orig, categoryManual)
    catalog: [],    // [{name, keywords[]}]
    selected: {},   // 품목명 → 선택 여부
    otherSelected: true,
    termText: '',
    view: 'included', // 표 보기: included(발주할 주문) / check(확인 필요) / excluded(빠진 주문)
    limit: 200,       // 한 번에 그리는 행 수 (대용량 대비)
    history: {},      // 해시 → { d: 'YYYY-MM-DD', s: 세션 }
    lastDownload: null,
    busy: false,
    gen: 0,
    pending: {}     // 처리 대기 중인 파일 시그니처
  };
  var seq = 0;
  var queue = Promise.resolve(); // 모든 업로드를 한 줄로 처리 (비밀번호 창이 겹치지 않게)
  var storageOk = true;

  $('appVersion').textContent = APP_VERSION;

  /* ---------------- 입력한 품목 (브라우저에 저장) ---------------- */

  var TERMS_KEY = 'po.terms.v1';
  var NOMATCH = '입력한 품목 아님';

  function applyTerms(text) {
    state.termText = text || '';
    state.catalog = P.parseTerms(state.termText);
    state.selected = {};
    state.catalog.forEach(function (c) { state.selected[c.name] = true; });
    // 품목을 하나도 적지 않았으면 전체 주문을 넣는다
    state.otherSelected = state.catalog.length === 0;
  }

  function loadTerms() {
    var text = '';
    try { text = localStorage.getItem(TERMS_KEY) || ''; } catch (e) { text = ''; }
    applyTerms(text);
    $('termInput').value = state.termText;
  }

  function saveTerms() {
    try { localStorage.setItem(TERMS_KEY, state.termText); } catch (e) { storageOk = false; }
  }

  /* ---------------- 지난 발주 기록 (여러 날에 걸친 중복 발주 방지) ---------------- */

  function today() {
    var d = new Date(), pad = function (n) { return (n < 10 ? '0' : '') + n; };
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  }

  function loadHistory() {
    var list = [];
    try { list = JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]'); } catch (e) { list = []; }
    if (!Array.isArray(list)) list = [];
    var cutoff = new Date(Date.now() - HISTORY_DAYS * 864e5).toISOString().slice(0, 10);
    state.history = {};
    list.forEach(function (x) {
      if (x && typeof x.h === 'string' && typeof x.d === 'string' && x.d >= cutoff) state.history[x.h] = { d: x.d, s: x.s };
    });
  }

  function saveHistory() {
    var list = Object.keys(state.history).map(function (h) { return { h: h, d: state.history[h].d, s: state.history[h].s }; });
    list.sort(function (a, b) { return a.d < b.d ? 1 : -1; });
    try { localStorage.setItem(HISTORY_KEY, JSON.stringify(list.slice(0, 20000))); } catch (e) { storageOk = false; }
  }

  function recordHistory(rows) {
    var d = today();
    rows.forEach(function (r) { if (r.hash) state.history[r.hash] = { d: d, s: batchId() }; });
    saveHistory();
  }

  // 개인정보(주문번호·이름·전화)를 그대로 남기지 않도록 SHA-256 해시 앞부분만 쓴다
  function hashRows(rows) {
    if (!(window.crypto && crypto.subtle)) return Promise.resolve();
    var enc = new TextEncoder();
    return Promise.all(rows.map(function (r) {
      return crypto.subtle.digest('SHA-256', enc.encode('po-history|' + r.key)).then(function (buf) {
        r.hash = Array.prototype.map.call(new Uint8Array(buf).subarray(0, 12), function (b) { return (b < 16 ? '0' : '') + b.toString(16); }).join('');
      });
    })).catch(function () { /* 기록 기능만 건너뜀 */ });
  }

  function historyHit(r) {
    var h = r.hash && state.history[r.hash];
    return h && h.s !== batchId() ? h : null; // 지금 작업 중에 받은 것은 제외 (고친 뒤 다시 받기 가능)
  }

  // '새로 시작'·'처음부터 다시' 이후는 다른 작업으로 본다
  function batchId() { return SESSION + ':' + state.gen; }

  function mmdd(d) { return d.slice(5).replace('-', '/'); }

  /* ---------------- 파일 읽기 ---------------- */

  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function (e) { resolve(new Uint8Array(e.target.result)); };
      reader.onerror = function () { reject(reader.error); };
      reader.readAsArrayBuffer(file);
    });
  }

  function isZip(d) { return d[0] === 0x50 && d[1] === 0x4b; }
  function isCfb(d) { return d[0] === 0xd0 && d[1] === 0xcf && d[2] === 0x11 && d[3] === 0xe0; }

  function decodeText(data) {
    // 엑셀 '유니코드 텍스트' 저장본(UTF-16) → UTF-8 → CP949 순서로 시도
    if (data[0] === 0xff && data[1] === 0xfe) return new TextDecoder('utf-16le').decode(data.subarray(2));
    if (data[0] === 0xfe && data[1] === 0xff) return new TextDecoder('utf-16be').decode(data.subarray(2));
    try { return new TextDecoder('utf-8', { fatal: true }).decode(data).replace(/^\uFEFF/, ''); }
    catch (e) { return new TextDecoder('euc-kr').decode(data); }
  }

  // 일부 쇼핑몰은 HTML 표를 .xls 이름으로 내려준다
  function looksLikeHtml(text) { return /^\s*<(!doctype|html|table|meta|head|body)/i.test(text); }

  var MESSAGES = {
    NOT_EXCEL: '엑셀 파일(.xlsx, .xls, .csv)이 아닙니다. 쇼핑몰에서 받은 주문 엑셀을 올려 주세요.',
    EMPTY: '빈 파일입니다. 파일을 다시 받아 주세요.',
    UNREADABLE: '파일을 열 수 없습니다. 엑셀에서 연 뒤 ‘다른 이름으로 저장(.xlsx)’ 해서 다시 올려 주세요.',
    PW_SKIPPED: '비밀번호를 입력하지 않아 건너뛰었습니다. 다시 올리면 비밀번호를 물어봅니다.',
    OLD_ENCRYPTION: '구형 방식으로 암호가 걸린 파일이라 열 수 없습니다. 엑셀에서 열어 비밀번호를 해제(파일 → 정보 → 통합 문서 보호 → 암호 설정에서 비우기)한 뒤 다시 올려 주세요.',
    NO_ROWS: '주문이 한 건도 없습니다(제목 줄만 있는 파일).'
  };

  function friendly(code, extra) {
    var err = new Error(MESSAGES[code] || extra || '알 수 없는 오류');
    err.code = code;
    return err;
  }

  var READ_OPTS = { type: 'array', cellDates: true };

  function workbookFromBytes(file, data) {
    var ext = (file.name.split('.').pop() || '').toLowerCase();
    if (!data.length) return Promise.reject(friendly('EMPTY'));
    if (isZip(data)) return Promise.resolve(XLSX.read(data, READ_OPTS));
    if (isCfb(data)) {
      var enc = PODecrypt.inspect(data);
      if (enc.encrypted && !enc.supported) return Promise.reject(friendly('OLD_ENCRYPTION'));
      if (enc.encrypted) return unlock(file, data).then(function (plain) { return XLSX.read(plain, READ_OPTS); });
      return Promise.resolve(XLSX.read(data, READ_OPTS));
    }
    if (['csv', 'txt', 'tsv', 'xls', 'htm', 'html'].indexOf(ext) !== -1) {
      var text = decodeText(data);
      if (ext === 'xls' && !looksLikeHtml(text)) return Promise.reject(friendly('UNREADABLE'));
      // raw: 긴 주문번호가 숫자로 바뀌며 자릿수가 깨지지 않도록 글자 그대로 읽는다
      return Promise.resolve(XLSX.read(text, { type: 'string', raw: true }));
    }
    return Promise.reject(friendly('NOT_EXCEL'));
  }

  function unlock(file, data, errorMsg) {
    return askPassword(file.name, errorMsg).then(function (pw) {
      if (pw == null) throw friendly('PW_SKIPPED');
      setBusy('비밀번호로 ‘' + file.name + '’ 여는 중…');
      return PODecrypt.decrypt(data, pw).catch(function (e) {
        if (e.code === 'BAD_PASSWORD') return unlock(file, data, '비밀번호가 맞지 않습니다. 다시 입력해 주세요.');
        if (e.code === 'UNSUPPORTED') throw friendly('OLD_ENCRYPTION');
        throw e;
      });
    });
  }

  function sheetsOf(wb) {
    // raw: 주문번호 같은 긴 숫자가 '2.03E+13' 처럼 표시 형식으로 바뀌지 않도록 원래 값으로 읽는다
    return wb.SheetNames.map(function (n) {
      return { name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }) };
    });
  }

  function fileSig(f) { return [f.name, f.size].join('|'); }

  function handleFiles(fileList) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return queue;
    var gen = state.gen;
    var dupNames = [];
    files = files.filter(function (f) {
      var sig = fileSig(f);
      var dup = state.pending[sig] || state.files.some(function (x) { return x.sig === sig; });
      if (dup) dupNames.push(f.name);
      else state.pending[sig] = true;
      return !dup;
    });
    if (dupNames.length) toast('이미 올린 파일이라 건너뛰었습니다: ' + dupNames.join(', '));
    if (!files.length) return queue;
    // 실제 파일을 올리면 샘플 데이터는 자동으로 치운다
    if (state.files.some(function (f) { return f.demo; })) {
      state.files.filter(function (f) { return f.demo; }).forEach(function (f) { removeFile(f.id, true); });
      toast('샘플 데이터를 치우고 올린 파일로 바꿨습니다.');
    }
    files.forEach(function (f) {
      queue = queue.then(function () {
        if (gen !== state.gen) return;
        setBusy('‘' + f.name + '’ 읽는 중…');
        return processFile(f, gen);
      }).catch(function (e) {
        toast('처리 중 오류가 발생했습니다: ' + (e && e.message));
      }).then(function () {
        delete state.pending[fileSig(f)];
      });
    });
    queue = queue.then(function () {
      if (gen !== state.gen) return;
      setBusy(null);
      render();
      if (state.rows.length) {
        var inc = included().length;
        $('srStatus').textContent = '주문 ' + state.rows.length + '건을 읽었습니다. 발주 ' + inc + '건, 빠진 주문 ' + (state.rows.length - inc) + '건.';
        scrollToCard('step-preview');
        // 포커스도 결과 쪽으로 옮겨, 다음 Tab 이 화면 밖으로 튀지 않게
        var a = document.activeElement;
        if (!a || a === document.body || a.id === 'pickBtn') $('h-preview').focus({ preventScroll: true });
      }
    });
    return queue;
  }

  function processFile(f, gen) {
    var entry = { id: ++seq, name: f.name, sig: fileSig(f) };
    var result = null;
    return readFile(f)
      .then(function (data) { return workbookFromBytes(f, data); })
      .then(function (wb) {
        var res = P.convertSheets(sheetsOf(wb));
        if (res.error) throw friendly(res.code, res.error);
        if (!res.rows.length) throw friendly('NO_ROWS');
        result = res;
        return hashRows(res.rows);
      })
      .catch(function (e) {
        entry.error = e && e.code ? e.message : MESSAGES.UNREADABLE;
      })
      .then(function () {
        if (gen !== state.gen) return; // 처리 도중 '처음부터 다시'를 눌렀으면 결과를 버린다
        if (result) {
          addResult(entry, result);
          try { localStorage.setItem(USED_KEY, '1'); } catch (e) { /* 무시 */ }
        }
        state.files.push(entry);
        render();
      });
  }

  function addResult(entry, res) {
    entry.market = res.market;
    entry.marketLabel = res.marketLabel;
    entry.count = res.rows.length;
    if (res.missingLabels.length) {
      entry.notice = '이 파일에는 ' + res.missingLabels.join(', ') + ' 열이 없습니다. 미리보기에서 직접 입력하거나 원본 파일을 확인해 주세요.';
    }
    res.rows.forEach(function (r) {
      r.uid = ++seq;
      r.fileId = entry.id;
      r.override = null;
      r.categoryManual = null;
      r.qtyDefaulted = r.issues.indexOf('qty-default') !== -1;
      r.orig = {};
      EDIT_FIELDS.forEach(function (f) { r.orig[f] = r[f]; });
      state.rows.push(r);
    });
  }

  /* ---------------- 행 상태 계산 ---------------- */

  var EDIT_FIELDS = ['name', 'item', 'qty', 'phone', 'address', 'memo'];

  function validate(r) {
    var keepDefault = r.qtyDefaulted && r.qty === r.orig.qty;
    var issues = [];
    if (!String(r.name || '').trim()) issues.push('name');
    if (!String(r.item || '').trim()) issues.push('item');
    var ph = P.normalizePhone(r.phone);
    if (!ph.value) issues.push('phone');
    else if (P.phoneTooShort(ph.value)) issues.push('phone-short');
    else if (!ph.valid) issues.push('phone-format');
    if (!String(r.address || '').trim()) issues.push('address');
    if (!(typeof r.qty === 'number' && r.qty > 0)) issues.push('qty');
    else if (keepDefault) issues.push('qty-default');
    r.issues = issues;
  }

  function catalogHas(name) {
    return state.catalog.some(function (c) { return c.name === name; });
  }

  function classifyRow(r) {
    // 직접 지정한 품목이 목록에서 삭제되었으면 자동 분류로 돌아간다
    if (r.categoryManual != null && (r.categoryManual === '' || catalogHas(r.categoryManual))) {
      return { name: r.categoryManual, matches: [] };
    }
    var edited = r.item !== r.orig.item;
    return edited ? P.classifyText(r.item, state.catalog) : P.classifyProduct(r.product, r.option, state.catalog);
  }

  function compute() {
    // 1) 어느 파일에서든 취소·반품·발송된 주문번호는 모두 제외 (오전 파일 결제완료 + 오후 파일 취소요청)
    var stopped = {};
    state.rows.forEach(function (r) {
      if (r.excluded && r.keyKind !== 'c' && !stopped[r.key]) stopped[r.key] = r.excluded;
    });
    // 2) 주문번호가 없는 파일: 먼저 올린 파일과 내용이 전부 같으면 파일째 중복
    var contentSeen = {}, fileDupCount = {}, fileRowCount = {};
    state.rows.forEach(function (r) {
      fileRowCount[r.fileId] = (fileRowCount[r.fileId] || 0) + 1;
      if (r.keyKind !== 'c') return;
      if (contentSeen[r.key] != null && contentSeen[r.key] !== r.fileId) fileDupCount[r.fileId] = (fileDupCount[r.fileId] || 0) + 1;
      else if (contentSeen[r.key] == null) contentSeen[r.key] = r.fileId;
    });
    var firstFile = {};
    state.rows.forEach(function (r) {
      var reasons = [];
      var warns = [];
      var status = r.excluded || stopped[r.key];
      if (status) {
        var label = P.isShippedStatus(status) ? '이미 발송' : '취소·반품';
        reasons.push(label + ' (' + (r.excluded ? '' : '같은 주문: ') + status + ')');
      }
      var hist = historyHit(r);
      if (hist) {
        // 주문번호가 같으면 이미 보낸 주문 → 기본 제외. 주문번호가 없으면 같은 손님의 재주문일 수 있어 확인만
        if (r.keyKind !== 'c') reasons.push('지난 발주에 있음 (' + mmdd(hist.d) + ')');
        else warns.push('history-suspect');
      }
      if (firstFile[r.key] == null) {
        firstFile[r.key] = r.fileId;
      } else if (r.keyKind !== 'c' && firstFile[r.key] !== r.fileId) {
        reasons.push('중복 주문 (먼저 올린 파일에 있음)');
      } else if (r.keyKind === 'c' && firstFile[r.key] !== r.fileId && fileDupCount[r.fileId] === fileRowCount[r.fileId]) {
        reasons.push('중복 파일 (먼저 올린 파일과 내용이 같음)');
      } else {
        // 같은 파일 안이거나 주문번호가 없으면 자동으로 빼지 않고 확인만 요청
        warns.push('dup-suspect');
      }
      var cls = classifyRow(r);
      r.category = cls.name;
      r.matches = cls.matches;
      if (cls.matches.length > 1) warns.push('ambiguous');
      // 포함 여부는 표에 보이는 대표 품목 하나로만 정한다 (체크를 끈 품목이 몰래 섞이지 않게)
      var catOn = r.category ? !!state.selected[r.category] : state.otherSelected;
      if (!catOn) reasons.push(NOMATCH);
      r.reasons = reasons;
      r.warns = warns;
      r.included = r.override != null ? r.override : reasons.length === 0;
    });
  }

  function blockingIssues(r) {
    return r.issues.filter(function (i) { return P.BLOCKING[i]; });
  }

  function needsCheck(r) {
    return (r.included && (r.issues.length > 0 || r.warns.length > 0 || forcedStopped(r))) || ambiguousOut(r);
  }

  // 여러 품목에 걸린 주문이 대표 품목이 체크 안 돼 빠졌는데, 걸린 다른 품목은 체크된 경우 (세트 상품 등)
  function ambiguousOut(r) {
    return !r.included && r.override == null && r.matches && r.matches.length > 1 &&
      r.reasons.length === 1 && r.reasons[0] === NOMATCH &&
      r.matches.some(function (m) { return m !== r.category && state.selected[m]; });
  }

  // 취소·반품·발송된 주문인데 직접 체크해서 포함시킨 경우
  function forcedStopped(r) {
    return r.override === true && r.reasons.some(function (x) { return /^(취소·반품|이미 발송)/.test(x); });
  }

  function isEdited(r) {
    if (r.categoryManual != null) return true;
    return EDIT_FIELDS.some(function (f) { return fmt(r[f]) !== fmt(r.orig[f]); });
  }

  /* ---------------- 렌더링 ---------------- */

  function el(tag, attrs, text) {
    var e = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      if (k === 'className') e.className = attrs[k];
      else e.setAttribute(k, attrs[k]);
    });
    if (text != null) e.textContent = text;
    return e;
  }

  function fmt(v) { return v == null ? '' : String(v); }

  function render() {
    compute();
    renderUpload();
    renderFiles();
    renderTerms();
    renderTable();
    renderStats();
  }

  // 칩·파일 목록·요약만 갱신 (표는 그대로 두어 편집 중인 칸과 포커스를 지킨다)
  function renderLight() {
    renderFiles();
    renderTerms();
    renderStats();
  }

  function renderUpload() {
    var has = state.files.length > 0;
    $('dropzone').classList.toggle('compact', has);
    $('pickBtn').textContent = has ? '+ 파일 추가' : '파일 선택';
    var used = false;
    try { used = !!localStorage.getItem(USED_KEY); } catch (e) { used = false; }
    $('demoBtn').hidden = has || !used;
    $('demoBtnBig').hidden = has || used;
    $('resetBtn').hidden = !has;
  }

  function renderFiles() {
    var list = $('fileList');
    list.innerHTML = '';
    state.files.forEach(function (f) {
      var li = el('li', { className: f.error ? 'error' : f.notice ? 'notice' : '' });
      if (f.error) {
        li.appendChild(el('span', { 'aria-hidden': 'true' }, '⛔'));
      } else {
        li.appendChild(el('span', { className: 'badge ' + f.market }, f.demo ? '샘플' : f.marketLabel));
      }
      li.appendChild(el('span', { className: 'file-name' }, f.name));
      if (!f.error) {
        var rows = state.rows.filter(function (r) { return r.fileId === f.id; });
        var inc = rows.filter(function (r) { return r.included; }).length;
        var meta = '주문 ' + rows.length + '건 → 발주 ' + inc + '건';
        if (rows.length - inc) meta += ' (빠짐 ' + (rows.length - inc) + '건)';
        li.appendChild(el('span', { className: 'file-meta' }, meta));
      }
      var rm = el('button', { type: 'button', className: 'icon-btn', 'aria-label': f.name + ' 빼기', title: '이 파일 빼기' }, '✕');
      rm.addEventListener('click', function () { removeFile(f.id); });
      li.appendChild(rm);
      if (f.error || f.notice) li.appendChild(el('p', { className: 'file-msg' }, (f.error ? '' : '⚠ ') + (f.error || f.notice)));
      list.appendChild(li);
    });
  }

  function candidateCount(pred) {
    return state.rows.filter(function (r) {
      return r.override !== false && !r.reasons.some(function (x) { return x !== NOMATCH; }) && pred(r);
    }).length;
  }

  // 입력한 품목별로 몇 건이 잡혔는지 바로 보여 준다
  function renderTerms() {
    var box = $('termResults');
    box.innerHTML = '';
    $('termClear').hidden = !state.termText;
    $('storageWarn').hidden = storageOk;
    if (!state.catalog.length) {
      if (state.rows.length) box.appendChild(el('p', { className: 'term-all' }, '품목을 적지 않아 전체 주문 ' + candidateCount(function () { return true; }) + '건이 발주서에 들어갑니다.'));
      return;
    }
    if (!state.rows.length) {
      box.appendChild(el('p', { className: 'muted small' }, '주문 파일을 올리면 품목별 건수가 여기에 나옵니다.'));
      return;
    }
    state.catalog.forEach(function (c) {
      var n = candidateCount(function (r) { return r.category === c.name; });
      var chipEl = el('span', { className: 'term-chip' + (n ? '' : ' zero') }, c.name + ' ');
      chipEl.appendChild(el('b', null, n + '건'));
      if (!n) chipEl.title = '상품명에 이 글자가 들어 있는 주문이 없습니다. 글자를 확인해 주세요.';
      box.appendChild(chipEl);
    });
    var none = candidateCount(function (r) { return !r.category; });
    if (none) {
      var noneEl = el('button', { type: 'button', className: 'term-chip none' }, '해당 없음 ' + none + '건 (빠짐)');
      noneEl.addEventListener('click', function () { setView('excluded', true); });
      box.appendChild(noneEl);
    }
  }

  function setTerms(text) {
    applyTerms(text);
    saveTerms();
    state.limit = PAGE;
    compute();
    renderLight();
    renderTable();
  }

  /* ---------------- 표 ---------------- */

  var PAGE = 200;
  var COLS = ['check', 'name', 'item', 'cat', 'qty', 'phone', 'address', 'memo', 'act'];
  var COL_LABEL = { cat: '품목', name: '수취자명', item: '구입품목', qty: '수량', phone: '전화', address: '주소', memo: '메세지' };
  var COL_HINT = { name: '이름 입력 필요', item: '품목 입력 필요', qty: '수량 확인', phone: '전화번호 입력 필요', address: '주소 입력 필요', memo: '' };
  var ISSUE_FIELD = { name: 'name', item: 'item', phone: 'phone', 'phone-short': 'phone', 'phone-format': 'phone', address: 'address', qty: 'qty', 'qty-default': 'qty' };

  var VIEWS = {
    included: { label: '발주할 주문', test: function (r) { return r.included; } },
    check: { label: '확인 필요', test: needsCheck },
    excluded: { label: '빠진 주문', test: function (r) { return !r.included; } }
  };

  function visibleRows() {
    var test = VIEWS[state.view].test;
    return state.rows.filter(test);
  }

  function rowByUid(uid) {
    for (var i = 0; i < state.rows.length; i++) if (String(state.rows[i].uid) === String(uid)) return state.rows[i];
    return null;
  }

  function renderTable() {
    var tbody = document.querySelector('#previewTable tbody');
    // 편집 중인 칸은 저장하고, 표를 다시 그린 뒤 같은 칸으로 포커스를 돌려준다
    var act = document.activeElement;
    var keep = null;
    if (catMenu.node && catMenu.td && catMenu.node.contains(act)) {
      keep = { uid: catMenu.td.parentNode.dataset.uid, col: 'cat' };
    } else if (act && tbody.contains(act)) {
      var td = act.closest('td[data-col]');
      if (td) {
        keep = { uid: td.parentNode.dataset.uid, col: td.dataset.col };
        if (td.isContentEditable) finishEdit(td.parentNode, rowByUid(keep.uid), td, true);
      }
    }
    closeCatMenu();
    tbody.innerHTML = '';
    var has = state.rows.length > 0;
    $('emptyPreview').hidden = has;
    $('previewBody').hidden = !has;
    var rows = visibleRows();
    var frag = document.createDocumentFragment();
    rows.slice(0, state.limit).forEach(function (r) { frag.appendChild(buildRow(r)); });
    tbody.appendChild(frag);
    if (rows.length > state.limit) {
      var more = el('tr', { className: 'more-row' });
      var mtd = el('td', { colspan: String(COLS.length) });
      var btn = el('button', { type: 'button', className: 'btn secondary' }, '다음 ' + Math.min(PAGE, rows.length - state.limit) + '건 더 보기 (남은 ' + (rows.length - state.limit) + '건)');
      btn.addEventListener('click', function () { state.limit += PAGE; renderTable(); });
      mtd.appendChild(btn);
      more.appendChild(mtd);
      tbody.appendChild(more);
    }
    if (has && !rows.length) {
      var tr = el('tr', { className: 'empty-row' });
      tr.appendChild(el('td', { colspan: String(COLS.length), className: 'empty' },
        state.view === 'check' ? '확인이 필요한 주문이 없습니다. 👍' : state.view === 'excluded' ? '빠진 주문이 없습니다.' : '발주할 주문이 없습니다. ② 에서 품목을 골라 주세요.'));
      tbody.appendChild(tr);
    }
    // 방향키 이동: 표 안에서 Tab 으로 멈추는 칸은 하나만
    var target = keep && tbody.querySelector('tr[data-uid="' + keep.uid + '"] td[data-col="' + keep.col + '"]');
    var first = target || tbody.querySelector('td[data-col="name"]');
    if (first) first.tabIndex = 0;
    if (target) target.focus();
  }

  function buildRow(r) {
    var tr = el('tr', { 'data-uid': r.uid });
    // 포함 체크 + 마켓 표시
    var tdc = el('td', { className: 'c-check', 'data-col': 'check', tabindex: '-1' });
    var cb = el('input', { type: 'checkbox', tabindex: '-1', 'aria-label': (r.name || '이름 없음') + ' 주문 발주 포함' });
    cb.addEventListener('change', function () { setIncluded(tr, r, cb.checked); });
    tdc.appendChild(cb);
    tdc.appendChild(el('span', { className: 'mk ' + r.market, title: r.source, 'aria-label': r.source }, r.market === 'smartstore' ? 'N' : r.market === 'coupang' ? 'C' : '·'));
    tr.appendChild(tdc);

    EDIT_FIELDS.forEach(function (f) {
      var td = el('td', { className: 'c-' + (f === 'address' ? 'addr' : f), 'data-col': f, 'data-field': f, 'data-label': COL_LABEL[f], tabindex: '-1' });
      td.addEventListener('click', function () { if (!td.isContentEditable) startEdit(td); });
      td.addEventListener('blur', function () { if (td.isContentEditable) finishEdit(tr, r, td, true); });
      tr.appendChild(td);
      if (f === 'item') tr.appendChild(buildCatCell(tr, r));
    });

    var tda = el('td', { className: 'c-act', 'data-col': 'act', tabindex: '-1' });
    var undo = el('button', { type: 'button', className: 'undo', tabindex: '-1', title: '이 주문을 원래 값으로 되돌리기' }, '↺ 되돌리기');
    undo.addEventListener('click', function () { revertRow(tr, r); });
    tda.appendChild(undo);
    tr.appendChild(tda);

    updateRow(tr, r);
    return tr;
  }

  // 품목은 평소에 글자 태그로만 보여주고, 누르면 메뉴를 연다 (행마다 드롭다운을 두지 않아 가볍다)
  function buildCatCell(tr, r) {
    var td = el('td', { className: 'c-cat', 'data-col': 'cat', tabindex: '-1' });
    var tag = el('button', { type: 'button', className: 'cat-tag', tabindex: '-1', 'aria-haspopup': 'listbox' });
    tag.addEventListener('click', function (e) {
      e.stopPropagation();
      if (catMenu.closedTd === td && Date.now() - catMenu.closedAt < 400) return; // 열린 메뉴의 태그를 다시 누르면 닫기만
      openCatMenu(td, tr, r);
    });
    td.appendChild(tag);
    td.appendChild(el('span', { className: 'row-reason' }));
    return td;
  }

  var catMenu = { node: null, td: null };

  function closeCatMenu(refocus) {
    if (!catMenu.node) return;
    catMenu.node.remove();
    var td = catMenu.td;
    catMenu.node = null; catMenu.td = null;
    if (refocus && td && document.body.contains(td)) td.focus();
  }

  function openCatMenu(td, tr, r) {
    closeCatMenu();
    var menu = el('div', { className: 'cat-menu', role: 'listbox', 'aria-label': (r.name || '') + ' 주문의 품목', tabindex: '-1' });
    if (!state.catalog.length) {
      toast('② 에 품목을 적으면 여기서 주문마다 품목을 바꿀 수 있습니다.');
      $('termInput').focus();
      return;
    }
    var opts = state.catalog.map(function (c) { return { value: c.name, label: c.name, on: true }; });
    opts.push({ value: '', label: '해당 없음', on: false });
    if (r.categoryManual != null) opts.push({ value: null, label: '↺ 자동 분류로 되돌리기', on: true });
    var cur = 0;
    opts.forEach(function (o, i) {
      var item = el('div', { className: 'cat-opt' + (o.value === (r.category || '') && r.categoryManual !== undefined ? ' current' : ''), role: 'option', id: 'catopt-' + i });
      item.appendChild(el('span', null, o.label));
      if (!o.on) item.appendChild(el('span', { className: 'cat-off' }, '발주에서 빠짐'));
      item.addEventListener('mousedown', function (e) { e.preventDefault(); choose(o); });
      menu.appendChild(item);
      if (o.value === (r.category || '')) cur = i;
    });
    function highlight(i) {
      cur = (i + opts.length) % opts.length;
      Array.prototype.forEach.call(menu.children, function (c, j) { c.classList.toggle('active', j === cur); c.setAttribute('aria-selected', String(j === cur)); });
      menu.setAttribute('aria-activedescendant', 'catopt-' + cur);
      menu.children[cur].scrollIntoView({ block: 'nearest' });
    }
    function choose(o) {
      r.categoryManual = o.value;
      closeCatMenu(true);
      refreshAfterChange(tr, r);
      if (!r.included && r.override == null) toast('입력한 품목이 아니라서 이 주문은 발주에서 빠집니다.');
    }
    menu.addEventListener('keydown', function (e) {
      if (e.key === 'ArrowDown') { e.preventDefault(); highlight(cur + 1); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); highlight(cur - 1); }
      else if (e.key === 'Home') { e.preventDefault(); highlight(0); }
      else if (e.key === 'End') { e.preventDefault(); highlight(opts.length - 1); }
      else if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); choose(opts[cur]); }
      else if (e.key === 'Escape' || e.key === 'Tab') { e.preventDefault(); closeCatMenu(true); }
    });
    menu.addEventListener('focusout', function () { setTimeout(function () { if (catMenu.node === menu && !menu.contains(document.activeElement)) closeCatMenu(); }, 0); });
    // 표 스크롤 영역에 잘리지 않도록 body 에 띄우고, 아래 공간이 부족하면 위로 펼친다
    document.body.appendChild(menu);
    placeMenu(menu, td.querySelector('.cat-tag'));
    catMenu.node = menu; catMenu.td = td;
    highlight(cur);
    menu.focus({ preventScroll: true });
  }

  var MENU_MAX = 460;

  function placeMenu(menu, anchor) {
    var a = anchor.getBoundingClientRect();
    var bar = $('actionbar').hidden ? 0 : $('actionbar').getBoundingClientRect().top;
    var bottomLimit = bar || window.innerHeight;
    menu.style.maxHeight = '';
    var h = Math.min(menu.scrollHeight, MENU_MAX);
    var below = bottomLimit - a.bottom - 8, above = a.top - 8;
    var up = below < h && above > below;
    var avail = Math.max(120, up ? above : below);
    menu.style.maxHeight = Math.min(MENU_MAX, avail) + 'px';
    menu.style.left = Math.max(8, Math.min(a.left, window.innerWidth - menu.offsetWidth - 8)) + 'px';
    menu.style.top = (up ? Math.max(8, a.top - Math.min(h, avail) - 4) : a.bottom + 4) + 'px';
  }

  // 스크롤하면 메뉴가 태그를 따라가고, 태그가 화면 밖으로 나가면 닫는다
  function followMenu() {
    if (!catMenu.node) return;
    var tag = catMenu.td && catMenu.td.querySelector('.cat-tag');
    var r = tag && tag.getBoundingClientRect();
    if (!r || r.bottom < 0 || r.top > window.innerHeight) closeCatMenu();
    else placeMenu(catMenu.node, tag);
  }
  window.addEventListener('scroll', function (e) {
    if (catMenu.node && catMenu.node.contains(e.target)) return; // 메뉴 안 스크롤
    followMenu();
  }, true);
  window.addEventListener('resize', followMenu);

  function startEdit(td) {
    td.contentEditable = 'plaintext-only';
    if (td.contentEditable !== 'plaintext-only') td.contentEditable = 'true';
    td.dataset.before = td.textContent;
    td.classList.add('editing');
    td.focus();
    // 수량은 통째로 바꾸기 쉽게 전체 선택, 글자 칸은 기존 내용 뒤에 커서
    var range = document.createRange();
    range.selectNodeContents(td);
    if (td.dataset.field !== 'qty') range.collapse(false);
    var s = window.getSelection();
    s.removeAllRanges();
    s.addRange(range);
  }

  function finishEdit(tr, r, td, save) {
    td.removeAttribute('contenteditable');
    td.classList.remove('editing');
    if (!save || !r) { td.textContent = td.dataset.before || ''; return; }
    var f = td.dataset.field;
    var text = td.textContent.replace(/\s*\n\s*/g, ' ').trim();
    if (f === 'qty') {
      var n = P.parseQty(text);
      r.qty = n == null ? text : n;
    } else if (f === 'phone') {
      r.phone = P.normalizePhone(text).value;
    } else {
      r[f] = text;
    }
    refreshAfterChange(tr, r);
  }

  function refreshAfterChange(tr, r) {
    validate(r);
    compute();
    if (document.body.contains(tr)) updateRow(tr, r);
    renderLight();
  }

  function setIncluded(tr, r, on) {
    r.override = on === (r.reasons.length === 0) ? null : on;
    refreshAfterChange(tr, r);
  }

  function revertRow(tr, r) {
    EDIT_FIELDS.forEach(function (f) { r[f] = r.orig[f]; });
    r.categoryManual = null;
    refreshAfterChange(tr, r);
    toast('원래 값으로 되돌렸습니다.');
  }

  function updateRow(tr, r) {
    tr.className = r.included ? '' : 'excluded';
    tr.querySelector('.c-check input').checked = r.included;
    var tag = tr.querySelector('.cat-tag');
    var catLabel = r.category || (state.catalog.length ? '해당 없음' : '전체');
    tag.textContent = catLabel + ' ▾';
    tag.classList.toggle('manual', r.categoryManual != null);
    tag.setAttribute('aria-label', '품목: ' + catLabel + (r.categoryManual != null ? ' (직접 지정)' : '') + ', 바꾸려면 누르세요');
    var reason = tr.querySelector('.c-cat .row-reason');
    var notes = [];
    if (r.included) blockingIssues(r).forEach(function (i) { notes.push('⚠ ' + P.ISSUE_LABEL[i]); });
    if (r.reasons.length) notes.push((r.included ? '직접 포함: ' : '빠짐: ') + r.reasons.join(', '));
    r.warns.forEach(function (w) {
      notes.push('⚠ ' + (w === 'ambiguous' ? '여러 품목(' + r.matches.join('·') + ')에 해당' : P.ISSUE_LABEL[w]));
    });
    reason.textContent = notes.join(' · ');
    reason.hidden = !notes.length;
    tr.querySelector('.c-cat').classList.toggle('soft', r.included && r.warns.length > 0);

    var bad = {};
    r.issues.forEach(function (i) { bad[ISSUE_FIELD[i]] = P.BLOCKING[i] ? 'bad' : (bad[ISSUE_FIELD[i]] || 'soft'); });
    EDIT_FIELDS.forEach(function (f) {
      var td = tr.querySelector('td[data-field="' + f + '"]');
      if (!td.isContentEditable) td.textContent = fmt(r[f]);
      td.classList.remove('bad', 'soft', 'edited');
      td.dataset.hint = bad[f] === 'bad' && r.included ? COL_HINT[f] : '';
      var tips = [];
      if (bad[f] && r.included) {
        td.classList.add(bad[f]);
        tips.push(r.issues.filter(function (i) { return ISSUE_FIELD[i] === f; }).map(function (i) { return P.ISSUE_LABEL[i]; }).join(', '));
      }
      if (fmt(r[f]) !== fmt(r.orig[f])) {
        td.classList.add('edited');
        tips.push('고친 칸 · 원래 값: ' + (fmt(r.orig[f]) || '(빈 칸)'));
      }
      if (tips.length) td.title = tips.join(' / '); else td.removeAttribute('title');
    });
    tr.querySelector('.c-act .undo').hidden = !isEdited(r);
  }

  function included() { return state.rows.filter(function (r) { return r.included; }); }

  function renderStats() {
    var inc = included();
    var out = state.rows.filter(function (r) { return !r.included; });
    var check = state.rows.filter(needsCheck);
    var problems = inc.filter(function (r) { return blockingIssues(r).length; }).length;
    var edited = state.rows.filter(isEdited).length;

    // 보기 탭 (발주할 주문 / 확인 필요 / 빠진 주문)
    var counts = { included: inc.length, check: check.length, excluded: out.length };
    var tabs = $('viewTabs');
    tabs.innerHTML = '';
    var keys = Object.keys(VIEWS);
    keys.forEach(function (v, i) {
      var on = state.view === v;
      var zeroCheck = v === 'check' && !counts[v];
      var b = el('button', {
        type: 'button', role: 'tab', id: 'tab-' + v, 'aria-controls': 'tableWrap', 'aria-selected': String(on),
        tabindex: on ? '0' : '-1', className: 'view-tab ' + v + (on ? ' active' : '') + (zeroCheck ? ' clear' : '')
      });
      if (zeroCheck) {
        b.appendChild(document.createTextNode('✓ 확인 필요 없음'));
      } else {
        b.appendChild(el('b', null, String(counts[v])));
        b.appendChild(document.createTextNode(' ' + VIEWS[v].label));
      }
      b.addEventListener('click', function () { setView(v, true); });
      b.addEventListener('keydown', function (e) {
        var d = e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0;
        if (!d) return;
        e.preventDefault();
        var next = keys[(i + d + keys.length) % keys.length];
        setView(next, true);
        $('tab-' + next).focus();
      });
      tabs.appendChild(b);
    });
    $('tableWrap').setAttribute('aria-labelledby', 'tab-' + state.view);
    var byMarket = {};
    inc.forEach(function (r) { byMarket[r.source] = (byMarket[r.source] || 0) + 1; });
    $('marketLine').textContent = Object.keys(byMarket).map(function (m) { return m + ' ' + byMarket[m] + '건'; }).join(' · ') + (edited ? ' · 고친 주문 ' + edited + '건' : '');

    renderExcludeNotice(out);

    var bar = $('actionbar');
    bar.hidden = state.rows.length === 0;
    var txt = $('actionText');
    txt.innerHTML = '';
    var ld = state.lastDownload;
    var done = !!(ld && inc.length && ld.sig === contentSig(inc));
    bar.classList.toggle('done', done);
    $('doneBox').hidden = !done || ld.demo;
    if (done) {
      txt.appendChild(el('b', null, '✅ ' + ld.count + '건 저장됨'));
      txt.appendChild(document.createTextNode(' · ' + ld.name));
    } else {
      txt.appendChild(document.createTextNode('발주 '));
      txt.appendChild(el('b', null, inc.length + '건'));
      var otherOut = out.filter(function (r) { return r.override == null && r.reasons.length === 1 && r.reasons[0] === NOMATCH; }).length;
      if (out.length) txt.appendChild(document.createTextNode(' · 빠짐 ' + out.length + '건' + (otherOut ? ' (입력한 품목 아님 ' + otherOut + '건)' : '')));
      if (problems) {
        txt.appendChild(document.createTextNode(' · '));
        txt.appendChild(el('span', { className: 'warn-text' }, '빨간 칸 ' + problems + '건 확인 필요'));
      }
      if (ld && inc.length) txt.appendChild(el('span', { className: 'changed-note' }, ' · 받은 뒤 내용이 바뀌었습니다 — 다시 받아 주세요'));
    }
    var demoOnly = isDemoOnly();
    $('downloadBtn').textContent = done ? '다시 받기' : demoOnly ? '샘플 발주서 받아보기' : '당일발주 엑셀 다운로드';
    $('downloadBtn').classList.toggle('secondary', done);
    $('newStartBtn').hidden = !done;
    $('demoBand').hidden = !demoOnly;
    $('downloadBtn').disabled = inc.length === 0;
    $('skipLink').hidden = state.rows.length === 0;

    var steps = [$('stepper-1'), $('stepper-2'), $('stepper-3')];
    var done = state.rows.length > 0;
    steps.forEach(function (b, i) {
      b.parentNode.className = done ? (i < 2 ? 'done' : 'active') : (i === 0 ? 'active' : '');
      b.querySelector('.step-mark').textContent = done && i < 2 ? '✓' : String(i + 1);
      b.tabIndex = done ? 0 : -1; // 데이터가 없을 때는 Tab 이 단계 표시에 걸리지 않게
    });
  }

  function setView(v, reveal) {
    state.view = v;
    state.limit = PAGE;
    renderTable(); renderStats();
    if (reveal) revealTable();
  }

  // 탭을 바꾸면 표 윗부분이 화면에 보이도록 (하단 바에 가리지 않게)
  function revealTable() {
    var tabs = $('viewTabs');
    var rect = tabs.getBoundingClientRect();
    var bar = $('actionbar').hidden ? 0 : $('actionbar').offsetHeight;
    if (rect.top < 0 || rect.top > window.innerHeight - bar - 240) {
      window.scrollBy({ top: rect.top - 12, behavior: 'smooth' });
    }
  }

  // 왜 빠졌는지 사유별로 요약 (특히 지난번 품목 선택 때문에 빠진 주문을 놓치지 않게)
  function renderExcludeNotice(out) {
    var box = $('excludeNotice');
    box.innerHTML = '';
    if (!out.length) { box.hidden = true; return; }
    var byReason = {}, byItem = {};
    out.forEach(function (r) {
      var reason = r.override === false ? '직접 뺌' : (r.reasons[0] || '').replace(/ \(.*\)$/, '');
      byReason[reason] = (byReason[reason] || 0) + 1;
      if (r.override !== false && r.reasons.length === 1 && r.reasons[0] === NOMATCH) {
        byItem[NOMATCH] = (byItem[NOMATCH] || 0) + 1;
      }
    });
    var itemKeys = Object.keys(byItem);
    box.className = 'notice-box' + (byReason['지난 발주에 있음'] ? ' warn' : '');
    if (state.view === 'excluded') {
      // 이미 빠진 주문을 보고 있으면 한 줄만
      box.appendChild(el('p', null, itemKeys.length
        ? 'ℹ 입력한 품목이 아닌 주문 ' + byItem[NOMATCH] + '건 — 넣으려면 ② 에 품목을 더 적거나 맨 앞 체크를 켜세요.'
        : 'ℹ 맨 앞 체크를 켜면 그 주문을 발주서에 다시 넣을 수 있습니다.'));
      box.hidden = false;
      return;
    }
    var p = el('p');
    p.appendChild(el('b', null, (byReason['지난 발주에 있음'] ? '⚠ ' : 'ℹ ') + '발주서에서 빠진 주문 ' + out.length + '건'));
    p.appendChild(document.createTextNode(' — ' + Object.keys(byReason).map(function (k) { return k + ' ' + byReason[k] + '건'; }).join(', ')));
    box.appendChild(p);
    if (byReason['지난 발주에 있음']) {
      box.appendChild(el('p', { className: 'small' }, '지난 발주에 이미 넣은 주문 ' + byReason['지난 발주에 있음'] + '건은 두 번 보내지 않도록 뺐습니다. 쇼핑몰에서 발주확인 처리를 했는지 확인해 주세요.'));
    }
    if (itemKeys.length) {
      box.appendChild(el('p', { className: 'small' }, '입력한 품목(' + state.catalog.map(function (c) { return c.name; }).join(', ') + ')이 아닌 주문 ' + byItem[NOMATCH] + '건은 뺐습니다. 더 넣으려면 ② 에 품목을 적으세요.'));
    }
    if (state.view !== 'excluded') {
      var btn = el('button', { type: 'button', className: 'link-btn' }, '빠진 주문 보기');
      btn.addEventListener('click', function () { setView('excluded', true); });
      box.appendChild(btn);
    }
    box.hidden = false;
  }

  /* ---------------- 표 키보드 조작 ---------------- */

  function cellAt(tr, col) { return tr && tr.querySelector('td[data-col="' + col + '"]'); }

  function focusCell(target) {
    if (!target) return;
    var tbody = target.closest('tbody');
    tbody.querySelectorAll('td[tabindex="0"]').forEach(function (c) { c.tabIndex = -1; });
    target.tabIndex = 0;
    target.focus();
  }

  function moveFocus(td, dRow, dCol) {
    var tr = td.parentNode;
    var ci = COLS.indexOf(td.dataset.col);
    if (dRow) {
      var sib = dRow > 0 ? tr.nextElementSibling : tr.previousElementSibling;
      focusCell(cellAt(sib, td.dataset.col));
    } else {
      focusCell(cellAt(tr, COLS[Math.max(0, Math.min(COLS.length - 1, ci + dCol))]));
    }
  }

  var tbodyEl = document.querySelector('#previewTable tbody');

  tbodyEl.addEventListener('keydown', function (e) {
    var td = e.target.closest('td[data-col]');
    if (!td || e.target.closest('.cat-menu')) return;
    var tr = td.parentNode;
    var r = rowByUid(tr.dataset.uid);
    if (td.isContentEditable) {
      if (e.key === 'Enter') { e.preventDefault(); finishEdit(tr, r, td, true); td.focus(); }
      else if (e.key === 'Escape') { e.preventDefault(); finishEdit(tr, r, td, false); td.focus(); }
      else if (e.key === 'Tab') {
        // 스프레드시트처럼 저장 후 옆 칸으로
        e.preventDefault();
        finishEdit(tr, r, td, true);
        moveFocus(td, 0, e.shiftKey ? -1 : 1);
      }
      return;
    }
    if (e.target !== td) return;
    var moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (moves[e.key]) { e.preventDefault(); moveFocus(td, moves[e.key][0], moves[e.key][1]); return; }
    if (e.key === 'Enter' || e.key === 'F2' || e.key === ' ') {
      e.preventDefault();
      var col = td.dataset.col;
      if (col === 'check') { setIncluded(tr, r, !r.included); }
      else if (col === 'cat') { openCatMenu(td, tr, r); }
      else if (col === 'act') { if (!td.querySelector('.undo').hidden) revertRow(tr, r); }
      else if (e.key !== ' ') startEdit(td);
    }
  });

  tbodyEl.addEventListener('focusin', function (e) {
    var td = e.target.closest('td[data-col]');
    if (!td || td.tabIndex === 0) return;
    this.querySelectorAll('td[tabindex="0"]').forEach(function (c) { c.tabIndex = -1; });
    td.tabIndex = 0;
  });

  document.addEventListener('mousedown', function (e) {
    if (catMenu.node && !catMenu.node.contains(e.target)) {
      catMenu.closedTd = catMenu.td; catMenu.closedAt = Date.now();
      closeCatMenu();
    }
  });

  /* ---------------- 동작 ---------------- */

  function removeFile(id, silent) {
    state.files = state.files.filter(function (f) { return f.id !== id; });
    state.rows = state.rows.filter(function (r) { return r.fileId !== id; });
    if (!silent) render();
  }

  function setBusy(text) {
    state.busy = !!text;
    $('busy').hidden = !text;
    if (text) $('busyText').textContent = text;
  }

  var toastTimer;
  function toast(msg, actionLabel, onAction) {
    var t = $('toast');
    $('toastText').textContent = msg;
    var btn = $('toastAction');
    btn.hidden = !actionLabel;
    btn.onclick = null;
    if (actionLabel) {
      btn.textContent = actionLabel;
      btn.onclick = function () { t.hidden = true; onAction(); };
    }
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.hidden = true; }, actionLabel ? 8000 : 4000);
  }

  function scrollToCard(id) {
    var node = $(id);
    if (node && node.scrollIntoView) node.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function askPassword(fileName, errorMsg) {
    var dlg = $('pwDialog');
    $('pwFile').textContent = fileName;
    $('pwInput').value = '';
    $('pwError').hidden = !errorMsg;
    $('pwError').textContent = errorMsg || '';
    setBusy(null);
    return new Promise(function (resolve) {
      function done(v) {
        $('pwForm').removeEventListener('submit', onSubmit);
        $('pwSkip').removeEventListener('click', onSkip);
        dlg.removeEventListener('cancel', onSkip);
        if (dlg.open) dlg.close();
        resolve(v);
      }
      function onSubmit(e) { e.preventDefault(); done($('pwInput').value); }
      function onSkip(e) { if (e) e.preventDefault(); done(null); }
      $('pwForm').addEventListener('submit', onSubmit);
      $('pwSkip').addEventListener('click', onSkip);
      dlg.addEventListener('cancel', onSkip);
      dlg.showModal();
      $('pwInput').focus();
    });
  }

  function confirmBox(title, bodyNodes, okLabel) {
    var dlg = $('confirmDialog');
    $('confirmTitle').textContent = title;
    var body = $('confirmBody');
    body.innerHTML = '';
    bodyNodes.forEach(function (n) { body.appendChild(n); });
    $('confirmOk').textContent = okLabel || '확인';
    return new Promise(function (resolve) {
      dlg.returnValue = '';
      dlg.addEventListener('close', function onClose() {
        dlg.removeEventListener('close', onClose);
        resolve(dlg.returnValue === 'ok');
      });
      dlg.showModal();
      $('confirmCancel').focus();
    });
  }

  // 다운로드 직전 점검: 빈 칸·오류, 중복 의심·품목 겹침, 목록에 없는 상품이 빠지는지
  function preflight(rows) {
    var nodes = [];
    var bad = rows.filter(function (r) { return blockingIssues(r).length; });
    var warn = rows.filter(function (r) { return r.warns.length; });
    var unknownOut = state.rows.filter(function (r) { return !r.included && r.override == null && r.reasons.length === 1 && r.reasons[0] === NOMATCH; });
    if (bad.length) {
      nodes.push(el('p', null, '🔴 빈 칸이나 오류가 있는 주문 ' + bad.length + '건 — 공급처에서 배송하지 못할 수 있습니다.'));
      var ul = el('ul');
      bad.slice(0, 6).forEach(function (r) {
        ul.appendChild(el('li', null, (r.name || '(이름 없음)') + ' — ' + blockingIssues(r).map(function (i) { return P.ISSUE_LABEL[i]; }).join(', ')));
      });
      if (bad.length > 6) ul.appendChild(el('li', null, '외 ' + (bad.length - 6) + '건'));
      nodes.push(ul);
    }
    if (warn.length) {
      var dup = warn.filter(function (r) { return r.warns.indexOf('dup-suspect') !== -1; }).length;
      var amb = warn.filter(function (r) { return r.warns.indexOf('ambiguous') !== -1; }).length;
      var parts = [];
      if (dup) parts.push('같은 내용이 또 있는 주문 ' + dup + '건');
      var histW = warn.filter(function (r) { return r.warns.indexOf('history-suspect') !== -1; }).length;
      if (histW) parts.push('지난 발주와 내용이 같은 주문 ' + histW + '건');
      if (amb) parts.push('여러 품목에 걸린 주문 ' + amb + '건');
      nodes.push(el('p', null, '🟡 ' + parts.join(', ') + ' — 두 번 발주되거나 다른 품목으로 갈 수 있습니다.'));
    }
    var ambOut = state.rows.filter(ambiguousOut);
    if (ambOut.length) {
      nodes.push(el('p', null, '🟡 여러 품목에 걸려 빠진 주문 ' + ambOut.length + '건 (' + ambOut.slice(0, 3).map(function (r) { return r.item; }).join(', ') + (ambOut.length > 3 ? ' 등' : '') + ')'));
    }
    var forced = rows.filter(forcedStopped);
    if (forced.length) {
      nodes.push(el('p', null, '🟠 취소·반품·발송된 주문인데 직접 포함한 주문 ' + forced.length + '건 (' + forced.slice(0, 3).map(function (r) { return r.name; }).join(', ') + (forced.length > 3 ? ' 등' : '') + ')'));
    }
    var count = bad.length + warn.length + forced.length + ambOut.length;
    // '기타'(목록에 없는 상품)가 빠지는 것은 판매자가 고른 설정이라 창을 띄우지 않고, 어차피 창이 뜰 때만 함께 알린다
    if (count && unknownOut.length) {
      nodes.push(el('p', { className: 'muted' }, 'ℹ 참고: 입력한 품목이 아닌 주문 ' + unknownOut.length + '건(' + unknownOut.slice(0, 3).map(function (r) { return r.product || r.item; }).join(', ') + (unknownOut.length > 3 ? ' 등' : '') + ')은 빠집니다.'));
    }
    return { nodes: nodes, count: count };
  }

  // Claude 페이지 안에서 열렸으면 그쪽 저장 기능을, 아니면 일반 브라우저 다운로드를 쓴다
  var hostDownloads = null;
  if (window.claude && typeof window.claude.use === 'function') {
    window.claude.use('downloads').then(function (ns) { hostDownloads = ns; }).catch(function () { hostDownloads = null; });
  }

  /** 파일 저장. 저장했으면 true, 사용자가 취소했으면 false */
  function saveFile(blob, fileName) {
    if (hostDownloads) {
      return hostDownloads.save({ filename: fileName, data: blob }).then(function () { return true; }, function (e) {
        if (e && e.code === 'declined') { toast('저장을 취소했습니다.'); return false; }
        if (e && e.code === 'rate_limited') { toast('저장 창이 이미 열려 있습니다. 잠시 뒤 다시 눌러 주세요.'); return false; }
        throw new Error((e && e.message) || '저장할 수 없습니다');
      });
    }
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    return Promise.resolve(true);
  }

  function isDemoOnly() {
    return state.files.length > 0 && state.files.every(function (f) { return f.demo || f.error; });
  }

  // 받은 뒤 내용이 바뀌었는지 알기 위한 간단한 지문
  function contentSig(rows) {
    var str = rows.map(function (r) { return [r.uid].concat(EDIT_FIELDS.map(function (f) { return fmt(r[f]); })).join('\u0001'); }).join('\u0002');
    var h = 5381;
    for (var i = 0; i < str.length; i++) h = ((h * 33) ^ str.charCodeAt(i)) >>> 0;
    return rows.length + ':' + h;
  }

  function download() {
    var rows = included();
    if (!rows.length) return;
    var pf = preflight(rows);
    var ask = pf.count
      ? confirmBox('다운로드 전에 확인해 주세요', pf.nodes.concat([el('p', null, '‘확인 필요’ 탭에서 고칠 수 있습니다. 그대로 다운로드할까요?')]), '그대로 다운로드')
      : Promise.resolve(true);
    ask.then(function (ok) {
      if (!ok) {
        if (pf.count) {
          setView('check', true);
          var first = document.querySelector('#previewTable tbody td.bad') || document.querySelector('#previewTable tbody td.c-cat');
          if (first) focusCell(first);
        }
        return;
      }
      var btn = $('downloadBtn');
      btn.disabled = true;
      var demo = isDemoOnly();
      var fileName = (demo ? '샘플_' : '') + POExporter.defaultFileName();
      return POExporter.buildWorkbook(ExcelJS, rows).xlsx.writeBuffer().then(function (buf) {
        var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        return saveFile(blob, fileName);
      }).then(function (saved) {
        if (!saved) return;
        if (!demo) recordHistory(rows);
        state.lastDownload = { name: fileName, count: rows.length, sig: contentSig(rows), demo: demo };
        renderStats();
        toast('✅ ' + fileName + ' (' + rows.length + '건) 저장했습니다.');
      }).catch(function (e) {
        toast('엑셀을 저장하지 못했습니다: ' + (e && e.message || e));
      }).then(function () { btn.disabled = false; });
    });
  }

  function reset() {
    confirmBox('처음부터 다시 할까요?', [el('p', null, '올린 파일과 미리보기에서 고친 내용이 모두 지워집니다. (품목 목록과 선택은 그대로 남습니다.)')], '모두 지우기')
      .then(function (ok) {
        if (!ok) return;
        state.gen++;
        state.files = [];
        state.rows = [];
        state.pending = {};
        state.view = 'included';
        state.limit = PAGE;
        state.lastDownload = null;
        setBusy(null);
        render();
        $('pickBtn').focus();
      });
  }

  // 샘플 데이터 (스마트스토어 형식, 가상 인물)
  function loadDemo(quiet) {
    var H = ['상품주문번호', '주문번호', '수취인명', '주문상태', '상품명', '옵션정보', '수량', '수취인연락처1', '통합배송지', '배송메세지'];
    var rows = [H,
      ['D1', 'A1', '김하늘', '결제완료', '[손칼국수] 쫄깃한 생칼국수 1kg', '4인분', '2', '010-1234-5678', '서울특별시 강남구 테헤란로 123 4층', '문 앞에 놓아주세요'],
      ['D2', 'A2', '이도윤', '결제완료', '감자 수제비 반죽 500g', '', '3', '01098765432', '부산광역시 해운대구 센텀중앙로 55 1203호', ''],
      ['D3', 'A3', '박서연', '결제완료', '기장 쪽파 1단 (산지직송)', '1kg', '1', '0504-1111-2222', '경기도 성남시 분당구 판교역로 235', '경비실에 맡겨주세요'],
      ['D4', 'A4', '최민준', '결제완료', '여수 돌산갓 2kg', '', '1', '010-2222-3333', '대구광역시 수성구 달구벌대로 2450', ''],
      ['D5', 'A5', '정예린', '취소요청', '청도 한재 미나리 1kg', '', '2', '010-4444-5555', '광주광역시 서구 상무중앙로 110', ''],
      ['D6', 'A6', '한지우', '결제완료', '국산 들기름 350ml', '', '1', '010-7777-8888', '', '부재 시 연락'],
      ['D7', 'A7', '오세훈', '결제완료', '오곡 곡물면 1kg', '', '2', '010-3030-4040', '울산광역시 남구 삼산로 300', '']
    ];
    var res = P.convertSheet(rows);
    var entry = { id: ++seq, name: '샘플 데이터 (스마트스토어 형식)', sig: 'demo|' + seq, demo: true };
    res.rows.forEach(function (r) { r.key = 'demo|' + r.key; }); // 샘플은 실제 발주 기록과 섞이지 않게
    addResult(entry, res);
    state.files.push(entry);
    render();
    if (quiet === true) return;
    toast('샘플 데이터를 불러왔습니다. 실제 파일을 올리면 샘플은 자동으로 사라집니다.');
    scrollToCard('step-preview');
    $('h-preview').focus({ preventScroll: true });
  }

  /* ---------------- 이벤트 연결 ---------------- */

  var fileInput = $('fileInput');
  var dropzone = $('dropzone');

  $('pickBtn').addEventListener('click', function () { fileInput.click(); });
  fileInput.addEventListener('change', function () {
    var files = Array.prototype.slice.call(fileInput.files);
    fileInput.value = '';
    handleFiles(files);
  });
  ['dragenter', 'dragover'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) { e.preventDefault(); dropzone.classList.add('drag'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) { e.preventDefault(); dropzone.classList.remove('drag'); });
  });
  dropzone.addEventListener('drop', function (e) { handleFiles(e.dataTransfer.files); });
  // 영역 밖에 떨어뜨려도 브라우저가 파일을 열어버리지 않게, 페이지 어디에 놓아도 업로드
  window.addEventListener('dragover', function (e) { e.preventDefault(); });
  window.addEventListener('drop', function (e) {
    e.preventDefault();
    if (!dropzone.contains(e.target) && e.dataTransfer && e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
  });

  document.querySelectorAll('.stepper button').forEach(function (b) {
    b.addEventListener('click', function () { scrollToCard(b.dataset.target); });
  });
  $('downloadBtn').addEventListener('click', download);
  $('resetBtn').addEventListener('click', reset);
  $('demoBtn').addEventListener('click', loadDemo);
  $('demoBtnBig').addEventListener('click', loadDemo);
  $('newStartBtn').addEventListener('click', function () {
    state.gen++;
    state.files = []; state.rows = []; state.pending = {};
    state.view = 'included'; state.limit = PAGE; state.lastDownload = null;
    render();
    window.scrollTo({ top: 0, behavior: 'smooth' });
    $('pickBtn').focus({ preventScroll: true });
  });
  $('clearHistoryBtn').addEventListener('click', function () {
    var n = Object.keys(state.history).length;
    confirmBox('발주 기록을 지울까요?', [el('p', null, '최근 ' + HISTORY_DAYS + '일 동안 받은 발주서의 기록 ' + n + '건을 지웁니다. 지우면 어제 보낸 주문이 다시 들어와도 알려주지 못합니다.')], '기록 지우기')
      .then(function (ok) {
        if (!ok) return;
        state.history = {};
        try { localStorage.removeItem(HISTORY_KEY); } catch (e) { /* 무시 */ }
        render();
        toast('발주 기록을 지웠습니다.');
      });
  });
  var termTimer;
  $('termInput').addEventListener('input', function () {
    clearTimeout(termTimer);
    termTimer = setTimeout(function () { setTerms($('termInput').value); }, 250);
  });
  $('termInput').addEventListener('keydown', function (e) {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    clearTimeout(termTimer);
    setTerms($('termInput').value);
  });
  $('termClear').addEventListener('click', function () {
    $('termInput').value = '';
    setTerms('');
    $('termInput').focus();
  });

  loadTerms();
  loadHistory();
  render();

  // 호스팅 버전에서만: 처음 여는 사람에게는 샘플이 채워진 화면을 보여 준다
  if (window.PO_AUTO_DEMO) {
    var usedBefore = false;
    try { usedBefore = !!localStorage.getItem(USED_KEY); } catch (e) { usedBefore = false; }
    if (!usedBefore) loadDemo(true);
  }

  // 테스트용 훅
  window.__po = { state: state, handleFiles: handleFiles };
})();
