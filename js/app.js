(function () {
  'use strict';

  var APP_VERSION = 'v3';
  var STORE_KEY = 'po.items.v1';

  var $ = function (id) { return document.getElementById(id); };
  var P = window.POParser;

  var state = {
    files: [],      // {id, name, sig, market, marketLabel, count, error, notice, demo}
    rows: [],       // 변환된 주문 행 (+ uid, fileId, override, orig, categoryManual)
    catalog: [],    // [{name, keywords[]}]
    selected: {},   // 품목명 → 선택 여부
    otherSelected: false,
    showExcluded: false,
    managing: false,
    busy: false,
    gen: 0,
    pending: {}     // 처리 대기 중인 파일 시그니처
  };
  var seq = 0;
  var queue = Promise.resolve(); // 모든 업로드를 한 줄로 처리 (비밀번호 창이 겹치지 않게)
  var storageOk = true;

  $('appVersion').textContent = APP_VERSION;

  /* ---------------- 품목 목록 (브라우저에 저장) ---------------- */

  function cleanCatalog(list) {
    if (!Array.isArray(list)) return null;
    var seen = {};
    var out = [];
    list.forEach(function (c) {
      if (!c || typeof c.name !== 'string' || !c.name.trim()) return;
      var name = c.name.trim();
      if (seen[name]) return;
      seen[name] = true;
      var kws = Array.isArray(c.keywords) ? c.keywords : String(c.keywords || '').split(',');
      out.push({ name: name, keywords: kws.map(function (k) { return String(k).trim(); }).filter(Boolean) });
    });
    return out;
  }

  function loadCatalog() {
    var saved = null;
    try { saved = JSON.parse(localStorage.getItem(STORE_KEY) || 'null'); } catch (e) { saved = null; }
    var cat = saved && cleanCatalog(saved.catalog);
    if (cat) {
      state.catalog = cat;
      state.selected = {};
      var sel = saved.selected && typeof saved.selected === 'object' ? saved.selected : {};
      cat.forEach(function (c) { state.selected[c.name] = sel[c.name] !== false; });
      state.otherSelected = saved.otherSelected === true;
    } else {
      resetCatalog();
    }
  }

  function resetCatalog() {
    state.catalog = P.DEFAULT_CATALOG.map(function (c) { return { name: c.name, keywords: c.keywords.slice() }; });
    state.selected = {};
    state.catalog.forEach(function (c) { state.selected[c.name] = true; });
    state.otherSelected = false;
  }

  function saveCatalog() {
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify({ catalog: state.catalog, selected: state.selected, otherSelected: state.otherSelected }));
    } catch (e) {
      if (storageOk) toast('이 브라우저에서는 품목 설정을 저장할 수 없어, 다음에 열면 기본 품목으로 돌아갑니다.');
      storageOk = false;
    }
  }

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
    try { return new TextDecoder('utf-8', { fatal: true }).decode(data).replace(/^﻿/, ''); }
    catch (e) { return new TextDecoder('euc-kr').decode(data); }
  }

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
    if (ext === 'csv' || ext === 'txt' || ext === 'tsv') {
      // raw: 긴 주문번호가 숫자로 바뀌며 자릿수가 깨지지 않도록 글자 그대로 읽는다
      return Promise.resolve(XLSX.read(decodeText(data), { type: 'string', raw: true }));
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
      if (state.rows.length) scrollToCard('step-preview');
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
      })
      .catch(function (e) {
        entry.error = e && e.code ? e.message : MESSAGES.UNREADABLE;
      })
      .then(function () {
        if (gen !== state.gen) return; // 처리 도중 '처음부터 다시'를 눌렀으면 결과를 버린다
        if (result) addResult(entry, result);
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
    else if (!ph.valid) issues.push('phone-format');
    if (!String(r.address || '').trim()) issues.push('address');
    if (!(typeof r.qty === 'number' && r.qty > 0)) issues.push('qty');
    else if (keepDefault) issues.push('qty-default');
    r.issues = issues;
  }

  function classifyRow(r) {
    if (r.categoryManual != null) return { name: r.categoryManual, matches: [] };
    var edited = r.item !== r.orig.item;
    return edited ? P.classifyText(r.item, state.catalog) : P.classifyProduct(r.product, r.option, state.catalog);
  }

  function compute() {
    // 1) 어느 파일에서든 취소·반품된 주문번호는 모두 제외 (오전 파일 결제완료 + 오후 파일 취소요청)
    var canceled = {};
    state.rows.forEach(function (r) {
      if (r.excluded && r.keyKind !== 'c') canceled[r.key] = r.excluded;
    });
    var firstFile = {};
    state.rows.forEach(function (r) {
      var reasons = [];
      var warns = [];
      if (r.excluded) {
        reasons.push(P.isShippedStatus(r.excluded) ? '이미 발송 (' + r.excluded + ')' : '취소·반품 (' + r.excluded + ')');
      } else if (canceled[r.key]) {
        reasons.push('다른 파일에서 ' + canceled[r.key]);
      }
      if (firstFile[r.key] == null) {
        firstFile[r.key] = r.fileId;
      } else if (r.keyKind !== 'c' && firstFile[r.key] !== r.fileId) {
        // 다른 파일에 같은 주문번호 → 같은 주문을 다시 받은 것
        reasons.push('중복 주문 (먼저 올린 파일에 있음)');
      } else {
        // 같은 파일 안이거나 주문번호가 없으면 자동으로 빼지 않고 확인만 요청
        warns.push('dup-suspect');
      }
      var cls = classifyRow(r);
      r.category = cls.name;
      if (cls.matches.length > 1) warns.push('ambiguous');
      var catOn = r.category ? !!state.selected[r.category] : state.otherSelected;
      if (!catOn) reasons.push('품목 미선택' + (r.category ? ' (' + r.category + ')' : ' (기타)'));
      r.reasons = reasons;
      r.warns = warns;
      r.included = r.override != null ? r.override : reasons.length === 0;
    });
  }

  function blockingIssues(r) {
    return r.issues.filter(function (i) { return P.BLOCKING[i]; });
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

  function render() {
    compute();
    renderUpload();
    renderFiles();
    renderChips();
    renderManage();
    renderTable();
    renderStats();
  }

  function renderUpload() {
    var has = state.files.length > 0;
    $('dropzone').classList.toggle('compact', has);
    $('pickBtn').textContent = has ? '+ 파일 추가' : '파일 선택';
    $('demoBtn').hidden = has;
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
      return r.override !== false && !r.reasons.some(function (x) { return x.indexOf('품목 미선택') !== 0; }) && pred(r);
    }).length;
  }

  function chip(label, checked, count, onToggle, extraClass) {
    var cls = 'chip' + (checked ? ' checked' : '') + (extraClass ? ' ' + extraClass : '') + (state.rows.length && !count ? ' zero' : '');
    var lab = el('label', { className: cls });
    var cb = el('input', { type: 'checkbox' });
    cb.checked = checked;
    cb.addEventListener('change', function () { onToggle(cb.checked); });
    lab.appendChild(cb);
    lab.appendChild(el('span', null, label));
    if (state.rows.length) lab.appendChild(el('span', { className: 'count', 'aria-label': '주문 ' + count + '건' }, String(count)));
    return lab;
  }

  function renderChips() {
    var box = $('itemChips');
    box.innerHTML = '';
    state.catalog.forEach(function (c) {
      box.appendChild(chip(c.name, !!state.selected[c.name],
        candidateCount(function (r) { return r.category === c.name; }),
        function (on) { state.selected[c.name] = on; saveCatalog(); render(); }));
    });
    box.appendChild(chip('기타 (목록에 없는 상품)', state.otherSelected,
      candidateCount(function (r) { return !r.category; }),
      function (on) { state.otherSelected = on; saveCatalog(); render(); }, 'other'));
  }

  function renderManage() {
    $('managePanel').hidden = !state.managing;
    $('manageBtn').setAttribute('aria-expanded', String(state.managing));
    if (!state.managing) return;
    var list = $('manageList');
    list.innerHTML = '';
    state.catalog.forEach(function (c, idx) {
      var li = el('li');
      li.appendChild(el('span', { className: 'manage-name' }, c.name));
      var id = 'kw-' + idx;
      li.appendChild(el('label', { for: id, className: 'sr-only' }, c.name + ' 키워드'));
      var input = el('input', { type: 'text', id: id, className: 'kw-input', placeholder: '키워드 (쉼표로 구분)' });
      input.value = c.keywords.join(', ');
      input.addEventListener('change', function () {
        c.keywords = input.value.split(',').map(function (k) { return k.trim(); }).filter(Boolean);
        saveCatalog(); render();
      });
      li.appendChild(input);
      var del = el('button', { type: 'button', className: 'btn secondary danger small-btn', 'aria-label': c.name + ' 품목 삭제' }, '삭제');
      del.addEventListener('click', function () { deleteItem(idx); });
      li.appendChild(del);
      list.appendChild(li);
    });
  }

  function deleteItem(idx) {
    var c = state.catalog[idx];
    var wasSelected = state.selected[c.name];
    var affected = state.rows.filter(function (r) { return r.category === c.name; }).length;
    state.catalog.splice(idx, 1);
    delete state.selected[c.name];
    saveCatalog(); render();
    var msg = '‘' + c.name + '’ 품목을 삭제했습니다.';
    if (affected) msg += ' 이 품목 주문 ' + affected + '건은 이제 ‘기타’로 분류됩니다.';
    toast(msg, '되돌리기', function () {
      state.catalog.splice(idx, 0, c);
      state.selected[c.name] = wasSelected;
      saveCatalog(); render();
      toast('‘' + c.name + '’ 품목을 되살렸습니다.');
    });
  }

  var COLS = ['check', 'cat', 'name', 'item', 'qty', 'phone', 'address', 'memo', 'act'];
  var COL_LABEL = { cat: '품목', name: '수취자명', item: '구입품목', qty: '수량', phone: '전화번호', address: '주소', memo: '배송메세지' };
  var ISSUE_FIELD = { name: 'name', item: 'item', phone: 'phone', 'phone-format': 'phone', address: 'address', qty: 'qty', 'qty-default': 'qty' };

  function visibleRows() {
    return state.rows.filter(function (r) { return r.included || state.showExcluded || r.override != null; });
  }

  function renderTable() {
    var tbody = document.querySelector('#previewTable tbody');
    tbody.innerHTML = '';
    var has = state.rows.length > 0;
    $('emptyPreview').hidden = has;
    $('previewBody').hidden = !has;
    var rows = visibleRows();
    rows.forEach(function (r) { tbody.appendChild(buildRow(r)); });
    if (has && !rows.length) {
      var tr = el('tr', { className: 'empty-row' });
      tr.appendChild(el('td', { colspan: String(COLS.length), className: 'empty' }, '발주할 주문이 없습니다. ② 에서 품목을 고르거나 ‘빠진 주문도 보기’를 켜 보세요.'));
      tbody.appendChild(tr);
    }
    // 방향키 이동: 표 안에서 Tab 으로 멈추는 칸은 하나만
    var cells = tbody.querySelectorAll('td[data-col]');
    var active = cells.length ? (tbody.querySelector('td[data-col="name"]') || cells[0]) : null;
    if (active) active.tabIndex = 0;
  }

  function buildRow(r) {
    var tr = el('tr', { 'data-uid': r.uid });
    // 포함 체크 + 마켓 표시
    var tdc = el('td', { className: 'c-check', 'data-col': 'check', tabindex: '-1' });
    var cb = el('input', { type: 'checkbox', tabindex: '-1', 'aria-label': (r.name || '이름 없음') + ' 주문 발주 포함' });
    cb.addEventListener('change', function () { setIncluded(tr, r, cb.checked); });
    tdc.appendChild(cb);
    var dot = el('span', { className: 'mk ' + r.market, title: r.source, 'aria-label': r.source }, r.market === 'smartstore' ? 'N' : r.market === 'coupang' ? 'C' : '·');
    tdc.appendChild(dot);
    tr.appendChild(tdc);

    // 품목 분류 (직접 바꿀 수 있음)
    var tdcat = el('td', { className: 'c-cat', 'data-col': 'cat', 'data-label': '품목', tabindex: '-1' });
    var sel = el('select', { tabindex: '-1', 'aria-label': (r.name || '') + ' 주문 품목' });
    state.catalog.forEach(function (c) { sel.appendChild(el('option', { value: c.name }, c.name)); });
    sel.appendChild(el('option', { value: '' }, '기타'));
    sel.addEventListener('change', function () {
      r.categoryManual = sel.value;
      refreshAfterChange(tr, r);
      tdcat.focus();
    });
    sel.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' || e.key === 'Tab') { e.preventDefault(); tdcat.focus(); }
    });
    tdcat.appendChild(sel);
    tdcat.appendChild(el('span', { className: 'row-reason' }));
    tr.appendChild(tdcat);

    EDIT_FIELDS.forEach(function (f) {
      var td = el('td', { className: 'c-' + (f === 'address' ? 'addr' : f), 'data-col': f, 'data-field': f, 'data-label': COL_LABEL[f], tabindex: '-1' });
      td.addEventListener('click', function () { if (!td.isContentEditable) startEdit(td); });
      td.addEventListener('blur', function () { if (td.isContentEditable) finishEdit(tr, r, td, true); });
      tr.appendChild(td);
    });

    var tda = el('td', { className: 'c-act', 'data-col': 'act', tabindex: '-1' });
    var undo = el('button', { type: 'button', className: 'icon-btn undo', tabindex: '-1', title: '이 주문을 원래 값으로 되돌리기', 'aria-label': '원래 값으로 되돌리기' }, '↺');
    undo.addEventListener('click', function () { revertRow(tr, r); });
    tda.appendChild(undo);
    tr.appendChild(tda);

    updateRow(tr, r);
    return tr;
  }

  function startEdit(td) {
    td.contentEditable = 'plaintext-only';
    if (td.contentEditable !== 'plaintext-only') td.contentEditable = 'true';
    td.dataset.before = td.textContent;
    td.classList.add('editing');
    td.focus();
    var range = document.createRange();
    range.selectNodeContents(td);
    var s = window.getSelection();
    s.removeAllRanges();
    s.addRange(range);
  }

  function finishEdit(tr, r, td, save) {
    td.contentEditable = 'false';
    td.removeAttribute('contenteditable');
    td.classList.remove('editing');
    if (!save) { td.textContent = td.dataset.before || ''; return; }
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
    updateRow(tr, r);
    renderStats(); renderChips(); renderFiles();
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

  function fmt(v) { return v == null ? '' : String(v); }

  function updateRow(tr, r) {
    var cls = [];
    if (!r.included) cls.push('excluded');
    tr.className = cls.join(' ');
    tr.querySelector('.c-check input').checked = r.included;
    var sel = tr.querySelector('.c-cat select');
    sel.value = r.category || '';
    sel.classList.toggle('manual', r.categoryManual != null);
    var reason = tr.querySelector('.c-cat .row-reason');
    var notes = [];
    if (r.reasons.length) notes.push((r.included ? '직접 포함: ' : '빠짐: ') + r.reasons.join(', '));
    r.warns.forEach(function (w) { notes.push('⚠ ' + P.ISSUE_LABEL[w]); });
    reason.textContent = notes.join(' · ');
    reason.hidden = !notes.length;
    tr.querySelector('.c-cat').classList.toggle('soft', r.included && r.warns.indexOf('ambiguous') !== -1);

    var bad = {};
    r.issues.forEach(function (i) { bad[ISSUE_FIELD[i]] = P.BLOCKING[i] ? 'bad' : (bad[ISSUE_FIELD[i]] || 'soft'); });
    var anyEdited = r.categoryManual != null;
    EDIT_FIELDS.forEach(function (f) {
      var td = tr.querySelector('td[data-field="' + f + '"]');
      if (!td.isContentEditable) td.textContent = fmt(r[f]);
      td.classList.remove('bad', 'soft', 'edited');
      var tips = [];
      if (bad[f] && r.included) {
        td.classList.add(bad[f]);
        tips.push(r.issues.filter(function (i) { return ISSUE_FIELD[i] === f; }).map(function (i) { return P.ISSUE_LABEL[i]; }).join(', '));
      }
      if (fmt(r[f]) !== fmt(r.orig[f])) {
        td.classList.add('edited');
        tips.push('고친 칸 · 원래 값: ' + (fmt(r.orig[f]) || '(빈 칸)'));
        anyEdited = true;
      }
      if (tips.length) td.title = tips.join(' / '); else td.removeAttribute('title');
    });
    tr.querySelector('.c-act .undo').hidden = !anyEdited;
  }

  function included() { return state.rows.filter(function (r) { return r.included; }); }

  function renderStats() {
    var inc = included();
    var byMarket = {};
    inc.forEach(function (r) { byMarket[r.source] = (byMarket[r.source] || 0) + 1; });
    var problems = inc.filter(function (r) { return blockingIssues(r).length; }).length;
    var soft = inc.filter(function (r) { return !blockingIssues(r).length && (r.issues.length || r.warns.length); }).length;
    var out = state.rows.filter(function (r) { return !r.included; });
    var stats = $('stats');
    stats.innerHTML = '';
    function stat(cls, num, label, onClick) {
      var d = el(onClick ? 'button' : 'div', { className: 'stat ' + cls });
      if (onClick) { d.type = 'button'; d.addEventListener('click', onClick); }
      d.appendChild(el('b', null, String(num)));
      d.appendChild(document.createTextNode(label));
      stats.appendChild(d);
    }
    stat('main', inc.length, '발주할 주문');
    Object.keys(byMarket).forEach(function (m) { stat('', byMarket[m], m); });
    if (out.length) stat('clickable', out.length, state.showExcluded ? '빠진 주문 숨기기' : '빠진 주문 보기', toggleExcluded);
    if (problems) stat('danger', problems, '빈 칸·오류');
    if (soft) stat('warn', soft, '확인 권장');

    renderExcludeNotice(out);

    var bar = $('actionbar');
    bar.hidden = state.rows.length === 0;
    var txt = $('actionText');
    txt.innerHTML = '';
    txt.appendChild(document.createTextNode('발주 '));
    txt.appendChild(el('b', null, inc.length + '건'));
    if (out.length) txt.appendChild(document.createTextNode(' · 빠짐 ' + out.length + '건'));
    if (problems) {
      txt.appendChild(document.createTextNode(' · '));
      txt.appendChild(el('span', { className: 'warn-text' }, '빨간 칸 ' + problems + '건 확인 필요'));
    }
    $('downloadBtn').disabled = inc.length === 0;
    $('skipLink').hidden = state.rows.length === 0;

    var steps = [$('stepper-1'), $('stepper-2'), $('stepper-3')];
    var done = state.rows.length > 0;
    steps.forEach(function (b, i) {
      b.parentNode.className = done ? (i < 2 ? 'done' : 'active') : (i === 0 ? 'active' : '');
      b.querySelector('.step-mark').textContent = done && i < 2 ? '✓' : String(i + 1);
    });
  }

  // 왜 빠졌는지 사유별로 요약 (특히 지난번 품목 선택 때문에 빠진 주문을 놓치지 않게)
  function renderExcludeNotice(out) {
    var box = $('excludeNotice');
    box.innerHTML = '';
    if (!out.length) { box.hidden = true; return; }
    var byReason = {}, byItem = {};
    out.forEach(function (r) {
      var reason = r.override === false ? '직접 뺌' : (r.reasons[0] || '').replace(/ \(.*\)$/, '');
      if (/^다른 파일에서/.test(reason)) reason = '취소·반품';
      byReason[reason] = (byReason[reason] || 0) + 1;
      if (r.override !== false && /^품목 미선택/.test(r.reasons[0] || '')) {
        var k = r.category || '기타';
        byItem[k] = (byItem[k] || 0) + 1;
      }
    });
    var itemKeys = Object.keys(byItem);
    box.className = 'notice-box' + (itemKeys.length ? ' warn' : '');
    var p = el('p');
    p.appendChild(el('b', null, (itemKeys.length ? '⚠ ' : 'ℹ ') + '발주서에서 빠진 주문 ' + out.length + '건'));
    p.appendChild(document.createTextNode(' — ' + Object.keys(byReason).map(function (k) { return k + ' ' + byReason[k] + '건'; }).join(', ')));
    box.appendChild(p);
    if (itemKeys.length) {
      box.appendChild(el('p', { className: 'small' }, '체크하지 않은 품목: ' + itemKeys.map(function (k) { return k + ' ' + byItem[k] + '건'; }).join(', ') + ' — 오늘 발주에 넣으려면 ② 에서 체크하세요.'));
    }
    var btn = el('button', { type: 'button', className: 'link-btn' }, state.showExcluded ? '빠진 주문 숨기기' : '빠진 주문 보기');
    btn.addEventListener('click', toggleExcluded);
    box.appendChild(btn);
    box.hidden = false;
  }

  function toggleExcluded() {
    state.showExcluded = !state.showExcluded;
    $('showExcluded').checked = state.showExcluded;
    renderTable(); renderStats();
  }

  /* ---------------- 표 키보드 조작 ---------------- */

  function cellAt(tr, col) { return tr && tr.querySelector('td[data-col="' + col + '"]'); }

  function moveFocus(td, dRow, dCol) {
    var tr = td.parentNode;
    var ci = COLS.indexOf(td.dataset.col);
    var target;
    if (dRow) {
      var sib = dRow > 0 ? tr.nextElementSibling : tr.previousElementSibling;
      target = cellAt(sib, td.dataset.col);
    } else {
      target = cellAt(tr, COLS[Math.max(0, Math.min(COLS.length - 1, ci + dCol))]);
    }
    if (!target) return;
    var tbody = tr.parentNode;
    tbody.querySelectorAll('td[tabindex="0"]').forEach(function (c) { c.tabIndex = -1; });
    target.tabIndex = 0;
    target.focus();
  }

  document.querySelector('#previewTable tbody').addEventListener('keydown', function (e) {
    var td = e.target.closest('td[data-col]');
    if (!td) return;
    var tr = td.parentNode;
    var r = state.rows.filter(function (x) { return String(x.uid) === tr.dataset.uid; })[0];
    if (td.isContentEditable) {
      if (e.key === 'Enter') { e.preventDefault(); finishEdit(tr, r, td, true); td.focus(); }
      else if (e.key === 'Escape') { e.preventDefault(); finishEdit(tr, r, td, false); td.focus(); }
      return;
    }
    if (e.target !== td) return; // select 등 내부 요소가 처리
    var moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (moves[e.key]) { e.preventDefault(); moveFocus(td, moves[e.key][0], moves[e.key][1]); return; }
    if (e.key === 'Enter' || e.key === 'F2' || e.key === ' ') {
      e.preventDefault();
      var col = td.dataset.col;
      if (col === 'check') { setIncluded(tr, r, !r.included); }
      else if (col === 'cat') { td.querySelector('select').focus(); }
      else if (col === 'act') { if (!td.querySelector('.undo').hidden) revertRow(tr, r); }
      else if (e.key !== ' ') startEdit(td);
    }
  });

  document.querySelector('#previewTable tbody').addEventListener('focusin', function (e) {
    var td = e.target.closest('td[data-col]');
    if (!td || td.tabIndex === 0) return;
    this.querySelectorAll('td[tabindex="0"]').forEach(function (c) { c.tabIndex = -1; });
    td.tabIndex = 0;
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

  function download() {
    var rows = included();
    if (!rows.length) return;
    var bad = rows.filter(function (r) { return blockingIssues(r).length; });
    var ask = Promise.resolve(true);
    if (bad.length) {
      var ul = el('ul');
      bad.slice(0, 8).forEach(function (r) {
        ul.appendChild(el('li', null, (r.name || '(이름 없음)') + ' — ' + blockingIssues(r).map(function (i) { return P.ISSUE_LABEL[i]; }).join(', ')));
      });
      if (bad.length > 8) ul.appendChild(el('li', null, '외 ' + (bad.length - 8) + '건'));
      ask = confirmBox('빈 칸이나 오류가 있는 주문이 ' + bad.length + '건 있습니다',
        [el('p', null, '그대로 받으면 공급처에서 배송하지 못할 수 있습니다.'), ul, el('p', null, '그래도 다운로드할까요?')],
        '그래도 다운로드');
    }
    ask.then(function (ok) {
      if (!ok) return;
      var btn = $('downloadBtn');
      btn.disabled = true;
      var fileName = POExporter.defaultFileName();
      return POExporter.buildWorkbook(ExcelJS, rows).xlsx.writeBuffer().then(function (buf) {
        var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        var a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
        toast('✅ ' + fileName + ' (' + rows.length + '건) 저장했습니다.');
      }).catch(function (e) {
        toast('엑셀을 만들지 못했습니다: ' + e.message);
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
        setBusy(null);
        render();
        $('pickBtn').focus();
      });
  }

  // 샘플 데이터 (스마트스토어 형식, 가상 인물)
  function loadDemo() {
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
    addResult(entry, res);
    state.files.push(entry);
    render();
    toast('샘플 데이터를 불러왔습니다. 실제 파일을 올리면 샘플은 자동으로 사라집니다.');
    scrollToCard('step-preview');
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
  $('showExcluded').addEventListener('change', function (e) {
    state.showExcluded = e.target.checked; renderTable(); renderStats();
  });
  $('selectAllBtn').addEventListener('click', function () {
    state.catalog.forEach(function (c) { state.selected[c.name] = true; });
    state.otherSelected = true; saveCatalog(); render();
  });
  $('selectNoneBtn').addEventListener('click', function () {
    state.catalog.forEach(function (c) { state.selected[c.name] = false; });
    state.otherSelected = false; saveCatalog(); render();
  });
  $('manageBtn').addEventListener('click', function () { state.managing = !state.managing; renderManage(); });
  $('manageDoneBtn').addEventListener('click', function () { state.managing = false; renderManage(); $('manageBtn').focus(); });
  $('addItemForm').addEventListener('submit', function (e) {
    e.preventDefault();
    var input = $('newItemInput');
    var name = input.value.trim();
    if (!name) { input.focus(); return; }
    var n = name.replace(/\s+/g, '');
    if (state.catalog.some(function (c) { return c.name.replace(/\s+/g, '') === n; })) {
      toast('‘' + name + '’ 은(는) 이미 목록에 있습니다.');
      return;
    }
    state.catalog.push({ name: name, keywords: [n] });
    state.selected[name] = true;
    input.value = '';
    saveCatalog(); render();
    toast('‘' + name + '’ 품목을 추가했습니다. 필요하면 키워드를 더 넣어 주세요.');
  });
  $('restoreCatalogBtn').addEventListener('click', function () {
    var before = { catalog: state.catalog, selected: state.selected, other: state.otherSelected };
    resetCatalog(); saveCatalog(); render();
    toast('기본 품목 목록으로 되돌렸습니다.', '취소', function () {
      state.catalog = before.catalog; state.selected = before.selected; state.otherSelected = before.other;
      saveCatalog(); render();
    });
  });

  loadCatalog();
  render();

  // 테스트용 훅
  window.__po = { state: state, handleFiles: handleFiles };
})();
