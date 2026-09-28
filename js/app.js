(function () {
  'use strict';

  var APP_VERSION = 'v2';
  var STORE_KEY = 'po.items.v1';

  var $ = function (id) { return document.getElementById(id); };
  var P = window.POParser;

  var state = {
    files: [],      // {id, name, sig, market, marketLabel, count, error, notice}
    rows: [],       // 변환된 주문 행 (+ uid, fileId, override)
    catalog: [],    // [{name, keywords}]
    selected: {},   // 품목명 → 선택 여부
    otherSelected: true,
    showExcluded: false,
    busy: false,
    gen: 0
  };
  var seq = 0;

  $('appVersion').textContent = APP_VERSION;

  /* ---------------- 품목 목록 (브라우저에 저장) ---------------- */

  function loadCatalog() {
    var saved = null;
    try { saved = JSON.parse(localStorage.getItem(STORE_KEY) || 'null'); } catch (e) { saved = null; }
    if (saved && Array.isArray(saved.catalog)) {
      state.catalog = saved.catalog;
      state.selected = saved.selected || {};
      state.otherSelected = saved.otherSelected !== false;
    } else {
      resetCatalog();
    }
  }

  function resetCatalog() {
    state.catalog = P.DEFAULT_CATALOG.map(function (c) { return { name: c.name, keywords: c.keywords.slice() }; });
    state.selected = {};
    state.catalog.forEach(function (c) { state.selected[c.name] = true; });
    state.otherSelected = true;
  }

  function saveCatalog() {
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify({ catalog: state.catalog, selected: state.selected, otherSelected: state.otherSelected }));
    } catch (e) { /* 저장 불가 환경은 무시 */ }
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

  function friendly(code, extra) {
    var msg = {
      NOT_EXCEL: '엑셀 파일(.xlsx, .xls, .csv)이 아닙니다. 쇼핑몰에서 받은 주문 엑셀을 올려 주세요.',
      EMPTY: '빈 파일입니다. 파일을 다시 받아 주세요.',
      UNREADABLE: '파일을 열 수 없습니다. 엑셀에서 연 뒤 ‘다른 이름으로 저장(.xlsx)’ 해서 다시 올려 주세요.',
      PW_SKIPPED: '비밀번호를 입력하지 않아 건너뛰었습니다. 다시 올리면 비밀번호를 물어봅니다.',
      NO_ROWS: '주문이 한 건도 없습니다(제목 줄만 있는 파일).'
    }[code];
    var err = new Error(msg || extra || '알 수 없는 오류');
    err.code = code;
    return err;
  }

  function workbookFromBytes(file, data) {
    var ext = (file.name.split('.').pop() || '').toLowerCase();
    if (!data.length) return Promise.reject(friendly('EMPTY'));
    if (isZip(data)) return Promise.resolve(XLSX.read(data, { type: 'array', cellDates: true }));
    if (isCfb(data)) {
      if (PODecrypt.isEncrypted(data)) return unlock(file, data).then(function (plain) { return XLSX.read(plain, { type: 'array', cellDates: true }); });
      return Promise.resolve(XLSX.read(data, { type: 'array', cellDates: true }));
    }
    if (ext === 'csv' || ext === 'txt' || ext === 'tsv') {
      return Promise.resolve(XLSX.read(decodeText(data), { type: 'string', cellDates: true }));
    }
    return Promise.reject(friendly('NOT_EXCEL'));
  }

  function unlock(file, data, errorMsg) {
    return askPassword(file.name, errorMsg).then(function (pw) {
      if (pw == null) throw friendly('PW_SKIPPED');
      setBusy('비밀번호로 파일을 여는 중…');
      return PODecrypt.decrypt(data, pw).catch(function (e) {
        if (e.code === 'BAD_PASSWORD') return unlock(file, data, '비밀번호가 맞지 않습니다. 다시 입력해 주세요.');
        throw e;
      });
    });
  }

  function sheetsOf(wb) {
    return wb.SheetNames.map(function (n) {
      return { name: n, rows: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '' }) };
    });
  }

  function fileSig(f) { return [f.name, f.size].join('|'); }

  function handleFiles(fileList) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return Promise.resolve();
    var gen = state.gen;
    var dupNames = [];
    files = files.filter(function (f) {
      var sig = fileSig(f);
      var dup = state.files.some(function (x) { return x.sig === sig; });
      if (dup) dupNames.push(f.name);
      return !dup;
    });
    if (dupNames.length) toast('이미 올린 파일이라 건너뛰었습니다: ' + dupNames.join(', '));
    // 순서를 지키며 하나씩 처리 (비밀번호 입력창이 겹치지 않도록)
    var chain = Promise.resolve();
    files.forEach(function (f) {
      chain = chain.then(function () {
        if (gen !== state.gen) return;
        setBusy('‘' + f.name + '’ 읽는 중…');
        return processFile(f, gen);
      });
    });
    return chain.catch(function (e) {
      toast('처리 중 오류가 발생했습니다: ' + (e && e.message));
    }).then(function () {
      if (gen === state.gen) { setBusy(null); render(); }
    });
  }

  function processFile(f, gen) {
    var entry = { id: ++seq, name: f.name, sig: fileSig(f) };
    return readFile(f)
      .then(function (data) { return workbookFromBytes(f, data); })
      .then(function (wb) {
        var res = P.convertSheets(sheetsOf(wb));
        if (res.error) throw friendly(res.code, res.error);
        if (!res.rows.length) throw friendly('NO_ROWS');
        addResult(entry, res);
      })
      .catch(function (e) {
        entry.error = e && e.code ? e.message : friendly('UNREADABLE').message;
      })
      .then(function () {
        if (gen !== state.gen) return;
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
      state.rows.push(r);
    });
  }

  /* ---------------- 행 상태 계산 ---------------- */

  function validate(r) {
    var keepDefault = r.issues && r.issues.indexOf('qty-default') !== -1 && !r.qtyEdited;
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

  function compute() {
    var seen = {};
    state.rows.forEach(function (r) {
      var reasons = [];
      if (r.excluded) reasons.push('취소·반품 (' + r.excluded + ')');
      if (seen[r.key]) reasons.push('중복 주문');
      seen[r.key] = true;
      r.category = P.classify(r.item, state.catalog);
      var catOn = r.category ? !!state.selected[r.category] : state.otherSelected;
      if (!catOn) reasons.push('품목 미선택');
      r.reasons = reasons;
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
    renderFiles();
    renderChips();
    renderTable();
    renderStats();
  }

  function renderFiles() {
    var list = $('fileList');
    list.innerHTML = '';
    state.files.forEach(function (f) {
      var li = el('li', { className: f.error ? 'error' : f.notice ? 'notice' : '' });
      if (f.error) {
        li.appendChild(el('span', { 'aria-hidden': 'true' }, '⛔'));
      } else {
        li.appendChild(el('span', { className: 'badge ' + f.market }, f.marketLabel));
      }
      li.appendChild(el('span', { className: 'file-name' }, f.name));
      if (!f.error) {
        var rows = state.rows.filter(function (r) { return r.fileId === f.id; });
        var inc = rows.filter(function (r) { return r.included; }).length;
        var meta = '주문 ' + rows.length + '건 → 발주 ' + inc + '건';
        if (rows.length - inc) meta += ' (제외 ' + (rows.length - inc) + '건)';
        li.appendChild(el('span', { className: 'file-meta' }, meta));
      }
      var rm = el('button', { type: 'button', className: 'icon-btn', 'aria-label': f.name + ' 빼기', title: '이 파일 빼기' }, '✕');
      rm.addEventListener('click', function () { removeFile(f.id); });
      li.appendChild(rm);
      if (f.error || f.notice) li.appendChild(el('p', { className: 'file-msg' }, (f.error ? '' : '⚠ ') + (f.error || f.notice)));
      list.appendChild(li);
    });
    $('resetBtn').hidden = state.files.length === 0;
  }

  function candidateCount(pred) {
    return state.rows.filter(function (r) {
      return r.override !== false && !r.excluded && r.reasons.indexOf('중복 주문') === -1 && pred(r);
    }).length;
  }

  function chip(label, checked, count, onToggle, onRemove, extraClass) {
    var lab = el('label', { className: 'chip' + (checked ? ' checked' : '') + (extraClass ? ' ' + extraClass : '') });
    var cb = el('input', { type: 'checkbox' });
    cb.checked = checked;
    cb.addEventListener('change', function () { onToggle(cb.checked); });
    lab.appendChild(cb);
    lab.appendChild(el('span', null, label));
    if (state.rows.length) lab.appendChild(el('span', { className: 'count', 'aria-label': count + '건' }, String(count)));
    if (onRemove) {
      var x = el('button', { type: 'button', className: 'remove', 'aria-label': label + ' 품목 삭제', title: '품목 삭제' }, '✕');
      x.addEventListener('click', function (e) { e.preventDefault(); onRemove(); });
      lab.appendChild(x);
    }
    return lab;
  }

  function renderChips() {
    var box = $('itemChips');
    box.innerHTML = '';
    state.catalog.forEach(function (c, idx) {
      box.appendChild(chip(c.name, !!state.selected[c.name],
        candidateCount(function (r) { return r.category === c.name; }),
        function (on) { state.selected[c.name] = on; saveCatalog(); render(); },
        function () {
          state.catalog.splice(idx, 1);
          delete state.selected[c.name];
          saveCatalog(); render();
          toast('‘' + c.name + '’ 품목을 목록에서 뺐습니다.');
        }));
    });
    box.appendChild(chip('기타 (목록에 없는 상품)', state.otherSelected,
      candidateCount(function (r) { return !r.category; }),
      function (on) { state.otherSelected = on; saveCatalog(); render(); }, null, 'other'));
  }

  var EDIT_FIELDS = ['name', 'item', 'qty', 'phone', 'address', 'memo'];
  var ISSUE_FIELD = { name: 'name', item: 'item', phone: 'phone', 'phone-format': 'phone', address: 'address', qty: 'qty', 'qty-default': 'qty' };

  function renderTable() {
    var tbody = document.querySelector('#previewTable tbody');
    tbody.innerHTML = '';
    var has = state.rows.length > 0;
    $('emptyPreview').hidden = has;
    $('previewBody').hidden = !has;
    state.rows.forEach(function (r) {
      if (!r.included && !state.showExcluded && r.override == null) return;
      tbody.appendChild(buildRow(r));
    });
    if (has && !tbody.children.length) {
      var tr = el('tr');
      var td = el('td', { colspan: '9', className: 'empty' }, '선택한 품목의 주문이 없습니다. ② 에서 품목을 골라 주세요.');
      tr.appendChild(td);
      tbody.appendChild(tr);
    }
  }

  function buildRow(r) {
    var tr = el('tr', { 'data-uid': r.uid });
    var tdc = el('td', { className: 'c-check' });
    var cb = el('input', { type: 'checkbox', 'aria-label': (r.name || '이름 없음') + ' 주문 발주 포함' });
    cb.checked = r.included;
    cb.addEventListener('change', function () {
      r.override = cb.checked === (r.reasons.length === 0) ? null : cb.checked;
      compute(); updateRow(tr, r); renderStats(); renderFiles(); renderChips();
    });
    tdc.appendChild(cb);
    tr.appendChild(tdc);
    tr.appendChild(el('td', null)).appendChild(el('span', { className: 'badge ' + r.market }, r.source));
    tr.appendChild(el('td', { className: 'c-cat' }));
    EDIT_FIELDS.forEach(function (f) {
      var td = el('td', { contenteditable: 'plaintext-only', 'data-field': f, spellcheck: 'false' }, r[f] == null ? '' : String(r[f]));
      if (td.contentEditable !== 'plaintext-only') td.contentEditable = 'true';
      if (f === 'qty') td.className = 'c-qty';
      td.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); td.blur(); }
        if (e.key === 'Escape') { td.textContent = r[f] == null ? '' : String(r[f]); td.blur(); }
      });
      td.addEventListener('blur', function () { commitEdit(tr, r, f, td.textContent); });
      tr.appendChild(td);
    });
    updateRow(tr, r);
    return tr;
  }

  function commitEdit(tr, r, field, text) {
    text = text.replace(/\s*\n\s*/g, ' ').trim();
    var before = r[field];
    if (field === 'qty') {
      var n = P.parseQty(text);
      r.qty = n == null ? text : n;
      if (String(before) !== String(r.qty)) r.qtyEdited = true;
    } else if (field === 'phone') {
      var ph = P.normalizePhone(text);
      r.phone = ph.value;
    } else {
      r[field] = text;
    }
    validate(r);
    compute();
    updateRow(tr, r);
    renderStats(); renderChips(); renderFiles();
  }

  function updateRow(tr, r) {
    tr.className = r.included ? '' : 'excluded';
    tr.querySelector('.c-check input').checked = r.included;
    var cat = tr.querySelector('.c-cat');
    cat.textContent = r.category || '기타';
    if (r.reasons.length) {
      cat.appendChild(el('span', { className: 'row-reason' }, (r.included ? '포함함: ' : '제외: ') + r.reasons.join(', ')));
    }
    var bad = {};
    r.issues.forEach(function (i) { bad[ISSUE_FIELD[i]] = P.BLOCKING[i] ? 'bad' : (bad[ISSUE_FIELD[i]] || 'soft'); });
    EDIT_FIELDS.forEach(function (f) {
      var td = tr.querySelector('td[data-field="' + f + '"]');
      if (document.activeElement !== td) td.textContent = r[f] == null ? '' : String(r[f]);
      td.classList.remove('bad', 'soft');
      td.removeAttribute('title');
      if (bad[f] && r.included) {
        td.classList.add(bad[f]);
        td.title = r.issues.filter(function (i) { return ISSUE_FIELD[i] === f; }).map(function (i) { return P.ISSUE_LABEL[i]; }).join(', ');
      }
    });
  }

  function included() { return state.rows.filter(function (r) { return r.included; }); }

  function renderStats() {
    var inc = included();
    var byMarket = {};
    inc.forEach(function (r) { byMarket[r.source] = (byMarket[r.source] || 0) + 1; });
    var problems = inc.filter(function (r) { return blockingIssues(r).length; }).length;
    var soft = inc.filter(function (r) { return !blockingIssues(r).length && r.issues.length; }).length;
    var excluded = state.rows.length - inc.length;
    var stats = $('stats');
    stats.innerHTML = '';
    function stat(cls, num, label) {
      var d = el('div', { className: 'stat ' + cls });
      d.appendChild(el('b', null, String(num)));
      d.appendChild(document.createTextNode(label));
      stats.appendChild(d);
    }
    stat('main', inc.length, '발주할 주문');
    Object.keys(byMarket).forEach(function (m) { stat('', byMarket[m], m); });
    if (excluded) stat('', excluded, '제외된 주문');
    if (problems) stat('danger', problems, '빈 칸·오류 있음');
    if (soft) stat('warn', soft, '확인 권장');

    var bar = $('actionbar');
    bar.hidden = state.rows.length === 0;
    var txt = $('actionText');
    txt.innerHTML = '';
    txt.appendChild(document.createTextNode('발주 '));
    txt.appendChild(el('b', null, inc.length + '건'));
    if (excluded) txt.appendChild(document.createTextNode(' · 제외 ' + excluded + '건'));
    if (problems) {
      txt.appendChild(document.createTextNode(' · '));
      txt.appendChild(el('span', { className: 'warn-text' }, '빨간 칸 ' + problems + '건 확인 필요'));
    }
    $('downloadBtn').disabled = inc.length === 0;

    var s = [$('stepper-1'), $('stepper-2'), $('stepper-3')];
    s.forEach(function (x) { x.className = ''; });
    if (state.rows.length) { s[0].className = 'done'; s[1].className = 'done'; s[2].className = 'active'; }
    else s[0].className = 'active';
  }

  /* ---------------- 동작 ---------------- */

  function removeFile(id) {
    state.files = state.files.filter(function (f) { return f.id !== id; });
    state.rows = state.rows.filter(function (r) { return r.fileId !== id; });
    render();
  }

  function setBusy(text) {
    state.busy = !!text;
    $('busy').hidden = !text;
    if (text) $('busyText').textContent = text;
  }

  var toastTimer;
  function toast(msg) {
    var t = $('toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.hidden = true; }, 4000);
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
    confirmBox('처음부터 다시 할까요?', [el('p', null, '올린 파일과 미리보기에서 고친 내용이 모두 지워집니다. (품목 목록은 그대로 남습니다.)')], '모두 지우기')
      .then(function (ok) {
        if (!ok) return;
        state.gen++;
        state.files = [];
        state.rows = [];
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
    var entry = { id: ++seq, name: '샘플 데이터 (스마트스토어 형식)', sig: 'demo|' + seq };
    addResult(entry, res);
    state.files.push(entry);
    render();
    toast('샘플 데이터를 불러왔습니다. ② 에서 품목을 바꿔 보세요.');
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
  // 영역 밖에 떨어뜨려도 브라우저가 파일을 열어버리지 않게
  ['dragover', 'drop'].forEach(function (ev) { window.addEventListener(ev, function (e) { e.preventDefault(); }); });

  $('downloadBtn').addEventListener('click', download);
  $('resetBtn').addEventListener('click', reset);
  $('demoBtn').addEventListener('click', loadDemo);
  $('showExcluded').addEventListener('change', function (e) { state.showExcluded = e.target.checked; renderTable(); });
  $('selectAllBtn').addEventListener('click', function () {
    state.catalog.forEach(function (c) { state.selected[c.name] = true; });
    state.otherSelected = true; saveCatalog(); render();
  });
  $('selectNoneBtn').addEventListener('click', function () {
    state.catalog.forEach(function (c) { state.selected[c.name] = false; });
    state.otherSelected = false; saveCatalog(); render();
  });
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
    toast('‘' + name + '’ 품목을 추가했습니다.');
  });
  $('restoreCatalogBtn').addEventListener('click', function () {
    resetCatalog(); saveCatalog(); render();
    toast('기본 품목 목록으로 되돌렸습니다.');
  });

  loadCatalog();
  render();

  // 테스트용 훅
  window.__po = { state: state, handleFiles: handleFiles };
})();
