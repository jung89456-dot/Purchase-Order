(function () {
  'use strict';

  var APP_VERSION = 'v1';

  var state = { files: [], rows: [] };

  var $ = function (id) { return document.getElementById(id); };
  var fileInput = $('fileInput');
  var dropzone = $('dropzone');

  $('appVersion').textContent = APP_VERSION;

  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function (e) { resolve(new Uint8Array(e.target.result)); };
      reader.onerror = function () { reject(reader.error); };
      reader.readAsArrayBuffer(file);
    });
  }

  function parseWorkbook(data) {
    var wb = XLSX.read(data, { type: 'array', cellDates: true });
    var sheet = wb.Sheets[wb.SheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' });
    return POParser.convertSheet(rows);
  }

  function handleFiles(fileList) {
    var files = Array.prototype.slice.call(fileList);
    return Promise.all(files.map(function (f) {
      return readFile(f).then(function (data) {
        var result;
        try { result = parseWorkbook(data); }
        catch (e) { result = { error: '파일을 읽을 수 없습니다: ' + e.message }; }
        state.files.push({ name: f.name, result: result });
      });
    })).then(render);
  }

  function render() {
    var list = $('fileList');
    list.innerHTML = '';
    state.rows = [];
    state.files.forEach(function (f) {
      var li = document.createElement('li');
      if (f.result.error) {
        li.innerHTML = '<span class="error"></span>';
        li.firstChild.textContent = f.name + ' — ' + f.result.error;
      } else {
        var b = document.createElement('span');
        b.className = 'badge ' + f.result.market;
        b.textContent = f.result.marketLabel;
        li.appendChild(b);
        var txt = f.name + ' — ' + f.result.rows.length + '건';
        if (f.result.skipped.length) txt += ' (취소/반품 ' + f.result.skipped.length + '건 제외)';
        if (f.result.missing.length) txt += ' ⚠ 누락 항목: ' + f.result.missing.join(', ');
        li.appendChild(document.createTextNode(txt));
        state.rows = state.rows.concat(f.result.rows);
      }
      list.appendChild(li);
    });

    var tbody = document.querySelector('#previewTable tbody');
    tbody.innerHTML = '';
    state.rows.forEach(function (r) {
      var tr = document.createElement('tr');
      ['name', 'item', 'qty', 'phone', 'address', 'memo'].forEach(function (k) {
        var td = document.createElement('td');
        td.textContent = r[k] == null ? '' : r[k];
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    $('summary').textContent = '총 ' + state.rows.length + '건의 발주 데이터가 준비되었습니다.';
    $('step-preview').hidden = state.rows.length === 0;
    $('step-download').hidden = state.rows.length === 0;
  }

  function download() {
    var wb = POExporter.buildWorkbook(ExcelJS, state.rows);
    wb.xlsx.writeBuffer().then(function (buf) {
      var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = POExporter.defaultFileName();
      document.body.appendChild(a);
      a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    });
  }

  fileInput.addEventListener('change', function () {
    handleFiles(fileInput.files);
    fileInput.value = '';
  });
  ['dragenter', 'dragover'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) { e.preventDefault(); dropzone.classList.add('drag'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    dropzone.addEventListener(ev, function (e) { e.preventDefault(); dropzone.classList.remove('drag'); });
  });
  dropzone.addEventListener('drop', function (e) { handleFiles(e.dataTransfer.files); });
  $('downloadBtn').addEventListener('click', download);
  $('resetBtn').addEventListener('click', function () { state.files = []; render(); });
})();
