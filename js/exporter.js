/*
 * 당일발주 양식 엑셀 생성 (ExcelJS). 첨부 템플릿의 서식을 그대로 재현한다.
 *  - 시트명: 당일발주
 *  - 1행 헤더: 굵게, 배경 #9BC2E6, 가운데 정렬, 얇은 테두리, 높이 17.4
 *  - 열 너비: A 16.22 / B 37 / D 21 / E 90.11 / F 29.22
 */
(function (root) {
  'use strict';

  var COLUMNS = [
    { header: '수취자명', key: 'name', width: 16.22, font: 'Arial' },
    { header: '구입품목 ', key: 'item', width: 37, font: 'Arial' },
    { header: '수량', key: 'qty', width: 9, font: 'Arial' },
    { header: '전화번호', key: 'phone', width: 21, font: '맑은 고딕' },
    { header: '주소', key: 'address', width: 90.11, font: '맑은 고딕' },
    { header: '배송메세지', key: 'memo', width: 29.22, font: 'Arial' }
  ];

  var THIN = { style: 'thin', color: { argb: 'FF000000' } };
  var BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };

  // 엑셀 시트 이름에 쓸 수 없는 글자를 빼고 31자로
  function sheetName(name) {
    return String(name || '추가 데이터').replace(/[\[\]:*?\/\\]/g, ' ').trim().slice(0, 31) || '추가 데이터';
  }

  // 숫자는 숫자로, 0으로 시작하는 번호(전화번호 등)는 글자 그대로
  function cellValue(v) {
    return /^-?(0|[1-9]\d{0,14})(\.\d+)?$/.test(String(v)) ? Number(v) : v;
  }

  /** 표(머리글 + 행)를 시트 하나로 추가. 발주서와 같은 머리글 서식 */
  function addTableSheet(wb, title, headers, rows) {
    var ws = wb.addWorksheet(sheetName(title));
    ws.columns = headers.map(function (h, i) {
      var w = String(h).length * 2 + 4;
      rows.slice(0, 200).forEach(function (r) { w = Math.max(w, String(r[i] == null ? '' : r[i]).length * 1.6 + 2); });
      return { key: 'c' + i, width: Math.min(60, Math.max(10, w)) };
    });
    var header = ws.getRow(1);
    headers.forEach(function (h, i) {
      var cell = header.getCell(i + 1);
      cell.value = h;
      cell.font = { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FF000000' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9BC2E6' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = BORDER;
    });
    header.height = 17.4;
    rows.forEach(function (r) {
      var row = ws.addRow(headers.map(function (_, i) { return cellValue(r[i] == null ? '' : r[i]); }));
      row.eachCell({ includeEmpty: true }, function (cell) { cell.font = { name: '맑은 고딕', size: 11 }; });
    });
    return ws;
  }

  /**
   * @param rows  발주할 주문
   * @param extra (선택) { title, headers, rows } — 발주서와 함께 넣을 추가 데이터. 별도 시트로 들어간다
   */
  function buildWorkbook(ExcelJS, rows, extra) {
    var wb = new ExcelJS.Workbook();
    wb.creator = '발주서 자동 변환기';
    wb.created = new Date();
    var ws = wb.addWorksheet('당일발주');
    ws.columns = COLUMNS.map(function (c) { return { key: c.key, width: c.width }; });

    var header = ws.getRow(1);
    COLUMNS.forEach(function (c, i) {
      var cell = header.getCell(i + 1);
      cell.value = c.header;
      cell.font = { name: c.font, size: 11, bold: true, color: { argb: 'FF000000' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9BC2E6' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = BORDER;
    });
    header.height = 17.4;

    rows.forEach(function (r) {
      var row = ws.addRow({
        name: r.name, item: r.item, qty: r.qty, phone: r.phone, address: r.address, memo: r.memo
      });
      row.eachCell({ includeEmpty: true }, function (cell) {
        cell.font = { name: '맑은 고딕', size: 11 };
        cell.alignment = { vertical: 'middle' };
      });
      row.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
      row.getCell(4).numFmt = '@'; // 전화번호 텍스트 유지
    });
    if (extra && extra.rows && extra.rows.length) addTableSheet(wb, extra.title, extra.headers, extra.rows);
    return wb;
  }

  function defaultFileName(date) {
    var d = date || new Date();
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
    // 하루에 여러 번 발주해도 파일명이 겹치지 않도록 시각(HHMM)을 붙인다
    return '당일발주_' + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '_' + pad(d.getHours()) + pad(d.getMinutes()) + '.xlsx';
  }

  var api = { COLUMNS: COLUMNS, buildWorkbook: buildWorkbook, addTableSheet: addTableSheet, sheetName: sheetName, defaultFileName: defaultFileName };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.POExporter = api;
})(typeof self !== 'undefined' ? self : this);
