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

  function buildWorkbook(ExcelJS, rows) {
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
    return wb;
  }

  function defaultFileName(date) {
    var d = date || new Date();
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
    return '당일발주_' + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '.xlsx';
  }

  var api = { COLUMNS: COLUMNS, buildWorkbook: buildWorkbook, defaultFileName: defaultFileName };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.POExporter = api;
})(typeof self !== 'undefined' ? self : this);
