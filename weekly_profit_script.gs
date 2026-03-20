// 주간 시간당 채산이익 — Google Apps Script 웹훅
// 사용법: 확장 프로그램 → Apps Script → 붙여넣기 → 배포 → 웹 앱으로 배포

const SHEET_ID = '1gwxAe3RMLaa1rROeWSk0yBAnxdumAkh6J-du65zzcoI';
const SHEET_NAME = '주간기록';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'save') {
      saveWeekData(data);
      return jsonResponse({ status: 'ok', message: '저장 완료' });
    }
    if (action === 'load') {
      const rows = loadAllData();
      return jsonResponse({ status: 'ok', data: rows });
    }
    if (action === 'delete') {
      deleteWeekData(data.weekKey);
      return jsonResponse({ status: 'ok', message: '삭제 완료' });
    }
    return jsonResponse({ status: 'error', message: '알 수 없는 action' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  const rows = loadAllData();
  return jsonResponse({ status: 'ok', data: rows });
}

function saveWeekData(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      '주차키', '주차라벨', '저장일시',
      '목표_매출', '목표_매출원가', '목표_변동비', '목표_고정비', '목표_영업외손익', '목표_근무시간',
      '목표_공헌이익', '목표_채산이익', '목표_시간당채산이익',
      '실적_매출', '실적_매출원가', '실적_변동비', '실적_고정비', '실적_영업외손익', '실적_근무시간',
      '실적_공헌이익', '실적_채산이익', '실적_시간당채산이익',
      '달성률_매출', '달성률_채산이익'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // 기존 주차 행 찾아서 업데이트
  const allData = sheet.getDataRange().getValues();
  let existingRow = -1;
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.weekKey) { existingRow = i + 1; break; }
  }

  const g = data.goal || {};
  const a = data.actual || {};
  const achSales = g.sales > 0 && a.sales > 0 ? (a.sales / g.sales * 100).toFixed(1) + '%' : '-';
  const achProfit = g.profit > 0 && a.profit ? (a.profit / g.profit * 100).toFixed(1) + '%' : '-';

  const row = [
    data.weekKey,
    data.weekLabel,
    new Date().toLocaleString('ko-KR'),
    g.sales || 0, g.cogs || 0, g.varC || 0, g.fixed || 0, g.other || 0, g.hours || 0,
    g.contrib || 0, g.profit || 0,
    g.hours > 0 ? Math.round((g.profit * 1000000) / g.hours) : 0,
    a.sales || 0, a.cogs || 0, a.varC || 0, a.fixed || 0, a.other || 0, a.hours || 0,
    a.contrib || 0, a.profit || 0,
    a.hours > 0 ? Math.round((a.profit * 1000000) / a.hours) : 0,
    achSales, achProfit
  ];

  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  // 숫자 셀 서식
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 4, 1, 9).setNumberFormat('#,##0');
  sheet.getRange(lastRow, 13, 1, 9).setNumberFormat('#,##0');
}

function loadAllData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function deleteWeekData(weekKey) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  const allData = sheet.getDataRange().getValues();
  for (let i = allData.length - 1; i >= 1; i--) {
    if (allData[i][0] === weekKey) { sheet.deleteRow(i + 1); break; }
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
