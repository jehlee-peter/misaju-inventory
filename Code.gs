/**
 * 미사주 재고 관리 - Google Apps Script 백엔드
 *
 * ── 배포 방법 ──────────────────────────────────────────
 * 1. https://script.google.com 접속 → 새 프로젝트 생성
 * 2. 이 코드 전체를 붙여넣기 (기존 내용 삭제 후)
 * 3. 상단 메뉴 [배포] → [새 배포]
 * 4. 유형: 웹 앱 / 실행 계정: 나 / 액세스 권한: 모든 사용자
 * 5. [배포] 클릭 → 표시된 URL을 앱 설정에 붙여넣기
 * ────────────────────────────────────────────────────────
 */

// ══════════════════════════════════════════
//  스프레드시트 ID (변경하지 마세요)
// ══════════════════════════════════════════
const STOCK_ID = '1Ij8KLANvBbP5zuo1b1zO7_dV-YiA0TkkYzfSoP_3Fds';
const IN_ID    = '1pQaujzLRJVbw6e8UzdHXdGTGmQX0tbGWzhuJQ3C9OnI';
const OUT_ID   = '13oGxnDWQ-9plcGn6T5_GWJwRwP0QQDpQuoa4pi11wds';

// ══════════════════════════════════════════
//  메인 진입점 (모든 요청 처리)
// ══════════════════════════════════════════
function doGet(e) {
  const p = e.parameter || {};
  let result;

  try {
    switch (p.action) {
      case 'getAll':    result = getAll();         break;
      case 'saveStock': result = saveStock(p);     break;
      case 'addIn':     result = addTxn(p, 'in');  break;
      case 'addOut':    result = addTxn(p, 'out'); break;
      case 'delete':    result = deleteTxn(p);     break;
      case 'ping':      result = { ok: true, msg: '연결 성공' }; break;
      default:          result = { ok: false, msg: 'action 파라미터가 없습니다: ' + p.action };
    }
  } catch (err) {
    result = { ok: false, msg: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════
//  전체 데이터 조회
// ══════════════════════════════════════════
function getAll() {
  return { ok: true, stock: readStock(), txns: readAllTxns() };
}

function readStock() {
  const sheet = SpreadsheetApp.openById(STOCK_ID).getActiveSheet();
  initStockSheet(sheet);
  const data  = sheet.getDataRange().getValues();
  const stock = { white: 0, red: 0 };

  for (let i = 1; i < data.length; i++) {
    const type  = String(data[i][0]).trim();
    const boxes = parseInt(data[i][1]) || 0;
    const bots  = parseInt(data[i][2]) || 0;
    if (type === '화이트') stock.white = boxes * 6 + bots;
    if (type === '레드')   stock.red   = boxes * 6 + bots;
  }
  return stock;
}

function readAllTxns() {
  const txns = [];

  // 입고 내역 읽기
  const inSheet = SpreadsheetApp.openById(IN_ID).getActiveSheet();
  initSheet(inSheet, ['ID', '일자', '종류', '구매박스', '구매낱병', '비고', '생성시각']);
  const inData = inSheet.getDataRange().getValues();

  for (let i = 1; i < inData.length; i++) {
    const r = inData[i];
    if (!r[0]) continue;
    const boxes = parseInt(r[3]) || 0;
    const bots  = parseInt(r[4]) || 0;
    txns.push({
      id:    String(r[0]),
      type:  'in',
      wine:  String(r[2]) === '화이트' ? 'white' : 'red',
      date:  fmtDate(r[1]),
      boxes, bottles: bots,
      total: boxes * 6 + bots,
      note:  String(r[5] || ''),
      ts:    String(r[6] || '')
    });
  }

  // 출고 내역 읽기
  const outSheet = SpreadsheetApp.openById(OUT_ID).getActiveSheet();
  initSheet(outSheet, ['ID', '일자', '종류', '출고박스', '출고낱병', '출고자', '수령자및비고', '생성시각']);
  const outData = outSheet.getDataRange().getValues();

  for (let i = 1; i < outData.length; i++) {
    const r = outData[i];
    if (!r[0]) continue;
    const boxes = parseInt(r[3]) || 0;
    const bots  = parseInt(r[4]) || 0;
    txns.push({
      id:     String(r[0]),
      type:   'out',
      wine:   String(r[2]) === '화이트' ? 'white' : 'red',
      date:   fmtDate(r[1]),
      boxes, bottles: bots,
      total:  boxes * 6 + bots,
      person: String(r[5] || ''),
      note:   String(r[6] || ''),
      ts:     String(r[7] || '')
    });
  }

  // 날짜 내림차순 정렬
  txns.sort((a, b) => {
    if (b.date !== a.date) return b.date > a.date ? 1 : -1;
    return b.ts > a.ts ? 1 : -1;
  });

  return txns;
}

// ══════════════════════════════════════════
//  초기 재고 저장
// ══════════════════════════════════════════
function saveStock(p) {
  const sheet = SpreadsheetApp.openById(STOCK_ID).getActiveSheet();
  initStockSheet(sheet);
  const white = parseInt(p.white) || 0;
  const red   = parseInt(p.red)   || 0;
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const type = String(data[i][0]).trim();
    if (type === '화이트') sheet.getRange(i+1,2,1,2).setValues([[Math.floor(white/6), white%6]]);
    if (type === '레드')   sheet.getRange(i+1,2,1,2).setValues([[Math.floor(red/6),   red%6]]);
  }
  return { ok: true };
}

// ══════════════════════════════════════════
//  입출고 내역 추가
// ══════════════════════════════════════════
function addTxn(p, type) {
  const id    = new Date().getTime() + '_' + Math.random().toString(36).slice(2, 7);
  const ts    = new Date().toISOString();
  const date  = p.date || todayStr();
  const wine  = p.wine === 'white' ? '화이트' : '레드';
  const boxes = parseInt(p.boxes)   || 0;
  const bots  = parseInt(p.bottles) || 0;

  if (type === 'in') {
    const sheet = SpreadsheetApp.openById(IN_ID).getActiveSheet();
    initSheet(sheet, ['ID', '일자', '종류', '구매박스', '구매낱병', '비고', '생성시각']);
    sheet.appendRow([id, date, wine, boxes, bots, p.note || '', ts]);
  } else {
    const sheet = SpreadsheetApp.openById(OUT_ID).getActiveSheet();
    initSheet(sheet, ['ID', '일자', '종류', '출고박스', '출고낱병', '출고자', '수령자및비고', '생성시각']);
    sheet.appendRow([id, date, wine, boxes, bots, p.person || '', p.note || '', ts]);
  }

  return { ok: true, id };
}

// ══════════════════════════════════════════
//  내역 삭제
// ══════════════════════════════════════════
function deleteTxn(p) {
  const id      = String(p.id);
  const txnType = p.txnType; // 'in' or 'out'
  const sheet   = SpreadsheetApp.openById(txnType === 'in' ? IN_ID : OUT_ID).getActiveSheet();
  const data    = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, msg: '해당 내역을 찾을 수 없습니다 (ID: ' + id + ')' };
}

// ══════════════════════════════════════════
//  헬퍼 함수
// ══════════════════════════════════════════

// 초기재고 시트 초기화 (헤더 + 기본 행)
function initStockSheet(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2 || String(data[0][0]) !== '종류') {
    sheet.clearContents();
    sheet.getRange(1, 1, 3, 3).setValues([
      ['종류', '박스', '병'],
      ['화이트', 0, 0],
      ['레드', 0, 0]
    ]);
  }
}

// 일반 시트 헤더 초기화
function initSheet(sheet, headers) {
  const data = sheet.getDataRange().getValues();
  if (data.length === 0 || String(data[0][0]) !== headers[0]) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

// 날짜 포맷 (Date 객체 또는 문자열 → yyyy-MM-dd)
function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val).slice(0, 10);
}

// 오늘 날짜 문자열
function todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
