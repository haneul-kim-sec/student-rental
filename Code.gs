/**
 * 스마트보안학과 학생회 물품 대여 — Google Apps Script
 *
 * 시트 구성 (setup()을 한 번 실행하면 자동으로 만들어져요)
 *  - "물품"     : 물품명, 메모, 총수량, 대여기간(일)을 학생회가 직접 관리
 *  - "대여기록" : 대여/반납이 자동으로 쌓이는 장부 (학번·이름 포함 → 시트는 학생회만 공유)
 *
 * 학생 화면에는 물품별 남은 수량과 반납 기한만 전달되고, 학번·이름은 절대 내려가지 않아요.
 */

const TZ = 'Asia/Seoul';
const ITEM_SHEET = '물품';
const LOG_SHEET = '대여기록';
const ITEM_HEADERS = ['ID(자동)', '물품명', '메모', '총수량', '대여기간(일)', '분류'];
const LOG_HEADERS = ['대여ID', '대여시각', '물품ID', '물품명', '학번', '이름', '반납기한', '반납시각', '상태'];
const COL = { id: 1, at: 2, itemId: 3, itemName: 4, sid: 5, name: 6, due: 7, returnedAt: 8, status: 9 };
const LENT = '대여중';
const RETURNED = '반납완료';

/* ---------- 웹 페이지 ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('학생회 물품 대여')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- 처음 한 번만 실행: 시트 만들기 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);

  let items = ss.getSheetByName(ITEM_SHEET);
  if (!items) {
    items = ss.insertSheet(ITEM_SHEET);
    items.getRange(1, 1, 1, ITEM_HEADERS.length).setValues([ITEM_HEADERS]).setFontWeight('bold');
    const examples = [
      ['charger-c', 'C타입 충전기', '케이블과 어댑터가 한 세트예요', 5, 1, '전자기기'],
      ['charger-8', '8핀 충전기', '아이폰용 · 케이블과 어댑터가 한 세트예요', 3, 1, '전자기기'],
      ['umbrella', '장우산', '비 오는 날 사용 후 말려서 반납해 주세요', 10, 3, '생활용품'],
      ['hairtie', '머리끈', '실험·체육 활동 때 쓰기 좋아요', 20, 7, '생활용품'],
      ['pen', '볼펜', '검정 볼펜', 20, 7, '학습용품'],
      ['mat', '돗자리', '2~4인용 · 털어서 반납해 주세요', 4, 3, '여가용품'],
      ['curler', '고데기', '사용 후 전원을 끄고 식혀서 반납해 주세요', 1, 1, '생활용품'],
    ];
    items.getRange(2, 1, examples.length, examples[0].length).setValues(examples);
    items.setFrozenRows(1);
    items.setColumnWidth(1, 120); items.setColumnWidth(2, 200); items.setColumnWidth(3, 220);
    items.getRange('A1').setNote('ID는 비워 두면 자동으로 채워져요. 대여 중인 물품의 ID는 바꾸지 마세요.');
  }

  let log = ss.getSheetByName(LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(LOG_SHEET);
    log.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold');
    log.setFrozenRows(1);
    log.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
    log.getRange('H:H').setNumberFormat('yyyy-mm-dd hh:mm');
    log.getRange('E:E').setNumberFormat('@'); // 학번은 텍스트로
    log.getRange('G:G').setNumberFormat('@'); // 반납기한 yyyy-mm-dd 텍스트
    // 반납 기한이 지났는데 아직 대여중이면 빨간색으로 표시
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($I2="' + LENT + '", $G2<>"", DATEVALUE($G2)<TODAY())')
      .setBackground('#F9DEDB').setFontColor('#B3261E')
      .setRanges([log.getRange('A2:I')]).build();
    log.setConditionalFormatRules([rule]);
  }

  const first = ss.getSheets()[0];
  if (first.getName() !== ITEM_SHEET && first.getName() !== LOG_SHEET && first.getLastRow() === 0) ss.deleteSheet(first);
}

/* ---------- 학생 화면용: 물품별 수량 (개인정보 없음) ---------- */
function getStatus() {
  const items = readItems_();
  const lent = countLent_();
  return items.map((it) => ({
    id: it.id, name: it.name, note: it.note, category: it.category, total: it.total, days: it.days,
    lent: lent[it.id] || 0,
    remaining: Math.max(0, it.total - (lent[it.id] || 0)),
  }));
}

/* ---------- 대여 ---------- */
function borrow(itemId, sid, name, consent) {
  const v = validate_(sid, name);
  if (!consent) throw new Error('개인정보 수집에 동의해야 대여할 수 있어요.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const it = readItems_().find((x) => x.id === itemId);
    if (!it) throw new Error('물품을 찾을 수 없어요. 새로고침해 주세요.');
    const lentNow = countLent_()[it.id] || 0;
    if (lentNow >= it.total) throw new Error('방금 모두 대여됐어요.');
    const now = new Date();
    const due = Utilities.formatDate(new Date(now.getTime() + it.days * 86400000), TZ, 'yyyy-MM-dd');
    const id = 'L' + now.getTime().toString(36) + Math.random().toString(36).slice(2, 6);
    sheet_(LOG_SHEET).appendRow([id, now, it.id, safe_(it.name), v.sid, v.name, due, '', LENT]);
    return { due: due, dueLabel: dueLabel_(due), itemName: it.name };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 반납 ---------- */
function returnItem(itemId, sid, name) {
  const v = validate_(sid, name);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(LOG_SHEET);
    const last = sh.getLastRow();
    if (last < 2) throw new Error('대여 기록을 찾을 수 없어요. 빌릴 때 적은 학번과 이름을 확인해 주세요.');
    const rows = sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues();
    const clean = (s) => String(s).replace(/^'/, '').replace(/\s+/g, '');
    let hit = -1, hitDue = null;
    rows.forEach((r, i) => {
      if (r[COL.status - 1] === LENT && r[COL.itemId - 1] === itemId &&
          clean(r[COL.sid - 1]) === v.sid && clean(r[COL.name - 1]) === clean(v.name)) {
        const due = String(r[COL.due - 1]);
        if (hit === -1 || due < hitDue) { hit = i; hitDue = due; }
      }
    });
    if (hit === -1) throw new Error('이 학번·이름으로 빌린 기록이 없어요. 빌릴 때 적은 그대로 입력해 주세요.');
    const row = hit + 2;
    sh.getRange(row, COL.returnedAt).setValue(new Date());
    sh.getRange(row, COL.status).setValue(RETURNED);
    const late = hitDue && hitDue < Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    return { late: !!late };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 내부 함수 ---------- */
function sheet_(n) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);
  if (!sh) throw new Error('시트가 준비되지 않았어요. 학생회에 문의해 주세요.');
  return sh;
}

function readItems_() {
  const sh = sheet_(ITEM_SHEET);
  // 예전 시트에 '분류' 열이 없으면 머리글을 자동으로 추가
  const headCell = sh.getRange(1, ITEM_HEADERS.length);
  if (!String(headCell.getValue()).trim()) headCell.setValue(ITEM_HEADERS[ITEM_HEADERS.length - 1]).setFontWeight('bold');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const range = sh.getRange(2, 1, last - 1, ITEM_HEADERS.length);
  const vals = range.getValues();
  let changed = false;
  const out = [];
  vals.forEach((r) => {
    const name = String(r[1]).trim();
    if (!name) return;
    if (!String(r[0]).trim()) { r[0] = 'i' + Utilities.getUuid().slice(0, 8); changed = true; }
    out.push({
      id: String(r[0]).trim(), name: name, note: String(r[2] || '').trim(),
      total: Math.max(0, parseInt(r[3], 10) || 0), days: Math.max(0, parseInt(r[4], 10) || 0),
      category: String(r[5] || '').trim(),
    });
  });
  if (changed) range.setValues(vals);
  return out;
}

function countLent_() {
  const sh = sheet_(LOG_SHEET);
  const last = sh.getLastRow();
  const counts = {};
  if (last < 2) return counts;
  sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues().forEach((r) => {
    if (r[COL.status - 1] === LENT) counts[r[COL.itemId - 1]] = (counts[r[COL.itemId - 1]] || 0) + 1;
  });
  return counts;
}

function validate_(sid, name) {
  sid = String(sid || '').trim();
  name = String(name || '').trim();
  if (!/^\d{8,10}$/.test(sid)) throw new Error('학번은 숫자 8~10자리로 적어 주세요.');
  if (name.length < 2 || name.length > 20) throw new Error('이름을 2~20자로 적어 주세요.');
  return { sid: sid, name: safe_(name) };
}

// 시트 수식 삽입 방지: =, +, -, @ 로 시작하는 값은 텍스트로 저장
function safe_(s) {
  s = String(s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function dueLabel_(ymd) {
  const p = ymd.split('-').map(Number);
  const wd = ['일', '월', '화', '수', '목', '금', '토'][new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay()];
  return p[1] + '월 ' + p[2] + '일(' + wd + ')';
}
