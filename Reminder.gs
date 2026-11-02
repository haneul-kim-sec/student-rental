/**
 * 반납 기한 알림 메일 — Reminder.gs
 *
 * 매일 아침 9시쯤, 학생회에 이런 메일을 보내요.
 *  - 내일이 반납 기한인 대여 목록
 *  - 오늘이 반납 기한인 대여 목록
 *  - 반납 기한이 이미 지난 대여 목록
 * 해당하는 대여가 하나도 없는 날에는 메일을 보내지 않아요.
 *
 * 이 파일은 Code.gs에 있는 TZ, LOG_SHEET, COL, LENT, sheet_ 를 그대로 같이 써요.
 * (Apps Script에서는 같은 프로젝트의 .gs 파일끼리 변수와 함수를 공유해요.)
 *
 * 처음 한 번만: 함수 선택 칸에서 installReminder 를 골라 '실행' → 권한 허용
 * 바로 테스트:  함수 선택 칸에서 testReminder 를 골라 '실행' → 메일함 확인
 * 알림 끄기:    함수 선택 칸에서 removeReminder 를 골라 '실행'
 */

// 알림을 받을 사람. 시트 주인(나)은 자동으로 받고, 다른 임원을 추가하려면 따옴표 안에 이메일을 넣어요.
// 예: const EXTRA_RECIPIENTS = ['vice@gmail.com', 'president@gmail.com'];
const EXTRA_RECIPIENTS = [];

const REMINDER_HOUR = 9; // 매일 몇 시에 보낼지 (한국 시간, 0~23)

/* ---------- 매일 자동으로 실행되는 함수 ---------- */
function sendDailyReminder() {
  const report = buildReport_();
  if (report.tomorrow.length + report.today.length + report.overdue.length === 0) return; // 보낼 게 없으면 조용히 끝
  sendReport_(report);
}

/* ---------- 테스트: 지금 바로 한 번 보내 보기 (목록이 비어 있어도 보냄) ---------- */
function testReminder() {
  sendReport_(buildReport_());
}

/* ---------- 처음 한 번만 실행: 매일 알림 예약 ---------- */
function installReminder() {
  removeReminder(); // 같은 알림이 두 번 예약되지 않게 먼저 지움
  ScriptApp.newTrigger('sendDailyReminder')
    .timeBased()
    .everyDays(1)
    .atHour(REMINDER_HOUR)
    .inTimezone(TZ)
    .create();
  Logger.log('매일 ' + REMINDER_HOUR + '시쯤 반납 알림을 보내도록 예약했어요.');
}

/* ---------- 알림 끄기 ---------- */
function removeReminder() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'sendDailyReminder')
    .forEach((t) => ScriptApp.deleteTrigger(t));
}

/* ---------- 내부 함수 ---------- */

// 대여기록 시트에서 '대여중'인 줄만 골라 내일/오늘/기한 지남으로 나눠요.
function buildReport_() {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const tomorrow = Utilities.formatDate(new Date(Date.now() + 86400000), TZ, 'yyyy-MM-dd');
  const report = { today: [], tomorrow: [], overdue: [], todayStr: today };

  const sh = sheet_(LOG_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return report;

  const rows = sh.getRange(2, 1, last - 1, COL.status).getValues();
  rows.forEach((r) => {
    if (r[COL.status - 1] !== LENT) return;
    const due = toYmd_(r[COL.due - 1]);
    if (!due) return;
    const row = {
      item: String(r[COL.itemName - 1]).replace(/^'/, ''),
      sid: String(r[COL.sid - 1]).replace(/^'/, ''),
      name: String(r[COL.name - 1]).replace(/^'/, ''),
      due: due,
      lateDays: daysBetween_(due, today),
    };
    if (due === tomorrow) report.tomorrow.push(row);
    else if (due === today) report.today.push(row);
    else if (due < today) report.overdue.push(row);
  });
  report.overdue.sort((a, b) => b.lateDays - a.lateDays); // 오래 밀린 순
  return report;
}

// 반납기한 칸이 글자('2026-11-03')든 날짜든 'yyyy-MM-dd'로 맞춰요.
function toYmd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function daysBetween_(from, to) {
  const a = Date.UTC(+from.slice(0, 4), +from.slice(5, 7) - 1, +from.slice(8, 10));
  const b = Date.UTC(+to.slice(0, 4), +to.slice(5, 7) - 1, +to.slice(8, 10));
  return Math.round((b - a) / 86400000);
}

function sendReport_(report) {
  const owner = Session.getEffectiveUser().getEmail();
  const to = [owner].concat(EXTRA_RECIPIENTS).filter((e, i, arr) => e && arr.indexOf(e) === i).join(',');
  const n = report.tomorrow.length + report.today.length + report.overdue.length;

  const subject = '[학생회 물품 대여] ' + dueLabel_(report.todayStr) + ' 반납 알림'
    + (report.overdue.length ? ' · 기한 지남 ' + report.overdue.length + '건' : '')
    + (n === 0 ? ' · 해당 없음' : '');

  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const table = (title, color, list, showLate) => {
    if (!list.length) return '';
    const rows = list.map((r) =>
      '<tr>' +
      '<td style="padding:6px 10px;border-bottom:1px solid #eee">' + esc(r.item) + '</td>' +
      '<td style="padding:6px 10px;border-bottom:1px solid #eee;font-family:monospace">' + esc(r.sid) + '</td>' +
      '<td style="padding:6px 10px;border-bottom:1px solid #eee">' + esc(r.name) + '</td>' +
      '<td style="padding:6px 10px;border-bottom:1px solid #eee">' + dueLabel_(r.due) + (showLate ? ' <b style="color:#b3261e">(' + r.lateDays + '일 지남)</b>' : '') + '</td>' +
      '</tr>').join('');
    return '<h3 style="margin:20px 0 8px;color:' + color + '">' + title + ' ' + list.length + '건</h3>' +
      '<table style="border-collapse:collapse;font-size:14px"><tr style="background:#f4f5f0">' +
      '<th style="padding:6px 10px;text-align:left">물품</th><th style="padding:6px 10px;text-align:left">학번</th>' +
      '<th style="padding:6px 10px;text-align:left">이름</th><th style="padding:6px 10px;text-align:left">반납 기한</th></tr>' +
      rows + '</table>';
  };

  const html =
    '<div style="font-family:Apple SD Gothic Neo,Malgun Gothic,sans-serif;color:#17211f">' +
    '<p>오늘(' + dueLabel_(report.todayStr) + ') 기준 반납 현황이에요.</p>' +
    (n === 0 ? '<p>내일·오늘 반납 예정이거나 기한이 지난 대여가 없어요. (테스트 메일)</p>' : '') +
    table('반납 기한 지남', '#b3261e', report.overdue, true) +
    table('오늘까지 반납', '#996716', report.today, false) +
    table('내일까지 반납', '#1e6b58', report.tomorrow, false) +
    '<p style="margin-top:24px;font-size:12px;color:#777">반납을 확인했는데 기록이 남아 있다면, 시트 \'대여기록\' 탭에서 상태를 \'반납완료\'로 바꿔 주세요.<br>' +
    '이 메일에는 학번과 이름이 들어 있어요. 학생회 밖으로 전달하지 마세요.</p></div>';

  MailApp.sendEmail({ to: to, subject: subject, htmlBody: html, name: '학생회 물품 대여' });
}
