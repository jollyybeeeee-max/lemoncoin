/**
 * 레몬코인 출석 도장 · 오늘의 퀴즈 서버
 * '레몬코인 홈페이지 데이터' 구글 시트 > 확장 프로그램 > Apps Script 에 그대로 붙여 넣고
 * 웹 앱(액세스: 모든 사용자)으로 배포하세요.
 *
 * 시트 탭 (없으면 자동으로 만들어져요)
 *  - stamps  : 출석 도장 기록 (날짜, 이름, 골드, 시각)
 *  - quiz    : 퀴즈 문제 은행 (문제, 보기1~4, 정답 번호 1~4, 해설)
 *  - answers : 퀴즈 응답 기록
 */
const TZ = 'Asia/Seoul';

function ymd_(d) { return Number(Utilities.formatDate(d || new Date(), TZ, 'yyyyMMdd')); }
function dayNum_() { return Math.floor((Date.now() + 9 * 3600000) / 86400000); }
function monday_() {
  const now = new Date();
  const dow = Number(Utilities.formatDate(now, TZ, 'u')); // 1=월 … 7=일
  return ymd_(new Date(now.getTime() - (dow - 1) * 86400000));
}
function tab_(name, head) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(head); s.setFrozenRows(1); }
  return s;
}
function rows_(s, cols) { const n = s.getLastRow(); return n < 2 ? [] : s.getRange(2, 1, n - 1, cols).getValues(); }
function members_() {
  try { return JSON.parse(SpreadsheetApp.getActive().getSheetByName('public').getRange('A1').getValue()).members.map(m => m.n); }
  catch (e) { return []; }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  const p = (e && e.parameter) || {};
  const a = p.a || 'state';
  const name = String(p.name || '').trim();
  const valid = name && members_().indexOf(name) >= 0;
  const lock = LockService.getScriptLock();
  try {
    if ((a === 'stamp' || a === 'answer') && !valid) return out_({ ok: false, err: 'name' });
    if (a === 'stamp') { lock.waitLock(10000); stamp_(name, false); }
    let res = null;
    if (a === 'answer') { lock.waitLock(10000); res = answer_(name, Number(p.c)); }
    return out_(Object.assign({ ok: true }, state_(valid ? name : null), res ? { result: res } : {}));
  } catch (err) {
    return out_({ ok: false, err: String(err) });
  } finally { try { lock.releaseLock(); } catch (e2) {} }
}

/* ---------- 출석 도장 ---------- */
function stampTab_() { return tab_('stamps', ['날짜', '이름', '골드', '시각']); }
function stamp_(name, gold) {
  const s = stampTab_(), t = ymd_(), rows = rows_(s, 3);
  for (let i = 0; i < rows.length; i++) {
    if (Number(rows[i][0]) === t && rows[i][1] === name) {
      if (gold && !rows[i][2]) s.getRange(i + 2, 3).setValue(1);
      return;
    }
  }
  s.appendRow([t, name, gold ? 1 : 0, new Date()]);
}
function state_(name) {
  const mon = monday_();
  const stamps = name ? rows_(stampTab_(), 3)
    .filter(r => r[1] === name && Number(r[0]) >= mon)
    .map(r => ({ d: Number(r[0]), g: !!r[2] })) : [];
  return { today: ymd_(), monday: mon, stamps: stamps, quiz: quizState_(name) };
}

/* ---------- 오늘의 퀴즈 ---------- */
function bank_() {
  const s = tab_('quiz', ['문제', '보기1', '보기2', '보기3', '보기4', '정답(1~4)', '해설']);
  return rows_(s, 7).map((r, i) => ({
    row: i + 2, q: String(r[0]).trim(),
    opts: [r[1], r[2], r[3], r[4]].map(String).map(x => x.trim()).filter(Boolean),
    ans: Number(r[5]), ex: String(r[6] || '').trim()
  })).filter(x => x.q && x.opts.length >= 2 && x.ans >= 1 && x.ans <= x.opts.length);
}
function todayQuiz_() {
  const b = bank_();
  if (!b.length) return null;
  // 고정된 순서로 섞은 뒤 하루에 하나씩 → 문제를 다 돌기 전까지 겹치지 않아요
  const order = b.map((x, i) => i);
  let seed = 20261120;
  for (let i = order.length - 1; i > 0; i--) {
    seed = (seed * 1103515245 + 12345) % 2147483648;
    const j = seed % (i + 1); const t = order[i]; order[i] = order[j]; order[j] = t;
  }
  return b[order[dayNum_() % b.length]];
}
function answerTab_() { return tab_('answers', ['날짜', '이름', '문제 행', '고른 답', '정답 여부', '시각']); }
function quizState_(name) {
  const q = todayQuiz_();
  if (!q) return null;
  const t = ymd_();
  const today = rows_(answerTab_(), 5).filter(r => Number(r[0]) === t && Number(r[2]) === q.row);
  const mine = name ? today.filter(r => r[1] === name)[0] : null;
  const st = { n: today.length, ok: today.filter(r => r[4]).length };
  const base = { id: t + '-' + q.row, q: q.q, opts: q.opts, stats: st };
  if (mine) Object.assign(base, { done: { c: Number(mine[3]), ok: !!mine[4], ans: q.ans, ex: q.ex } });
  return base;
}
function answer_(name, c) {
  const q = todayQuiz_();
  if (!q) return { err: 'noquiz' };
  const t = ymd_(), s = answerTab_();
  const prev = rows_(s, 5).filter(r => Number(r[0]) === t && Number(r[2]) === q.row && r[1] === name)[0];
  if (prev) return { already: true };
  if (!(c >= 1 && c <= q.opts.length)) return { err: 'choice' };
  const ok = c === q.ans;
  s.appendRow([t, name, q.row, c, ok ? 1 : 0, new Date()]);
  stamp_(name, ok); // 퀴즈를 풀면 출석 도장, 맞히면 골드 도장
  return { ok: ok };
}
