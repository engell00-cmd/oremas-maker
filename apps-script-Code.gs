/**
 * 오레마스 포켓몬 생성기 — 공유 저장소 (구글 시트)
 * 1) 새 구글 시트 → 확장 프로그램 → Apps Script → 이 코드를 붙여넣고 저장
 * 2) 배포 → 새 배포 → 유형: 웹 앱 → 실행: 나 / 액세스: 모든 사용자 → 배포 → 나온 주소(…/exec)를 복사
 * (시트에서 Apps Script 메뉴를 못 찾으면: script.google.com 에서 새 프로젝트를 만들고 아래 SHEET_ID 에 시트 주소의 /d/ 와 /edit 사이 글자를 넣어도 됨)
 * 시트 「모은 포켓몬」「채점」이 처음 저장할 때 생김. 「숨김」 칸에 TRUE(체크)를 하면 그 줄은 페이지에서 안 보이고 학습에서도 빠짐.
 */
const SHEET_ID = '';   // 비워 두면 이 코드를 연 시트를 씀
const SHEETS = {entries: '모은 포켓몬', ratings: '채점'};
const HEAD = ['id', '시각', 'by', '닉네임', '숨김', '이름', '요약', 'json'];
const MAX_JSON = 45000;          // 셀 하나 한도(5만 자) 아래
const RATE_N = 30, RATE_SEC = 600; // 한 사람 10분에 30번까지 저장

function sheet_(col) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEETS[col]);
  if (!s) { s = ss.insertSheet(SHEETS[col]); s.appendRow(HEAD); s.setFrozenRows(1); s.getRange('E2:E').insertCheckboxes(); }
  return s;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hidden_(v) { return v === true || String(v).toUpperCase() === 'TRUE' || v === '숨김'; }

function doGet(e) {
  const res = {};
  for (const col in SHEETS) {
    const v = sheet_(col).getDataRange().getValues(); const rows = [];
    for (let i = 1; i < v.length; i++) {
      const r = v[i]; if (!r[0] || hidden_(r[4])) continue;
      try { const d = JSON.parse(r[7]); d.id = String(r[0]); d.by = String(r[2]); d.nick = String(r[3] || ''); rows.push(d); } catch (err) {}
    }
    res[col] = rows;
  }
  return out_(res);
}

function summary_(col, d) {
  if (col === 'entries') return [d.role || '', (d.types || []).join('/'), '합 ' + (d.total || ''), d.src || ''].filter(String).join(' · ');
  return ['점수 ' + d.score, (d.tags || []).join(', '), d.role || '', (d.essay || '').slice(0, 120)].filter(String).join(' · ');
}

function doPost(e) {
  let q; try { q = JSON.parse(e.postData.contents); } catch (err) { return out_({ok: false, error: '형식 오류', code: 'invalid_argument'}); }
  const col = q.col, by = String(q.by || '').slice(0, 40), id = String(q.id || '').replace(/[^\w\-]/g, '').slice(0, 60);
  if (!SHEETS[col] || !by || !id) return out_({ok: false, error: '빠진 값이 있어요', code: 'invalid_argument'});
  const cache = CacheService.getScriptCache(); const rk = 'r_' + by; const n = Number(cache.get(rk) || 0);
  if (n >= RATE_N) return out_({ok: false, error: '잠시 뒤에 다시 저장해 주세요', code: 'resource_exhausted'});
  cache.put(rk, String(n + 1), RATE_SEC);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const s = sheet_(col); const v = s.getDataRange().getValues();
    let row = -1; for (let i = 1; i < v.length; i++) if (String(v[i][0]) === id) { row = i + 1; break; }
    if (row > 0 && String(v[row - 1][2]) !== by) return out_({ok: false, error: '다른 사람의 기록이에요', code: 'invalid_argument'});
    if (q.action === 'del') { if (row > 0) s.deleteRow(row); return out_({ok: true}); }
    if (q.action !== 'put') return out_({ok: false, error: '모르는 동작', code: 'invalid_argument'});
    const d = q.doc || {}; delete d.id; d.by = by;
    const js = JSON.stringify(d); if (js.length > MAX_JSON) return out_({ok: false, error: '너무 길어요', code: 'invalid_argument'});
    const vals = [id, new Date(), by, String(q.nick || '').slice(0, 20), false, String(d.name || '').slice(0, 60), summary_(col, d), js];
    if (row > 0) { vals[4] = v[row - 1][4]; s.getRange(row, 1, 1, vals.length).setValues([vals]); }
    else { s.appendRow(vals); s.getRange(s.getLastRow(), 5).insertCheckboxes(); }
    return out_({ok: true, id: id});
  } finally { lock.releaseLock(); }
}
