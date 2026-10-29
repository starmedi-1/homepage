/**
 * 김나영내과의원 홈페이지 — 구글 시트 연결 코드 (Google Apps Script)
 *
 * 하는 일
 *  - 홈페이지가 공지사항·진료시간·휴진일을 읽어 갈 수 있게 해 줍니다. (누구나)
 *  - 관리자가 홈페이지에서 쓴 글을 이 시트에 저장합니다. (비밀번호가 맞을 때만)
 *
 * 설치 순서는 「설치안내」 문서를 따라 주세요.
 *  1) 이 코드를 시트의 [확장 프로그램 → Apps Script]에 붙여넣기
 *  2) 위쪽 함수 선택에서 setup 을 골라 한 번 실행 (시트 탭이 자동으로 만들어집니다)
 *  3) [프로젝트 설정 → 스크립트 속성]에 ADMIN_PASSWORD 추가
 *  4) [배포 → 새 배포 → 웹 앱], 액세스 권한 "모든 사용자"로 배포
 */

const TZ = 'Asia/Seoul';
const WD = ['일', '월', '화', '수', '목', '금', '토'];

const SHEETS = {
  notices: { name: '공지사항', headers: ['id', '날짜', '고정(Y)', '제목', '내용', '사진(자동)'] },
  hours:   { name: '진료시간', headers: ['요일번호', '요일', '진료(Y/N)', '시작', '종료', '점심시작', '점심끝'] },
  closed:  { name: '휴진일',   headers: ['날짜', '사유'] },
};

// 처음 설치할 때 들어가는 진료시간 (2026년 10월 기준)
const DEFAULT_HOURS = [
  [1, 'Y', '09:00', '18:00', '13:00', '14:00'],
  [2, 'Y', '09:00', '18:00', '13:00', '14:00'],
  [3, 'Y', '09:00', '18:00', '13:00', '14:00'],
  [4, 'Y', '09:00', '13:00', '', ''],
  [5, 'Y', '09:00', '18:00', '13:00', '14:00'],
  [6, 'Y', '09:00', '13:00', '', ''],
  [0, 'N', '', '', '', ''],
];

const IMAGE_FOLDER_NAME = '김나영내과 홈페이지 사진';
const MAX_IMAGES = 5;            // 공지 하나에 올릴 수 있는 사진 수
const MAX_IMAGE_BYTES = 4 * 1024 * 1024;

const TOKEN_HOURS = 8;          // 로그인 유지 시간
const MAX_FAILS = 10;           // 비밀번호를 이만큼 틀리면
const LOCK_MINUTES = 10;        // 이 시간 동안 로그인 잠금

/* =========================================================
 * 설치용: 한 번만 실행
 * ========================================================= */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(key => {
    const def = SHEETS[key];
    const sh = ss.getSheetByName(def.name) || ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold').setBackground('#e2ebf7');
    sh.setFrozenRows(1);
    // 날짜·시간이 자동 변환되지 않도록 글자 형식으로 고정
    sh.getRange(1, 1, sh.getMaxRows(), def.headers.length).setNumberFormat('@');
  });
  const hs = ss.getSheetByName(SHEETS.hours.name);
  if (hs.getLastRow() < 2) {
    const rows = DEFAULT_HOURS.map(r => [String(r[0]), WD[r[0]], r[1], r[2], r[3], r[4], r[5]]);
    hs.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  // 비어 있는 기본 탭(시트1) 정리
  ['시트1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  imageFolder();   // 사진 폴더를 미리 만들고 드라이브 권한도 함께 승인받습니다
  const hasPw = !!PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  Logger.log(hasPw ? '설치 완료. 이제 [배포 → 새 배포]를 진행하세요.'
                   : '설치 완료. [프로젝트 설정 → 스크립트 속성]에 ADMIN_PASSWORD를 추가하세요.');
}

/* =========================================================
 * 홈페이지가 부르는 입구
 * ========================================================= */
function doGet(e) {
  return respond(() => ({ data: getData() }));
}

function doPost(e) {
  return respond(() => {
    let body;
    try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
    catch (_) { throw userError('요청 형식이 올바르지 않습니다.'); }

    if (body.action === 'login') return { token: login(String(body.password || '')) };
    if (body.action === 'data') return { data: getData() };

    requireToken(body.token);
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      switch (body.action) {
        case 'logout':       CacheService.getScriptCache().remove('tok_' + body.token); return {};
        case 'saveNotice':   saveNotice(body.item || {}); break;
        case 'deleteNotice': deleteNotice(body.id); break;
        case 'saveHours':    saveHours(body.hours); break;
        case 'addClosed':    addClosed(body.day, body.reason); break;
        case 'deleteClosed': deleteClosed(body.day); break;
        default: throw userError('알 수 없는 요청입니다.');
      }
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return { data: getData() };
  });
}

function respond(fn) {
  let out;
  try {
    out = Object.assign({ ok: true }, fn());
  } catch (err) {
    out = { ok: false, error: err.userMessage || '저장하지 못했습니다. 잠시 후 다시 시도해 주세요.', auth: !!err.auth };
    if (!err.userMessage) console.error(err);
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function userError(msg, auth) {
  const e = new Error(msg);
  e.userMessage = msg;
  e.auth = !!auth;
  return e;
}

/* =========================================================
 * 관리자 로그인
 * ========================================================= */
function login(password) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) throw userError(`비밀번호를 여러 번 틀려서 ${LOCK_MINUTES}분 동안 로그인할 수 없습니다.`);
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real) throw userError('관리자 비밀번호가 아직 설정되지 않았습니다.');
  if (password !== real) {
    cache.put('fails', String(fails + 1), LOCK_MINUTES * 60);
    throw userError('비밀번호가 맞지 않습니다.');
  }
  cache.remove('fails');
  const token = Utilities.getUuid();
  cache.put('tok_' + token, '1', TOKEN_HOURS * 3600);
  return token;
}

function requireToken(token) {
  if (!token || !CacheService.getScriptCache().get('tok_' + token)) {
    throw userError('로그인 시간이 지났습니다. 다시 로그인해 주세요.', true);
  }
}

/* =========================================================
 * 읽기
 * ========================================================= */
function sheetOf(key) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS[key].name);
  if (!sh) throw userError('시트 탭이 없습니다. setup을 먼저 실행해 주세요.');
  return sh;
}

function readRows(key) {
  const sh = sheetOf(key);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, SHEETS[key].headers.length).getDisplayValues();
}

// 시트에서 직접 줄을 추가해 id가 비어 있으면 채워 넣습니다.
function fillMissingIds(key) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return false;
  try {
    const sh = sheetOf(key);
    const rows = readRows(key);
    const textCols = [3, 4];
    let changed = false;
    const ids = rows.map(r => {
      if (!r[0] && textCols.some(c => r[c])) { changed = true; return [Utilities.getUuid().slice(0, 8)]; }
      return [r[0]];
    });
    if (changed) { sh.getRange(2, 1, ids.length, 1).setValues(ids); SpreadsheetApp.flush(); }
    return changed;
  } finally {
    lock.releaseLock();
  }
}

function getData(retried) {
  if (!retried) {
    const needN = readRows('notices').some(r => !r[0] && (r[3] || r[4]));
    if (needN && fillMissingIds('notices')) return getData(true);
  }

  const notices = readRows('notices')
    .filter(r => r[0] && (r[3] || r[4]))
    .map(r => ({ id: r[0], date: normDate(r[1]), pinned: isYes(r[2]), title: r[3], body: r[4],
                 images: splitIds(r[5]).map(id => ({ id: id, url: imageUrl(id) })) }));

  const hours = readRows('hours')
    .filter(r => r[0] !== '' && !isNaN(Number(r[0])))
    .map(r => ({ weekday: Number(r[0]), is_open: isYes(r[2]), open: normTime(r[3]), close: normTime(r[4]), bs: normTime(r[5]), be: normTime(r[6]) }));

  const closed = readRows('closed')
    .filter(r => r[0])
    .map(r => ({ day: normDate(r[0]), reason: r[1] }))
    .filter(c => /^\d{4}-\d{2}-\d{2}$/.test(c.day));

  return { notices, hours, closed };
}

/* =========================================================
 * 쓰기
 * ========================================================= */
function saveNotice(item) {
  const title = clean(item.title), body = clean(item.body, true);
  const sh = sheetOf('notices');
  const row = item.id ? findRow('notices', item.id) : 0;
  if (item.id && !row) throw userError('수정할 공지를 찾지 못했습니다. 새로고침 후 다시 시도해 주세요.');
  const oldIds = row ? splitIds(sh.getRange(row, 6).getDisplayValue()) : [];

  // images 가 없으면(고정만 바꿀 때 등) 기존 사진을 그대로 둡니다.
  const imgs = Array.isArray(item.images) ? item.images : null;
  if (imgs && imgs.length > MAX_IMAGES) throw userError(`사진은 ${MAX_IMAGES}장까지 올릴 수 있습니다.`);
  const willHaveImages = imgs ? imgs.length > 0 : oldIds.length > 0;

  if (!title) throw userError('제목을 입력해 주세요.');
  if (!body && !willHaveImages) throw userError('내용을 쓰거나 사진을 올려 주세요.');
  if (title.length > 60) throw userError('제목은 60자 이내로 써 주세요.');
  if (body.length > 3000) throw userError('내용은 3000자 이내로 써 주세요.');
  const pinned = item.pinned ? 'Y' : '';

  let ids = oldIds;
  if (imgs) {
    ids = imgs.map(im => {
      if (im && im.id && oldIds.indexOf(String(im.id)) >= 0) return String(im.id);   // 그대로 둔 사진
      if (im && im.data) return uploadImage(im.data, im.type, title);              // 새 사진
      throw userError('사진 정보가 올바르지 않습니다. 다시 첨부해 주세요.');
    });
    oldIds.filter(id => ids.indexOf(id) < 0).forEach(trashImage);                  // 뺀 사진은 휴지통으로
  }

  if (row) {
    sh.getRange(row, 3, 1, 4).setValues([[pinned, safe(title), safe(body), ids.join(',')]]);
  } else {
    sh.appendRow([Utilities.getUuid().slice(0, 8), today(), pinned, safe(title), safe(body), ids.join(',')]);
  }
}

function deleteNotice(id) {
  const row = findRow('notices', id);
  if (!row) throw userError('이미 삭제되었거나 찾을 수 없습니다. 새로고침해 주세요.');
  const sh = sheetOf('notices');
  splitIds(sh.getRange(row, 6).getDisplayValue()).forEach(trashImage);
  sh.deleteRow(row);
}

/* =========================================================
 * 사진 (구글 드라이브 폴더에 저장)
 * ========================================================= */
function imageFolder() {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('IMAGE_FOLDER_ID');
  if (saved) {
    try { const f = DriveApp.getFolderById(saved); if (!f.isTrashed()) return f; } catch (_) {}
  }
  const folder = DriveApp.createFolder(IMAGE_FOLDER_NAME);
  props.setProperty('IMAGE_FOLDER_ID', folder.getId());
  return folder;
}

function uploadImage(base64, type, title) {
  const mime = /^image\/(jpeg|png|webp|gif)$/.test(String(type)) ? String(type) : 'image/jpeg';
  let bytes;
  try { bytes = Utilities.base64Decode(String(base64).replace(/^data:[^,]*,/, '')); }
  catch (_) { throw userError('사진 파일을 읽지 못했습니다. 다른 사진으로 시도해 주세요.'); }
  if (bytes.length > MAX_IMAGE_BYTES) throw userError('사진 용량이 너무 큽니다.');
  const ext = mime.split('/')[1].replace('jpeg', 'jpg');
  const name = `${today()}_${String(title).replace(/[\\/:*?"<>|]/g, '').slice(0, 30)}_${Utilities.getUuid().slice(0, 4)}.${ext}`;
  const file = imageFolder().createFile(Utilities.newBlob(bytes, mime, name));
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    file.setTrashed(true);
    throw userError('사진 공유 설정이 막혀 있습니다. 구글 계정의 공유 설정을 확인해 주세요.');
  }
  return file.getId();
}

function trashImage(id) {
  try { DriveApp.getFileById(id).setTrashed(true); } catch (_) {}
}

function imageUrl(id) { return 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w1600'; }

function splitIds(s) { return String(s || '').split(',').map(x => x.trim()).filter(Boolean); }

function saveHours(hours) {
  if (!Array.isArray(hours) || hours.length !== 7) throw userError('요일별 진료시간 7개가 모두 필요합니다.');
  const byDay = {};
  hours.forEach(h => {
    const w = Number(h.weekday);
    if (!(w >= 0 && w <= 6)) throw userError('요일 정보가 올바르지 않습니다.');
    const t = ['open', 'close', 'bs', 'be'].map(k => normTime(h[k] || ''));
    if (t.some(v => v && !/^\d{2}:\d{2}$/.test(v))) throw userError(`${WD[w]}요일 시간 형식을 확인해 주세요.`);
    if (h.is_open && (!t[0] || !t[1] || t[0] >= t[1])) throw userError(`${WD[w]}요일 진료 시작·종료 시간을 확인해 주세요.`);
    if ((t[2] && !t[3]) || (!t[2] && t[3]) || (t[2] && t[3] && t[2] >= t[3])) throw userError(`${WD[w]}요일 점심시간을 확인해 주세요.`);
    byDay[w] = [String(w), WD[w], h.is_open ? 'Y' : 'N', t[0], t[1], t[2], t[3]];
  });
  if (Object.keys(byDay).length !== 7) throw userError('요일별 진료시간 7개가 모두 필요합니다.');
  const rows = [1, 2, 3, 4, 5, 6, 0].map(w => byDay[w]);
  const sh = sheetOf('hours');
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 7).clearContent();
  sh.getRange(2, 1, 7, 7).setNumberFormat('@').setValues(rows);
}

function addClosed(day, reason) {
  day = normDate(String(day || ''));
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day)) throw userError('날짜를 선택해 주세요.');
  if (day < today()) throw userError('오늘 이후 날짜를 선택해 주세요.');
  const r = clean(reason);
  if (r.length > 40) throw userError('사유는 40자 이내로 써 주세요.');
  const sh = sheetOf('closed');
  const row = findRow('closed', day);
  if (row) sh.getRange(row, 2).setValue(safe(r));
  else sh.appendRow([day, safe(r)]);
  sortClosed();
}

function deleteClosed(day) {
  const row = findRow('closed', normDate(String(day || '')));
  if (row) sheetOf('closed').deleteRow(row);
}

function deleteById(key, id) {
  const row = findRow(key, id);
  if (!row) throw userError('이미 삭제되었거나 찾을 수 없습니다. 새로고침해 주세요.');
  sheetOf(key).deleteRow(row);
}

/* =========================================================
 * 도우미
 * ========================================================= */
function findRow(key, id) {
  if (!id) return 0;
  const rows = readRows(key);
  for (let i = 0; i < rows.length; i++) {
    const v = key === 'closed' ? normDate(rows[i][0]) : rows[i][0];
    if (v === String(id)) return i + 2;
  }
  return 0;
}

function sortClosed() {
  const sh = sheetOf('closed');
  const n = sh.getLastRow() - 1;
  if (n > 1) sh.getRange(2, 1, n, 2).sort({ column: 1, ascending: true });
}

function today() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function isYes(v) { return /^(y|yes|o|예|네|true|1)$/i.test(String(v || '').trim()); }

function clean(v, multiline) {
  let s = String(v == null ? '' : v).replace(/\r\n?/g, '\n');
  s = multiline ? s.replace(/[ \t]+\n/g, '\n').trim() : s.replace(/\s+/g, ' ').trim();
  return s;
}

// =, +, -, @ 로 시작하는 글이 수식으로 바뀌지 않도록 막습니다.
function safe(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }

function normDate(s) {
  s = String(s || '').trim();
  const m = s.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? `${m[1]}-${('0' + m[2]).slice(-2)}-${('0' + m[3]).slice(-2)}` : s;
}

function normTime(s) {
  s = String(s || '').trim();
  if (!s) return '';
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (!m) return s;
  let h = Number(m[1]);
  if (/오후|pm/i.test(s) && h < 12) h += 12;
  if (/오전|am/i.test(s) && h === 12) h = 0;
  return `${('0' + h).slice(-2)}:${m[2]}`;
}
