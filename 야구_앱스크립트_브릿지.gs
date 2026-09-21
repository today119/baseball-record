/**
 * ⚾ 야구 수행평가 — 구글 시트 브릿지
 * ══════════════════════════════════════════════════════════
 *  설치 (5분)
 *   1. 구글 스프레드시트를 하나 만든다
 *   2. 확장 프로그램 → Apps Script
 *   3. 기본 코드를 지우고 이 파일을 통째로 붙여넣는다
 *   4. 💾 저장 → 함수 목록에서 authorize 선택 → ▶ 실행
 *        (권한 창: 계정 선택 → 고급 → 안전하지 않음 이동 → 허용)
 *   5. 배포 → 새 배포 → ⚙ → 웹 앱
 *        실행 사용자 : 나
 *        액세스 권한 : 링크가 있는 모든 사용자   ← 꼭!
 *      → 나오는 /exec 주소를 복사
 *   6. 앱 [설정 → 구글 시트 연동] 에 붙여넣고 「연결 확인」 → 「시트 양식 만들기」
 *
 *  ※ 코드를 고치면 반드시 배포 → 배포 관리 → ✏️ → 「새 버전」 으로 다시 배포해야 반영됩니다.
 *
 *  ══ 시트 구조 ══════════════════════════════════════════
 *  탭 하나 = 학급 하나 (1반, 2반, …). 명단과 기록이 한 표에 들어간다.
 *
 *   A    B    C    D   │ 명단 (사람이 넣어도 되고 앱이 올려도 된다)
 *   번호 이름 성별 조
 *
 *   E~H   탁구포핸드 탁구백핸드 스트로크합계 스트로크점수      ← 탁구 스트로크 (20점)
 *   I~M   탁구경기수 탁구총득점 탁구평균득점 탁구카드 탁구리그점수 ← 탁구 리그전 (20점)
 *   N~Z   경기 이닝 타수 안타 타율 득점 타점 수비참여 수비아웃
 *         이닝당기여 경기당기여 야구카드 야구점수              ← 야구 (40점)
 *   AA~AD 보고서충족 보고서점수 총점 갱신일시                  ← 보고서 (20점) + 합계
 *
 *  ※ 점수 환산 기준 (2026-2학기 평가계획)
 *    · 스트로크 : 합계 9↑20 / 7~8 16 / 5~6 12 / 3~4 8 / 2↓ 4        (기본 4)
 *    · 탁구리그 : 평균득점 9.0↑20 / 7.0~ 16 / 5.0~ 12 / 3.0~ 8 / 미만 4
 *                + 카드 1장 1점(최대 4), 합계 최대 20
 *    · 야구     : 타율 .500↑30 / .400~25 / .300~20 / .200~15 / 미만 10
 *                + 경기당기여 2.0↑10 / 1.5~8 / 1.0~6 / 0.5~4 / 미만 2
 *                + 카드 1장 1점(최대 5), 합계 최대 40             (기본 12)
 *    · 보고서   : 충족 6개 20 / 5개 16 / 4개 12 / 3개 8 / 2개↓ 4    (기본 2)
 *
 *  ※ 기여도는 두 가지를 모두 저장한다.
 *    · 이닝당기여 = (득점+타점+수비참여+수비아웃) ÷ 참여 이닝   ← 기록·순위·특기사항용
 *    · 경기당기여 = (득점+타점+수비참여+수비아웃) ÷ 참여 경기   ← 평가 점수 환산용(계획서 기준)
 *
 *  · 명단(A~D)은 시트에서 직접 입력해도 되고 앱에서 올려도 된다
 *  · 앱의 「내보내기」는 번호·이름으로 학생을 찾아 **조(D) + E열 이후**를 갱신한다
 *    → 번호·이름·성별은 시트 것을 그대로 둔다
 * ══════════════════════════════════════════════════════════
 */

/* 이 스크립트를 시트 안에서 만들었다면 그대로 두세요. */
var SHEET_ID = '';

/* 배포 버전 — 앱이 이 값을 보고 「옛 버전이 배포돼 있다」를 알아챈다.
   코드를 고쳤는데 이 값이 앱에 안 뜨면 재배포가 안 된 것이다. */
var VER = '2026-09-21-plan2';

var HDR = ['번호', '이름', '성별', '조',
           /* 탁구 스트로크 */ '탁구포핸드', '탁구백핸드', '스트로크합계', '스트로크점수',
           /* 탁구 리그전  */ '탁구경기수', '탁구총득점', '탁구평균득점', '탁구카드', '탁구리그점수',
           /* 야구        */ '경기', '이닝', '타수', '안타', '타율', '득점', '타점',
                             '수비참여', '수비아웃', '이닝당기여', '경기당기여', '야구카드', '야구점수',
           /* 보고서·합계 */ '보고서충족', '보고서점수', '총점', '갱신일시'];
var NAME_COLS = 4;                 // A~D 가 명단
var RESERVED = ['시트1', 'Sheet1'];

var _SS = null;
function SS_() {
  if (_SS) return _SS;
  _SS = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!_SS) throw new Error('시트를 찾을 수 없습니다. SHEET_ID 를 넣어주세요.');
  return _SS;
}

/** ▶ 이 함수를 한 번 실행해서 권한을 승인하세요 */
function authorize() {
  var n = SS_().getName();
  Logger.log('✅ 시트 연결 성공: ' + n);
  return n;
}

/* 머리글 색 구역 — [시작열(1부터), 길이, 색]
   명단 / 탁구 스트로크 / 탁구 리그전 / 야구 / 보고서·합계 */
var BANDS = [
  [1,  4, '#1f3d4d'],   // A~D  명단
  [5,  4, '#4a3b6b'],   // E~H  탁구 스트로크
  [9,  5, '#6b3b5c'],   // I~M  탁구 리그전
  [14, 13, '#2d5c3a'],  // N~Z  야구
  [27, 4, '#5c4a1f']    // AA~AD 보고서·총점·갱신일시
];

/** 머리글 줄을 HDR 에 맞춰 다시 쓴다 (1행만 건드린다) */
function setupHeader_(s) {
  s.getRange(1, 1, 1, HDR.length).setValues([HDR])
    .setFontWeight('bold').setFontColor('#FFFFFF').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrap(true);
  BANDS.forEach(function (b) {
    if (b[0] <= HDR.length) {
      s.getRange(1, b[0], 1, Math.min(b[1], HDR.length - b[0] + 1)).setBackground(b[2]);
    }
  });
  s.setFrozenRows(1);
  s.setFrozenColumns(2);
  s.setColumnWidth(1, 48); s.setColumnWidth(2, 84);
  s.setColumnWidth(3, 48); s.setColumnWidth(4, 44);
  for (var c = NAME_COLS + 1; c <= HDR.length; c++) s.setColumnWidth(c, 62);
  s.setRowHeight(1, 34);
  return s;
}

/** 학급 탭 가져오기 (없으면 머리글까지 만들어서) */
function tab_(cn) {
  var ss = SS_(), s = ss.getSheetByName(cn);
  if (!s) {
    s = ss.insertSheet(cn);
    setupHeader_(s);
  }
  return s;
}

/** 이미 있는 탭의 머리글을 새 양식으로 갈아끼운다.
    ⚠️ 열 구성이 바뀌었으므로 기존 기록(E열 이후)은 자리가 어긋난다.
       wipe=true 면 기록 칸을 비우고 명단(A~D)만 남긴다. */
function migrateTabs_(list, wipe) {
  var done = [];
  list.forEach(function (cn) {
    var s = SS_().getSheetByName(cn);
    if (!s) { s = tab_(cn); done.push(cn + '(새로 만듦)'); return; }
    setupHeader_(s);
    if (wipe) {
      var last = s.getLastRow();
      if (last > 1) {
        s.getRange(2, NAME_COLS + 1, last - 1, HDR.length - NAME_COLS).clearContent();
      }
    }
    done.push(cn);
  });
  return done;
}

function rowsOf_(s) {
  var last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, HDR.length).getDisplayValues();
}
function key_(num, name) {
  return String(num).trim() + '|' + String(name).trim();
}

/** 명단 읽기 — A~D 만 */
function getRoster_(cn) {
  var out = [];
  rowsOf_(tab_(cn)).forEach(function (r) {
    if (!String(r[1]).trim()) return;                // 이름 없으면 건너뜀
    out.push({ num: String(r[0]).trim(), name: String(r[1]).trim(),
               gender: String(r[2]).trim(), group: String(r[3]).trim() || 'A' });
  });
  return out;
}

/** 명단 쓰기 — A~D 만 (기록 열은 건드리지 않음) */
function saveRoster_(cn, students) {
  var s = tab_(cn), cur = rowsOf_(s), idx = {};
  cur.forEach(function (r, i) { idx[key_(r[0], r[1])] = i + 2; });
  var added = 0;
  students.forEach(function (st) {
    var k = key_(st.num, st.name), row = idx[k];
    if (!row) { row = s.getLastRow() + 1; added++; }
    s.getRange(row, 1, 1, NAME_COLS)
     .setValues([[st.num, st.name, st.gender, st.group || 'A']]);
  });
  return { total: students.length, added: added };
}

/** 기록 쓰기 — 번호·이름으로 찾아 E열 이후만 갱신. 없으면 줄을 새로 만든다. */
function saveRecords_(cn, rows) {
  var s = tab_(cn), cur = rowsOf_(s), idx = {};
  cur.forEach(function (r, i) { idx[key_(r[0], r[1])] = i + 2; });
  var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
  var updated = 0, added = 0;

  rows.forEach(function (r) {
    // r = [번호,이름,성별,조, 경기,이닝,타수,안타,타율,득점,타점,수비아웃,이닝당기여,
    //      탁구포핸드,탁구백핸드,태도카드,보고서충족]
    var k = key_(r[0], r[1]), row = idx[k];
    if (!row) {
      row = s.getLastRow() + 1;
      s.getRange(row, 1, 1, NAME_COLS).setValues([r.slice(0, NAME_COLS)]);
      added++;
    } else {
      // 「조」(D열)는 앱이 원본이므로 함께 갱신한다. 번호·이름·성별은 건드리지 않는다.
      s.getRange(row, 4).setValue(r[3]);
      updated++;
    }
    var rec = r.slice(NAME_COLS).concat([stamp]);
    s.getRange(row, NAME_COLS + 1, 1, rec.length).setValues([rec]);
  });
  return { updated: updated, added: added };
}

/* ── 앱이 부르는 입구 ─────────────────────────────────── */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (p.action === 'ping') {
      out = { ok: true, ver: VER, msg: '연결 정상 · ' + SS_().getName() + ' (v' + VER + ')' };

    } else if (p.action === 'getRoster') {
      out = { ok: true, data: getRoster_(p.cn) };

    } else if (p.action === 'migrate') {
      // 이미 있는 탭의 머리글을 새 양식으로 교체한다.  classes=1반,2반,…  wipe=1 이면 기록 칸 비움
      var ml = String(p.classes || '').split(',')
                 .map(function (x) { return x.trim(); }).filter(Boolean);
      if (!ml.length) throw new Error('학급 목록이 비어 있습니다');
      var moved = migrateTabs_(ml, String(p.wipe || '') === '1');
      out = { ok: true, ver: VER, msg: moved.length + '개 탭 양식 갱신 완료', sheets: moved };

    } else if (p.action === 'initSheets') {
      // 학급 탭을 한꺼번에 만든다.  classes=1반,2반,…
      var list = String(p.classes || '').split(',')
                   .map(function (x) { return x.trim(); }).filter(Boolean);
      if (!list.length) throw new Error('학급 목록이 비어 있습니다');
      // 이미 있는 탭도 머리글을 새 양식으로 맞춘다 (기록은 건드리지 않는다)
      list.forEach(function (cn) { var t = tab_(cn); setupHeader_(t); });
      // 예시 줄 (첫 학급 탭이 완전히 비어 있을 때만)
      var f = tab_(list[0]);
      if (f.getLastRow() < 2) {
        f.getRange(2, 1, 1, NAME_COLS).setValues([['1', '홍길동', '남', 'A']])
         .setFontColor('#999999').setFontStyle('italic');
        f.getRange(2, 1).setNote('예시 줄입니다. 지우고 실제 명단을 넣으세요.');
      }
      // 기본 「시트1」이 비어 있으면 정리
      RESERVED.forEach(function (n) {
        var x = SS_().getSheetByName(n);
        if (x && x.getLastRow() <= 1 && SS_().getSheets().length > 1) SS_().deleteSheet(x);
      });
      out = { ok: true, ver: VER, msg: list.length + '개 학급 탭 준비 완료', sheets: list };

    } else if (p.action === 'getCount') {
      var n = 0;
      rowsOf_(tab_(p.cn)).forEach(function (r) { if (String(r[1]).trim()) n++; });
      out = { ok: true, count: n };

    } else {
      out = { ok: false, error: '알 수 없는 요청: ' + p.action };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  var body = JSON.stringify(out);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var out;
  try {
    var action = e.parameter.action;
    var payload = JSON.parse(e.parameter.payload || '{}');

    if (action === 'saveBaseball') {
      out = { ok: true, result: saveRecords_(payload.cn, payload.rows || []) };

    } else if (action === 'saveRoster') {
      out = { ok: true, result: saveRoster_(payload.cn, payload.students || []) };

    } else {
      out = { ok: false, error: '알 수 없는 요청: ' + action };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
