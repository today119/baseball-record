/**
 * ⚾ 야구 수행평가 — 구글 시트 브릿지
 * ══════════════════════════════════════════════════════════
 *  설치 (5분)
 *   1. 구글 스프레드시트를 하나 새로 만든다
 *   2. 상단 메뉴  확장 프로그램 → Apps Script
 *   3. 기본으로 열린 코드를 지우고 이 파일 내용을 통째로 붙여넣는다
 *   4. 💾 저장 → 함수 목록에서 authorize 선택 → ▶ 실행
 *        (권한 요청이 뜨면 계정 선택 → 고급 → 안전하지 않음 이동 → 허용)
 *        실행 로그에 시트 이름이 뜨면 성공
 *   5. 오른쪽 위 배포 → 새 배포 → ⚙ → 웹 앱
 *        설명       : 아무거나
 *        실행 사용자 : 나
 *        액세스 권한 : 링크가 있는 모든 사용자
 *      → 배포 → 나오는 /exec 로 끝나는 주소를 복사
 *   6. 앱 [설정 → 구글 시트 연동]에 붙여넣고 「연결 확인」
 *
 *  ※ 코드를 고치면 반드시 배포 → 배포 관리 → ✏️ → 버전 「새 버전」 으로 다시 배포해야
 *    반영됩니다. (이걸 안 해서 안 되는 경우가 제일 많습니다)
 *
 *  만들어지는 시트 (자동 생성)
 *   명단      학급 | 번호 | 이름 | 성별 | 조
 *   야구기록   학급 | 번호 | 이름 | 성별 | 경기 | 이닝 | 타수 | 안타 | 타율 |
 *             득점 | 타점 | 수비아웃 | 이닝당기여 | 탁구포핸드 | 탁구백핸드 |
 *             태도카드 | 보고서충족 | 갱신일시
 * ══════════════════════════════════════════════════════════
 */

/* 이 스크립트를 시트 안에서 만들었다면 아래는 그대로 두면 됩니다.
   따로 만든 경우에만 시트 ID를 넣으세요. */
var SHEET_ID = '';

var _SS = null;
function SS_() {
  if (_SS) return _SS;
  if (SHEET_ID) _SS = SpreadsheetApp.openById(SHEET_ID);
  else _SS = SpreadsheetApp.getActiveSpreadsheet();
  if (!_SS) throw new Error('시트를 찾을 수 없습니다. SHEET_ID 를 넣어주세요.');
  return _SS;
}

/** ▶ 이 함수를 한 번 실행해서 권한을 승인하세요 */
function authorize() {
  var n = SS_().getName();
  Logger.log('✅ 시트 연결 성공: ' + n);
  return n;
}

var ROSTER_H = ['학급', '번호', '이름', '성별', '조'];
var REC_H = ['학급', '번호', '이름', '성별', '경기', '이닝', '타수', '안타', '타율',
             '득점', '타점', '수비아웃', '이닝당기여', '탁구포핸드', '탁구백핸드',
             '태도카드', '보고서충족', '갱신일시'];

function sh_(name, header) {
  var s = SS_().getSheetByName(name);
  if (!s) {
    s = SS_().insertSheet(name);
    s.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#1f3d4d').setFontColor('#ffffff');
    s.setFrozenRows(1);
    s.setColumnWidth(1, 70);
    s.setColumnWidth(3, 90);
  }
  return s;
}

/** 학급 명단 읽기 */
function getRoster_(cn) {
  var s = sh_('명단', ROSTER_H);
  var last = s.getLastRow();
  if (last < 2) return [];
  var vals = s.getRange(2, 1, last - 1, ROSTER_H.length).getDisplayValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (String(r[0]).trim() !== String(cn).trim()) continue;
    if (!String(r[2]).trim()) continue;
    out.push({
      num: String(r[1]).trim(),
      name: String(r[2]).trim(),
      gender: String(r[3]).trim(),
      group: String(r[4]).trim() || 'A'
    });
  }
  return out;
}

/** 같은 학급 행만 지우고 새로 넣는다 (다른 학급 기록은 건드리지 않음) */
function replaceRows_(sheetName, header, cn, rows) {
  var s = sh_(sheetName, header);
  var last = s.getLastRow();
  if (last > 1) {
    var col = s.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (var i = col.length - 1; i >= 0; i--) {
      if (String(col[i][0]).trim() === String(cn).trim()) s.deleteRow(i + 2);
    }
  }
  if (rows && rows.length) {
    s.getRange(s.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return rows ? rows.length : 0;
}

/* ── 앱이 부르는 입구 ─────────────────────────────────── */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    if (p.action === 'ping') {
      out = { ok: true, msg: '연결 정상 · ' + SS_().getName() };
    } else if (p.action === 'getRoster') {
      out = { ok: true, data: getRoster_(p.cn) };
    } else if (p.action === 'initSheets') {
      // 앱의 「시트 양식 만들기」 — 두 장을 머리글까지 갖춰 만든다
      var a = sh_('명단', ROSTER_H), b = sh_('야구기록', REC_H);
      // 예시 한 줄 (명단이 완전히 비어 있을 때만)
      if (a.getLastRow() < 2) {
        a.getRange(2, 1, 1, ROSTER_H.length)
         .setValues([['1반', '1', '홍길동', '남', 'A']])
         .setFontColor('#999999').setFontStyle('italic');
        a.getRange(2, 1, 1, 1).setNote('예시 줄입니다. 지우고 실제 명단을 넣으세요.');
      }
      out = { ok: true, msg: '명단 · 야구기록 시트 준비 완료',
              sheets: [a.getName(), b.getName()] };

    } else if (p.action === 'getCount') {
      // 내보내기가 실제로 들어갔는지 확인용
      var s = sh_('야구기록', REC_H), n = 0, last = s.getLastRow();
      if (last > 1) {
        var col = s.getRange(2, 1, last - 1, 1).getDisplayValues();
        for (var i = 0; i < col.length; i++) {
          if (String(col[i][0]).trim() === String(p.cn).trim()) n++;
        }
      }
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
      var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
      var rows = (payload.rows || []).map(function (r) {
        return [payload.cn].concat(r).concat([stamp]);
      });
      out = { ok: true, saved: replaceRows_('야구기록', REC_H, payload.cn, rows) };

    } else if (action === 'saveRoster') {
      var rr = (payload.students || []).map(function (s) {
        return [payload.cn, s.num, s.name, s.gender, s.group || 'A'];
      });
      out = { ok: true, saved: replaceRows_('명단', ROSTER_H, payload.cn, rr) };

    } else {
      out = { ok: false, error: '알 수 없는 요청: ' + action };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
