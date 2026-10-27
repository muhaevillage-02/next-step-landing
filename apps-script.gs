/**
 * 청년 NEXT STEP 프로젝트 신청서 → 구글 시트 자동 연동용 Apps Script
 *
 * [연결 방법]
 * 1. 구글 시트를 새로 만들고 [확장 프로그램] > [Apps Script] 를 연다.
 * 2. Code.gs 내용을 모두 지우고 이 파일 전체를 붙여넣은 뒤 저장한다.
 * 3. [배포] > [새 배포] > 유형 선택(톱니바퀴): 웹 앱
 *    실행 계정: 나 / 액세스 권한이 있는 사용자: 모든 사용자 → [배포] → 권한 승인
 * 4. 웹 앱 URL(https://script.google.com/macros/s/.../exec)을 index.html 의 CONFIG.SCRIPT_URL 에 넣는다.
 *
 * - 헤더 행은 첫 응답이 도착할 때 자동으로 만들어진다.
 * - 같은 휴대폰 번호로 이미 접수된 건이 있으면 저장하지 않고 {ok:false, error:'already'} 를 돌려준다.
 * - 필수 항목이 비어 있으면 {ok:false, error:'invalid'} 를 돌려준다.
 * - 성공 시 {ok:true} 를 돌려준다. (신청서 화면이 이 응답을 기다린다)
 *
 * 코드를 수정한 뒤에는 [배포] > [배포 관리] > 연필 > 버전: 새 버전 > [배포] 를 눌러야
 * 같은 URL로 반영된다.
 */
var HEADERS = [
  "접수시각", "지원 전형", "전형 관련 상황", "요즘 나는", "요즘 가장 큰 고민", "고민 글자수", "얻고 싶은 것", "더 하고 싶은 말",
  "이름", "출생연도", "휴대폰", "시·도", "시·군·구", "알게 된 경로", "개인정보 동의", "연구 활용 동의",
  "작성 시간(초)", "신청 페이지",
  "utm_source", "utm_medium", "utm_campaign", "utm_content", "utm_term"
];

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var p = (e && e.parameter) || {};
    function v(k) { return String(p[k] || "").trim(); }
    // 링크나 입력값이 수식(=, +, -, @ 로 시작)으로 실행되지 않도록 글자로 저장
    function txt(k) { var s = v(k); return /^[=+\-@]/.test(s) ? "'" + s : s; }

    if (!v("track") || !v("status") || !v("worry") || !v("hope") || !v("name") ||
        !/^\d{4}$/.test(v("birth_year")) || !v("phone") || !v("sido") || v("consent") !== "동의") {
      return out({ ok: false, error: "invalid" });
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    // 빈 시트이거나, 나중에 추가된 열의 제목이 비어 있으면 그 칸만 채운다
    var head = sheet.getRange(1, 1, 1, HEADERS.length);
    var cur = head.getDisplayValues()[0];
    if (cur.some(function (h) { return h === ""; })) {
      head.setValues([HEADERS.map(function (h, i) { return cur[i] || h; })]).setFontWeight("bold").setBackground("#F8F25A");
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(HEADERS.indexOf("요즘 가장 큰 고민") + 1, 360);
    }

    // 휴대폰 번호 중복 확인
    var phoneCol = HEADERS.indexOf("휴대폰") + 1;
    var digits = v("phone").replace(/\D/g, "");
    var last = sheet.getLastRow();
    if (last > 1) {
      var phones = sheet.getRange(2, phoneCol, last - 1, 1).getDisplayValues();
      for (var i = 0; i < phones.length; i++) {
        if (String(phones[i][0]).replace(/\D/g, "") === digits) return out({ ok: false, error: "already" });
      }
    }

    // 휴대폰·출생연도가 숫자로 바뀌지 않도록 글자로 저장
    sheet.appendRow([
      new Date(), txt("track"), txt("detail"), txt("status"), txt("worry"), v("worry").length, txt("hope"), txt("memo"),
      txt("name"), "'" + v("birth_year"), "'" + v("phone"), txt("sido"), txt("gu"), txt("source"),
      txt("consent"), txt("research"), txt("seconds"), txt("page"),
      txt("utm_source"), txt("utm_medium"), txt("utm_campaign"), txt("utm_content"), txt("utm_term")
    ]);
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: "error", message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 배포 주소를 브라우저로 열어 확인할 때 보이는 문구
function doGet() {
  return out({ ok: true, service: "next-step-apply" });
}
