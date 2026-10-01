// ============================================================
// 우리 반 출결 기록 DB — 구글 Apps Script (어렵지 않아요!)
// ------------------------------------------------------------
// 이 코드는 '출결기록' 스프레드시트를 작은 데이터베이스처럼 쓰게
// 해주는 통역사예요. 앱에서 "출결 저장해줘"라고 하면 시트에 적어주고,
// "오늘 출결 보여줘"라고 하면 시트에서 읽어다 줍니다.
// ★ 이 코드는 '출결기록' 전용 스프레드시트에 묶어서 쓰세요!
//   (학생명단 파일이 아닙니다. 파일 2개를 따로 만드세요.)
// 코딩을 몰라도 돼요. 아래 2곳만 확인하세요!
//   ① SHEET_NAME: 첫 시트 이름 (꼭 '출결기록'으로 해주세요)
//   ② 열 순서: 기록시간 | 날짜 | 번호 | 이름 | 상태 | 입력방법 | 비고
// ============================================================

// 출결기록 전용 스프레드시트의 첫 시트 이름 (똑같은 이름으로 해주세요!)
const SHEET_NAME = '출결기록';

/**
 * [읽기] 앱에서 "오늘 출결 보여줘"라고 요청하면 실행돼요.
 * 주소 예) https://script.google.com/.../exec?date=2026-10-15
 * ?date= 를 비우면 전체 기록을 다 줍니다.
 */
function doGet(e) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) {
    return jsonOut({ ok: false, msg: '출결기록 시트를 찾지 못했어요. 첫 시트 이름이 출결기록인지 확인해주세요!' });
  }
  const rows = sh.getDataRange().getValues();
  const wantDate = (e.parameter.date || '').trim();  // 원하는 날짜 (없으면 전체)
  const list = [];
  for (let i = 1; i < rows.length; i++) {  // 0번 행은 제목줄이라 건너뜀
    const r = rows[i];
    if (!wantDate || String(r[1]).trim() === wantDate) {
      list.push({
        time: r[0],     // 기록시간
        date: r[1],     // 날짜 (예: 2026-10-15)
        no: r[2],       // 번호
        name: r[3],     // 이름
        status: r[4],   // 출석 / 지각 / 조퇴 / 결석
        by: r[5],       // 입력방법
        note: r[6]      // 비고
      });
    }
  }
  return jsonOut({ ok: true, count: list.length, data: list });
}

/**
 * [쓰기] 앱에서 출결 버튼을 누르면 실행돼요.
 * 앱이 {날짜, 번호, 이름, 상태}를 보내주면 시트 맨 아래에 한 줄 추가합니다.
 */
function doPost(e) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) {
    return jsonOut({ ok: false, msg: '출결기록 시트를 찾지 못했어요. 첫 시트 이름이 출결기록인지 확인해주세요!' });
  }
  const d = JSON.parse(e.postData.contents);  // 앱이 보낸 데이터 꺼내기
  sh.appendRow([
    new Date(),        // 기록시간: 지금 시각 자동 기록
    d.date || '',      // 날짜 (예: 2026-10-15)
    d.no || '',        // 번호
    d.name || '',      // 이름
    d.status || '',    // 출석 / 지각 / 조퇴 / 결석
    '앱입력',           // 입력방법
    d.note || ''       // 비고
  ]);
  return jsonOut({ ok: true, msg: '저장됐어요!' });
}

/**
 * 결과를 JSON 형식으로 예쁘게 돌려주는 도우미예요.
 * (doGet, doPost 안에서만 쓰여요. 신경 안 써도 됩니다!)
 */
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
