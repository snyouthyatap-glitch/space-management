// 네이버 예약 예약자명 정리 규칙.
// 같은 규칙이 staff-dashboard/NaverBookingParser.cs(NormalizeBookerName)와
// APPS_SCRIPT_READY.gs(cleanNaverBookerName_)에도 있다. 바꿀 때는 세 곳을 같이 바꾸고
// `npm test`로 확인한다.

const PHONE_PATTERN = /01[0-9]-?\d{3,4}-?\d{4}/g;
const PROXY_PATTERN = /대리\s*예약/;

// 예) "김예약 [대리예약] 방문자: 이방문" → "이방문", "최관리 [관리자등록]" → "최관리"
function cleanNaverBookerName(value) {
  let text = String(value || "").replace(/\s+/g, " ").trim();
  text = text.replace(/01[0-9]-?\d{3,4}-?\d{4}.*$/, "").replace(/\d{9,}.*$/, "");
  const parts = text.split(/(?:예약자|방문자)\s*[:：]?/);
  for (let index = parts.length - 1; index >= 0; index -= 1) {
    const name = parts[index]
      .replace(/\[[^\]]*\]/g, " ")
      .replace(/대리\s*예약|관리자\s*등록/g, " ")
      .replace(/[^가-힣A-Za-z\s]/g, "")
      .replace(/\s+/g, " ")
      .trim();
    if (name) return name;
  }
  return "";
}

function isNaverProxyBooking(value) {
  return PROXY_PATTERN.test(String(value || ""));
}

function findPhones(value) {
  return String(value || "").match(PHONE_PATTERN) || [];
}

// 목록 행에서 등록할 이름과 연락처를 정한다.
// 대리예약이면 아래쪽 방문자 이름과 아래쪽(두 번째) 연락처를 쓴다.
function resolveNaverBooker(bookerText, phones) {
  const list = (phones || []).filter(Boolean);
  const proxy = isNaverProxyBooking(bookerText);
  return {
    bookerName: cleanNaverBookerName(bookerText) || null,
    phone: (proxy && list[1]) || list[0] || null,
    proxyBooking: proxy,
  };
}

// 목록 행 글자에서 이용일시를 찾는다. 네이버 화면 형태가 두 가지라 둘 다 읽는다.
//   26. 8. 14.(금) 오후 6:00~9:00   /   26.11.11.수 17:30~19:30 (24시간제)
// 날짜 없이 시간만 있는 행(오늘이용)도 있다. 예약번호 바로 뒤를 먼저 본다.
const DATE_PART = String.raw`\d{2}\.\s*\d{1,2}\.\s*\d{1,2}\.\s*(?:\([^)]+\)|[월화수목금토일])`;
const TIME_PART = String.raw`(?:(?:오전|오후)\s*)?\d{1,2}:\d{2}\s*~\s*(?:(?:오전|오후)\s*)?\d{1,2}:\d{2}`;

function extractUseDateTime(text, reservationNo) {
  const source = String(text || "");
  if (reservationNo) {
    const afterNo = source.match(
      new RegExp(`${reservationNo}(?:\\s*변경)?\\s*((?:${DATE_PART}\\s*)?${TIME_PART})`)
    );
    if (afterNo) return afterNo[1].replace(/\s+/g, " ").trim();
  }
  const anywhere = source.match(new RegExp(`(${DATE_PART}\\s*${TIME_PART})`));
  return anywhere ? anywhere[1].replace(/\s+/g, " ").trim() : null;
}

// 네이버는 바뀐 예약의 예약번호 아래에 '변경'을 붙인다.
function isChangedReservation(text, reservationNo) {
  if (!reservationNo) return false;
  return new RegExp(`${reservationNo}\\s*변경`).test(String(text || ""));
}

module.exports = {
  cleanNaverBookerName,
  isNaverProxyBooking,
  findPhones,
  resolveNaverBooker,
  extractUseDateTime,
  isChangedReservation,
};
