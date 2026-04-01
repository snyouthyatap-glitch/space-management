/**
 * 야탑 청년이봄 - 통합 관리 시스템 구글 앱스크립트 (최종 검증 완료 + Calendar API 수정)
 * 기능: 웹앱 데이터 수신(doPost), CSV 자동/수동 처리, 캘린더 연동, Excel 다운로드
 */

// ==================== 상수 정의 ====================
const SHEET_NAMES = {
  VISIT: '방문일지',
  PRINTER: '프린터',
  CAREER_ZONE: '커리어존',
  CONNECT_ROOM: '커넥트룸'
};

// ==================== [최종 수정] 데이터 수신 및 기록 (doPost) ====================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetType = data.sheetType || 'visit';
    let sheetName = SHEET_NAMES.VISIT;

    if (sheetType === 'printer') sheetName = SHEET_NAMES.PRINTER;
    else if (sheetType === 'career') sheetName = SHEET_NAMES.CAREER_ZONE;
    else if (sheetType === 'connect') sheetName = SHEET_NAMES.CONNECT_ROOM;

    const sheet = getOrCreateSheet(ss, sheetName, sheetType);
    const lastRow = sheet.getLastRow();
    const seq = lastRow > 0 ? (lastRow) : 1;

    let values = [];

    if (sheetType === 'visit') {
      values = [[seq, data.date || '', data.name || '', data.gender || '', data.age || '']];
    } else if (sheetType === 'printer') {
      // 프린터 일지: [순번, 날짜, 이름, 매수, 성별] 순으로 저장 (나이 제외)
      values = [[seq, data.date || '', data.name || '', data.count || '', data.gender || '']];
    } else if (sheetType === 'career') {
      values = [[
        seq,
        data.date || '',
        data.place || '',
        data.startTime || '',
        data.endTime || '',
        data.maleCount || 0,
        data.femaleCount || 0,
        data.name || '',
        data.purpose || '',
        data.companionsDetail || ''
      ]];
    } else if (sheetType === 'connect') {
      values = [[
        seq,
        data.date || '',
        data.startTime || '',
        data.endTime || '',
        data.m20 || 0,
        data.f20 || 0,
        data.m30 || 0,
        data.f30 || 0,
        data.m19 || 0,
        data.f19 || 0,
        data.purpose || '',
        data.name || '',
        data.companionsDetail || ''
      ]];
    }

    sheet.getRange(lastRow + 1, 1, 1, values[0].length).setValues(values);

    // [추가] 커리어존/커넥트룸인 경우 캘린더 및 달력 시트 연동
    if (sheetType === 'career' || sheetType === 'connect') {
      try {
        syncManualRecordToCalendar(data, sheetType);

        // 해당 월의 달력 시트 새로고침
        const datePart = data.date || ''; // YYYY-MM-DD
        if (datePart.length >= 7) {
          const monthKey = datePart.substring(0, 7);
          const calSheet = createCalendarSheet(ss, monthKey);
          const cal = CalendarApp.getDefaultCalendar();
          const [year, month] = monthKey.split('-');
          createCalendarView(calSheet, cal, parseInt(year), parseInt(month) - 1);
        }
      } catch (err) {
        Logger.log('❌ 상시 이용 건 캘린더 연동 오류: ' + err.toString());
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    Logger.log('❌ doPost 오류: ' + e.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: e.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// CORS 요청 대응
function doOptions(e) {
  const output = ContentService.createTextOutput();
  output.setHeader('Access-Control-Allow-Origin', '*');
  output.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  output.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  return output;
}

// 시트 확인 및 생성 (헤더 자동 정의)
function getOrCreateSheet(ss, sheetName, type) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    let headers = [];
    if (type === 'visit') headers = ['순번', '날짜', '성명', '성별', '나이'];
    else if (type === 'printer') headers = ['순번', '날짜', '이름', '매수', '성별'];
    else if (type === 'career') headers = ['순번', '날짜', '이용공간', '이용 시작 시간', '이용 종료 시간', '남', '여', '대표자성명', '이용 목적', '동반자정보상세(숨김)'];
    else if (type === 'connect') headers = ['순번', '날짜', '시작시간', '종료시간', '20~29세(남)', '20~29세(여)', '30~39세(남)', '30~39세(여)', '~19세(남)', '~19세(여)', '이용목적', '대표자성명', '동반자정보상세(숨김)'];

    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sheet;
}

// ==================== UI 메뉴 & 트리거 제어 ====================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('야탑 유스센터 관리')
    .addItem('🚀 통합 예약 연동 (CSV + 시트)', 'syncAll')
    .addSeparator()
    .addItem('📅 이번 달 달력 새로고침', 'refreshCurrentMonth')
    .addItem('🔍 특정 달 달력 보기', 'refreshSelectMonth')
    .addSeparator()
    .addItem('🔄 기존 시트 데이터만 동기화', 'syncExistingDataToCalendar')
    .addItem('📥 CSV 파일만 수동 파싱', 'manualProcessCSV')
    .addSeparator()
    .addItem('📊 데이터 Excel 다운로드', 'downloadAsExcel')
    .addToUi();
}


// ==================== CSV 자동 처리 (드라이브 감시) ====================
function processUploadedFilesAuto() {
  Logger.log('=== 자동 CSV 처리 시작 ===');
  coreCSVProcess(false);
  Logger.log('=== 자동 CSV 처리 완료 ===');
}

// 24H 포맷 변환 유틸리티
function convertTo24H(timeStr) {
  if (!timeStr) return '';
  const match = timeStr.match(/(오전|오후)\s*(\d+):(\d+)/);
  if (!match) return timeStr;
  let h = parseInt(match[2]);
  const m = match[3];
  if (match[1] === '오후' && h !== 12) h += 12;
  else if (match[1] === '오전' && h === 12) h = 0;
  return h.toString().padStart(2, '0') + ':' + m;
}

// ==================== 수동 CSV 파싱 ====================
function manualProcessCSV() {
  Logger.log('=== 수동 CSV 파싱 시작 ===');
  const result = coreCSVProcess(true);
  Logger.log('=== 수동 CSV 파싱 완료 ===');

  if (result) {
    SpreadsheetApp.getUi().alert(
      '✅ CSV 파싱 완료\n\n' +
      '처리된 파일: ' + result.fileCount + '개\n' +
      '총 데이터: ' + result.processCount + '개\n' +
      '오류: ' + result.errorCount + '개'
    );
  }
}

function syncAll() {
  Logger.log('=== [통합 연동] 시작 ===');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 드라이브 내 CSV 파일 파싱 및 캘린더 연동 (내부적으로 시트 갱신 포함)
  coreCSVProcess(false);
  
  // 2. 시트의 기존 기록들 캘린더 연동 및 업데이트
  syncExistingDataToCalendarInternal();
  
  // 3. 최근 3개월 달력 시트 미리 갱신 (현재달 포함)
  const today = new Date();
  for (let i = -1; i <= 1; i++) {
    const d = new Date(today.getFullYear(), today.getMonth() + i, 1);
    const monthKey = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM");
    const sheet = createCalendarSheet(ss, monthKey);
    const cal = CalendarApp.getDefaultCalendar();
    createCalendarView(sheet, cal, d.getFullYear(), d.getMonth());
  }
  
  Logger.log('=== [통합 연동] 완료 ===');
  SpreadsheetApp.getUi().alert('✅ 통합 예약 연동이 완료되었습니다.\n(드라이브 CSV -> 캘린더 -> 시트 3개월치 갱신)');
}

function coreCSVProcess(isManual) {
  try {
    const folders = DriveApp.getFoldersByName('예약');
    if (!folders.hasNext()) {
      if (isManual) SpreadsheetApp.getUi().alert('❌ "예약" 폴더가 없습니다.');
      return null;
    }

    const folder = folders.next();
    const files = folder.getFilesByType(MimeType.CSV);
    let processCount = 0;
    let errorCount = 0;
    let fileCount = 0;

    while (files.hasNext()) {
      const file = files.next();
      fileCount++;
      let deleteFile = false;
      try {
        const csv = file.getBlob().getDataAsString('UTF-8').replace(/^\uFEFF/, '');
        const rows = splitCSVRows(csv);
        const data = parseCSVData(rows);
        if (data.length > 0) {
          addToCalendarOptimized(data);
          addToCalendarSheet(data);
          deleteFile = true;
          processCount += data.length;
        }
      } catch (e) {
        Logger.log('❌ ' + file.getName() + ' 오류: ' + e.toString());
        errorCount++;
      }

      if (deleteFile) {
        file.setTrashed(true);
      }
    }
    return { fileCount, processCount, errorCount };
  } catch (e) {
    Logger.log('❌ coreCSVProcess 에러: ' + e.toString());
    return null;
  }
}

// ==================== 파싱 및 유틸리티 기능 ====================
function parseCSVData(rows) {
  const headerRow = rows.find(row => row.join(',').includes('예약번호'));
  if (!headerRow) {
    Logger.log('❌ 헤더 없음');
    return [];
  }
  const headerIndex = rows.indexOf(headerRow);
  const results = [];
  const seen = new Set();

  for (let i = headerIndex + 1; i < rows.length; i++) {
    const row = rows[i];
    const parsed = formatCSVRow(row);
    
    if (!parsed.장소 || !parsed.이용일시) {
      Logger.log('⚠️ 필드 누락 행 건너뜀 (' + (i+1) + '행)');
      continue;
    }

    parsed.isCancelled = parsed.상태 && (parsed.상태.includes('취소') || parsed.상태.includes('노쇼'));

    const key = `${parsed.장소}|${parsed.이용일시}|${parsed.성함}`;
    if (seen.has(key)) {
      Logger.log('⚠️ 중복 제거: ' + key);
      continue;
    }
    seen.add(key);

    const allowedProducts = ['청년 커넥트룸', 'AI 커리어존'];
    if (!allowedProducts.some(p => parsed.장소.includes(p))) {
      Logger.log('⚠️ 필터링: ' + parsed.장소);
      continue;
    }

    parsed.isCancelled = parsed.상태 && (parsed.상태.includes('취소') || parsed.상태.includes('노쇼'));
    parsed.statusTag = '';
    if (parsed.상태 && parsed.상태.includes('취소')) parsed.statusTag = '[취소]';
    else if (parsed.상태 && parsed.상태.includes('노쇼')) parsed.statusTag = '[노쇼]';
    
    results.push(parsed);
    Logger.log('✅ 파싱: ' + parsed.성함 + ' (' + parsed.이용일시 + ')');
  }

  Logger.log('📊 총 ' + results.length + '개 항목 파싱 완료 (취소 포함)');
  return results;
}

function formatCSVRow(parts) {
  const result = {
    장소: '', 이용일시: '', 성함: '', 인원수: '', 사용목적: '', 전화번호: '', 상태: '', 예약자: ''
  };
  
  if (parts.length > 3) result.상태 = parts[3].trim();
  if (parts.length > 5) result.예약자 = parts[5].trim();
  if (parts.length > 6) result.전화번호 = parts[6].trim();
  if (parts.length > 9 && parts[9].trim()) {
    result.성함 = parts[9].trim(); // 방문자 우선
  } else {
    result.성함 = result.예약자;
  }
  
  if (parts.length > 11) result.이용일시 = parts[11].trim();
  if (parts.length > 13) result.장소 = parts[13].trim();
  if (parts.length > 26) result.인원수 = parts[26].trim();
  if (parts.length > 29) result.사용목적 = parts[29].trim();
  
  return result;
}

function parseKoreanDate(dateTimeStr) {
  // YYYY.MM.DD 또는 YYYY-MM-DD 형식 지원
  const dateMatch = dateTimeStr.match(/(\d+)[.-]\s*(\d+)[.-]\s*(\d+)/);
  if (!dateMatch) {
    Logger.log('❌ 날짜 파싱 실패: ' + dateTimeStr);
    return null;
  }

  let year = parseInt(dateMatch[1]);
  if (year < 100) year += 2000; // 2자리 연도 처리
  const month = parseInt(dateMatch[2]);
  const day = parseInt(dateMatch[3]);

  if (month < 1 || month > 12 || day < 1 || day > 31) {
    Logger.log('❌ 잘못된 날짜: ' + year + '-' + month + '-' + day);
    return null;
  }

  const timeMatch = dateTimeStr.match(/(오전|오후)\s+(\d+):(\d+)/);
  let hour = 0, minute = 0;
  if (timeMatch) {
    hour = parseInt(timeMatch[2]);
    minute = parseInt(timeMatch[3]);
    if (timeMatch[1] === '오후' && hour !== 12) hour += 12;
    else if (timeMatch[1] === '오전' && hour === 12) hour = 0;
  }

  const timeStr = `${hour.toString().padStart(2, '0')}:${minute.toString().padStart(2, '0')}`;
  
  const result = new Date(year, month - 1, day, hour, minute, 0);
  if (isNaN(result.getTime())) {
    Logger.log('❌ Invalid Date 생성됨: ' + dateTimeStr);
    return null;
  }
  
  result.formattedTime = timeStr; // 24H 포맷 저장
  return result;
}

function splitCSVRows(csv) {
  const result = [];
  let row = [];
  let cell = '';
  let inQuotes = false;
  
  for (let i = 0; i < csv.length; i++) {
    const char = csv[i];
    const nextChar = csv[i+1];
    
    if (inQuotes) {
      if (char === '"' && nextChar === '"') {
        cell += '"'; 
        i++;
      } else if (char === '"') {
        inQuotes = false;
      } else {
        cell += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === ',') {
        row.push(cell);
        cell = '';
      } else if (char === '\n' || char === '\r') {
        row.push(cell);
        if (row.length > 1) result.push(row);
        row = [];
        cell = '';
        if (char === '\r' && nextChar === '\n') i++;
      } else {
        cell += char;
      }
    }
  }
  if (row.length > 0 || cell) {
    row.push(cell);
    if (row.length > 1) result.push(row);
  }
  return result;
}

// ==================== 캘린더 추가 (수정 기능 포함) ====================
function addToCalendarOptimized(arr) {
  const cal = CalendarApp.getDefaultCalendar();

  arr.forEach(item => {
    const start = parseKoreanDate(item.이용일시);
    if (!start) return;

    const prefix = item.statusTag || '';
    const title = `${prefix}[${item.장소}] ${item.성함}`;
    
    let endHour = start.getHours() + 2, endMinute = start.getMinutes();

    let timeStr24 = start.formattedTime + '~';
    const timeMatch = item.이용일시.match(/(오전|오후)\s+(\d+):(\d+)~(\d+):(\d+)/);
    if (timeMatch) {
      let endH = parseInt(timeMatch[4]);
      const endM = parseInt(timeMatch[5]);
      if (timeMatch[1] === '오후' && endH !== 12) endH += 12;
      else if (timeMatch[1] === '오전' && endH === 12) endH = 0;
      
      endHour = endH;
      endMinute = endM;
      timeStr24 += `${endH.toString().padStart(2, '0')}:${endM.toString().padStart(2, '0')}`;
    } else {
      timeStr24 += `${endHour.toString().padStart(2, '0')}:${endMinute.toString().padStart(2, '0')}`;
    }

    const end = new Date(start.getTime());
    end.setHours(endHour, endMinute, 0);

    // [중요] 종료 시간이 시작 시간보다 빠르면(예: 11:00~1:00), 종료 시간을 오후로 간주(+12시간)
    if (end <= start) {
      end.setHours(end.getHours() + 12);
    }
    
    // 최종 확인: 만약 그럼에도 불구하고 종료가 빠르다면 기본 2시간 예약으로 강제 설정
    if (end <= start) {
      end.setTime(start.getTime() + (2 * 60 * 60 * 1000));
    }

    const desc = `시간: ${timeStr24}\n연락처: ${item.전화번호}\n인원수: ${item.인원수}\n사용목적: ${item.사용목적}`;

    // 기존 이벤트 찾기 (방문자 또는 예약자 이름으로 검색)
    const baseTitleVisitor = `[${item.장소}] ${item.성함}`;
    const baseTitleBooker = `[${item.장소}] ${item.예약자 || item.성함}`;
    
    const existingEvents = cal.getEventsForDay(start).filter(e => {
      const eventTitle = e.getTitle();
      const startTimeMatch = e.getStartTime().getTime() === start.getTime();
      const nameMatch = eventTitle.includes(item.성함) || (item.예약자 && eventTitle.includes(item.예약자));
      const roomMatch = eventTitle.includes(`[${item.장소}]`);
      
      return startTimeMatch && nameMatch && roomMatch;
    });

    if (existingEvents.length > 0) {
      const existingEvent = existingEvents[0];
      try {
        existingEvent.setTitle(title);
        existingEvent.setTime(start, end);
        existingEvent.setDescription(desc);

        try {
          if (item.isCancelled) {
            existingEvent.setColor("8"); // GRAPHITE (String)
          } else {
            existingEvent.setColor("9"); // BLUE (String)
          }
        } catch (colorErr) {
          Logger.log('⚠️ 색상 변경 불가: ' + colorErr.message);
        }

        Logger.log('✏️ 이벤트 수정: ' + title);
      } catch (e) {
        Logger.log('⚠️ 이벤트 수정 실패: ' + e.toString());
      }

      for (let i = 1; i < existingEvents.length; i++) {
        try {
          cal.deleteEvent(existingEvents[i]);
          Logger.log('🗑️ 중복 이벤트 삭제: ' + existingEvents[i].getTitle());
        } catch (e) {
          Logger.log('⚠️ 중복 이벤트 삭제 실패: ' + e.toString());
        }
      }
    } else {
      try {
        const event = cal.createEvent(title, start, end, { description: desc });
        try {
          if (item.isCancelled) {
            event.setColor("8"); 
          } else {
            event.setColor("9"); 
          }
        } catch (colorErr) {
          Logger.log('⚠️ 색상 설정 불가: ' + colorErr.message);
        }
        Logger.log('✅ 이벤트 생성: ' + title);
      } catch (e) {
        Logger.log('❌ 이벤트 생성 실패: ' + e.toString());
      }
    }
    
    // API 할당량/속도 제한 방지
    Utilities.sleep(200);
  });
}

// ==================== 시트 및 캘린더 뷰 생성 ====================
function addToCalendarSheet(arr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthMap = {};

  arr.forEach(item => {
    const date = parseKoreanDate(item.이용일시);
    if (!date) return;

    const monthKey = date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
    if (!monthMap[monthKey]) monthMap[monthKey] = [];
    monthMap[monthKey].push(item);
  });

  Object.keys(monthMap).forEach(monthKey => {
    const sheet = createCalendarSheet(ss, monthKey);
    const cal = CalendarApp.getDefaultCalendar();
    const [year, month] = monthKey.split('-');
    createCalendarView(sheet, cal, parseInt(year), parseInt(month) - 1);
  });
}

// 캘린더 시트 생성 (데이터 시트와 별도)
function createCalendarSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  return sheet;
}

// ==================== 달력 시각화 ====================
function createCalendarView(sheet, calendar, year, month) {
  if (!sheet || !calendar || !year || month === undefined) {
    Logger.log('❌ 입력값 오류');
    return;
  }

  Logger.log('달력 생성: ' + year + '-' + String(month + 1).padStart(2, '0'));

  // 제목
  const titleRange = sheet.getRange(1, 1, 1, 7);
  titleRange.merge();
  titleRange.setValue(year + '년 ' + (month + 1) + '월')
    .setFontSize(16)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF');

  // 요일 헤더
  const days = ['일', '월', '화', '수', '목', '금', '토'];
  const headerRow = 3;
  for (let i = 0; i < 7; i++) {
    const cell = sheet.getRange(headerRow, i + 1);
    cell.setValue(days[i])
      .setFontWeight('bold')
      .setBackground('#E8F0FE')
      .setHorizontalAlignment('center');

    if (i === 0) cell.setFontColor('#FF0000');
    else if (i === 6) cell.setFontColor('#0000FF');
  }

  // 이벤트 데이터
  const startOfMonth = new Date(year, month, 1);
  const endOfMonth = new Date(year, month + 1, 0, 23, 59, 59);
  let events = [];
  try {
    events = calendar.getEvents(startOfMonth, endOfMonth);
    Logger.log('이벤트 수: ' + events.length);
  } catch (e) {
    Logger.log('❌ Calendar 오류: ' + e.toString());
  }

  const eventMap = {};
  events.forEach(ev => {
    const d = ev.getStartTime();
    const key = d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
    if (!eventMap[key]) eventMap[key] = [];

    const startStr = String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
    const endStr = String(ev.getEndTime().getHours()).padStart(2, '0') + ':' + String(ev.getEndTime().getMinutes()).padStart(2, '0');
    const isCancelled = ev.getTitle().includes('[취소]') || ev.getTitle().includes('[노쇼]');

    eventMap[key].push({
      time: startStr + '~' + endStr,
      title: ev.getTitle(),
      description: ev.getDescription(),
      isCancelled: isCancelled
    });
  });

  // 캘린더 그리드 생성
  let currentRow = 4;
  const firstDay = new Date(year, month, 1).getDay();
  const lastDay = new Date(year, month + 1, 0).getDate();
  let currentDay = 1 - firstDay; // 시작 날짜 (음수일 수 있음)

  for (let week = 0; week < 6; week++) {
    if (currentDay > lastDay) break;

    // 1. 이번 주 각 날짜의 이벤트 수 확인하여 최대 행 결정
    const weekEventCounts = [];
    for (let d = 0; d < 7; d++) {
      const tempDay = currentDay + d;
      if (tempDay > 0 && tempDay <= lastDay) {
        const tempKey = year + '-' + (month + 1) + '-' + tempDay;
        weekEventCounts.push((eventMap[tempKey] || []).length);
      } else {
        weekEventCounts.push(0);
      }
    }
    const maxEventsInWeek = Math.max(4, ...weekEventCounts);

    // 2. 날짜 행 및 이벤트 행들 그리기
    for (let dayOfWeek = 0; dayOfWeek < 7; dayOfWeek++) {
      const col = dayOfWeek + 1;
      const virtualDay = currentDay + dayOfWeek;

      if (virtualDay <= 0 || virtualDay > lastDay) {
        // 이번 달이 아닌 공백 칸
        sheet.getRange(currentRow, col, 1 + maxEventsInWeek, 1)
          .setBackground('#F9F9F9')
          .setBorder(true, true, true, true, null, null);
      } else {
        // 날짜 셀
        const dateCell = sheet.getRange(currentRow, col);
        dateCell.setValue(virtualDay)
          .setFontSize(12)
          .setFontWeight('bold')
          .setHorizontalAlignment('left')
          .setBackground('#FFFFFF')
          .setBorder(true, true, true, true, null, null);

        if (dayOfWeek === 0) dateCell.setFontColor('#FF0000');
        else if (dayOfWeek === 6) dateCell.setFontColor('#0000FF');

        const dayKey = year + '-' + (month + 1) + '-' + virtualDay;
        const dayEvents = eventMap[dayKey] || [];

        // 이벤트 배치
        dayEvents.forEach((ev, idx) => {
          const eventRow = currentRow + 1 + idx;
          const eventCell = sheet.getRange(eventRow, col);
          eventCell.setValue(ev.time + ' ' + ev.title)
            .setBackground(ev.isCancelled ? '#F0F0F0' : '#FFFACD')
            .setFontSize(9)
            .setWrap(true)
            .setVerticalAlignment('middle')
            .setBorder(true, true, true, true, null, null);

          if (ev.isCancelled) {
            eventCell.setFontColor('#999999').setFontLine('line-through');
          }

          if (ev.description) {
            eventCell.setNote(ev.description);
          }
        });
        
        // 빈 공간 테두리 처리
        if (dayEvents.length < maxEventsInWeek) {
          sheet.getRange(currentRow + 1 + dayEvents.length, col, maxEventsInWeek - dayEvents.length, 1)
            .setBackground('#FFFFFF')
            .setBorder(true, true, true, true, null, null);
        }
      }
    }

    // 3. 행 높이 설정
    sheet.setRowHeight(currentRow, 25); // 날짜행
    for (let i = 1; i <= maxEventsInWeek; i++) {
      sheet.setRowHeight(currentRow + i, 45); // 이벤트행
    }

    currentRow += (1 + maxEventsInWeek);
    currentDay += 7;
  }

  // 열 너비 설정
  for (let i = 1; i <= 7; i++) {
    sheet.setColumnWidth(i, 150);
  }

  Logger.log('✅ 달력 생성 완료');
}

// ==================== 달력 새로고침 ====================
function refreshCurrentMonth() {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = year + '-' + String(month + 1).padStart(2, '0');
  const sheet = createCalendarSheet(ss, sheetName);

  const cal = CalendarApp.getDefaultCalendar();
  createCalendarView(sheet, cal, year, month);

  SpreadsheetApp.getUi().alert('✅ ' + sheetName + ' 달력 새로고침 완료');
}

function refreshSelectMonth() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('달력을 표시할 월을 입력하세요.\n형식: YYYY-MM\n예: 2026-03');

  if (!response || response.getSelectedButton() === ui.Button.CANCEL) {
    return;
  }

  const input = response.getResponseText().trim();
  if (!input || !input.match(/^\d{4}-\d{2}$/)) {
    ui.alert('❌ 형식이 잘못되었습니다. YYYY-MM 형식으로 입력하세요.');
    return;
  }

  const [yearStr, monthStr] = input.split('-');
  const year = parseInt(yearStr);
  const month = parseInt(monthStr) - 1;

  if (month < 0 || month > 11) {
    ui.alert('❌ 월이 1~12 사이여야 합니다.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = year + '-' + String(month + 1).padStart(2, '0');
  const sheet = createCalendarSheet(ss, sheetName);

  const cal = CalendarApp.getDefaultCalendar();
  createCalendarView(sheet, cal, year, month);

  ui.alert('✅ ' + sheetName + ' 달력 생성 완료');
}

// ==================== Excel 다운로드 ====================
function downloadAsExcel() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const sheetName = sheet.getName();

    Logger.log('📊 Excel 다운로드 시작: ' + sheetName);

    const range = sheet.getDataRange();
    if (!range || range.getNumRows() === 0) {
      SpreadsheetApp.getUi().alert('❌ 데이터가 없습니다.');
      return;
    }

    const values = range.getValues();
    const backgrounds = range.getBackgrounds();
    const fontColors = range.getFontColors();
    const fontWeights = range.getFontWeights();
    const fontLines = range.getFontLines();

    Logger.log('데이터 크기: ' + values.length + ' rows × ' + values[0].length + ' cols');

    // 임시 시트 생성
    let tempSheet = ss.getSheetByName('_Export_Temp');
    if (tempSheet) ss.deleteSheet(tempSheet);
    tempSheet = ss.insertSheet('_Export_Temp', ss.getSheets().length);

    // 데이터 복사
    tempSheet.getRange(1, 1, values.length, values[0].length).setValues(values);

    // 서식 적용
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i].length; j++) {
        const cell = tempSheet.getRange(i + 1, j + 1);

        if (backgrounds[i][j] && backgrounds[i][j] !== '#ffffff') {
          cell.setBackground(backgrounds[i][j]);
        }

        if (fontColors[i][j] && fontColors[i][j] !== '#000000') {
          cell.setFontColor(fontColors[i][j]);
        }

        if (fontWeights[i][j] === 'bold') {
          cell.setFontWeight('bold');
        }

        if (fontLines[i][j] === 'line-through') {
          cell.setFontLine('line-through');
        }
      }
    }

    // 병합 셀 복사
    try {
      const mergedRanges = sheet.getMergedRanges();
      mergedRanges.forEach(mergedRange => {
        try {
          const row = mergedRange.getRow();
          const col = mergedRange.getColumn();
          const numRows = mergedRange.getNumRows();
          const numCols = mergedRange.getNumColumns();
          tempSheet.getRange(row, col, numRows, numCols).merge();
          Logger.log('✅ 병합: ' + row + ',' + col);
        } catch (e) {
          Logger.log('⚠️ 병합 실패: ' + e.toString());
        }
      });
    } catch (e) {
      Logger.log('⚠️ 병합 셀 복사 건너뜀: ' + e.toString());
    }

    // 열 너비 복사
    try {
      for (let i = 1; i <= values[0].length; i++) {
        const width = sheet.getColumnWidth(i);
        if (width) tempSheet.setColumnWidth(i, width);
      }
    } catch (e) {
      Logger.log('⚠️ 열 너비 복사 실패: ' + e.toString());
    }

    // 행 높이 복사
    try {
      for (let i = 1; i <= values.length; i++) {
        const height = sheet.getRowHeight(i);
        if (height) tempSheet.setRowHeight(i, height);
      }
    } catch (e) {
      Logger.log('⚠️ 행 높이 복사 실패: ' + e.toString());
    }

    // 달력 시트 특별 처리
    if (sheetName.match(/^\d{4}-\d{2}$/)) {
      Logger.log('📅 달력 시트 감지: ' + sheetName);
      try {
        tempSheet.setRowHeight(1, 25);
        tempSheet.setRowHeight(3, 25);
        for (let row = 4; row <= values.length; row++) {
          tempSheet.setRowHeight(row, 60);
        }
        Logger.log('✅ 달력 행 높이 설정 완료');
      } catch (e) {
        Logger.log('⚠️ 달력 행 높이 설정 실패: ' + e.toString());
      }
    }

    Logger.log('✅ 서식 적용 완료');
    SpreadsheetApp.getUi().alert(
      '✅ Excel 파일 준비 완료\n\n' +
      '다음 단계:\n' +
      '1. 파일 > 다운로드 > Microsoft Excel (.xlsx)\n' +
      '2. 임시 시트 "_Export_Temp"가 생성되었습니다.\n' +
      '3. 다운로드 후 수동으로 삭제해주세요.'
    );

  } catch (e) {
    Logger.log('❌ 다운로드 오류: ' + e.toString());
    SpreadsheetApp.getUi().alert('❌ 오류: ' + e.toString());
  }
}

// ==================== [신설] 웹앱 수동 제출 건 캘린더 연동 ====================
function syncManualRecordToCalendar(data, type) {
  const cal = CalendarApp.getDefaultCalendar();
  const name = data.name;
  const place = data.place;
  const dateStr = data.date;

  const start = parseKoreanDate(`${dateStr} 오전 0:00`); 
  if (!start) return;

  const sTime24 = convertTo24H(data.startTime);
  const eTime24 = convertTo24H(data.endTime);
  
  const [sH, sM] = sTime24.split(':').map(Number);
  const [eH, eM] = eTime24.split(':').map(Number);
  
  start.setHours(sH, sM, 0);
  const end = new Date(start.getTime());
  end.setHours(eH, eM, 0);

  // [중요] 종료 시간 안전 보정
  if (end <= start) {
    end.setHours(end.getHours() + 12);
  }
  if (end <= start) {
    end.setTime(start.getTime() + (2 * 60 * 60 * 1000));
  }

  const title = `[${place}] ${name}`;
  const desc = `시간: ${sTime24}~${eTime24}\n연락처: ${data.phone || ''}\n인원수: ${data.count || ''}\n사용목적: ${data.purpose || ''}`;

  try {
    const sameDayEvents = cal.getEventsForDay(start);
    const existingEvent = sameDayEvents.find(e => {
      const eTitle = e.getTitle();
      const eStart = e.getStartTime().getTime();
      return Math.abs(eStart - start.getTime()) < 1000 && eTitle.includes(name);
    });

    if (existingEvent) {
      existingEvent.setTitle(title);
      existingEvent.setDescription(desc);
      existingEvent.setTime(start, end);
      Logger.log('✏️ 업데이트: ' + title);
    } else {
      const event = cal.createEvent(title, start, end, { description: desc });
      event.setColor("9");
      Logger.log('✅ 등록: ' + title);
    }
    Utilities.sleep(200); 
  } catch (e) {
    Logger.log('❌ syncManualRecordToCalendar 오류: ' + e.toString());
  }
}

// syncAll에서 사용되는 무인터랙션 버전
function syncExistingDataToCalendarInternal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const careerSheet = ss.getSheetByName('커리어존');
  const connectSheet = ss.getSheetByName('커넥트룸');

  if (careerSheet) {
    const data = careerSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[1] || !row[7]) continue;
      const formatted = formatRecordForSync({
        date: row[1], 
        place: row[2], 
        startTime: row[3], 
        endTime: row[4], 
        name: row[7], 
        purpose: row[8],
        count: (Number(row[5]) || 0) + (Number(row[6]) || 0)
      });
      syncManualRecordToCalendar(formatted, 'career');
    }
  }

  if (connectSheet) {
    const data = connectSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[1] || !row[11]) continue;
      const formatted = formatRecordForSync({
        date: row[1], 
        place: '청년 커넥트룸',
        startTime: row[2], 
        endTime: row[3], 
        name: row[11], 
        purpose: row[10],
        count: (Number(row[4]) || 0) + (Number(row[5]) || 0) + (Number(row[6]) || 0) + 
               (Number(row[7]) || 0) + (Number(row[8]) || 0) + (Number(row[9]) || 0)
      });
      syncManualRecordToCalendar(formatted, 'connect');
    }
  }
}

function syncExistingDataToCalendar() {
  const ui = SpreadsheetApp.getUi();
  Logger.log('=== 기존 데이터 동기화 시작 ===');
  syncExistingDataToCalendarInternal();
  Logger.log('=== 기존 데이터 동기화 완료 ===');
  ui.alert('✅ 기존 시트 데이터가 캘린더와 동기화되었습니다.');
}

// 데이터 포맷 보정 유틸리티 (Sheet -> Calendar 연동용)
function formatRecordForSync(data) {
  // 날짜 보정
  if (data.date instanceof Date) {
    data.date = Utilities.formatDate(data.date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } else if (typeof data.date === 'string' && (data.date.includes('.') || data.date.includes('-'))) {
    const parts = data.date.split(/[.-]/).map(p => p.trim()).filter(p => p);
    if (parts.length >= 3) {
      const y = parts[0].length === 2 ? '20' + parts[0] : parts[0];
      const m = parts[1].padStart(2, '0');
      const d = parts[2].padStart(2, '0');
      data.date = `${y}-${m}-${d}`;
    }
  }

  // 시간 보정
  const fixTime = (t) => {
    if (!t) return '00:00';
    if (t instanceof Date) {
      return Utilities.formatDate(t, Session.getScriptTimeZone(), "HH:mm");
    }
    let s = String(t).trim();
    if (s.includes(':')) {
      const parts = s.split(':');
      if (parts.length >= 2) {
        return parts[0].padStart(2, '0') + ':' + parts[1].padStart(2, '0');
      }
      return s;
    }
    if (s.length === 4) return s.substring(0, 2) + ':' + s.substring(2);
    if (s.length === 3) return '0' + s.substring(0, 1) + ':' + s.substring(1);
    return s;
  };

  data.startTime = fixTime(data.startTime);
  data.endTime = fixTime(data.endTime);

  return data;
}
