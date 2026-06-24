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
    if (data && data.action) {
      return createJsonResponse_(dispatchWebAction_(data.action, data.payload || {}));
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetType = data.sheetType || 'visit';
    let sheetName = SHEET_NAMES.VISIT;

    if (sheetType === 'printer') sheetName = SHEET_NAMES.PRINTER;
    else if (sheetType === 'career') sheetName = SHEET_NAMES.CAREER_ZONE;
    else if (sheetType === 'connect') sheetName = SHEET_NAMES.CONNECT_ROOM;

    const sheet = getOrCreateSheet(ss, sheetName, sheetType);
    const lastRow = sheet.getLastRow();
    
    // 순번(seq) 계산 고도화: 마지막 행의 순번 값에 +1 (헤더만 있을 경우 1부터 시작)
    let seq = 1;
    if (lastRow > 1) {
      const lastSeq = sheet.getRange(lastRow, 1).getValue();
      seq = (typeof lastSeq === 'number') ? lastSeq + 1 : lastRow;
    }

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

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || '').trim();
    if (!action) {
      return ContentService
        .createTextOutput('Space management Apps Script is running.')
        .setMimeType(ContentService.MimeType.TEXT);
    }

    const callback = String((e.parameter && e.parameter.callback) || '').trim();
    const payloadText = String((e.parameter && e.parameter.payload) || '{}');
    const payload = JSON.parse(payloadText);
    const result = dispatchWebAction_(action, payload);

    if (callback) {
      return createJsonpResponse_(callback, result);
    }

    return createJsonResponse_(result);
  } catch (error) {
    const callback = String((e && e.parameter && e.parameter.callback) || '').trim();
    const result = {
      ok: false,
      error: error && error.message ? error.message : String(error)
    };
    return callback ? createJsonpResponse_(callback, result) : createJsonResponse_(result);
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

// ==================== 네이버 예약 CSV 동기화 ====================
const NAVER_SYNC_CONFIG = {
  folderName: '네이버예약CSV',
  sheetName: '네이버예약_동기화',
  timeZone: 'Asia/Seoul'
};

const NAVER_SYNC_HEADERS = [
  'reservationId',
  'status',
  'inflowPath',
  'reservationType',
  'bookerName',
  'maskedPhone',
  'email',
  'usageDate',
  'startTime',
  'endTime',
  'placeName',
  'normalizedPlace',
  'quantity',
  'ageRange',
  'headcount',
  'clubYn',
  'clubName',
  'purpose',
  'requestNote',
  'requestedAt',
  'confirmedAt',
  'completedAt',
  'cancelledAt',
  'cancelReason',
  'calendarEventId',
  'calendarStatus',
  'lastSyncedAt',
  'sourceFileName',
  'sourceModifiedAt',
  'rawUsageText'
];

function syncNaverReservationCsv() {
  const result = syncNaverReservationCsvInternal_(true);
  SpreadsheetApp.getUi().alert(
    '네이버 예약 동기화 완료\n\n' +
    '처리 파일: ' + result.fileCount + '개\n' +
    '신규 예약: ' + result.createdCount + '건\n' +
    '수정 예약: ' + result.updatedCount + '건\n' +
    '취소 반영: ' + result.cancelledCount + '건\n' +
    '오류: ' + result.errorCount + '건'
  );
}

function syncNaverReservationCsvInternal_(trashProcessedFiles) {
  const folder = getRequiredFolder_(NAVER_SYNC_CONFIG.folderName);
  const files = folder.getFilesByType(MimeType.CSV);
  const sheet = getOrCreateNaverSyncSheet_();
  const existingMap = buildExistingReservationMap_(sheet);
  const calendar = CalendarApp.getDefaultCalendar();
  const summary = {
    fileCount: 0,
    createdCount: 0,
    updatedCount: 0,
    cancelledCount: 0,
    errorCount: 0
  };

  while (files.hasNext()) {
    const file = files.next();
    summary.fileCount += 1;

    try {
      const csv = file.getBlob().getDataAsString('UTF-8').replace(/^\uFEFF/, '');
      const rows = splitCSVRows(csv);
      const headerIndex = rows.findIndex(row => row[0] && String(row[0]).includes('예약번호'));
      if (headerIndex < 0) {
        throw new Error('예약번호 헤더를 찾을 수 없습니다.');
      }

      const headers = rows[headerIndex];
      for (let i = headerIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || !row[0]) continue;

        const record = mapNaverReservationRow_(headers, row, file);
        if (!record.reservationId) continue;

        const existing = existingMap[record.reservationId];
        const syncType = upsertNaverReservation_(sheet, calendar, existing, record);

        if (syncType === 'created') summary.createdCount += 1;
        else if (syncType === 'updated') summary.updatedCount += 1;
        else if (syncType === 'cancelled') summary.cancelledCount += 1;

        existingMap[record.reservationId] = {
          rowNumber: existing ? existing.rowNumber : sheet.getLastRow(),
          values: buildSheetRowFromReservation_(record)
        };
      }

      if (trashProcessedFiles) {
        file.setTrashed(true);
      }
    } catch (error) {
      summary.errorCount += 1;
      Logger.log('❌ 네이버 예약 CSV 처리 오류 (' + file.getName() + '): ' + error.toString());
    }
  }

  return summary;
}

function getRequiredFolder_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) {
    throw new Error('드라이브 폴더를 찾을 수 없습니다: ' + folderName);
  }
  return folders.next();
}

function getOrCreateNaverSyncSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(NAVER_SYNC_CONFIG.sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(NAVER_SYNC_CONFIG.sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, NAVER_SYNC_HEADERS.length).setValues([NAVER_SYNC_HEADERS]);
    sheet.getRange(1, 1, 1, NAVER_SYNC_HEADERS.length)
      .setBackground('#0f766e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }

  return sheet;
}

function buildExistingReservationMap_(sheet) {
  const map = {};
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return map;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  values.forEach((row, index) => {
    const reservationId = String(row[0] || '').trim();
    if (!reservationId) return;
    map[reservationId] = {
      rowNumber: index + 2,
      values: row
    };
  });
  return map;
}

function mapNaverReservationRow_(headers, row, file) {
  const rowMap = {};
  headers.forEach((header, index) => {
    rowMap[String(header || '').trim()] = String(row[index] || '').trim();
  });

  const usageInfo = parseNaverUsageDateTime_(rowMap['이용일시']);
  const normalizedPlace = normalizeNaverPlaceName_(rowMap['상품']);
  const lastSyncedAt = Utilities.formatDate(new Date(), NAVER_SYNC_CONFIG.timeZone, 'yyyy-MM-dd HH:mm:ss');
  const sourceModifiedAt = Utilities.formatDate(file.getLastUpdated(), NAVER_SYNC_CONFIG.timeZone, 'yyyy-MM-dd HH:mm:ss');
  const status = rowMap['상태'];
  const isCancelled = status === '취소';

  return {
    reservationId: rowMap['예약번호'],
    status: status,
    inflowPath: rowMap['유입경로'],
    reservationType: rowMap['예약유형'],
    bookerName: rowMap['예약자'],
    maskedPhone: rowMap['전화번호'] || rowMap['방문자전화번호'],
    email: rowMap['이메일'],
    usageDate: usageInfo.date,
    startTime: usageInfo.startTime,
    endTime: usageInfo.endTime,
    placeName: rowMap['상품'],
    normalizedPlace: normalizedPlace,
    quantity: rowMap['수량'],
    ageRange: firstFilledValue_(
      rowMap['예약자입력정보1-이용자 나이 범위 (예: 14~16세)'],
      rowMap['예약자입력정보7-User age range (e.g. 14-16 years old)']
    ),
    headcount: firstFilledValue_(
      rowMap['예약자입력정보2-이용 인원(명)'],
      rowMap['예약자입력정보8-Number of users (number of employees)']
    ),
    clubYn: firstFilledValue_(
      rowMap['예약자입력정보3-야탑유스센터 동아리 소속 여부'],
      rowMap['예약자입력정보9-Whether you belong to the Yatap Youth Center club or not']
    ),
    clubName: firstFilledValue_(
      rowMap['예약자입력정보4-동아리명(소속 동아리일 경우에만 작성)'],
      rowMap['예약자입력정보10-Name of the club (only if it\'s a club to which it belongs)']
    ),
    purpose: firstFilledValue_(
      rowMap['예약자입력정보5-이용 목적'],
      rowMap['예약자입력정보11-Purpose of Use']
    ),
    requestNote: rowMap['요청사항'],
    requestedAt: rowMap['예약신청일시'],
    confirmedAt: rowMap['예약확정일시'],
    completedAt: rowMap['이용완료일시'],
    cancelledAt: rowMap['예약취소일시'],
    cancelReason: rowMap['취소사유'],
    calendarEventId: '',
    calendarStatus: isCancelled ? 'cancelled' : 'active',
    lastSyncedAt: lastSyncedAt,
    sourceFileName: file.getName(),
    sourceModifiedAt: sourceModifiedAt,
    rawUsageText: rowMap['이용일시']
  };
}

function parseNaverUsageDateTime_(text) {
  const raw = String(text || '').trim();
  if (!raw) {
    return { date: '', startTime: '', endTime: '' };
  }

  const dateMatch = raw.match(/(\d{2,4})\.\s*(\d{1,2})\.\s*(\d{1,2})/);
  const timeMatch = raw.match(/(오전|오후)\s*(\d{1,2}):(\d{2})~(\d{1,2}):(\d{2})/);

  if (!dateMatch || !timeMatch) {
    return { date: '', startTime: '', endTime: '' };
  }

  let year = Number(dateMatch[1]);
  if (year < 100) year += 2000;
  const month = String(dateMatch[2]).padStart(2, '0');
  const day = String(dateMatch[3]).padStart(2, '0');
  const period = timeMatch[1];

  const startHour = convertPeriodHour_(period, Number(timeMatch[2]));
  const startMinute = Number(timeMatch[3]);
  let endHour = convertPeriodHour_(period, Number(timeMatch[4]));
  const endMinute = Number(timeMatch[5]);

  if (endHour < startHour || (endHour === startHour && endMinute <= startMinute)) {
    endHour += 12;
  }

  return {
    date: year + '-' + month + '-' + day,
    startTime: String(startHour).padStart(2, '0') + ':' + String(startMinute).padStart(2, '0'),
    endTime: String(endHour).padStart(2, '0') + ':' + String(endMinute).padStart(2, '0')
  };
}

function convertPeriodHour_(period, hour) {
  if (period === '오후' && hour !== 12) return hour + 12;
  if (period === '오전' && hour === 12) return 0;
  return hour;
}

function normalizeNaverPlaceName_(placeName) {
  const raw = String(placeName || '').trim();
  if (!raw) return '';
  return raw
    .replace('[동아리실] ', '')
    .replace('[Club room] ', '')
    .replace('3층 ', '')
    .replace('3rd floor ', '')
    .trim();
}

function firstFilledValue_() {
  for (let i = 0; i < arguments.length; i++) {
    const value = String(arguments[i] || '').trim();
    if (value) return value;
  }
  return '';
}

function buildSheetRowFromReservation_(record) {
  return NAVER_SYNC_HEADERS.map(header => record[header] || '');
}

function upsertNaverReservation_(sheet, calendar, existing, record) {
  const existingValues = existing ? existing.values : null;
  if (existingValues) {
    record.calendarEventId = existingValues[24] || '';
  }

  record.calendarEventId = upsertNaverCalendarEvent_(calendar, record);
  const rowValues = buildSheetRowFromReservation_(record);

  if (!existing) {
    sheet.appendRow(rowValues);
    return record.calendarStatus === 'cancelled' ? 'cancelled' : 'created';
  }

  const rowNumber = existing.rowNumber;
  const hasChanged = JSON.stringify(existingValues) !== JSON.stringify(rowValues);
  if (hasChanged) {
    sheet.getRange(rowNumber, 1, 1, rowValues.length).setValues([rowValues]);
  }

  return record.calendarStatus === 'cancelled' ? 'cancelled' : 'updated';
}

function upsertNaverCalendarEvent_(calendar, record) {
  if (!record.usageDate || !record.startTime || !record.endTime || !record.normalizedPlace) {
    return record.calendarEventId || '';
  }

  const start = buildDateTime_(record.usageDate, record.startTime);
  const end = buildDateTime_(record.usageDate, record.endTime);
  const titlePrefix = record.calendarStatus === 'cancelled' ? '[취소]' : '[예약]';
  const title = `${titlePrefix} ${record.normalizedPlace} / ${record.bookerName}`;
  const description =
    '예약번호: ' + record.reservationId + '\n' +
    '상태: ' + record.status + '\n' +
    '예약자: ' + record.bookerName + '\n' +
    '전화번호: ' + record.maskedPhone + '\n' +
    '이용일시: ' + record.rawUsageText + '\n' +
    '인원: ' + record.headcount + '\n' +
    '목적: ' + record.purpose + '\n' +
    '취소사유: ' + record.cancelReason;

  let event = null;
  if (record.calendarEventId) {
    try {
      event = calendar.getEventById(record.calendarEventId);
    } catch (e) {
      event = null;
    }
  }

  if (!event) {
    const sameDayEvents = calendar.getEvents(start, end);
    event = sameDayEvents.find(function(item) {
      return String(item.getDescription() || '').includes('예약번호: ' + record.reservationId);
    }) || null;
  }

  if (!event) {
    event = calendar.createEvent(title, start, end, { description: description });
  } else {
    event.setTitle(title);
    event.setTime(start, end);
    event.setDescription(description);
  }

  try {
    event.setColor(record.calendarStatus === 'cancelled' ? '8' : '9');
  } catch (e) {
    Logger.log('예약 이벤트 색상 변경 실패: ' + e.toString());
  }

  return event.getId();
}

function buildDateTime_(dateStr, timeStr) {
  const dateParts = String(dateStr).split('-').map(Number);
  const timeParts = String(timeStr).split(':').map(Number);
  return new Date(
    dateParts[0],
    dateParts[1] - 1,
    dateParts[2],
    timeParts[0],
    timeParts[1],
    0
  );
}

// ==================== 웹사이트용 API ====================
const APP_API_SHEETS = {
  MEMBERS: 'APP_MEMBERS',
  LOUNGE_GUESTS: 'APP_LOUNGE_GUESTS',
  SETTINGS: 'APP_SETTINGS',
  VISIT_LOGS: 'APP_VISIT_LOGS',
  PRINTER_LOGS: 'APP_PRINTER_LOGS',
  CAREER_LOGS: 'APP_CAREER_LOGS',
  CONNECT_LOGS: 'APP_CONNECT_LOGS'
};

const APP_API_HEADERS = {
  APP_MEMBERS: ['memberId', 'name', 'gender', 'age', 'phone', 'phoneLastDigits', 'status', 'role', 'createdAt', 'updatedAt'],
  APP_LOUNGE_GUESTS: ['guestId', 'name', 'gender', 'birthDate', 'age', 'role', 'validDate', 'createdAt'],
  APP_SETTINGS: ['key', 'value', 'updatedAt'],
  APP_VISIT_LOGS: ['logId', 'userId', 'memberType', 'name', 'gender', 'age', 'phone', 'phoneLastDigits', 'date', 'source', 'createdBy', 'createdAt'],
  APP_PRINTER_LOGS: ['logId', 'userId', 'name', 'gender', 'age', 'phone', 'phoneLastDigits', 'date', 'count', 'createdBy', 'createdAt'],
  APP_CAREER_LOGS: ['logId', 'userId', 'name', 'gender', 'age', 'phone', 'phoneLastDigits', 'date', 'startTime', 'endTime', 'purpose', 'companionsJson', 'companionCount', 'createdBy', 'createdAt'],
  APP_CONNECT_LOGS: ['logId', 'userId', 'name', 'gender', 'age', 'phone', 'phoneLastDigits', 'date', 'startTime', 'endTime', 'purpose', 'companionsJson', 'companionCount', 'createdBy', 'createdAt']
};

function dispatchWebAction_(action, payload) {
  switch (action) {
    case 'bootstrap':
      return handleBootstrap_();
    case 'resolveFacilityEntry':
      return handleResolveFacilityEntry_(payload);
    case 'getFacilityMemberById':
      return handleGetFacilityMemberById_(payload);
    case 'registerFacilityMember':
      return handleRegisterFacilityMember_(payload);
    case 'getLoungeGuestById':
      return handleGetLoungeGuestById_(payload);
    case 'registerLoungeGuest':
      return handleRegisterLoungeGuest_(payload);
    case 'submitVisitLog':
      return handleSubmitVisitLog_(payload);
    case 'submitUsageRecord':
      return handleSubmitUsageRecord_(payload);
    case 'searchMembers':
      return handleSearchMembers_(payload);
    case 'getStats':
      return handleGetStats_(payload);
    default:
      throw new Error('지원하지 않는 요청입니다: ' + action);
  }
}

function createJsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function createJsonpResponse_(callback, payload) {
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function handleBootstrap_() {
  return {
    ok: true,
    adminSheetUrl: getAppSetting_('ADMIN_SHEET_URL')
  };
}

function handleResolveFacilityEntry_(payload) {
  const rawInput = String((payload && payload.inputValue) || '').trim();
  const adminPassword = getAppSetting_('ADMIN_PASSWORD');
  if (adminPassword && rawInput === adminPassword) {
    return {
      ok: true,
      mode: 'admin',
      adminSheetUrl: getAppSetting_('ADMIN_SHEET_URL')
    };
  }

  const digits = normalizeDigits_(rawInput);
  if (!digits) {
    throw new Error('숫자 또는 관리자 비밀번호를 입력해 주세요.');
  }

  const members = findMembersByPhoneInput_(digits);
  if (members.length === 0) {
    return {
      ok: true,
      mode: 'signup',
      suggestedPhone: digits
    };
  }

  const exactPhone = members.find(function(member) {
    return member.phone === digits;
  });
  if (exactPhone) {
    return {
      ok: true,
      mode: 'member',
      member: exactPhone
    };
  }

  if (members.length === 1) {
    return {
      ok: true,
      mode: 'member',
      member: members[0]
    };
  }

  return {
    ok: true,
    mode: 'ambiguous',
    message: '같은 끝자리 4자리를 사용하는 회원이 있습니다. 전체 전화번호를 다시 입력해 주세요.'
  };
}

function handleGetFacilityMemberById_(payload) {
  const memberId = String((payload && payload.memberId) || '').trim();
  if (!memberId) {
    throw new Error('회원 식별값이 없습니다.');
  }

  const member = getMemberById_(memberId);
  return {
    ok: true,
    member: member
  };
}

function handleRegisterFacilityMember_(payload) {
  const memberInput = payload && payload.member ? payload.member : {};
  const member = normalizeFacilityMember_(memberInput);
  if (!member.name || !member.gender || !member.age || !member.phone) {
    throw new Error('성함, 성별, 나이, 연락처를 모두 입력해 주세요.');
  }

  const existing = findMembersByExactPhone_(member.phone);
  if (existing.length > 0) {
    return {
      ok: true,
      member: existing[0],
      existing: true
    };
  }

  const now = formatDateTimeNow_();
  const newMember = {
    memberId: Utilities.getUuid(),
    name: member.name,
    gender: member.gender,
    age: member.age,
    phone: member.phone,
    phoneLastDigits: member.phoneLastDigits,
    status: 'approved',
    role: 'user',
    createdAt: now,
    updatedAt: now
  };

  appendSheetRecord_(APP_API_SHEETS.MEMBERS, APP_API_HEADERS.APP_MEMBERS, newMember);
  return {
    ok: true,
    member: convertMemberRecord_(newMember)
  };
}

function handleGetLoungeGuestById_(payload) {
  const guestId = String((payload && payload.guestId) || '').trim();
  if (!guestId) {
    throw new Error('임시회원 식별값이 없습니다.');
  }

  const sheet = getOrCreateAppSheet_(APP_API_SHEETS.LOUNGE_GUESTS, APP_API_HEADERS.APP_LOUNGE_GUESTS);
  const rows = getSheetObjects_(sheet);
  const guest = rows.find(function(row) {
    return row.guestId === guestId;
  });

  return {
    ok: true,
    guest: guest ? convertGuestRecord_(guest) : null
  };
}

function handleRegisterLoungeGuest_(payload) {
  const guestInput = payload && payload.guest ? payload.guest : {};
  const birthDate = String(guestInput.birthDate || '').trim();
  const gender = String(guestInput.gender || '').trim();
  const age = Number(guestInput.age || calculateAgeFromBirthDate_(birthDate) || 0);

  if (!birthDate || !gender || !age) {
    throw new Error('성별과 생년월일을 확인해 주세요.');
  }

  const guest = {
    guestId: Utilities.getUuid(),
    name: '라운지 이용자',
    gender: gender,
    birthDate: birthDate,
    age: age,
    role: 'lounge_guest',
    validDate: todayString_(),
    createdAt: formatDateTimeNow_()
  };

  appendSheetRecord_(APP_API_SHEETS.LOUNGE_GUESTS, APP_API_HEADERS.APP_LOUNGE_GUESTS, guest);
  return {
    ok: true,
    guest: convertGuestRecord_(guest)
  };
}

function handleSubmitVisitLog_(payload) {
  const member = normalizeSubjectForLogs_(payload && payload.member ? payload.member : {});
  const date = String((payload && payload.date) || todayString_()).trim();
  const source = String((payload && payload.source) || 'system').trim();
  const createdBy = String((payload && payload.createdBy) || 'self').trim();
  const allowDuplicate = Boolean(payload && payload.allowDuplicate);

  if (!member.name) {
    throw new Error('방문일지 대상 정보를 찾을 수 없습니다.');
  }

  if (!allowDuplicate && hasExistingVisitLog_(member.id, date)) {
    return {
      ok: true,
      created: false
    };
  }

  appendInternalVisitLog_(member, {
    date: date,
    source: source,
    createdBy: createdBy
  });
  appendPublicVisitLog_(member, date);

  return {
    ok: true,
    created: true
  };
}

function handleSubmitUsageRecord_(payload) {
  const type = String((payload && payload.type) || '').trim();
  const member = normalizeSubjectForLogs_(payload && payload.member ? payload.member : {});
  const data = payload && payload.data ? payload.data : {};
  const createdBy = String((payload && payload.createdBy) || 'self').trim();
  const date = String(data.date || todayString_()).trim();

  if (!type) {
    throw new Error('일지 종류가 없습니다.');
  }
  if (!member.name) {
    throw new Error('대상 회원 정보를 찾을 수 없습니다.');
  }

  if (type === 'printer') {
    const count = Number(data.count || 0);
    if (!count) {
      throw new Error('프린터 사용 매수를 입력해 주세요.');
    }

    const todayTotal = member.id ? getPrinterCountForDateFromSheet_(member.id, date) : 0;
    if (todayTotal + count > 10) {
      throw new Error('프린터는 하루 최대 10장까지 기록할 수 있습니다. 남은 수량을 확인해 주세요.');
    }

    appendInternalPrinterLog_(member, {
      date: date,
      count: count,
      createdBy: createdBy
    });
    appendPublicPrinterLog_(member, {
      date: date,
      count: count
    });
    return { ok: true };
  }

  const roomData = {
    date: date,
    startTime: normalizeTimeString_(data.startTime),
    endTime: normalizeTimeString_(data.endTime),
    purpose: String(data.purpose || '').trim(),
    companions: Array.isArray(data.companions) ? data.companions : [],
    createdBy: createdBy
  };

  if (!roomData.startTime || !roomData.endTime || !roomData.purpose) {
    throw new Error('시작 시간, 종료 시간, 이용 목적을 모두 입력해 주세요.');
  }

  if (type === 'careerZone') {
    appendInternalRoomLog_(APP_API_SHEETS.CAREER_LOGS, APP_API_HEADERS.APP_CAREER_LOGS, member, roomData);
    appendPublicCareerLog_(member, roomData);
    return { ok: true };
  }

  if (type === 'connectRoom') {
    appendInternalRoomLog_(APP_API_SHEETS.CONNECT_LOGS, APP_API_HEADERS.APP_CONNECT_LOGS, member, roomData);
    appendPublicConnectLog_(member, roomData);
    return { ok: true };
  }

  throw new Error('지원하지 않는 일지 종류입니다.');
}

function handleSearchMembers_(payload) {
  const term = String((payload && payload.term) || '').trim();
  if (!term) {
    return { ok: true, members: [] };
  }

  const digits = normalizeDigits_(term);
  const sheet = getOrCreateAppSheet_(APP_API_SHEETS.MEMBERS, APP_API_HEADERS.APP_MEMBERS);
  const rows = getSheetObjects_(sheet)
    .map(convertMemberRecord_)
    .filter(function(member) {
      const byName = String(member.name || '').indexOf(term) !== -1;
      const byPhone = digits && (
        String(member.phone || '').indexOf(digits) !== -1 ||
        String(member.phoneLastDigits || '').indexOf(digits) !== -1
      );
      return byName || byPhone;
    });

  return {
    ok: true,
    members: rows.slice(0, 50)
  };
}

function handleGetStats_(payload) {
  const period = String((payload && payload.period) || 'daily').trim();
  const range = getPeriodRange_(period);

  return {
    ok: true,
    range: range,
    totals: {
      visit: countRowsInRange_(APP_API_SHEETS.VISIT_LOGS, APP_API_HEADERS.APP_VISIT_LOGS, range),
      printer: sumPrinterInRange_(range),
      careerZone: countRowsInRange_(APP_API_SHEETS.CAREER_LOGS, APP_API_HEADERS.APP_CAREER_LOGS, range),
      connectRoom: countRowsInRange_(APP_API_SHEETS.CONNECT_LOGS, APP_API_HEADERS.APP_CONNECT_LOGS, range)
    }
  };
}

function getAppSetting_(key) {
  const fromProperties = PropertiesService.getScriptProperties().getProperty(key);
  if (fromProperties) {
    return String(fromProperties).trim();
  }

  const sheet = getOrCreateAppSheet_(APP_API_SHEETS.SETTINGS, APP_API_HEADERS.APP_SETTINGS);
  const rows = getSheetObjects_(sheet);
  const found = rows.find(function(row) {
    return row.key === key;
  });
  return found ? String(found.value || '').trim() : '';
}

function getOrCreateAppSheet_(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#111827')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }

  return sheet;
}

function getSheetObjects_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return [];
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values.map(function(row) {
    const obj = {};
    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });
    return obj;
  });
}

function appendSheetRecord_(sheetName, headers, record) {
  const sheet = getOrCreateAppSheet_(sheetName, headers);
  const row = headers.map(function(header) {
    return record[header] !== undefined ? record[header] : '';
  });
  sheet.appendRow(row);
}

function normalizeDigits_(value) {
  return String(value || '').replace(/\D/g, '');
}

function normalizeFacilityMember_(member) {
  const phone = normalizeDigits_(member.phone);
  return {
    name: String(member.name || '').trim(),
    gender: String(member.gender || '').trim(),
    age: Number(member.age || 0),
    phone: phone,
    phoneLastDigits: phone.slice(-4)
  };
}

function normalizeSubjectForLogs_(member) {
  const phone = normalizeDigits_(member.phone);
  return {
    id: String(member.id || member.memberId || member.guestId || '').trim(),
    name: String(member.name || '').trim(),
    gender: String(member.gender || '').trim(),
    age: Number(member.age || 0),
    phone: phone,
    phoneLastDigits: String(member.phoneLastDigits || phone.slice(-4) || '').trim(),
    role: String(member.role || 'user').trim()
  };
}

function convertMemberRecord_(row) {
  if (!row) return null;
  return {
    id: String(row.memberId || '').trim(),
    name: String(row.name || '').trim(),
    gender: String(row.gender || '').trim(),
    age: Number(row.age || 0),
    phone: normalizeDigits_(row.phone),
    phoneLastDigits: String(row.phoneLastDigits || '').trim(),
    status: String(row.status || 'approved').trim(),
    role: String(row.role || 'user').trim()
  };
}

function convertGuestRecord_(row) {
  if (!row) return null;
  return {
    id: String(row.guestId || '').trim(),
    name: String(row.name || '라운지 이용자').trim(),
    gender: String(row.gender || '').trim(),
    birthDate: String(row.birthDate || '').trim(),
    age: Number(row.age || 0),
    role: String(row.role || 'lounge_guest').trim(),
    validDate: String(row.validDate || '').trim()
  };
}

function findMembersByPhoneInput_(digits) {
  const exact = findMembersByExactPhone_(digits);
  if (exact.length > 0) {
    return exact;
  }

  const sheet = getOrCreateAppSheet_(APP_API_SHEETS.MEMBERS, APP_API_HEADERS.APP_MEMBERS);
  return getSheetObjects_(sheet)
    .map(convertMemberRecord_)
    .filter(function(member) {
      return member.phoneLastDigits === digits;
    });
}

function findMembersByExactPhone_(phone) {
  const sheet = getOrCreateAppSheet_(APP_API_SHEETS.MEMBERS, APP_API_HEADERS.APP_MEMBERS);
  return getSheetObjects_(sheet)
    .map(convertMemberRecord_)
    .filter(function(member) {
      return member.phone === phone;
    });
}

function getMemberById_(memberId) {
  const sheet = getOrCreateAppSheet_(APP_API_SHEETS.MEMBERS, APP_API_HEADERS.APP_MEMBERS);
  const row = getSheetObjects_(sheet).find(function(item) {
    return String(item.memberId || '').trim() === memberId;
  });
  return row ? convertMemberRecord_(row) : null;
}

function calculateAgeFromBirthDate_(birthDate) {
  const date = new Date(birthDate);
  if (isNaN(date.getTime())) {
    return 0;
  }

  const today = new Date();
  let age = today.getFullYear() - date.getFullYear();
  const monthDiff = today.getMonth() - date.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < date.getDate())) {
    age -= 1;
  }
  return age;
}

function todayString_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDateTimeNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function hasExistingVisitLog_(userId, date) {
  if (!userId) return false;
  const sheet = getOrCreateAppSheet_(APP_API_SHEETS.VISIT_LOGS, APP_API_HEADERS.APP_VISIT_LOGS);
  return getSheetObjects_(sheet).some(function(row) {
    return String(row.userId || '') === String(userId) && String(row.date || '') === String(date);
  });
}

function appendInternalVisitLog_(member, options) {
  appendSheetRecord_(APP_API_SHEETS.VISIT_LOGS, APP_API_HEADERS.APP_VISIT_LOGS, {
    logId: Utilities.getUuid(),
    userId: member.id || '',
    memberType: member.role || 'user',
    name: member.name,
    gender: member.gender,
    age: member.age,
    phone: member.phone || '',
    phoneLastDigits: member.phoneLastDigits || '',
    date: options.date,
    source: options.source || 'system',
    createdBy: options.createdBy || 'self',
    createdAt: formatDateTimeNow_()
  });
}

function appendPublicVisitLog_(member, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.VISIT, 'visit');
  const seq = nextSequence_(sheet);
  sheet.appendRow([seq, date, member.gender, member.age]);
}

function appendInternalPrinterLog_(member, options) {
  appendSheetRecord_(APP_API_SHEETS.PRINTER_LOGS, APP_API_HEADERS.APP_PRINTER_LOGS, {
    logId: Utilities.getUuid(),
    userId: member.id || '',
    name: member.name,
    gender: member.gender,
    age: member.age,
    phone: member.phone || '',
    phoneLastDigits: member.phoneLastDigits || '',
    date: options.date,
    count: Number(options.count || 0),
    createdBy: options.createdBy || 'self',
    createdAt: formatDateTimeNow_()
  });
}

function appendPublicPrinterLog_(member, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.PRINTER, 'printer');
  const seq = nextSequence_(sheet);
  sheet.appendRow([seq, data.date, member.gender, member.age]);
}

function appendInternalRoomLog_(sheetName, headers, member, data) {
  appendSheetRecord_(sheetName, headers, {
    logId: Utilities.getUuid(),
    userId: member.id || '',
    name: member.name,
    gender: member.gender,
    age: member.age,
    phone: member.phone || '',
    phoneLastDigits: member.phoneLastDigits || '',
    date: data.date,
    startTime: data.startTime,
    endTime: data.endTime,
    purpose: data.purpose,
    companionsJson: JSON.stringify(data.companions || []),
    companionCount: Array.isArray(data.companions) ? data.companions.length : 0,
    createdBy: data.createdBy || 'self',
    createdAt: formatDateTimeNow_()
  });
}

function appendPublicCareerLog_(member, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.CAREER_ZONE, 'career');
  const seq = nextSequence_(sheet);
  const members = [member].concat(data.companions || []);
  let maleCount = 0;
  let femaleCount = 0;
  members.forEach(function(item) {
    if (String(item.gender || '') === '남성') maleCount += 1;
    else if (String(item.gender || '') === '여성') femaleCount += 1;
  });

  const payload = {
    date: data.date,
    place: 'AI 커리어존',
    startTime: data.startTime,
    endTime: data.endTime,
    maleCount: maleCount,
    femaleCount: femaleCount,
    name: member.name,
    purpose: data.purpose,
    companionsDetail: buildCompanionsDetail_(data.companions || [])
  };

  sheet.appendRow([
    seq,
    payload.date,
    payload.place,
    payload.startTime,
    payload.endTime,
    payload.maleCount,
    payload.femaleCount,
    payload.name,
    payload.purpose,
    payload.companionsDetail
  ]);

  try {
    syncManualRecordToCalendar(payload, 'career');
  } catch (error) {
    Logger.log('커리어존 캘린더 동기화 실패: ' + error.toString());
  }
}

function appendPublicConnectLog_(member, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.CONNECT_ROOM, 'connect');
  const seq = nextSequence_(sheet);
  const members = [member].concat(data.companions || []);
  const counts = { m20: 0, f20: 0, m30: 0, f30: 0, m19: 0, f19: 0 };
  members.forEach(function(item) {
    const category = categorizeMemberForSheet_(item.gender, item.age);
    counts[category] = (counts[category] || 0) + 1;
  });

  const payload = {
    date: data.date,
    startTime: data.startTime,
    endTime: data.endTime,
    m20: counts.m20,
    f20: counts.f20,
    m30: counts.m30,
    f30: counts.f30,
    m19: counts.m19,
    f19: counts.f19,
    purpose: data.purpose,
    name: member.name,
    companionsDetail: buildCompanionsDetail_(data.companions || [])
  };

  sheet.appendRow([
    seq,
    payload.date,
    payload.startTime,
    payload.endTime,
    payload.m20,
    payload.f20,
    payload.m30,
    payload.f30,
    payload.m19,
    payload.f19,
    payload.purpose,
    payload.name,
    payload.companionsDetail
  ]);

  try {
    syncManualRecordToCalendar(payload, 'connect');
  } catch (error) {
    Logger.log('커넥트존 캘린더 동기화 실패: ' + error.toString());
  }
}

function buildCompanionsDetail_(companions) {
  if (!companions || companions.length === 0) {
    return '';
  }
  return '동반 ' + companions.length + '명 (' + companions.map(function(item) {
    return String(item.gender || '') + '/' + String(item.age || '');
  }).join(', ') + ')';
}

function categorizeMemberForSheet_(gender, age) {
  const ageNum = Number(age || 0);
  const isMale = String(gender || '') === '남성';
  const prefix = isMale ? 'm' : 'f';
  if (ageNum <= 19) return prefix + '19';
  if (ageNum <= 29) return prefix + '20';
  return prefix + '30';
}

function normalizeTimeString_(value) {
  const digits = normalizeDigits_(value);
  if (!digits) return '';
  const padded = digits.padStart(4, '0').slice(0, 4);
  const hh = Math.min(Number(padded.substring(0, 2)), 23);
  const mm = Math.min(Number(padded.substring(2, 4)), 59);
  return String(hh).padStart(2, '0') + ':' + String(mm).padStart(2, '0');
}

function getPrinterCountForDateFromSheet_(userId, date) {
  const sheet = getOrCreateAppSheet_(APP_API_SHEETS.PRINTER_LOGS, APP_API_HEADERS.APP_PRINTER_LOGS);
  return getSheetObjects_(sheet).reduce(function(total, row) {
    if (String(row.userId || '') !== String(userId) || String(row.date || '') !== String(date)) {
      return total;
    }
    return total + Number(row.count || 0);
  }, 0);
}

function nextSequence_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return 1;
  }
  const lastValue = sheet.getRange(lastRow, 1).getValue();
  return Number(lastValue || 0) + 1;
}

function getPeriodRange_(period) {
  const today = new Date();
  const end = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const start = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  if (period === 'weekly') {
    const day = start.getDay();
    const diff = start.getDate() - day + (day === 0 ? -6 : 1);
    start.setDate(diff);
  } else if (period === 'monthly') {
    start.setDate(1);
  }

  return {
    start: Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    end: end
  };
}

function countRowsInRange_(sheetName, headers, range) {
  const sheet = getOrCreateAppSheet_(sheetName, headers);
  return getSheetObjects_(sheet).filter(function(row) {
    const date = String(row.date || '');
    return date >= range.start && date <= range.end;
  }).length;
}

function sumPrinterInRange_(range) {
  const sheet = getOrCreateAppSheet_(APP_API_SHEETS.PRINTER_LOGS, APP_API_HEADERS.APP_PRINTER_LOGS);
  return getSheetObjects_(sheet).reduce(function(total, row) {
    const date = String(row.date || '');
    if (date < range.start || date > range.end) {
      return total;
    }
    return total + Number(row.count || 0);
  }, 0);
}

// ==================== 웹사이트 공개 일지 최종 양식 ====================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data && data.action) {
      return createJsonResponse_(dispatchWebAction_(data.action, data.payload || {}));
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetType = data.sheetType || 'visit';
    let sheetName = SHEET_NAMES.VISIT;

    if (sheetType === 'printer') sheetName = SHEET_NAMES.PRINTER;
    else if (sheetType === 'career') sheetName = SHEET_NAMES.CAREER_ZONE;
    else if (sheetType === 'connect') sheetName = SHEET_NAMES.CONNECT_ROOM;

    const sheet = getOrCreateSheet(ss, sheetName, sheetType);
    const seq = nextSequence_(sheet);
    let values = [];

    if (sheetType === 'visit') {
      values = [[seq, data.date || '', data.gender || '', data.age || '']];
    } else if (sheetType === 'printer') {
      values = [[seq, data.date || '', data.gender || '', data.age || '']];
    } else if (sheetType === 'career' || sheetType === 'connect') {
      values = [[
        seq,
        data.date || '',
        data.purpose || '',
        data.headcount || 0,
        data.companionsDetail || '',
        data.startTime || '',
        data.endTime || ''
      ]];
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, values[0].length).setValues(values);

    if (sheetType === 'career' || sheetType === 'connect') {
      try {
        syncManualRecordToCalendar(data, sheetType);
      } catch (err) {
        Logger.log('공간 이용 기록 캘린더 동기화 오류: ' + err.toString());
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    Logger.log('doPost 오류: ' + e.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: e.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet(ss, sheetName, type) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    let headers = [];
    if (type === 'visit') headers = ['연번', '날짜', '성별', '나이'];
    else if (type === 'printer') headers = ['연번', '날짜', '성별', '나이'];
    else if (type === 'career') headers = ['연번', '날짜', '사용목적', '인원수', '각 인원의 나이와 성별', '시작시간', '종료시간'];
    else if (type === 'connect') headers = ['연번', '날짜', '사용목적', '인원수', '각 인원의 나이와 성별', '시작시간', '종료시간'];

    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  return sheet;
}

function appendPublicVisitLog_(member, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.VISIT, 'visit');
  const seq = nextSequence_(sheet);
  sheet.appendRow([seq, date, member.gender, member.age]);
}

function appendPublicPrinterLog_(member, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.PRINTER, 'printer');
  const seq = nextSequence_(sheet);
  sheet.appendRow([seq, data.date, member.gender, member.age]);
}

function appendPublicCareerLog_(member, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.CAREER_ZONE, 'career');
  const seq = nextSequence_(sheet);
  const members = [member].concat(data.companions || []);

  const payload = {
    date: data.date,
    purpose: data.purpose,
    headcount: members.length,
    companionsDetail: buildParticipantDetail_(members),
    startTime: data.startTime,
    endTime: data.endTime
  };

  sheet.appendRow([
    seq,
    payload.date,
    payload.purpose,
    payload.headcount,
    payload.companionsDetail,
    payload.startTime,
    payload.endTime
  ]);

  try {
    syncManualRecordToCalendar(payload, 'career');
  } catch (error) {
    Logger.log('커리어존 캘린더 동기화 실패: ' + error.toString());
  }
}

function appendPublicConnectLog_(member, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, SHEET_NAMES.CONNECT_ROOM, 'connect');
  const seq = nextSequence_(sheet);
  const members = [member].concat(data.companions || []);

  const payload = {
    date: data.date,
    purpose: data.purpose,
    headcount: members.length,
    companionsDetail: buildParticipantDetail_(members),
    startTime: data.startTime,
    endTime: data.endTime
  };

  sheet.appendRow([
    seq,
    payload.date,
    payload.purpose,
    payload.headcount,
    payload.companionsDetail,
    payload.startTime,
    payload.endTime
  ]);

  try {
    syncManualRecordToCalendar(payload, 'connect');
  } catch (error) {
    Logger.log('커넥트존 캘린더 동기화 실패: ' + error.toString());
  }
}

function buildParticipantDetail_(members) {
  return (members || []).map(function(item) {
    return String(item.gender || '') + '/' + String(item.age || '');
  }).join(', ');
}

function getPublicSheetHeadersByType_(type) {
  if (type === 'visit') return ['연번', '날짜', '성별', '나이'];
  if (type === 'printer') return ['연번', '날짜', '성별', '나이'];
  if (type === 'career') return ['연번', '날짜', '사용목적', '인원수', '각 인원의 나이와 성별', '시작시간', '종료시간'];
  if (type === 'connect') return ['연번', '날짜', '사용목적', '인원수', '각 인원의 나이와 성별', '시작시간', '종료시간'];
  return [];
}

function archiveAndResetMonthlySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const propertyKey = 'MONTHLY_SHEET_ROLLOVER_LAST_RUN';
  const scriptProperties = PropertiesService.getScriptProperties();
  const currentMonthKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const lastRun = scriptProperties.getProperty(propertyKey);

  if (lastRun === currentMonthKey) {
    SpreadsheetApp.getUi().alert('이번 달 시트 전환은 이미 완료되었습니다.');
    return;
  }

  const archiveMonthLabel = getPreviousMonthLabel_();
  const sheetConfigs = [
    { name: SHEET_NAMES.VISIT, type: 'visit' },
    { name: SHEET_NAMES.PRINTER, type: 'printer' },
    { name: SHEET_NAMES.CAREER_ZONE, type: 'career' },
    { name: SHEET_NAMES.CONNECT_ROOM, type: 'connect' }
  ];

  const archived = [];
  const created = [];

  sheetConfigs.forEach(function(config) {
    const baseSheet = ss.getSheetByName(config.name);
    if (baseSheet) {
      const archivedName = config.name + '_' + archiveMonthLabel;
      const existingArchive = ss.getSheetByName(archivedName);
      if (!existingArchive) {
        baseSheet.setName(archivedName);
        baseSheet.hideSheet();
        archived.push(archivedName);
      }
    }

    const freshSheet = ss.getSheetByName(config.name) || ss.insertSheet(config.name);
    if (freshSheet.getLastRow() > 0) {
      freshSheet.clear();
    }
    const headers = getPublicSheetHeadersByType_(config.type);
    freshSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    freshSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    created.push(config.name);
  });

  scriptProperties.setProperty(propertyKey, currentMonthKey);

  SpreadsheetApp.getUi().alert(
    '월별 시트 전환이 완료되었습니다.\n\n' +
    '보관 시트: ' + (archived.length ? archived.join(', ') : '없음') + '\n' +
    '새 시트: ' + created.join(', ')
  );
}

function runMonthlySheetRolloverIfNeeded() {
  const propertyKey = 'MONTHLY_SHEET_ROLLOVER_LAST_RUN';
  const scriptProperties = PropertiesService.getScriptProperties();
  const currentMonthKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const lastRun = scriptProperties.getProperty(propertyKey);
  if (lastRun === currentMonthKey) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetConfigs = [
    { name: SHEET_NAMES.VISIT, type: 'visit' },
    { name: SHEET_NAMES.PRINTER, type: 'printer' },
    { name: SHEET_NAMES.CAREER_ZONE, type: 'career' },
    { name: SHEET_NAMES.CONNECT_ROOM, type: 'connect' }
  ];
  const archiveMonthLabel = getPreviousMonthLabel_();

  sheetConfigs.forEach(function(config) {
    const baseSheet = ss.getSheetByName(config.name);
    if (baseSheet) {
      const archivedName = config.name + '_' + archiveMonthLabel;
      if (!ss.getSheetByName(archivedName)) {
        baseSheet.setName(archivedName);
        baseSheet.hideSheet();
      }
    }

    const freshSheet = ss.getSheetByName(config.name) || ss.insertSheet(config.name);
    if (freshSheet.getLastRow() > 0) {
      freshSheet.clear();
    }
    const headers = getPublicSheetHeadersByType_(config.type);
    freshSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    freshSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  });

  scriptProperties.setProperty(propertyKey, currentMonthKey);
}

function getPreviousMonthLabel_() {
  const now = new Date();
  const previousMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return Utilities.formatDate(previousMonth, Session.getScriptTimeZone(), 'M월');
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('청년이봄홈페이지 관리')
    .addItem('월별 시트 전환 실행', 'archiveAndResetMonthlySheets')
    .addItem('네이버 예약 CSV 동기화', 'syncNaverReservationCsv')
    .addItem('통합 예약 동기화', 'syncAll')
    .addSeparator()
    .addItem('이번 달 일정표 새로고침', 'refreshCurrentMonth')
    .addItem('특정 달 일정표 보기', 'refreshSelectMonth')
    .addSeparator()
    .addItem('CSV 수동 처리', 'manualProcessCSV')
    .addItem('Excel 다운로드', 'downloadAsExcel')
    .addToUi();
}
